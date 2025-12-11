
<#
.SYNOPSIS
  Safe SharePoint Online version trim tool using PnP.PowerShell only.

.DESCRIPTION
  - Uses ONLY PnP.PowerShell (no SPO Management Shell needed).
  - Shows current PnP site version policy and offers to update it.
  - Enumerates document libraries (all, specific, or from CSV).
  - Trims file versions older than N days.
  - ALWAYS runs as DryRun on the first execution (per-user state file).
  - Uses time-based batching (MaxBatchMinutes) to avoid huge long-running calls.
  - Respects retention, eDiscovery, records (SharePoint blocks invalid deletes).
  - Logs warnings/errors to CSV if provided.
  - Optionally logs text events to a rolling text log.
  - Reports total size before/after, and estimated storage reclaimed.

.REQUIREMENTS
  - PnP.PowerShell
  - You must already be connected, e.g.:
      Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/site" -Interactive

.NOTES
  After the first dry run, you can run again with -Delete.
#>

# -----------------------------------------------------------------------------
#   Helper: Write-TrimEvent (text log + console levels)
# -----------------------------------------------------------------------------
function Write-TrimEvent {
    param(
        [string]$Level,     # Info / Warn / Error
        [string]$Message
    )

    $ts   = (Get-Date).ToString("o")
    $line = "[$ts] [$Level] $Message"

    # Text log path is passed into Invoke-PnPVersionTrimTool and stored script-scoped
    if ($script:TextLogPath) {
        $dir = Split-Path $script:TextLogPath -Parent
        if ($dir -and -not (Test-Path $dir)) {
            New-Item -ItemType Directory -Force -Path $dir | Out-Null
        }
        Add-Content -Path $script:TextLogPath -Value $line
    }

    switch ($Level) {
        'Error' { Write-Host $line -ForegroundColor Red }
        'Warn'  { Write-Host $line -ForegroundColor Yellow }
        default { } # Info stays in file only
    }
}

# -----------------------------------------------------------------------------
#   Helper: Show-PnPSiteVersionPolicy (+ cooldown marker)
# -----------------------------------------------------------------------------
function Show-PnPSiteVersionPolicy {
    [CmdletBinding()]
    param()

    try {
        $policy = Get-PnPSiteVersionPolicy
    } catch {
        Write-Warning "Failed to get site version policy: $($_.Exception.Message)"
        return
    }

    if (-not $policy) {
        Write-Host "No site version policy returned." -ForegroundColor Yellow
        return
    }

    Write-Host "========== Current Site Version Policy (Raw) ==========" -ForegroundColor Cyan
    $policy | Format-List *

    # Derive a status line
    $autoStatus = if ($policy.EnableAutoExpirationVersionTrim) { "ENABLED" } else { "DISABLED" }
    $color      = if ($policy.EnableAutoExpirationVersionTrim) { 'Green' } else { 'Yellow' }

    Write-Host ""
    Write-Host "----------------- Policy Status Summary -----------------" -ForegroundColor Cyan
    Write-Host (" Auto-expiration : {0}" -f $autoStatus) -ForegroundColor $color
    Write-Host (" Major versions  : {0}" -f $policy.MajorVersions)
    Write-Host (" Expire after    : {0} days" -f $policy.ExpireVersionsAfterDays)
    Write-Host "---------------------------------------------------------" -ForegroundColor Cyan

    Write-Host ""
    Write-Host "Review this carefully before trimming versions." -ForegroundColor Yellow

    $answer = Read-Host "Change/update policy now? (y/N)"
    if ($answer -eq 'y') {
        $enable = Read-Host "Enable auto expiration version trim? (true/false) [current: $($policy.EnableAutoExpirationVersionTrim)]"
        if ([string]::IsNullOrWhiteSpace($enable)) { $enable = $policy.EnableAutoExpirationVersionTrim }

        $major = Read-Host "Max major versions? [current: $($policy.MajorVersions)]"
        if (-not [int]::TryParse($major, [ref]0)) { $major = $policy.MajorVersions }

        $days  = Read-Host "Expire versions after how many days? [current: $($policy.ExpireVersionsAfterDays)]"
        if (-not [int]::TryParse($days, [ref]0)) { $days = $policy.ExpireVersionsAfterDays }

        try {
            Set-PnPSiteVersionPolicy `
                -EnableAutoExpirationVersionTrim ([bool]$enable) `
                -MajorVersions ([int]$major) `
                -ExpireVersionsAfterDays ([int]$days) `
                -ApplyToExistingDocumentLibraries `
                -ApplyToNewDocumentLibraries

            Write-Host "Policy updated." -ForegroundColor Green

            # Record the policy update for cooldown logic
            $global:PnPVersionTrim_LastPolicyUpdateUtc = (Get-Date).ToUniversalTime()
        } catch {
            Write-Warning "Failed to update policy: $($_.Exception.Message)"
        }
    }
}

# -----------------------------------------------------------------------------
#   Helper: Get-PnPListSizeBytes (fast-ish list size via File_x0020_Size)
# -----------------------------------------------------------------------------
function Get-PnPListSizeBytes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ListTitle
    )

    $items = Get-PnPListItem -List $ListTitle -PageSize 2000 -Fields "FileLeafRef","File_x0020_Size","FSObjType"

    $total = 0L
    foreach ($item in $items) {
        # FSObjType: 0 = File, 1 = Folder
        if ($item["FSObjType"] -eq 0 -and $item["File_x0020_Size"]) {
            $total += [int64]$item["File_x0020_Size"]
        }
    }

    return $total
}

# -----------------------------------------------------------------------------
#   Helper: Write-PnPSizeLog (before/after size snapshots)
# -----------------------------------------------------------------------------
function Write-PnPSizeLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LogPath,
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$LibraryTitle,
        [Parameter(Mandatory)][ValidateSet('Before','After')][string]$Phase,
        [Parameter(Mandatory)][long]$SizeBytes
    )

    $dir = Split-Path $LogPath -Parent
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }

    if (-not (Test-Path $LogPath)) {
        "Timestamp,RunId,SiteUrl,LibraryTitle,Phase,SizeBytes,SizeMB" |
            Out-File -FilePath $LogPath -Encoding UTF8
    }

    $mb   = [math]::Round($SizeBytes / 1MB, 2)
    $line = "{0},{1},{2},{3},{4},{5},{6}" -f (Get-Date).ToString("o"),
            $RunId, $SiteUrl, $LibraryTitle, $Phase, $SizeBytes, $mb

    Add-Content -Path $LogPath -Value $line
}

# -----------------------------------------------------------------------------
#   Main: Invoke-PnPVersionTrimTool
# -----------------------------------------------------------------------------
function Invoke-PnPVersionTrimTool {
    [CmdletBinding()]
    param(
        [int]$OlderThanDays = 45,
        [string]$LibraryTitle,
        [string]$LibraryCsvPath,
        [switch]$Delete,
        [string]$LogPath,
        [string]$TextLogPath,
        [int]$MaxBatchMinutes = 5,
        [switch]$AutoContinue,
        [switch]$BypassBatching,

        # per-file version batching controls
        [int]$VersionBatchSize = 50,
        [int]$VersionBatchPauseMs = 500,
        [int]$MaxRetries = 5
    )

    # Make TextLogPath visible to Write-TrimEvent
    $script:TextLogPath = $TextLogPath

    #
    # --- SAFETY: FIRST RUN MUST ALWAYS BE DRY RUN ---
    #
    $stateRoot = Join-Path $env:LOCALAPPDATA "PnPVersionTrim"
    if (-not (Test-Path $stateRoot)) {
        New-Item -ItemType Directory -Force -Path $stateRoot | Out-Null
    }
    $stateFile = Join-Path $stateRoot "state.json"

    $firstRun        = -not (Test-Path $stateFile)
    $effectiveDryRun = $true

    if ($firstRun) {
        Write-Host "First run detected. THIS WILL BE A DRY RUN ONLY." -ForegroundColor Yellow
    }
    elseif ($Delete) {
        $effectiveDryRun = $false
    }

    if ($effectiveDryRun) {
        Write-Host "DRY RUN: No versions will be deleted." -ForegroundColor Cyan
    }
    else {
        Write-Host "DELETE MODE: Actual deletions will occur." -ForegroundColor Red
        $confirm = Read-Host "Type DELETE to proceed"
        if ($confirm -ne 'DELETE') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    #
    # --- OPTIONAL CSV LOGGING ---
    #
    if ($LogPath) {
        if (-not (Test-Path (Split-Path $LogPath -Parent))) {
            New-Item -ItemType Directory -Force -Path (Split-Path $LogPath -Parent) | Out-Null
        }
        if (-not (Test-Path $LogPath)) {
            "Timestamp,Action,LibraryTitle,FileRef,VersionId,VersionLabel,VersionCreated,Result,Message" |
                Out-File -FilePath $LogPath -Encoding UTF8
        }
        function Write-TrimLog {
            param($Action,$Library,$FileRef,$VersionId,$Label,$Created,$Result,$Message)
            $line = "{0},{1},{2},{3},{4},{5},{6},{7},{8}" -f (Get-Date).ToString("o"),
                    $Action,$Library,$FileRef,$VersionId,$Label,$Created.ToString("o"),$Result,$Message
            Add-Content -Path $LogPath -Value $line
        }
    }
    else {
        function Write-TrimLog { }
    }

    #
    # --- SHOW POLICY & COOLDOWN ---
    #
    Show-PnPSiteVersionPolicy

    if ($global:PnPVersionTrim_LastPolicyUpdateUtc) {
        $minutesSinceChange = (New-TimeSpan -Start $global:PnPVersionTrim_LastPolicyUpdateUtc -End (Get-Date).ToUniversalTime()).TotalMinutes
        if ($minutesSinceChange -lt 30) {
            Write-Host "Policy was just changed ($([math]::Round($minutesSinceChange,1)) minutes ago)." -ForegroundColor Yellow
            Write-Host "Skipping trim to avoid running while policy changes may be pending." -ForegroundColor Yellow
            return
        }
    }

    #
    # --- HELPERS: per-call retry ---
    #
    function Invoke-WithRetry([scriptblock]$Action, [int]$Attempts = $MaxRetries) {
        $try = 1
        while ($true) {
            try {
                & $Action
                break
            }
            catch {
                if ($try -ge $Attempts) { throw }
                $delay = [math]::Pow(2, $try)  # 2,4,8,16...
                Write-Warning "ExecuteQuery failed (attempt $try): $($_.Exception.Message). Retrying in $delay sec..."
                Start-Sleep -Seconds $delay
                $try++
            }
        }
    }

    #
    # --- DISCOVER TARGET LIBRARIES ---
    #
    $ctx    = Get-PnPContext
    $cutoff = (Get-Date).AddDays(-1 * $OlderThanDays)

    # Increase request timeout to reduce timeouts during large deletes (min 10 minutes)
    $ctx.RequestTimeout = [Math]::Max(($MaxBatchMinutes * 60000), 600000)

    $libFilter = @()
    if ($LibraryCsvPath) {
        if (-not (Test-Path $LibraryCsvPath)) {
            Write-Warning "CSV not found: $LibraryCsvPath"
            return
        }
        $csv = Import-Csv $LibraryCsvPath
        foreach ($row in $csv) {
            if ($row.LibraryTitle) { $libFilter += $row.LibraryTitle }
            elseif ($row.Title)   { $libFilter += $row.Title }
        }
        $libFilter = $libFilter | Select-Object -Unique
    }

    if ($LibraryTitle) {
        $lists = Get-PnPList -Identity $LibraryTitle
    }
    else {
        $lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden }
        if ($libFilter.Count -gt 0) {
            $lists = $lists | Where-Object { $libFilter -contains $_.Title }
        }
    }

    if (-not $lists -or $lists.Count -eq 0) {
        Write-Warning "No document libraries found."
        return
    }

    Write-Host ""
    Write-Host ("Target libraries: {0}" -f (($lists | Select-Object -ExpandProperty Title) -join ', ')) -ForegroundColor Green

    #
    # --- SIZE SNAPSHOT (BEFORE) ---
    #
    $runId          = [Guid]::NewGuid().ToString()
    $siteUrl        = (Get-PnPContext).Url
    $startSizeBytes = 0L

    foreach ($list in $lists) {
        $size = Get-PnPListSizeBytes -ListTitle $list.Title
        $startSizeBytes += $size
        if ($LogPath) {
            Write-PnPSizeLog -LogPath $LogPath -RunId $runId `
                -SiteUrl $siteUrl -LibraryTitle $list.Title -Phase 'Before' -SizeBytes $size
        }
    }

    Write-Host ("Starting total size: {0:N2} MB" -f ($startSizeBytes / 1MB)) -ForegroundColor Cyan
    Write-TrimEvent -Level 'Info' -Message "Starting trim: OlderThanDays=$OlderThanDays; Cutoff=$cutoff"

    #
    # --- GLOBAL COUNTERS ---
    #
    $script:ProcessedCount       = 0
    $script:FilesWithOldVersions = 0
    $script:FailedDeletes        = 0
    $script:SkippedByError       = 0
    $script:FilesSeen            = 0

    [int]$MaxFilesPerBatch   = 2000
    [int]$MaxFilesThreshold  = 200000  # safety cutoff

    #
    # --- STREAMING ENUMERATION + TRIM ---
    #
    foreach ($list in $lists) {
        Write-Host "Processing: $($list.Title)" -ForegroundColor Yellow
        $currentListTitle = $list.Title

        Get-PnPListItem -List $list -PageSize 2000 -ScriptBlock {
            param($items)

            if (-not $script:BatchStopwatch) {
                $script:BatchStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                $script:FilesInBatch   = 0
                $script:BatchNumber    = 1
                Write-Host "===== Batch $($script:BatchNumber) =====" -ForegroundColor Magenta
            }

            foreach ($i in $items) {
                if ($i.FileSystemObjectType -ne "File") { continue }

                $script:FilesSeen++
                $script:FilesInBatch++
                $script:ProcessedCount++

                # Safety: too many files? bail.
                if ($script:FilesSeen -gt $MaxFilesThreshold) {
                    throw "FilesSeen ($($script:FilesSeen)) exceeded MaxFilesThreshold ($MaxFilesThreshold). Aborting for safety."
                }

                # Progress (unknown total -> 0% but status is still useful)
                Write-Progress -Activity "Trimming file versions" `
                               -Status "Processed $($script:ProcessedCount) files; $($script:FilesWithOldVersions) had old versions" `
                               -PercentComplete 0

                # Load file + versions using existing list item
                $ctxInner = Get-PnPContext
                $ctxInner.Load($i.File)
                $ctxInner.Load($i.File.Versions)

                try {
                    Invoke-WithRetry { $ctxInner.ExecuteQuery() }
                }
                catch {
                    $msg = $_.Exception.Message
                    Write-TrimEvent -Level 'Error' -Message ("Failed to load file for item ID {0} in {1}: {2}" -f $i.Id, $currentListTitle, $msg)
                    $script:SkippedByError++
                    continue
                }

                $file     = $i.File
                $versions = $file.Versions

                # Filter versions older than cutoff (extra safety: not current)
                $old = @()
                foreach ($v in $versions) {
                    if (-not $v.IsCurrentVersion -and $v.Created -lt $cutoff) { $old += $v }
                }
                if ($old.Count -eq 0) { continue }

                $script:FilesWithOldVersions++

                # DRY RUN: only log what would be deleted
                if ($effectiveDryRun) {
                    foreach ($v in $old) {
                        Write-TrimLog "DryRun" $currentListTitle $i["FileRef"] $v.ID $v.VersionLabel $v.Created "Planned" "DryRun - would delete"
                    }
                    continue
                }

                # DELETE MODE: delete in chunks to avoid huge CSOM calls
                $chunkSize = $VersionBatchSize
                if ($chunkSize -le 0) { $chunkSize = 50 }

                for ($idx = 0; $idx -lt $old.Count; $idx += $chunkSize) {
                    $chunk = $old[$idx..([math]::Min($idx + $chunkSize - 1, $old.Count - 1))]

                    foreach ($v in $chunk) {
                        $v.DeleteObject()
                    }

                    try {
                        Invoke-WithRetry { $ctxInner.ExecuteQuery() }

                        foreach ($v in $chunk) {
                            Write-TrimLog "Delete" $currentListTitle $i["FileRef"] $v.ID $v.VersionLabel $v.Created "Deleted" "Deleted"
                        }

                        if ($VersionBatchPauseMs -gt 0) {
                            Start-Sleep -Milliseconds $VersionBatchPauseMs
                        }
                    }
                    catch {
                        $script:FailedDeletes++
                        $msg = $_.Exception.Message

                        if ($msg -match 'retention|hold|record') {
                            $script:SkippedByError++
                            Write-TrimEvent -Level 'Warn' -Message ("Skipped {0} due to retention/hold: {1}" -f $i["FileRef"], $msg)
                        }
                        else {
                            Write-TrimEvent -Level 'Error' -Message ("Failed to delete versions for {0}: {1}" -f $i["FileRef"], $msg)
                        }

                        foreach ($v in $chunk) {
                            Write-TrimLog "Delete" $currentListTitle $i["FileRef"] $v.ID $v.VersionLabel $v.Created "Failed" $msg
                        }
                    }
                }

                # Batch timing / user prompt
                if (-not $BypassBatching -and $script:BatchStopwatch.Elapsed.TotalMinutes -ge $MaxBatchMinutes) {
                    Write-Host "Batch $($script:BatchNumber) time limit reached." -ForegroundColor Yellow

                    if (-not $AutoContinue) {
                        $resp = Read-Host "Press Enter to continue, or 'q' to quit"
                        if ($resp -eq 'q') {
                            throw "User aborted batches."
                        }
                    }

                    $script:BatchStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    $script:FilesInBatch   = 0
                    $script:BatchNumber++
                    Write-Host "===== Batch $($script:BatchNumber) =====" -ForegroundColor Magenta
                }
            }
        }
    }

    # Finish progress bar
    Write-Progress -Activity "Trimming file versions" -Completed

    #
    # --- SIZE SNAPSHOT (AFTER) ---
    #
    $endSizeBytes = 0L
    foreach ($list in $lists) {
        $size = Get-PnPListSizeBytes -ListTitle $list.Title
        $endSizeBytes += $size
        if ($LogPath) {
            Write-PnPSizeLog -LogPath $LogPath -RunId $runId `
                -SiteUrl $siteUrl -LibraryTitle $list.Title -Phase 'After' -SizeBytes $size
        }
    }

    $reclaimed = $startSizeBytes - $endSizeBytes

    #
    # --- SUMMARY ---
    #
    Write-Host ""
    Write-Host "===== Version trim summary =====" -ForegroundColor Cyan
    Write-Host ("  Total files scanned        : {0}" -f $script:ProcessedCount)
    Write-Host ("  Files with old versions    : {0}" -f $script:FilesWithOldVersions)
    Write-Host ("  Failed deletions           : {0}" -f $script:FailedDeletes)
    Write-Host ("  Skipped due to errors/holds: {0}" -f $script:SkippedByError)
    if ($LogPath)     { Write-Host ("  CSV log                    : {0}" -f $LogPath) }
    if ($TextLogPath) { Write-Host ("  Text log                   : {0}" -f $TextLogPath) }

    Write-Host ("Ending total size   : {0:N2} MB" -f ($endSizeBytes / 1MB)) -ForegroundColor Cyan
    Write-Host ("Estimated reclaimed : {0:N2} MB" -f ($reclaimed / 1MB)) -ForegroundColor Green

    #
    # WRITE STATE FILE (mark that first dry run has occurred)
    #
    $stateOut = @{ LastDryRunUtc = (Get-Date).ToUniversalTime().ToString("o") } | ConvertTo-Json
    Set-Content -Path $stateFile -Value $stateOut -Encoding UTF8

    Write-Host ""
    if ($effectiveDryRun) {
        Write-Host "DONE (Dry Run). Next run can use -Delete." -ForegroundColor Cyan
    }
    else {
        Write-Host "DONE (Deleted permitted versions)." -ForegroundColor Green
    }
}
