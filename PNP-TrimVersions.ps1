<# 
.SYNOPSIS
  PNP-TrimVersions - Safe SharePoint Online file version trim tool (PnP.PowerShell + CSOM)

.DESCRIPTION
  - Uses PnP.PowerShell only (no SPO Management Shell).
  - Assumes YOU manage authentication: connect first with Connect-PnPOnline.
  - Reads site version policy and blocks execution if policy indicates Pending (when available).
  - First run is always Dry Run (per-user state file).
  - Deletes versions older than -OlderThanDays in per-file chunks (reduces timeouts).
  - Streams list items (low memory) unless you explicitly opt into building a work list.
  - Logs ONLY exceptions / skips (not every successful delete) to keep logging cheap.

REQUIREMENTS
  - PnP.PowerShell
  - Connected session:
      Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/site" -Interactive

NOTES
  - SharePoint may block deletes due to retention, eDiscovery holds, records, etc.
  - This tool is designed to be crash-safe: if it stops mid-run, you can rerun; it does not "lock" files.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Invoke-PnPVersionTrim {
  [CmdletBinding(SupportsShouldProcess=$true)]
  param(
    [Parameter(Mandatory=$false)]
    [int]$OlderThanDays = 45,

    # Target selection: specify one, some, or all. If none supplied, you MUST choose via -AllLibraries.
    [Parameter(Mandatory=$false)]
    [string[]]$LibraryTitle,

    [Parameter(Mandatory=$false)]
    [string]$LibraryCsvPath,

    [Parameter(Mandatory=$false)]
    [switch]$AllLibraries,

    # Deletion mode (first run is forced DryRun even if -Delete is set)
    [Parameter(Mandatory=$false)]
    [switch]$Delete,

    # Guardrails
    [Parameter(Mandatory=$false)]
    [ValidateRange(5,3650)]
    [int]$MinOlderThanDays = 5,

    [Parameter(Mandatory=$false)]
    [int]$MaxFilesThreshold = 200000,

    # Streaming "batch" guardrails (not pre-enumeration batching)
    [Parameter(Mandatory=$false)]
    [ValidateRange(1,2000000)]
    [int]$PauseEveryFiles = 5000,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1,240)]
    [int]$PauseEveryMinutes = 10,

    [Parameter(Mandatory=$false)]
    [switch]$AutoContinue,

    # Per-file version chunking + backoff
    [Parameter(Mandatory=$false)]
    [ValidateRange(1,500)]
    [int]$VersionBatchSize = 50,

    [Parameter(Mandatory=$false)]
    [ValidateRange(0,60000)]
    [int]$VersionBatchPauseMs = 250,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1,10)]
    [int]$MaxRetries = 5,

    # Request timeout for CSOM ExecuteQuery (ms)
    [Parameter(Mandatory=$false)]
    [ValidateRange(60000,3600000)]
    [int]$RequestTimeoutMs = 600000,

    # Optional reporting
    [Parameter(Mandatory=$false)]
    [switch]$MeasureLibrarySizes,

    [Parameter(Mandatory=$false)]
    [string]$ExceptionCsvPath = ".\SPO-TrimVersions-Exceptions.csv",

    [Parameter(Mandatory=$false)]
    [string]$TextLogPath = ".\SPO-TrimVersions.log",

    # Optional: skip files whose names contain any token from CSV (column: Token)
    [Parameter(Mandatory=$false)]
    [string]$SkipNameContainsCsvPath
  )

  # -----------------------------
  # Helpers
  # -----------------------------
  function Write-TrimEvent {
    param(
      [ValidateSet('Info','Warn','Error')]
      [string]$Level,
      [string]$Message
    )
    $ts = (Get-Date).ToString('o')
    $line = "[{0}] [{1}] {2}" -f $ts, $Level.ToUpper(), $Message

    if ($TextLogPath) {
      $dir = Split-Path -Parent $TextLogPath
      if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
      Add-Content -Path $TextLogPath -Value $line
    }

    if ($Level -eq 'Warn') { Write-Host $line -ForegroundColor Yellow }
    if ($Level -eq 'Error') { Write-Host $line -ForegroundColor Red }
  }

  function Ensure-ExceptionCsv {
    if (-not $ExceptionCsvPath) { return }
    $dir = Split-Path -Parent $ExceptionCsvPath
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    if (-not (Test-Path $ExceptionCsvPath)) {
      "Timestamp,SiteUrl,LibraryTitle,FileRef,ItemId,Action,Result,Message" | Out-File -FilePath $ExceptionCsvPath -Encoding UTF8
    }
  }

  function Write-ExceptionRow {
    param(
      [string]$SiteUrl,[string]$Library,[string]$FileRef,[int]$ItemId,
      [string]$Action,[string]$Result,[string]$Message
    )
    if (-not $ExceptionCsvPath) { return }
    Ensure-ExceptionCsv
    $line = "{0},{1},{2},{3},{4},{5},{6},{7}" -f (Get-Date).ToString('o'),
      ($SiteUrl -replace ',', ' '),
      ($Library -replace ',', ' '),
      ($FileRef -replace ',', ' '),
      $ItemId,
      ($Action -replace ',', ' '),
      ($Result -replace ',', ' '),
      (($Message -replace "`r|`n",' ') -replace ',', ' ')
    Add-Content -Path $ExceptionCsvPath -Value $line
  }

  function Invoke-WithRetry {
    param(
      [scriptblock]$Action,
      [int]$Attempts = 5
    )
    $try = 1
    while ($true) {
      try {
        & $Action
        return
      } catch {
        if ($try -ge $Attempts) { throw }
        $delaySec = [math]::Pow(2, $try)
        Write-TrimEvent -Level 'Warn' -Message ("ExecuteQuery failed on attempt {0}. Waiting {1}s. {2}" -f $try, $delaySec, $_.Exception.Message)
        Start-Sleep -Seconds $delaySec
        $try++
      }
    }
  }

  function Get-SkipTokens {
    if (-not $SkipNameContainsCsvPath) { return @() }
    if (-not (Test-Path $SkipNameContainsCsvPath)) {
      throw "SkipNameContainsCsvPath not found: $SkipNameContainsCsvPath"
    }
    $rows = Import-Csv -Path $SkipNameContainsCsvPath
    $tokens = @()
    foreach ($r in $rows) {
      if ($r.Token) { $tokens += [string]$r.Token }
    }
    $tokens | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
  }

  function Get-PnPListSizeBytes {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ListTitle)

    # Uses list item field File_x0020_Size (fast-ish, avoids Folder.Files enumeration).
    # Still enumerates all items, so this is optional.
    $total = 0L
    Get-PnPListItem -List $ListTitle -PageSize 2000 -Fields "FSObjType","File_x0020_Size" -ScriptBlock {
      param($items)
      foreach ($it in $items) {
        if ($it["FSObjType"] -eq 0 -and $it["File_x0020_Size"]) {
          $script:__sizeTotal += [int64]$it["File_x0020_Size"]
        }
      }
    } | Out-Null
    return $script:__sizeTotal
  }

  function Show-PnPSiteVersionPolicy {
    [CmdletBinding()]
    param()

    try { $policy = Get-PnPSiteVersionPolicy } catch {
      Write-TrimEvent -Level 'Warn' -Message ("Failed to get site version policy. {0}" -f $_.Exception.Message)
      return $null
    }
    if (-not $policy) { return $null }

    Write-Host "========== Site Version Policy ==========" -ForegroundColor Cyan
    $policy | Format-List *
    Write-Host "=========================================" -ForegroundColor Cyan

    # If policy exposes a Status field and it is Pending, block.
    if ($policy.PSObject.Properties.Name -contains 'Status') {
      $st = [string]$policy.Status
      if ($st -match 'Pending') {
        Write-TrimEvent -Level 'Warn' -Message ("Policy Status is Pending. Refusing to run deletes right now. Status={0}" -f $st)
        return $policy
      }
    }
    return $policy
  }

  function Get-StateFilePath {
    param([string]$SiteUrl)

    $root = Join-Path $env:LOCALAPPDATA "PNP-TrimVersions"
    if (-not (Test-Path $root)) { New-Item -ItemType Directory -Force -Path $root | Out-Null }
    $safe = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($SiteUrl)).TrimEnd('=').Replace('/','_').Replace('+','-')
    Join-Path $root ("state-{0}.json" -f $safe)
  }

  function Read-State {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try { return (Get-Content -Raw -Path $Path | ConvertFrom-Json) } catch { return $null }
  }

  function Write-State {
    param([string]$Path,[hashtable]$Obj)
    $Obj | ConvertTo-Json | Set-Content -Encoding UTF8 -Path $Path
  }

  # -----------------------------
  # Preconditions
  # -----------------------------
  if ($OlderThanDays -lt $MinOlderThanDays) {
    throw "OlderThanDays must be >= ${MinOlderThanDays}."
  }

  $ctx = Get-PnPContext
  if (-not $ctx) { throw "No PnP context found. Connect first with Connect-PnPOnline." }
  $ctx.RequestTimeout = $RequestTimeoutMs

  $siteUrl = $ctx.Url
  $stateFile = Get-StateFilePath -SiteUrl $siteUrl
  $state = Read-State -Path $stateFile

  $firstRun = -not $state
  $effectiveDryRun = $true
  if (-not $firstRun -and $Delete) { $effectiveDryRun = $false }

  if ($firstRun) {
    Write-Host "First run for this site. FORCING DRY RUN." -ForegroundColor Yellow
  }

  if (-not $effectiveDryRun) {
    $confirm = Read-Host "DELETE MODE. Type DELETE to proceed"
    if ($confirm -ne 'DELETE') { Write-Host "Cancelled." -ForegroundColor Yellow; return }
  } else {
    Write-Host "DRY RUN: No versions will be deleted." -ForegroundColor Cyan
  }

  # Policy info + pending check
  $policy = Show-PnPSiteVersionPolicy
  if ($policy -and ($policy.PSObject.Properties.Name -contains 'Status')) {
    if ([string]$policy.Status -match 'Pending') { return }
  }

  # If state has a recent policy update timestamp, block deletes for cooldown.
  if ($state -and $state.LastPolicyUpdateUtc) {
    $mins = (New-TimeSpan -Start ([datetime]$state.LastPolicyUpdateUtc) -End (Get-Date).ToUniversalTime()).TotalMinutes
    if ($mins -lt 30) {
      Write-TrimEvent -Level 'Warn' -Message ("Policy was changed recently ({0} min). Refusing to run." -f [math]::Round($mins,1))
      return
    }
  }

  # -----------------------------
  # Determine target libraries
  # -----------------------------
  $targetTitles = @()
  if ($LibraryCsvPath) {
    if (-not (Test-Path $LibraryCsvPath)) { throw "LibraryCsvPath not found: $LibraryCsvPath" }
    $rows = Import-Csv -Path $LibraryCsvPath
    foreach ($r in $rows) {
      if ($r.LibraryTitle) { $targetTitles += [string]$r.LibraryTitle }
      elseif ($r.Title) { $targetTitles += [string]$r.Title }
    }
  }
  if ($LibraryTitle) { $targetTitles += $LibraryTitle }
  $targetTitles = $targetTitles | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

  if (-not $AllLibraries -and $targetTitles.Count -eq 0) {
    throw "No libraries selected. Use -LibraryTitle, -LibraryCsvPath, or -AllLibraries."
  }

  $lists = @()
  if ($AllLibraries) {
    $lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden }
  } else {
    foreach ($t in $targetTitles) {
      try { $lists += Get-PnPList -Identity $t } catch { throw "Library not found: $t" }
    }
  }
  $lists = $lists | Select-Object -Unique

  if (-not $lists -or $lists.Count -eq 0) { throw "No document libraries found." }

  Write-Host ("Target libraries: {0}" -f (($lists | Select-Object -ExpandProperty Title) -join ', ')) -ForegroundColor Green

  # -----------------------------
  # Optional: size reporting
  # -----------------------------
  $startTotalBytes = 0L
  if ($MeasureLibrarySizes) {
    foreach ($l in $lists) {
      $script:__sizeTotal = 0L
      try {
        $b = Get-PnPListSizeBytes -ListTitle $l.Title
        $startTotalBytes += $b
      } catch {
        Write-TrimEvent -Level 'Warn' -Message ("Size read failed for {0}. {1}" -f $l.Title, $_.Exception.Message)
      }
    }
    Write-Host ("Starting size (est.): {0:N2} MB" -f ($startTotalBytes / 1MB)) -ForegroundColor Cyan
  }

  # -----------------------------
  # Main run (streaming)
  # -----------------------------
  $cutoff = (Get-Date).AddDays(-1 * $OlderThanDays)
  $skipTokens = Get-SkipTokens

  $processed = 0
  $filesWithOld = 0
  $failedDeletes = 0
  $skipped = 0

  $runStart = Get-Date
  $lastPauseAt = Get-Date
  $sincePauseFiles = 0

  foreach ($l in $lists) {
    $libTitle = $l.Title
    Write-Host ("Scanning library: {0}" -f $libTitle) -ForegroundColor Yellow

    # Keep fields minimal for performance. (FileRef is needed for logging; FileLeafRef for skip tokens.)
    Get-PnPListItem -List $l -PageSize 2000 -Fields "FileRef","FileLeafRef","FSObjType" -ScriptBlock {
      param($items)

      foreach ($it in $items) {
        if ($it.FileSystemObjectType -ne "File") { continue }

        $script:processed++
        $script:sincePauseFiles++

        if ($script:processed -gt $MaxFilesThreshold) {
          throw ("MaxFilesThreshold exceeded ({0}). Aborting for safety." -f $MaxFilesThreshold)
        }

        # Progress (single dynamic output)
        if (($script:processed % 50) -eq 0) {
          $elapsed = (New-TimeSpan -Start $script:runStart -End (Get-Date)).ToString()
          Write-Progress -Activity "Trimming file versions" -Status ("Processed {0} files. OldVersionFiles {1}. Failed {2}. Skipped {3}. Elapsed {4}" -f $script:processed,$script:filesWithOld,$script:failedDeletes,$script:skipped,$elapsed) -PercentComplete 0
        }

        $fileRef = [string]$it["FileRef"]
        $fileLeaf = [string]$it["FileLeafRef"]
        $itemId = [int]$it.Id

        # Optional skip tokens
        if ($script:skipTokens -and $script:skipTokens.Count -gt 0) {
          foreach ($tok in $script:skipTokens) {
            if ($fileLeaf -and $fileLeaf.IndexOf($tok, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
              $script:skipped++
              Write-ExceptionRow -SiteUrl $script:siteUrl -Library $script:libTitle -FileRef $fileRef -ItemId $itemId -Action "Skip" -Result "Skipped" -Message ("Name contains token '{0}'" -f $tok)
              return
            }
          }
        }

        # Load by list + item ID (avoids URL normalization issues)
        $listObj = $script:ctx.Web.Lists.GetByTitle($script:libTitle)
        $li = $listObj.GetItemById($itemId)
        $script:ctx.Load($li)
        $script:ctx.Load($li.File)
        $script:ctx.Load($li.File.Versions)

        try {
          Invoke-WithRetry -Action { $script:ctx.ExecuteQuery() } -Attempts $script:MaxRetries
        } catch {
          $script:skipped++
          Write-ExceptionRow -SiteUrl $script:siteUrl -Library $script:libTitle -FileRef $fileRef -ItemId $itemId -Action "Load" -Result "Failed" -Message $_.Exception.Message
          return
        }

        $file = $li.File
        $versions = $file.Versions
        $old = @()
        foreach ($v in $versions) {
          if (-not $v.IsCurrentVersion -and $v.Created -lt $script:cutoff) { $old += $v }
        }
        if ($old.Count -eq 0) { return }

        $script:filesWithOld++

        if ($script:effectiveDryRun) {
          # Only log summary of dry run (exceptions file stays about exceptions).
          return
        }

        # Delete in chunks
        for ($i = 0; $i -lt $old.Count; $i += $script:VersionBatchSize) {
          $end = [math]::Min($i + $script:VersionBatchSize - 1, $old.Count - 1)
          $chunk = $old[$i..$end]
          foreach ($v in $chunk) { $v.DeleteObject() }

          try {
            Invoke-WithRetry -Action { $script:ctx.ExecuteQuery() } -Attempts $script:MaxRetries
          } catch {
            $script:failedDeletes++
            $msg = $_.Exception.Message
            # Most common "expected" blocks are retention/hold/record.
            $result = if ($msg -match 'retention|hold|record') { "Blocked" } else { "Failed" }
            Write-ExceptionRow -SiteUrl $script:siteUrl -Library $script:libTitle -FileRef $fileRef -ItemId $itemId -Action "Delete" -Result $result -Message $msg
            break
          }

          if ($script:VersionBatchPauseMs -gt 0) { Start-Sleep -Milliseconds $script:VersionBatchPauseMs }
        }

        # Pause guardrail
        $minsSincePause = (New-TimeSpan -Start $script:lastPauseAt -End (Get-Date)).TotalMinutes
        if ($script:sincePauseFiles -ge $script:PauseEveryFiles -or $minsSincePause -ge $script:PauseEveryMinutes) {
          $script:sincePauseFiles = 0
          $script:lastPauseAt = Get-Date
          if (-not $script:AutoContinue) {
            $resp = Read-Host "Checkpoint reached. Press Enter to continue, or type q to stop"
            if ($resp -eq 'q') { throw "User aborted." }
          }
        }
      }
    } | Out-Null
  }

  Write-Progress -Activity "Trimming file versions" -Completed

  # -----------------------------
  # Optional: end size reporting
  # -----------------------------
  $endTotalBytes = 0L
  if ($MeasureLibrarySizes) {
    foreach ($l in $lists) {
      $script:__sizeTotal = 0L
      try {
        $b = Get-PnPListSizeBytes -ListTitle $l.Title
        $endTotalBytes += $b
      } catch {
        Write-TrimEvent -Level 'Warn' -Message ("Size read failed for {0}. {1}" -f $l.Title, $_.Exception.Message)
      }
    }
    Write-Host ("Ending size (est.): {0:N2} MB" -f ($endTotalBytes / 1MB)) -ForegroundColor Cyan
    Write-Host ("Estimated reclaimed: {0:N2} MB" -f (($startTotalBytes - $endTotalBytes) / 1MB)) -ForegroundColor Green
  }

  # -----------------------------
  # Persist state
  # -----------------------------
  $newState = @{
    LastRunUtc = (Get-Date).ToUniversalTime().ToString('o')
    LastDryRunUtc = if ($effectiveDryRun) { (Get-Date).ToUniversalTime().ToString('o') } else { $null }
  }
  if ($state -and $state.LastPolicyUpdateUtc) { $newState.LastPolicyUpdateUtc = $state.LastPolicyUpdateUtc }
  Write-State -Path $stateFile -Obj $newState

  Write-Host ""
  Write-Host "===== Summary =====" -ForegroundColor Cyan
  Write-Host ("Site: {0}" -f $siteUrl)
  Write-Host ("Processed files      : {0}" -f $processed)
  Write-Host ("Files with old vers. : {0}" -f $filesWithOld)
  Write-Host ("Failed/Blocked deletes: {0}" -f $failedDeletes)
  Write-Host ("Skipped (load/filters): {0}" -f $skipped)
  if ($effectiveDryRun) { Write-Host "Mode: DRY RUN (first run is forced)" -ForegroundColor Cyan }
  else { Write-Host "Mode: DELETE" -ForegroundColor Green }

  if ($ExceptionCsvPath) { Write-Host ("Exceptions CSV: {0}" -f $ExceptionCsvPath) }
  if ($TextLogPath) { Write-Host ("Text log: {0}" -f $TextLogPath) }
}

Export-ModuleMember -Function Invoke-PnPVersionTrim -ErrorAction SilentlyContinue
