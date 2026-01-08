<#
.SYNOPSIS
  PerfMode - optional local performance boost + restore (no virtualization changes)

.DESCRIPTION
  - Writes all changes to a small JSON state file under LOCALAPPDATA.
  - Restore attempts to revert everything it changed, and prints anything it could not revert.
  - Includes a hard confirmation if you combine aggressive options.

NOTES
  This is intentionally conservative. It does NOT disable Hyper-V / VM services.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

param(
  [ValidateSet('Boost','Restore')]
  [string]$Mode = 'Boost',

  [ValidateSet('MSInfra','DevBox','Minimal')]
  [string]$Profile = 'Minimal',

  [switch]$Aggressive,
  [switch]$KillExplorer,

  [string]$LogPath = ".\PerfMode.log"
)

function Write-PerfLog {
  param(
    [string]$Message,
    [ValidateSet('Info','Warn','Error','Verbose')]
    [string]$Level = 'Info'
  )
  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  $line = "[{0}] [{1}] {2}" -f $ts, $Level.ToUpper(), $Message
  $dir = Split-Path -Parent $LogPath
  if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
  Add-Content -Path $LogPath -Value $line
  if ($Level -eq 'Error') { Write-Host $line -ForegroundColor Red }
  elseif ($Level -eq 'Warn') { Write-Host $line -ForegroundColor Yellow }
  elseif ($Level -eq 'Verbose') { Write-Host $line -ForegroundColor DarkGray }
  else { Write-Host $line }
}

function Get-StatePath {
  $root = Join-Path $env:LOCALAPPDATA "PerfMode"
  if (-not (Test-Path $root)) { New-Item -ItemType Directory -Force -Path $root | Out-Null }
  Join-Path $root "state.json"
}

function Save-State($obj) {
  $p = Get-StatePath
  ($obj | ConvertTo-Json -Depth 6) | Set-Content -Encoding UTF8 -Path $p
}

function Load-State {
  $p = Get-StatePath
  if (-not (Test-Path $p)) { return $null }
  try { Get-Content -Raw -Path $p | ConvertFrom-Json } catch { return $null }
}

function Get-CurrentPowerPlanGuid {
  $out = (powercfg /getactivescheme 2>$null)
  # Expected: "Power Scheme GUID: XXXXXXXX-....  (Name)"
  if ($out -match 'GUID:\s+([0-9a-fA-F-]+)') { return $matches[1] }
  return $null
}

function Set-PowerPlanGuid([string]$Guid) {
  if (-not $Guid) { return }
  powercfg /setactive $Guid | Out-Null
}

function Stop-ProcessSafe([string]$Name) {
  $p = Get-Process -Name $Name -ErrorAction SilentlyContinue
  if ($p) {
    try { $p | Stop-Process -Force -ErrorAction Stop; return $true } catch { return $false }
  }
  return $false
}

# Guard: hard confirmation
if ($Mode -eq 'Boost' -and $Aggressive -and $KillExplorer -and $Profile -ne 'Minimal') {
  $resp = Read-Host "Aggressive + KillExplorer + Profile ${Profile}. Type I UNDERSTAND to continue"
  if ($resp -ne 'I UNDERSTAND') { throw "Cancelled." }
}

if ($Mode -eq 'Boost') {
  $state = [ordered]@{
    CreatedUtc = (Get-Date).ToUniversalTime().ToString('o')
    OriginalPowerPlanGuid = (Get-CurrentPowerPlanGuid)
    ExplorerWasRunning = $false
    Profile = $Profile
    Changes = @()
  }

  Write-PerfLog "Boost starting. Profile ${Profile}. Aggressive=${Aggressive}. KillExplorer=${KillExplorer}" 'Info'
  Write-PerfLog ("Original power plan GUID: {0}" -f $state.OriginalPowerPlanGuid) 'Verbose'

  # Power plan: try High performance if present, else leave as-is
  $highPerfGuid = $null
  $schemes = (powercfg /list 2>$null)
  if ($schemes -match 'High performance\s+\*\*') { } # ignore
  if ($schemes -match '([0-9a-fA-F-]{36}).*High performance') { $highPerfGuid = $matches[1] }

  if ($highPerfGuid) {
    try {
      Set-PowerPlanGuid -Guid $highPerfGuid
      $state.Changes += "PowerPlan->HighPerformance"
      Write-PerfLog ("Switched power plan to High performance ({0})" -f $highPerfGuid) 'Info'
    } catch {
      Write-PerfLog ("Failed to switch power plan. {0}" -f $_.Exception.Message) 'Warn'
    }
  } else {
    Write-PerfLog "High performance plan not found. Leaving current plan." 'Verbose'
  }

  if ($KillExplorer) {
    $state.ExplorerWasRunning = [bool](Get-Process explorer -ErrorAction SilentlyContinue)
    if ($state.ExplorerWasRunning) {
      $ok = Stop-ProcessSafe -Name "explorer"
      if ($ok) {
        $state.Changes += "Stopped explorer.exe"
        Write-PerfLog "Stopped explorer.exe" 'Warn'
      } else {
        Write-PerfLog "Failed to stop explorer.exe" 'Warn'
      }
    }
  }

  Save-State $state
  Write-PerfLog ("Boost complete. State saved: {0}" -f (Get-StatePath)) 'Info'
  return
}

if ($Mode -eq 'Restore') {
  $state = Load-State
  if (-not $state) { throw "No PerfMode state file found to restore." }

  $notReverted = @()

  Write-PerfLog ("Restoring from state created {0}" -f $state.CreatedUtc) 'Info'

  if ($state.OriginalPowerPlanGuid) {
    try {
      Set-PowerPlanGuid -Guid $state.OriginalPowerPlanGuid
      Write-PerfLog ("Restored power plan GUID: {0}" -f $state.OriginalPowerPlanGuid) 'Info'
    } catch {
      $notReverted += "Power plan"
      Write-PerfLog ("Failed to restore power plan. {0}" -f $_.Exception.Message) 'Warn'
    }
  }

  if ($state.ExplorerWasRunning) {
    try {
      Start-Process explorer.exe | Out-Null
      Write-PerfLog "Restarted explorer.exe" 'Info'
    } catch {
      $notReverted += "explorer.exe"
      Write-PerfLog ("Failed to restart explorer.exe. {0}" -f $_.Exception.Message) 'Warn'
    }
  }

  if ($notReverted.Count -gt 0) {
    Write-PerfLog ("Restore finished with items not reverted: {0}" -f ($notReverted -join '; ')) 'Warn'
  } else {
    Write-PerfLog "Restore complete. All tracked changes reverted." 'Info'
  }
}
