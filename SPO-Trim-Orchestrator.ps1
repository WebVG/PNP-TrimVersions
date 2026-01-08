<#
.SYNOPSIS
  SPO Trim Orchestrator - queues multiple site URLs for manual Connect-PnPOnline + Invoke-PnPVersionTrim.

.DESCRIPTION
  - Reads a Sites CSV (column: SiteUrl) and optional Library targeting columns.
  - DOES NOT handle credentials or tokens.
  - For each site, it prompts you to Connect-PnPOnline -Url <site> -Interactive, then runs Invoke-PnPVersionTrim.
  - Supports a "pilot key" guardrail so you can't jump straight into multi-site deletes by accident.

CSV Columns (minimum):
  SiteUrl

Optional columns:
  Enabled (true/false, default true)
  AllLibraries (true/false)
  LibraryTitles (semicolon-separated list)
  LibraryCsvPath (path to per-site library csv)
  OlderThanDays (int)
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

param(
  [Parameter(Mandatory)]
  [string]$SitesCsvPath,

  [Parameter(Mandatory)]
  [string]$TrimScriptPath,

  [switch]$Delete,

  # Guardrail: require a pilot key for multi-site delete mode
  [string]$PilotKey,

  # Same key every time is fine; goal is "stop and think", not security
  [string]$ExpectedPilotKey = "I_RAN_A_PILOT",

  [switch]$AutoContinue
)

if (-not (Test-Path $SitesCsvPath)) { throw "SitesCsvPath not found: $SitesCsvPath" }
if (-not (Test-Path $TrimScriptPath)) { throw "TrimScriptPath not found: $TrimScriptPath" }

. $TrimScriptPath

$sites = Import-Csv -Path $SitesCsvPath

if ($Delete -and $sites.Count -gt 1) {
  if ([string]::IsNullOrWhiteSpace($PilotKey) -or $PilotKey -ne $ExpectedPilotKey) {
    throw "Multi-site DELETE is blocked. Provide -PilotKey $ExpectedPilotKey after you have proven a pilot run."
  }
}

foreach ($s in $sites) {
  $enabled = $true
  if ($s.Enabled) { $enabled = [bool]::Parse([string]$s.Enabled) }
  if (-not $enabled) { continue }

  $url = [string]$s.SiteUrl
  if ([string]::IsNullOrWhiteSpace($url)) { continue }

  Write-Host ""
  Write-Host ("=== Site: {0} ===" -f $url) -ForegroundColor Cyan
  Write-Host "Connect now in THIS session, then press Enter:" -ForegroundColor Yellow
  Write-Host ("  Connect-PnPOnline -Url ""{0}"" -Interactive" -f $url) -ForegroundColor Yellow
  [void](Read-Host)

  $args = @{
    OlderThanDays = if ($s.OlderThanDays) { [int]$s.OlderThanDays } else { 45 }
    Delete = $Delete
    AutoContinue = $AutoContinue
  }

  if ($s.AllLibraries -and ([string]$s.AllLibraries).ToLower() -eq 'true') {
    $args.AllLibraries = $true
  } elseif ($s.LibraryTitles) {
    $args.LibraryTitle = ([string]$s.LibraryTitles).Split(';') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  } elseif ($s.LibraryCsvPath) {
    $args.LibraryCsvPath = [string]$s.LibraryCsvPath
  } else {
    throw "Site ${url} has no library targeting. Set AllLibraries=true or LibraryTitles or LibraryCsvPath."
  }

  Invoke-PnPVersionTrim @args
}
