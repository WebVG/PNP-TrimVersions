<#
.SYNOPSIS
  Generates templates for SPO trimming: Sites CSV, Library CSV, Skip tokens CSV, Guardrails YAML.

.NOTES
  This does NOT call SharePoint. It is just scaffolding so admins don't fat-finger parameters.
#>

param(
  [string]$OutDir = ".\Trim-Config"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }

@"
SiteUrl,Enabled,AllLibraries,LibraryTitles,LibraryCsvPath,OlderThanDays
https://tenant.sharepoint.com/sites/Finance,true,false,Documents;Shared Documents,,60
https://tenant.sharepoint.com/sites/HR,true,true,,,45
"@ | Set-Content -Encoding UTF8 -Path (Join-Path $OutDir "Sites.csv")

@"
LibraryTitle
Documents
Shared Documents
"@ | Set-Content -Encoding UTF8 -Path (Join-Path $OutDir "Libraries.csv")

@"
Token
Budget
Quarterly
Checklist
2024
"@ | Set-Content -Encoding UTF8 -Path (Join-Path $OutDir "SkipNameContains.csv")

@"
hostProfiles:
  Laptop:
    maxFilesThreshold: 50000
    pauseEveryFiles: 3000
    pauseEveryMinutes: 10
    versionBatchSize: 25
    versionBatchPauseMs: 750

  Desktop:
    maxFilesThreshold: 100000
    pauseEveryFiles: 5000
    pauseEveryMinutes: 10
    versionBatchSize: 50
    versionBatchPauseMs: 250

trimProfiles:
  Normal:
    olderThanDays: 45
    versionBatchSize: 50
    versionBatchPauseMs: 250

sites: []
"@ | Set-Content -Encoding UTF8 -Path (Join-Path $OutDir "Guardrails.template.yml")

Write-Host ("Wrote templates to {0}" -f (Resolve-Path $OutDir)) -ForegroundColor Green
