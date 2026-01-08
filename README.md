# PNP Trim Suite (SharePoint Online version trimming)

## What this is
A small, opinionated set of scripts to trim **file version history** in SharePoint Online document libraries using **PnP.PowerShell + CSOM**.

It is designed to be:
- **Safe-ish by default** (first run per site is forced Dry Run)
- **Low memory** (streams list items, does not build a giant work list)
- **Resilient** (per-file deletion in chunks + retry/backoff)
- **Quiet** (logs only exceptions/skips, not every successful delete)

## Files
- `PNP-TrimVersions.ps1`  
  Defines `Invoke-PnPVersionTrim` (dot-source this).
- `SPO-Trim-Orchestrator.ps1`  
  Loops over many sites from a CSV. Does **not** manage auth; it prompts you to run `Connect-PnPOnline` per site.
- `New-TrimTemplates.ps1`  
  Generates template CSV/YAML config files.
- `PerfMode.ps1`  
  Optional local "Boost/Restore" helper. Conservative by design.

## Install prerequisites
```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

## Basic usage (single site)
```powershell
# 1) Connect
Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/Finance" -Interactive

# 2) Load the tool
. .\PNP-TrimVersions.ps1

# 3) Dry run (first run is ALWAYS dry-run)
Invoke-PnPVersionTrim -LibraryTitle "Documents" -OlderThanDays 45

# 4) Real delete (2nd run+)
Invoke-PnPVersionTrim -LibraryTitle "Documents" -OlderThanDays 45 -Delete
```

## All libraries (explicit)
```powershell
Invoke-PnPVersionTrim -AllLibraries -OlderThanDays 60
Invoke-PnPVersionTrim -AllLibraries -OlderThanDays 60 -Delete
```

## Use a library list CSV
`Libraries.csv`
```csv
LibraryTitle
Documents
Shared Documents
```

```powershell
Invoke-PnPVersionTrim -LibraryCsvPath .\Trim-Config\Libraries.csv -OlderThanDays 45
Invoke-PnPVersionTrim -LibraryCsvPath .\Trim-Config\Libraries.csv -OlderThanDays 45 -Delete
```

## Exceptions-only logging
By default, only exceptions/skips are logged:

- `-ExceptionCsvPath` (CSV)
- `-TextLogPath` (text)

Example:
```powershell
Invoke-PnPVersionTrim -AllLibraries -OlderThanDays 60 -Delete `
  -ExceptionCsvPath .\Logs\Exceptions.csv `
  -TextLogPath .\Logs\Run.log
```

## Optional size reporting (can be slow)
```powershell
Invoke-PnPVersionTrim -AllLibraries -OlderThanDays 60 -Delete -MeasureLibrarySizes
```

## Skip files by name tokens
`SkipNameContains.csv`
```csv
Token
Budget
Quarterly
Checklist
2024
```

```powershell
Invoke-PnPVersionTrim -AllLibraries -OlderThanDays 60 -Delete -SkipNameContainsCsvPath .\Trim-Config\SkipNameContains.csv
```

## Multi-site orchestration
1) Create `Sites.csv` (template via `New-TrimTemplates.ps1`)
2) Run orchestrator:

```powershell
.\SPO-Trim-Orchestrator.ps1 `
  -SitesCsvPath .\Trim-Config\Sites.csv `
  -TrimScriptPath .\PNP-TrimVersions.ps1
```

### Multi-site delete guardrail (pilot key)
Multi-site delete requires:
```powershell
.\SPO-Trim-Orchestrator.ps1 `
  -SitesCsvPath .\Trim-Config\Sites.csv `
  -TrimScriptPath .\PNP-TrimVersions.ps1 `
  -Delete `
  -PilotKey I_RAN_A_PILOT
```

## Notes on safety / crash behavior
- Deletes are done via CSOM `File.Versions` and `DeleteObject()` calls, committed with `ExecuteQuery()`.
- If the tool stops mid-run, you can rerun it. It is not a "transaction" and does not corrupt libraries.
- SharePoint will block deletes under retention/holds/records. Those show up in the Exceptions CSV.

