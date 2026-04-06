# Pharmacy Tools Add-In

Excel add-in for pharmacy billing analysis. Queries the CMS Medicaid NADAC (National Average Drug Acquisition Cost) API and populates per-unit pricing and effective dates directly into your spreadsheet — row by row, matched to the correct annual dataset by dispense date.

## Features

- **NADAC Lookup** — Batch-queries the CMS Medicaid NADAC API for each NDC in your spreadsheet. Matches each row to the correct annual dataset based on the dispense date year. Populates `NADAC` and `NADAC Effective Date` columns automatically.
- **Manage Datasets** — View, add, edit, and override annual NADAC dataset UUIDs. Built-in years (2013–2025) are pre-configured. Custom years can be added without reinstalling.
- **Smart column detection** — Auto-detects NDC and date columns by header name. Falls back to a prompt if not found.
- **Re-run safe** — Row highlights reset on each run; red (salmon) highlights reflect only the current run's unmatched NDCs.

## Requirements

- Microsoft Excel (Office 2016 or later, Windows)
- PowerShell 5.1+
- Internet access to `data.medicaid.gov`

## Installation

1. Close Excel if it is open.
2. Right-click `Install-NADACAddin.ps1` → **Run with PowerShell**.  
   *(No Administrator privileges required.)*
3. The installer will:
   - Build the `.xlam` add-in file
   - Inject the custom Pharmacy Tools ribbon tab and icon
   - Register the add-in so it loads automatically on Excel start
4. Open Excel. A **Pharmacy Tools** tab will appear in the ribbon.

## Usage

1. Open your pharmacy claim spreadsheet.
2. Click **Pharmacy Tools → NADAC Lookup**.
3. Confirm or select your NDC and dispense date columns.
4. Click **OK**. Progress is shown in the Excel status bar.
5. Results are written to `NADAC` and `NADAC Effective Date` columns.

Rows highlighted in salmon red indicate NDCs not found in the NADAC dataset for that year.

## Dataset Management

Click **Pharmacy Tools → Manage Datasets** to:
- View all built-in and custom dataset entries
- Add a new year (requires Dataset ID and Distribution UUID from `data.medicaid.gov`)
- Edit or override an existing year's UUIDs
- Reset a built-in year to factory defaults
- Delete a custom year

## Log File

A log is written to `%USERPROFILE%\Desktop\NADAC_Lookup_Log.txt` on each run. Rotates at 5 MB.

## Uninstall

1. Excel → **File → Options → Add-ins → Manage: Excel Add-ins → Go**
2. Uncheck **NADACLookup**
3. Delete `%APPDATA%\PharmacyTools\NADACLookup.xlam`

## Roadmap

See [ROADMAP.md](ROADMAP.md) for planned features including per-row status codes, unfound-row segregation, margin calculation, and a full UserForm Dataset Manager.

## Version History

| Version | Date | Notes |
|---------|------|-------|
| 2.3 | 2026-04-05 | Dataset Manager (MsgBox/InputBox UI), re-run colour reset, About screen |
| 2.2 | — | NADAC column, icon, batch POST with GET fallback |
| 2.1 | — | Initial ribbon integration |

## Contact

**Jeremiah Maudlin**  
Omega Integrated Management  
jmaudlin@omegaim.com | +1 (641) 660-6367

---

*Data source: [data.medicaid.gov](https://data.medicaid.gov)*
