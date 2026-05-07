# TeleStaff Export Automation for CRVTS

Replaces the 40-step manual Excel process documented in "TeleStaff Export Procedures for the Vacancy Tracking System."

Takes two TeleStaff downloads (Assignment Report + People CSV), joins and transforms them, and outputs `TS EXP.xlsx` ready to drop into SharePoint for the CRVTS Power Query.

## Requirements

```
pip install pandas openpyxl
```

Python 3.8+. No internet connection required at runtime.

## Getting the Source Files from TeleStaff

### Assignment Report

1. Switch **Institution** to **None**
2. Go to **Reports** > **Assignment Report**
3. Download with default options — select **Assignment Report (worksheet)**

### People CSV

1. Switch **Institution** back to your institution
2. Go to **People** and **let it fully load** (exporting before load completes gives you an empty CSV)
3. Click the **gear icon** > **Export to CSV** (`people.csv`)

## Usage

Run the script:

```
python ts_export.py
```

Three file dialogs will appear in sequence:

1. **Select the Assignment Report** — the `.xls` file from TeleStaff (it's actually XML, the script handles this)
2. **Select the People CSV** — the `.csv` export from TeleStaff People. You can cancel this dialog if you don't have it; Promoted and IDPH Status columns will just be empty.
3. **Save As** — choose where to save `TS EXP.xlsx`. Defaults to your Downloads folder with the filename pre-filled.

After saving, a validation summary prints to the console so you can eyeball record counts, IDPH match rates, and sample rows before uploading.

## What It Does

The script parses the Assignment Report XML (Microsoft SpreadsheetML disguised as `.xls`), handles the `MergeAcross=1` quirk on the Institution column, and derives several columns that the manual process required formulas or Power Query steps for:

- **Name** — Person field with the parenthetical unit code stripped out
- **TS Assignment** — the unit code extracted from the parenthetical
- **PLT** — Daley value if present, otherwise EMS platoon number from the shift name, otherwise defaults to 5
- **Promoted** — date pulled from the People CSV via Payroll ID join
- **IDPH Status** — license type (PAR, EMT, EMT (PM Drop), or NONE) derived from the People CSV Specialty field, with a fallback parse of the Name field

## Output Workbook

`TS EXP.xlsx` contains four sheets:

| Sheet | Purpose |
|---|---|
| **TS Assign** | Base data with "Text Between Delimiters" column (replaces the Power Query extraction step) |
| **TS EXP** | The main sheet CRVTS Power Query reads from |
| **TS Promoted** | Full People export preserved for reference |
| **IDPH Lic** | Payroll ID + license type pairs (no headers), for reference |

All data sheets are formatted as Excel tables.

## After Running

Drag `TS EXP.xlsx` to the SharePoint location and refresh the CRVTS Power Query.
