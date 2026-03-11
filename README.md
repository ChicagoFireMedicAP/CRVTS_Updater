eleStaff Export Automation for CRVTS

This script replaces the old manual Excel cleanup process for the CRVTS TeleStaff export. It takes two exports from TeleStaff, joins them together, builds the fields CRVTS needs, and writes out a finished TS_EXP.xlsx file ready to drop into SharePoint for the CRVTS Power Query. Credit to Paul Clark for getting the original spreadsheet process started.

Overview

The script takes an Assignment Report and a People export from TeleStaff, reads both files, matches them together, derives a few missing fields, and writes the finished workbook to your Downloads folder. The goal is to replace the old multi step Excel cleanup with one run of the script.

Required files

You need two exports from TeleStaff: the Assignment Report, which is the main staffing file and includes the current assignment rows the script builds from, and the People export, which provides the extra employee data used to fill in Promoted and IDPH Status; the script matches File from the Assignment Report to Payroll ID from the People export, and while the Assignment Report usually comes out as an .xls file that is really XML underneath, the People export can be .csv, .xlsx, or .xls.

Requirements

The script requires Python 3 along with pandas and openpyxl.

Install
pip install pandas openpyxl
Run
python ts_export.py

When it runs, it will first ask you to choose the TeleStaff Assignment Report and then the TeleStaff People export. If you skip the People export, the script will still finish, but Promoted will stay blank and IDPH Status will default to NONE.

What the script builds

The Assignment Report is the base file. From that, the script reads the assignment rows and keeps the fields CRVTS needs, including things like institution, region, station, unit, person, employee ID, file, shift, Daley, from, and rank. It then derives Name by stripping the parenthetical assignment code out of the Person field, derives TS Assignment from the value inside those parentheses, and builds PLT by using Daley if it exists, otherwise checking Shift for an EMS platoon value, and defaulting to 5 if neither applies.

The People export is used to enrich the base data. It provides the Promoted value and the information used to determine IDPH Status. The script checks the People file for Specialty first and falls back to Name if needed. The possible IDPH results are PAR, EMT, EMT (PM Drop), and NONE.

Output

The finished output file is TS_EXP.xlsx, saved to your Downloads folder. The workbook contains four sheets. TS Assign holds the base assignment data plus the derived columns. TS EXP is the main output sheet read by the CRVTS Power Query. TS Promoted keeps the full People export for reference. IDPH Lic is a simple two column reference sheet containing file or payroll ID and IDPH status, with no headers.

Validation

At the end of the run, the script prints a summary in the terminal so you can sanity check the output before uploading it. That summary includes total row count, IDPH status counts, PLT counts, rank fill counts, promoted fill counts, and a few sample rows. That makes it easier to catch a bad export, shifted columns, or a mismatched People file before the workbook gets uploaded.

Typical workflow

The normal process is to export the Assignment Report from TeleStaff, export the People file from TeleStaff, run ts_export.py, select both files when prompted, let the script build TS_EXP.xlsx, upload that file to SharePoint, and then refresh the CRVTS Power Query.

Notes

Set institution to none. The script assumes the Assignment Report contains a header row beginning with Institution and Region, and it assumes that File from the Assignment Report matches Payroll ID from the People export. If something breaks, the first things to check are whether the wrong TeleStaff export was selected, whether the header row changed, whether the People export is missing Payroll ID, whether the Assignment Report columns shifted, or whether TeleStaff changed the export format.

Libraries used

The script uses tkinter for the file picker, xml.etree.ElementTree for reading the TeleStaff XML export, pandas for joins and transformations, and openpyxl for writing the final workbook.
