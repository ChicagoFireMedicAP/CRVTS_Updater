TeleStaff Export Automation for CRVTS

This script replaces the old manual Excel cleanup process for the CRVTS TeleStaff export.

It takes two TeleStaff exports, joins them, builds the fields CRVTS needs, and writes out a finished TS_EXP.xlsx file ready to drop into SharePoint for the CRVTS Power Query.

Credit to Paul Clark for getting the original spreadsheet process started.

What it does

Opens a file picker so you can choose the two TeleStaff exports

Reads the Assignment Report export, which is really XML even though it comes out as .xls

Pulls out the rows and columns needed for CRVTS

Derives Name, PLT, and TS Assignment

Loads the People export

Matches People data to the Assignment Report using Payroll ID and File

Fills in Promoted and IDPH Status

Writes a final Excel workbook with four sheets:

TS Assign

TS EXP

TS Promoted

IDPH Lic

Why this exists

The old process was a long manual Excel cleanup just to get the TeleStaff data into a format CRVTS could use.

This script does that in one run and cuts out the repetitive cleanup.

Requirements

Python 3.x

pandas

openpyxl

Install

pip install pandas openpyxl

Run

python ts_export.py

A file picker will pop up.

First pick the TeleStaff Assignment Report.

Then pick the TeleStaff People export.

If you skip the People file, the script still runs, but Promoted will be blank and IDPH Status will default to NONE.

Input files
Assignment Report

Expected export from the TeleStaff Assignment Report.

Important note:

TeleStaff exports this as .xls

It is actually XML underneath

The script looks for a header row starting with Institution and Region, then parses the rows after that.

People export

Expected from the TeleStaff People export.

Accepted formats:

.csv

.xlsx

.xls

This file is used for Promoted and IDPH Status.

Matching logic

The join is:

File from Assignment Report

Payroll ID from People export

For this process, payroll ID and current file number are treated as the same value.

Derived fields
Name

Strips the parenthetical assignment code out of the Person field.

Example: Blow(515), Joe F. → Blow, Joe F.

TS Assignment

Pulls the value inside the parentheses from Person.

Example: Blow(515), Joe F. → 515

PLT

PLT is built in this order:

Use Daley if present

If not, check Shift for EMS Platoon N

If neither applies, default to 5

IDPH Status logic

The script checks the People export for medic license type.

It looks at:

Specialty

Name if needed

Possible outputs:

PAR

EMT

EMT (PM Drop)

NONE

Output workbook

The script writes TS_EXP.xlsx with these sheets:

TS Assign

Base assignment data plus derived columns.

TS EXP

Main output read by the CRVTS Power Query.

TS Promoted

Full People export preserved for reference.

IDPH Lic

Two-column reference sheet with file/payroll ID and IDPH status, with no headers.

Validation

At the end, the script prints:

Total row count

IDPH status counts

PLT counts

Rank fill counts

Promoted fill counts

Sample rows

This helps catch bad exports or mismatched columns before upload.

Typical workflow

Export Assignment Report from TeleStaff

Export People CSV from TeleStaff

Run ts_export.py

Select both files

Let it build TS_EXP.xlsx

Drop that file into SharePoint

Refresh CRVTS Power Query

Important reminder

Set institution to none.

If something breaks

Check these first:

Wrong TeleStaff export selected

Header row changed

People export missing Payroll ID

Assignment Report columns shifted

Bad or empty Person values

TeleStaff export format changed

Dependencies used

tkinter for the file picker

xml.etree.ElementTree for reading the TeleStaff XML export

pandas for joins and transforms

openpyxl for writing the final workbook
