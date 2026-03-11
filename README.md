TeleStaff Export Automation for CRVTS

This script replaces the old manual Excel cleanup process for the CRVTS TeleStaff export.

It takes two TeleStaff exports, joins them together, builds the fields CRVTS needs, and writes out a finished TS_EXP.xlsx file ready to drop into SharePoint for the CRVTS Power Query.

Credit to Paul Clark for getting the original spreadsheet process started.


The script uses tkinter for the file picker, xml.etree.ElementTree for reading the TeleStaff XML export, pandas for joins and transforms, and openpyxl for writing the final workbook.
