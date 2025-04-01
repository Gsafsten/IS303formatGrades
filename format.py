"""
Garrett Safsten, Jack Mair, Ryan Baldwin, Tanner Crookston
Description:
"""
# This pulls in the libraries we will need.
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

myWorkbook = openpyxl.load_workbook("Poorly_Organized_Data_1.xlsx")
newWorkbook = Workbook()

currentSheet = myWorkbook.active

# Iterate through rows in the original data sheet
for row in currentSheet.iter_rows(min_row=2, values_only=True):
    value = row[0]  # Column A value

    # Check if the sheet already exists, if not, create one
    if value not in newWorkbook.sheetnames:
        newWorkbook.create_sheet(title=str(value))
    if 'Sheet' in newWorkbook.sheetnames:
        del newWorkbook['Sheet']


