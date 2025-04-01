"""
Garrett Safsten, Jack Mair, Ryan Baldwin, Tanner Crookston
Description:
"""
# This pulls in the libraries we will need.
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

# Load the original workbook
myWorkbook = openpyxl.load_workbook("Poorly_Organized_Data_1.xlsx")
currentSheet = myWorkbook.active

# Create a new workbook
newWorkbook = Workbook()
newWorkbook.remove(newWorkbook.active)  # Remove default sheet

# Dictionary to track the last row for each sheet
row_counters = {}

# Iterate through rows in the original data sheet
for row in currentSheet.iter_rows(min_row=2, values_only=True):
    category = row[0]  # Column A value (determines sheet name)
    full_name = row[1]  # Column B value (contains names)
    grade = row[2]  # Column C value (grade)

    # Ensure the sheet exists
    if category not in newWorkbook.sheetnames:
        newSheet = newWorkbook.create_sheet(title=str(category))
        newSheet.append(["Last Name", "First Name", "Student ID", "Grade"])  # Add headers
        row_counters[category] = 2  # Start row count at 2

    # Split name into parts
    lstOrganizedData = full_name.split("_")

    # Store values in the appropriate sheet
    newSheet = newWorkbook[category]
    newSheet[f"A{row_counters[category]}"] = lstOrganizedData[0]
    newSheet[f"B{row_counters[category]}"] = lstOrganizedData[1]
    newSheet[f"C{row_counters[category]}"] = lstOrganizedData[2]
    newSheet[f"D{row_counters[category]}"] = grade

    # Increment row counter for this sheet
    row_counters[category] += 1
