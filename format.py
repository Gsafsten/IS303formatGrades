"""
Garrett Safsten, Jack Mair, Ryan Baldwin, Tanner Crookston
Description:
"""
# This pulls in the librries we will need.
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

myWorkbook = openpyxl.load_workbook('Poorly_Organized_Data_1')

newWorkbook = Workbook("Organized")



