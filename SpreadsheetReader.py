# Function to read data from spreadsheet

import os
from openpyxl import load_workbook

def getSpreadsheet(filename):
    path = os.path.join('data', filename) 
    wb = load_workbook(filename = path)
    
    sheet = wb.worksheets[0]
    column_list = [cell.value for cell in sheet[1]]
    return column_list
        
        
        

print(getSpreadsheet('T 336_002.xlsx'))