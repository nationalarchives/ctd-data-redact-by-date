# Function to read data from spreadsheet

import os
from openpyxl import load_workbook

def getSpreadsheet(filename):
    path = os.path.join('data', filename) 
    wb = load_workbook(filename = path)
    
    sheet = wb.worksheets[0]
    column_list = [cell.value for cell in sheet[1] if cell.value is not None]
    return column_list
        


def test_loadfile(column_headings):
    expected_columns = ['Letter','Series','Piece', 'Item', 'Treasury Case number', 'Home Office case number', 'First names/Initials', 'Surname', 'Age', 'Occupation', 'Award granted', 'Brief summary of grounds for recommendation'];
    
    assert column_headings == expected_columns;        
        

test_loadfile(getSpreadsheet('T 336_002.xlsx'))