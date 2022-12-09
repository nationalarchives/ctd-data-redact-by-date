# Function to read data from spreadsheet

import os
from pprintpp import pprint as pp
from openpyxl import load_workbook

def getSpreadsheetValues(filename):
    ''' Gets spreadsheet by name and returns the spreadsheet as a worksheet and a list of column headings '''
    path = os.path.join('data', filename) 
    wb = load_workbook(filename = path)
    
    sheet = wb.worksheets[0]
    values={}
    
    for col in sheet.columns:
        column = [cell.value for cell in col if cell.value is not None]
        
        if len(column) > 0:
            values[column[0]] = column[1:]
    
    return (values)
        


def getAgeFromColumn(sheet, columnname):
    ''' Get age from named column, if no age given then assume age is 18, and return a list of ages '''
    age = 18
    # if value is number then age otherwise default value
    
    return ageList

def getYearFromColumn(sheet, columnname):
    ''' Get year from named column and return a list of years'''
    # regex for dddd in text value
    
    pass

def createCoveringDateField(sheet, dateList):
    ''' print out a new spreadsheet with an extra column listing the covering date as specified by the dateList '''
    # openpyxl.worksheet.worksheet.Worksheet.insert_cols()
    pass

def redactColumns(sheet, columnList, year):
    ''' print out a new spreadsheet with the values in the specified columns replaced with the redaction text if year is not over 100 years since birth '''
    pass

def unredactColumns(sheet, year):
    ''' print out a new spreadsheet with the full text for all columns for just the rows where the year is 100 years since birth'''
    pass


### Tests ####


def test_loadfile(column_headings):
    expected_columns = ['Letter','Series','Piece', 'Item', 'Treasury Case number', 'Home Office case number', 'First names/Initials', 'Surname', 'Age', 'Occupation', 'Award granted', 'Brief summary of grounds for recommendation']
    
    assert column_headings == expected_columns  
    
    
def test_age():
    pass

def test_year():
    pass   

def test_dateList():
    pass

def test_addDatecolumn():
    pass   

def test_redacted():
    pass
        

test_loadfile(list(getSpreadsheetValues('T 336_002.xlsx').keys()))