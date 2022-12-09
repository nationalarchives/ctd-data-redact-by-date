# Function to read data from spreadsheet

import os, re
from pprintpp import pprint as pp
from openpyxl import load_workbook
from datetime import date

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
       


def getAgeFromColumn(column):
    ''' Get age from named column, if no age given then assume age is 18, and return a list of ages '''
    # if value is number then age otherwise default value   
    return [entry if str(entry).strip().isnumeric() else 18 for entry in column]         


def getYearFromColumn(column):
    ''' Get year from named column and return a list of years'''
    # regex for dddd in text value
    # since Python 3.8 := allows you to name an evaluation as a variable which you can use int he list comparhension see https://stackoverflow.com/questions/26672532/how-to-set-local-variable-in-list-comprehension
    years = [int(years[0]) if len((years := re.findall(r'\d{4}', entry))) == 1 else years for entry in column] 
    #pp(years)
    return years


def createCoveringDateField(sheet, dateList):
    ''' print out a new spreadsheet with an extra column listing the covering date as specified by the dateList '''
    # openpyxl.worksheet.worksheet.Worksheet.insert_cols()
    pass


def openingCalculation(age, year):
    return (year - age) + 100 


def createOpeningList(ages, years):   
    return list(map(openingCalculation, ages, years))


def sheetRedactionNeededCheck(openingList):
    return False if max(openingList) <= date.today().year else True
    

def redactColumns(columnsToRedact, openingList, lastYearInSeries, year=date.today().year):
    ''' 
    col1 & col2 -> {"base": [col1, col2], year: [col1_redacted, col2_redacted], year+1 [col1_opening, col2_opening]... max_year: [col1_openning, col2_opening]}
    '''
    boilerplate = "[Additional information regarding this case will be added to the catalogue when the case becomes over 100 years old. In cases when the date is not known, the latest date in the series (" + str(lastYearInSeries) + ") will be used]"
    
    processedColumns = {"base":columnsToRedact}
    
    yearsToPublish = list(range(year, max(openingList)+1))
    
    redactedColumnsByYear = {}
    
    for currentYear in yearsToPublish:
        print(currentYear)
        
        toRedact = [True if currentYear < openingYear else False for openingYear in openingList]
        test_redactByYear_testFile(toRedact, currentYear)
        redactedColumnsByYear[currentYear] = {}
    
        for columnName, column in columnsToRedact.items():
            newColumn = [boilerplate if record[1] else record[0] for record in zip(column, toRedact)]   
            redactedColumnsByYear[currentYear][columnName]=newColumn
            
        


def unredactColumns(sheet, year):
    ''' print out a new spreadsheet with the full text for all columns for just the rows where the year is 100 years since birth'''
    pass


### Tests ####


def test_loadfile(columnHeadings):
    expectedColumns = ['Letter','Series','Piece', 'Item', 'Treasury Case number', 'Home Office case number', 'First names/Initials', 'Surname', 'Age', 'Occupation', 'Award granted', 'Brief summary of grounds for recommendation']   
    assert columnHeadings == expectedColumns  
    
    
def test_all_ints(list):
    assert all(isinstance(x, int) for x in list)
    
def test_age_testFile(ageList):
    expectedAges = [16,18,18,20,25,45,16,18,20,16,18,20,16,18,18,20,16,18,20,36,16,18,20,30,25,23]   
    assert expectedAges == ageList

def test_year_testFile(yearList):
    expectedYears = [1936,1938,1938,1940,1945,1945,1937,1939,1941,1938,1940,1942,1938,1940,1941,1943,1940,1942,1944,1940,1941,1943,1945,1944,1943,1942]    
    assert expectedYears == yearList

def test_openingList_testFile(openingList):
    expectedOpening = [2020,2020,2020,2020,2020,2000,2021,2021,2021,2022,2022,2022,2022,2022,2023,2023,2024,2024,2024,2004,2025,2025,2025,2014,2018,2019]    
    assert expectedOpening == openingList 

def test_redactByYear_testFile(toRedactList, year):    
    if year == 2022:
        expectedRedactionList = [False,False,False,False,False,False,False,False,False,False,False,False,False,False,True,True,True,True,True,False,True,True,True,False,False,False]
        assert expectedRedactionList == toRedactList
    elif year == 2023:
        expectedRedactionList = [False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,True,True,True,False,True,True,True,False,False,False]
        assert expectedRedactionList == toRedactList
    elif year == 2024:
        expectedRedactionList = [False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,True,True,True,False,False,False]
        assert expectedRedactionList == toRedactList
    else:
        expectedRedactionList = [False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False]
        assert expectedRedactionList == toRedactList


### Main        

currentSpreadsheet = getSpreadsheetValues('test.xlsx')
test_loadfile(list(currentSpreadsheet.keys()))

ageList = getAgeFromColumn(currentSpreadsheet['Age'])
test_all_ints(ageList)
test_age_testFile(ageList)

yearList = getYearFromColumn(currentSpreadsheet['Brief summary of grounds for recommendation'])
test_all_ints(yearList)
test_year_testFile(yearList)

openingList = createOpeningList(ageList, yearList)
test_all_ints(openingList)
test_openingList_testFile(openingList)

if(sheetRedactionNeededCheck(openingList)):
    newColumns = redactColumns(dict((key, currentSpreadsheet[key]) for key in ['Occupation', 'Brief summary of grounds for recommendation']), openingList, 1945)
    print(newColumns)