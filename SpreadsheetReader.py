# Function to read data from spreadsheet

import os, re
from pprintpp import pprint as pp
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
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


def insertCoveringDateValues(sheet, dateList):
    ''' Insert covering date values into dictionary of spreadsheet values '''
    sheet["Covering Dates"] = dateList
    return sheet


def openingCalculation(age, year):
    ''' return the year the record can be opened (100 years after birth) given an age in a given year'''
    return (year - age) + 100 


def createOpeningList(agesList, yearsList):   
    ''' return a list of years in which the record will be open given a list of ages and list of years'''
    return list(map(openingCalculation, agesList, yearsList))


def sheetRedactionNeededCheck(openingList):
    ''' return False if no opening dates in the given list are later than the current year'''
    return False if max(openingList) <= date.today().year else True

        
def filterByYear(previousRedactionList, currentRedactionList):
    ''' Return filter of changed redaction'''    
    return [True if a == b else False for a, b in zip(previousRedactionList, currentRedactionList)]    


def redactColumns(columnsToRedact, openingList, lastYearInSeries, year=date.today().year):
    ''' given a dictionary containung the columns that may need redacting, return a dict containing the original record values and
    the processed values for each year, by year, until all records have been opened. 
    
    {"base": [{col1Name:col1, col2Name:col2}], 
    year: [{col1Name: col1_redacted, col2Name: col2_redacted], 
    year+1 [{col1Name: col1_redacted, col2Name: col2_redacted]... 
    max_year: [{col1Name: col1_redacted, col2Name: col2_redacted]}
    
    '''
    
    boilerplate = "[Additional information regarding this case will be added to the catalogue when the case becomes over 100 years old. In cases when the date is not known, the latest date in the series (" + str(lastYearInSeries) + ") will be used]"
    
    processedColumns = {"base":columnsToRedact}
    
    yearsToPublish = list(range(year, max(openingList)+1))
    previousRedactions = []
    
    for currentYear in yearsToPublish:
        #print(currentYear)
        
        toRedact = [True if currentYear < openingYear else False for openingYear in openingList]
        test_redactByYear_testFile(toRedact, currentYear)
        #print(toRedact)
        
        filter = [True] * len(openingList)
        
        if previousRedactions == []:
            previousRedactions = toRedact
        else:
            filter = filterByYear(previousRedactions, toRedact)
            previousRedactions = toRedact
        
        test_filterByYear_testFile(filter, currentYear)
        
        processedColumns[currentYear] = {}
    
        for columnName, column in columnsToRedact.items():
            newColumn = [boilerplate if record[1] else record[0] for record in zip(column, toRedact)]   
            processedColumns[currentYear][columnName]=newColumn
    
    return processedColumns


def unredactByYear(filename, values, newValues, year):
    ''' print out a new spreadsheet with the full text for all columns for just the rows where the year is 100 years since birth'''
    
    wb = Workbook()
    newFilename = str(year) + "_" + filename
    pathToFile = os.path.join('data', 'converted' + str(year))
    
    if not os.path.exists(pathToFile):
        os.makedirs(pathToFile)
    
    newFile = os.path.join(pathToFile, newFilename) 
    
    newSheet = wb.active
    
    y = 1
    
    #print(newValues[year].keys())
    
    for title, row in values.items():
        newSheet.cell(1, y, title).font = Font(bold=True)
    
        x = 2
        
        if title in newValues[year].keys():
            row = newValues[year][title]
           

        while x < (len(row) + 2):
            newSheet.cell(x, y, row[x - 2])
            x+=1
            
        y+=1 
    
    wb.save(newFile)


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

def test_filterByYear_testFile(filterList, year):    
    if year == 2022:
        expectedFilterList = [True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True]
        assert expectedFilterList == filterList
    elif year == 2023:
        expectedFilterList = [True,True,True,True,True,True,True,True,True,True,True,True,True,True,False,False,True,True,True,True,True,True,True,True,True,True]
        assert expectedFilterList == filterList
    elif year == 2024:
        expectedFilterList = [True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,False,False,False,True,True,True,True,True,True,True]
        assert expectedFilterList == filterList
    else:
        expectedFilterList = [True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,False,False,False,True,True,True]
        assert expectedFilterList == filterList


### Main        

currentSpreadsheet = getSpreadsheetValues('test.xlsx')
test_loadfile(list(currentSpreadsheet.keys()))

ageList = getAgeFromColumn(currentSpreadsheet['Age'])
test_all_ints(ageList)
test_age_testFile(ageList)

yearList = getYearFromColumn(currentSpreadsheet['Brief summary of grounds for recommendation'])
test_all_ints(yearList)
test_year_testFile(yearList)

insertCoveringDateValues(currentSpreadsheet, yearList)

openingList = createOpeningList(ageList, yearList)
test_all_ints(openingList)
test_openingList_testFile(openingList)

if(sheetRedactionNeededCheck(openingList)):
    newColumns = redactColumns(dict((key, currentSpreadsheet[key]) for key in ['Occupation', 'Brief summary of grounds for recommendation']), openingList, 1945)
    #pp(newColumns)
    
#unredactByYear("test.xlsx", currentSpreadsheet, newColumns, 2022)