# Function to read data from spreadsheet

import os, re, shutil
from pprintpp import pprint as pp
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from datetime import date
from pathlib import Path

def getSpreadsheetValues(filename):
    ''' Gets spreadsheet by name and returns the spreadsheet as a worksheet and a list of column headings '''
    #path = os.path.join('data', filename) 
    wb = load_workbook(filename)
    
    sheet = wb.worksheets[0]
    values={}
    
    for col in sheet.columns:
        #column = [cell.value for cell in col if cell.value is not None]
        column = [cell.value if cell.value is not None else "" for cell in col]
        
        if len(column) > 0 and column.count("") != len(column):
            values[str(column[0]).strip()] = column[1:]
            
    return (values)
       
def removeBlanksFromColumn(column):
    return [value for value in column if value != ""]  

def getAgeFromColumn(column):
    ''' Get age from named column, if no age given then assume age is 18, and return a list of ages '''
    # if value is number then age otherwise default value      
    return [entry if str(entry).strip().isnumeric() else 18 for entry in removeBlanksFromColumn(column)]         


def getYearFromColumn(column):
    ''' Get year from named column and return a dictionary of years and covering dates'''
    # regex for dddd in text value
    # since Python 3.8 := allows you to name an evaluation as a variable which you can use int he list comparhension see https://stackoverflow.com/questions/26672532/how-to-set-local-variable-in-list-comprehension 
    years = [int(years[0]) if len(years := re.findall(r'\d{4}', entry)) == 1 else years for entry in removeBlanksFromColumn(column)] 
    #pp(years)
       
    return codifyYears(years)

def getDateFromList(dateList, earliest, latest, default, max=True):
    foundDate = -1
 
    for date in dateList:
        date = int(date)
        if (foundDate == -1 and foundDate > earliest and foundDate < latest) or (max and date < latest and date > foundDate) or (not(max) and date > earliest and date < foundDate):
            foundDate = date
            
    if foundDate == -1:
        foundDate = default
        
    return foundDate


def codifyYears(yearsList):
    defaultYear = 1946
    earliest = 1935
    latest = 1946
    codifiedYears = []
    coveringDates = []
       
    for year in yearsList:
        if type(year) is not int:
            if len(year) > 1:
                latestDate = "" 
                
                if int(max(year)) > earliest and int(max(year)) < latest:
                    latestDate = int(max(year))
                else:
                    latestDate = getDateFromList(year, earliest, latest, defaultYear)
                
                if int(min(year)) > earliest and int(min(year)) < latest:
                    earliestDate = int(min(year))
                else:
                    earliestDate = getDateFromList(year, earliest, latest, defaultYear, False)
                
                codifiedYears.append(latestDate)
                coveringDates.append(str(earliestDate) + " - " + str(latestDate))
            else:
                codifiedYears.append(defaultYear)
                coveringDates.append(defaultYear)
        elif year < earliest or year > latest:
            codifiedYears.append(defaultYear)
            coveringDates.append(defaultYear)           
        else:
            codifiedYears.append(year)
            coveringDates.append(year)
    
    #print(codifiedYears)
    #print(coveringDates)
    return zip(codifiedYears, coveringDates)
             

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

        
def selectByYear(previousRedactionList, currentRedactionList):
    ''' Return filter of changed redaction'''    
    return [False if a == b else True for a, b in zip(previousRedactionList, currentRedactionList)]    

def yearsToPublish(openingList, year=date.today().year):
    return list(range(year, max(openingList)+1))

def redactColumns(columnsToRedact, openingList, lastYearInSeries, year=date.today().year, minimum=True):
    ''' given a dictionary containung the columns that may need redacting, return a dict containing the original record values and
    the processed values for each year, by year, until all records have been opened. 
    
    {"base": [{col1Name:col1, col2Name:col2}], 
    year: [{filter: filter, col1Name: col1_redacted, col2Name: col2_redacted], 
    year+1 [{filter: filter, col1Name: col1_redacted, col2Name: col2_redacted]... 
    max_year: [{filter: filter, col1Name: col1_redacted, col2Name: col2_redacted]}
    
    '''
    
    boilerplate = "[Additional information regarding this case will be added to the catalogue when the case becomes over 100 years old. In cases when the date is not known, the latest date in the series (" + str(lastYearInSeries) + ") will be used]"
    
    processedColumns = {"base":columnsToRedact}
    
    yearsToPublishList = yearsToPublish(openingList)
    previousRedactions = []
    
    for currentYear in yearsToPublishList:
        #print(currentYear)
        
        toRedact = [True if currentYear < openingYear else False for openingYear in openingList]
        #test_redactByYear_testFile(toRedact, currentYear)
        #print(toRedact)
        
        filter = [True] * len(openingList)
        
        if previousRedactions == []:
            previousRedactions = toRedact
        else:
            filter = selectByYear(previousRedactions, toRedact)
            previousRedactions = toRedact
        
        #test_selectByYear_testFile(filter, currentYear)
        
        processedColumns[currentYear] = {"filter": filter}
    
        for columnName, column in columnsToRedact.items():
            newColumn = [boilerplate if record[1] and record[0] != "" else record[0] for record in zip(column, toRedact)]   
            processedColumns[currentYear][columnName]=newColumn
    
    return processedColumns

def pathToFile(year):
    path = os.path.join('data', 'converted',str(year))
    if not os.path.exists(path):
        os.makedirs(path)
    return path

def unredactByYear(filename, values, newValues, year, min=True):
    ''' print out a new spreadsheet with the full text for all columns for just the rows where the year is 100 years since birth'''
    
    wb = Workbook()
    newSheet = wb.active
    
    col = 1
    
    #print(values)
    
    #print(newValues[year].keys())
    
    filter = newValues[year]["filter"]
    
    for title, column in values.items():
        newSheet.cell(1, col, title).font = Font(bold=True)
    
        row = 2
        
        if title in newValues[year].keys():
            column = newValues[year][title]

        filteredColumn = zip(column, filter)           

        for filteredRow in filteredColumn:
            #print(str(row[1]) + ": " + str(x) + ", " + str(y))
            if (min and filteredRow[1]) or not min:
                newSheet.cell(row, col, filteredRow[0])
                row+=1
            
        col+=1 
    
    #print("Last cel written for " + str(year) + " x:" + str(row) + ", y: " + str(col))
    if row > 2:   
        path = pathToFile(year)  
        newFilename = os.path.splitext(os.path.basename(filename))[0] + "_" + str(year) + os.path.splitext(os.path.basename(filename))[1]
        newFile = os.path.join(path, newFilename)  
        wb.save(newFile)


def generateSpreadsheets(filename, values, newValues, openingList):
    yearsToPublishList = yearsToPublish(openingList)
    
    for currentYear in yearsToPublishList:
        unredactByYear(filename, values, newValues, currentYear)

def generateSummary(filename, ageList, coveringDatesList, openingList):
    if os.path.exists(os.path.join('data', 'converted','summary.xlsx')):
        wb = load_workbook(os.path.join('data', 'converted', 'summary.xlsx'))
        ws = wb.create_sheet()
    else:
        wb = Workbook()        
        ws = wb.active
        
    ws.title = filename
    colHeadings = ["Item", "Age", "Covering Dates", "Opening Year"]
    
    col = 1
    count = 1
    
    for heading in colHeadings:
        row = 2
        ws.cell(1, col, heading).font = Font(bold=True)
        
        for age, coveringDate, opening in zip(ageList, coveringDatesList, openingList):
            ws.cell(row, 1, row-1)
            ws.cell(row, 2, age)
            ws.cell(row, 3, coveringDate)
            ws.cell(row, 4, opening)
            
            row += 1
            
        col += 1
            
    wb.save(os.path.join('data', 'converted', 'summary.xlsx'))
            
    
    
        
    
        
def getFileList(myDir):
    return [file for file in myDir.glob("[!~.]*.xlsx")]

def generateFiles(reset=True,output=True,summary=True):
    
    if reset:
        shutil.rmtree(os.path.join('data', 'converted'))
        
    if summary:
        if os.path.exists(os.path.join('data', 'summary.xlsx')):
            os.remove(os.path.join('data', 'summary.xlsx'))
    
    for file in getFileList(Path('data')):
        print("Processing " + os.path.basename(file))
        currentSpreadsheet = getSpreadsheetValues(file)
        #print(currentSpreadsheet.keys())
        
        try:        
            #test_loadfile(list(currentSpreadsheet.keys()))
        
            ageList = getAgeFromColumn(currentSpreadsheet['Age'])
            test_all_ints(ageList)
        
            dates = getYearFromColumn(currentSpreadsheet['Brief summary of grounds for recommendation'])        
            
            yearList = []
            coveringDatesList = []
            for parts in dates:
                yearList.append(parts[0])
                coveringDatesList.append(parts[1])
            
            insertCoveringDateValues(currentSpreadsheet, coveringDatesList)
            print("Adding covering dates for " + os.path.basename(file))
            
            openingList = createOpeningList(ageList, yearList)
            test_all_ints(openingList)
            
        except AssertionError as e:
            print("Issue with " + os.path.basename(file) + "skipping")
            print(e)
            continue
        
    
        if output:
            if(sheetRedactionNeededCheck(openingList)):
                newColumnValues = redactColumns(dict((key, currentSpreadsheet[key]) for key in ['Occupation', 'Brief summary of grounds for recommendation']), openingList, 1946)
                generateSpreadsheets(os.path.basename(file), currentSpreadsheet, newColumnValues, openingList)
                print(os.path.basename(file) + " redacted. Spreadsheets with redacted descriptions and unredactions generated.")
            else:
                path = pathToFile(date.today().year)
                filename = os.path.splitext(os.path.basename(file))[0] + '_NoRedactions' + os.path.splitext(os.path.basename(file))[1]
                try: 
                    shutil.copyfile(file, os.path.join(path, filename))
                except shutil.SameFileError:
                    pass
                print(os.path.basename(file) + " copied over, no redactions needed")
                
        if summary:   
            generateSummary(os.path.splitext(os.path.basename(file))[0], ageList, coveringDatesList, openingList)   
            

      
        

### Tests ####


def test_loadfile(columnHeadings):
    expectedColumns = ['Letter','Series','Piece', 'Item', 'Treasury Case number', 'Home Office case number', 'First names/Initials', 'Surname', 'Age', 'Occupation', 'Award granted', 'Brief summary of grounds for recommendation']   
    assert columnHeadings == expectedColumns, "Error in expected columns: " + str(columnHeadings)  
       
def test_all_ints(list):
    assert all(isinstance(x, int) for x in list), "Error in expected data types: " + str(list)

'''    
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

def test_selectByYear_testFile(selectList, year):    
    if year == 2022:
        expectedSelectList = [True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True,True]
        assert expectedSelectList == selectList
    elif year == 2023:
        expectedSelectList = [False,False,False,False,False,False,False,False,False,False,False,False,False,False,True,True,False,False,False,False,False,False,False,False,False,False]
        assert expectedSelectList == selectList
    elif year == 2024:
        expectedSelectList = [False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,True,True,True,False,False,False,False,False,False,False]
        assert expectedSelectList == selectList
    else:
        expectedSelectList = [False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,True,True,True,False,False,False]
        assert expectedSelectList == selectList
'''

### Main        

'''
testSpreadsheet = getSpreadsheetValues('test.xlsx')
test_loadfile(list(testSpreadsheet.keys()))

ageList = getAgeFromColumn(testSpreadsheet['Age'])
test_all_ints(ageList)
test_age_testFile(ageList)

yearList = getYearFromColumn(testSpreadsheet['Brief summary of grounds for recommendation'])
test_all_ints(yearList)
test_year_testFile(yearList)

insertCoveringDateValues(testSpreadsheet, yearList)

openingList = createOpeningList(ageList, yearList)
test_all_ints(openingList)
test_openingList_testFile(openingList)

if(sheetRedactionNeededCheck(openingList)):
    newColumns = redactColumns(dict((key, currentSpreadsheet[key]) for key in ['Occupation', 'Brief summary of grounds for recommendation']), openingList, 1945)
    #pp(newColumns)
    
#generateSpreadsheets("test.xlsx", currentSpreadsheet, newColumns, openingList)
'''

generateFiles(False, False, True)