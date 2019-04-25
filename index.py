import xlwings as xw
from os import walk
from dataRow import DataRow
import re

years = [
        {'range': '1980-2000', 'idNumber': '19802000'},
        {'range': '2000-2015', 'idNumber': '20002015'}
        ]
results = xw.Book('Table.xlsx')

def getRow64(datasheet):
    rows = datasheet.range('B2:B30')       
    for row in rows:
        if 64 == row.value:
            val64Row = row.address[3::]
            break
    return val64Row

def getCol64(datasheet):
    for col in datasheet.range('L1:Z1'):
        if 'VALUE_64' in col.value:
            val64Col = col.address[1]
            break
    return val64Col

def mapRow64(datasheet, val64Row):
    currentResults = DataRow()
    val64ToN = datasheet.range(f'C{val64Row}:Z{val64Row}')
    for n in val64ToN:
        if  datasheet.range(f'{n.address[1]}1').value is not None:
            headerCol = n.address[1]
            header = float((datasheet.range(f'{headerCol}1').value)[6::])
            currentResults.data[header] = n.value
    return currentResults

def mapCol64(datasheet, val64Col):
    currentResults = DataRow()
    nToVal64 = datasheet.range(f'{val64Col}2:{val64Col}26')
    for n in nToVal64:
        headerRow = n.address[2::]
        if datasheet.range(f'B{headerRow}').value is not None:
            header = float((datasheet.range(f'B{headerRow}').value))
            currentResults.data[header] = n.value
    return currentResults

def setResultsToTarget(resultsSheet, target, resultData):
    for col in target:
        headerRow = re.sub('[$]', '', col.address[1:3:])
        header = resultsSheet.range(f'{headerRow}2').value
        if (resultData.data[header] is not None):
            col.value = resultData.data[header]
        else:
            col.value = 'null'

def row64Helper(currentSheet, resultsSheet, resultsRow):
    val64Row = getRow64(currentSheet)
    currentResults = mapRow64(currentSheet, val64Row)
    target = resultsSheet.range(f'B{resultsRow}:AE{resultsRow}')
    setResultsToTarget(resultsSheet, target, currentResults)

def col64Helper(currentSheet, resultsSheet, resultsRow):
    val64Col = getCol64(currentSheet)
    currentResults = mapCol64(currentSheet, val64Col)
    target = resultsSheet.range(f'AG{resultsRow}:BJ{resultsRow}')
    setResultsToTarget(resultsSheet, target, currentResults)

def inputData(year):
    currentRow = 3
    yearRange = year['range']
    idNumber = year['idNumber']
    for (dirpath, dirnames, filenames) in walk(f'./{yearRange}'):
        resultsSheet = results.sheets[f'{yearRange}']
        for dir in dirnames:
            print(dir)
            currentFile = xw.Book(f'{yearRange}/{dir}/{dir}-{idNumber}.xlsx')
            currentSheet = currentFile.sheets[f'{dir}-{idNumber}']
            row64Helper(currentSheet, resultsSheet, currentRow)
            col64Helper(currentSheet, resultsSheet, currentRow)
            currentRow += 1
            currentFile.close()

for year in years:
    inputData(year)
results.save()