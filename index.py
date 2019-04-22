import xlwings as xw
from os import walk

years = [
        {'range': '1980-2000', 'idNumber': '19802000'},
        {'range': '2000-2015', 'idNumber': '20002015'}
        ]
results = xw.Book('Table.xlsx')

def getValue64(file, filename, idNumber):
    sheet = file.sheets[f'{filename}-{idNumber}']
    for col in sheet.range('N1:Z1'):
        if 'VALUE_64' in col.value:
            val64col = col.address[1]
            break
    for row in sheet.range('B10:B30'):
        if 64 == row.value:
            val64row = row.address[3::]
            break
    return sheet.range(f'{val64col}{val64row}').value

def inputData(year):
    currentRow = 3
    yearRange = year['range']
    idNumber = year['idNumber']
    for (dirpath, dirnames, filenames) in walk(f'./{yearRange}'):
        currentYear = results.sheets[f'{yearRange}']
        for dir in dirnames:
            print(dir)
            currentFile = xw.Book(f'{yearRange}/{dir}/{dir}-{idNumber}.xlsx')
            currentYear.range(f'R{currentRow}').value = getValue64(currentFile, dir, idNumber)
            currentRow += 1
            currentFile.close()

for year in years:
    inputData(year)
    
results.save()