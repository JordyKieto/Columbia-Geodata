import xlwings as xw
from os import walk

years = ['1980-2000', '2000-2015']
results = xw.Book('Table.xlsx')

def getValue64(file, filename):
    file = file.sheets[f'{filename}-19802000']
    val64col = ""
    val64row = ""
    cols = file.range('A1:Z1')
    rows = file.range('B2:B30')       
    for col in cols:
        if 'VALUE_64' in col.value:
            val64col = col.address[1]
            break
    for row in rows:
        if 64 == row.value:
            val64row = row.address[3::]
            break
    return file.range(f'{val64col}{val64row}').value

def inputData(year):
    currentRow = 3
    for (dirpath, dirnames, filenames) in walk(f'./{year}'):
        currentYear = results.sheets[f'{year}']
        for dir in dirnames:
            print(dir)
            currentFile = xw.Book(f'{year}/{dir}/{dir}-19802000.xlsx')
            currentYear.range(f'R{currentRow}').value = getValue64(currentFile, dir)
            currentRow += 1
            currentFile.close()

for year in years:
    inputData(year)
    
results.save()