import xlrd
import datetime
import subprocess as sp
#screen clearing
clear = sp.call('clear',shell=True)
#Error Functions
def noInputError():
    print('Input was required but none was given....exiting')
    exit()
#Information
print('### The name of the excel document will be used as the data base\'s name ###')
print('### Sheet names will be used as table names ###\n')
#variables
j = 0
document = input('Excel document\'s name:  ') or noInputError()
defaultExtension = input('File extension (default: "xlsx"):  ') or 'xlsx'
excelDoc = xlrd.open_workbook(document + '.' + defaultExtension)
sheet = excelDoc.sheet_by_index(j)
sheets = excelDoc.sheet_names()
sql = document+'.sql'
sqlFile = open(sql,'w+')
#Document Quarrying variables
rows = sheet.nrows
columns = sheet.ncols
#Sql writing function
def main():
    sheetsNumber = len(sheets)
    j = 0
    while j <= sheetsNumber - 1:
        tableCreation(sheets[j])
        j += 1
#table making
def tableCreation(name):
    x = 0
    keypair = {}
    sheet = excelDoc.sheet_by_name(name)
    columns = sheet.ncols
    rows = sheet.nrows
    names = '(id'
    keyChain = '(id INT NOT NULL PRIMARY KEY AUTO_INCREMENT'
    while x <= columns -1:
        vcell = str(sheet.cell_value(0,x))
        vcell = vcell.replace(" ", "_")
        names = names+','+' '+vcell
        clear = sp.call('clear',shell=True)
        key = input(f'[1]String\n[2]Numeric Data\n[3]Date and/or Time\n\nHow would you like {vcell} to be formatted(1,2 or 3):    ')
        if key == '1':
            keyChain = keyChain+', '+vcell+' CHAR(255)'
            keypair[x] = 'CHAR'
        elif key == '2':
            clear = sp.call('clear',shell=True)
            number = input(f'[1]INT\n[2]Float\n\nWould you like {vcell} to be formatted(1 or 2):    ')
            if number == '1':
                keyChain = keyChain+', '+vcell+' INT'
                keypair[x] = 'INT'
            elif number == '2':
                keyChain = keyChain+', '+vcell+' FLOAT'
                keypair[x] = 'FLOAT'
            else:
                noInputError()
        elif key == '3':
            clear = sp.call('clear',shell=True)
            date = input(f'[1]Year(4 digit)\n[2]Year(2 digit)\n[3]Time and Data\n[4]Time(HH:MM:SS)\n[5]Date(YYY-MM-DD)\n\nHow would you like the {vcell} to be formatted?(1,2,3,4 or 5):    ')
            if date == '1':
                keyChain = keyChain+', '+vcell+' YEAR(4)'
                keypair[x] = 'YEAR4'
            elif date == '2':
                keyChain = keyChain+', '+vcell+' YEAR(2)'
                keypair[x] = 'YEAR2'
            elif date == '3':
                keyChain = keyChain+', '+vcell+' DATETIME'
                keypair[x] = 'DATETIME'
            elif date == '4':
                keyChain = keyChain+', '+vcell+' TIME'
                keypair[x] = 'TIME'
            elif date  == '5':
                keyChain = keyChain+', '+vcell+' DATE'
                keypair[x] = 'DATE'
            else:
                noInputError()
        else:
            noInputError()
        x += 1
    sqlFile.write(f'CREATE TABLE {name}'+keyChain+');\n')
    names = names+')'+' '+'VALUES'+' '+'(NULL'
    names = 'INSERT INTO'+' '+name+' '+names
    read(sheet, names, keypair, rows, columns)
#Read table function
def read(sheet, names, keypair, rows, columns):
    x = 1
    y = 0
    namereset = names
    while x <= rows -1:
        cellv = sheet.cell_value(x,y)
        if keypair[y] == 'CHAR':
            names = names+','+f' \'{cellv}\''
        elif keypair[y] == 'INT':
            ivalue = str(int(cellv))
            names = names+','+f' \'{ivalue}\''
        elif keypair[y] == 'FLOAT':
            fvalue = str(cellv)
            names = names+','+f' \'{fvalue}\''
        elif keypair[y] == 'YEAR4':
            names = names+','+f' \'{cellv}\''
        elif keypair[y] == 'YEAR2':
            if len(cellv) == 2:
                names = names+','+f' \'{cellv}\''
            elif len(cellv) == 4:
                cellv = cellv[2]+cellv[3]
                names = names+','+f' \'{cellv}\''
            else:
                noInputError()
        elif keypair[y] == 'DATETIME':
            ms_date_number = sheet.cell_value(x,y)
            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ms_date_number,excelDoc.datemode)
            py_date = datetime.datetime(year, month, day, hour, minute, second)
            py_date = str(py_date)
            names = names+','+f' \'{py_date}\''
        elif keypair[y] == 'TIME':
            ms_date_number = sheet.cell_value(x,y)
            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ms_date_number,excelDoc.datemode)
            py_date = datetime.datetime(year, month, day, hour, minute, second)
            py_date = str(py_date)
            py_date = py_date[10:]
            names = names+','+f' \'{py_date}\''
        elif keypair[y] == 'DATE':
            ms_date_number = sheet.cell_value(x,y)
            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ms_date_number,excelDoc.datemode)
            py_date = datetime.datetime(year, month, day, hour, minute, second)
            py_date = str(py_date)
            py_date = py_date[:10]
            names = names+','+f' \'{py_date}\''
        if y == columns - 1:
            names = names+');'
            sqlFile.write(f'{names}\n')
            names = namereset
            y = -1
            x += 1
        y += 1
#FileWriting
document = document.replace(" ", "_")
sqlFile.write(f'CREATE DATABASE {document};\n')
sqlFile.write(f'USE {document};\n')
#run
main()
sqlFile.close()
clear = sp.call('clear',shell=True)
