filesName = 'looker-files.txt'
workbooksToSearch = []
print(f'Workbooks to search:')
with open(filesName) as file:
    lines = file.readlines()
    for line in lines:
        line = line.strip('\n')
        line = line.strip('"')
        print(line)
        workbooksToSearch.append(line)

from openpyxl import *
import sys
import os

outputWorkbookName = 'lex-lme.xlsx'

lexHeader = {'name':'LEX (ALL DOMESTIC VOYAGES ONLY) REFUNDS/COMMISSIONS TO BE REIMBURSED BY CHECK/ WIRE (Payments to be processed through AS400)', 'tabTitle':'LEX'}
lmeHeader = {'name':'LME (All FOREIGN VOYAGES ONLY) REFUNDS/COMMISSIONS TO BE REIMBURSED BY CHECK/WIRE (Payments to be processed through AS400)', 'tabTitle':'LME'}

columns = ['BOOKING #', 'NAME', 'AMOUNT', 'COMMENTS', ' AMT PAID ', 'DATE PAID', 'CHK/CC', 'VOID DATE']

sectionsToFind = [lexHeader,lmeHeader]
tabName = 'Sales'

# sampleWorkbookName = 'C:/Users/reece.Holzhauser/Downloads/Payments for Week Ending 120222.xlsx'
# workbooksToSearch = [sampleWorkbookName]

outputWorkbook = Workbook()
for sheet in outputWorkbook.sheetnames:
    outputWorkbook.remove(outputWorkbook[sheet])
for section in sectionsToFind:
    newsheet = outputWorkbook.create_sheet(section['tabTitle'])
    for i in range(len(columns)):
        newsheet.cell(row=1, column=i+1, value=columns[i])

activeSection = None
headerColumn = None

for workbookName in workbooksToSearch:
    workbook = load_workbook(filename=workbookName, read_only=True)
    searchSheet = workbook[tabName]
    for row in searchSheet.values:
        if activeSection is None:
            for section in sectionsToFind:
                if section['name'] in row:
                    #extract section
                    activeSection = section
                    headerColumn = row.index(activeSection['name'])
                    sectionsToFind.remove(section)
                    break
        else:
            firstcell = row[headerColumn]
            if row[headerColumn] == columns[0]:
                continue
            if row[headerColumn] is None and row[headerColumn+1] is None:
                activeSection = None
                headerColumn = None
                continue
            outputSheet = outputWorkbook[activeSection['tabTitle']]
            valuesToAdd = []
            for i in range (headerColumn,headerColumn+len(columns)):
                valuesToAdd.append(row[i])
            outputSheet.append(valuesToAdd)

# print('\n\n------------------------------')
# print('---- OUTPUT ----')
# print('------------------------------')
# for sheetName in outputWorkbook.sheetnames:
#     print(f'Tab title: {sheetName}')
#     sheet = outputWorkbook[sheetName]
#     for row in sheet.values:
#         print(row)


outputWorkbook.save(f'{outputWorkbookName}')
print(f'Output stored in file: {os.path.join(os.getcwd(),outputWorkbookName)}')
