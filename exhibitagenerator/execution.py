## copied from https://pbpython.com/python-word-template.html ##
from __future__ import print_function
from mailmerge import MailMerge
import os, xlrd
import sys

# Define the variables
dir_path = os.path.dirname(os.path.realpath(__file__))
outputPath = sys.argv[5]
spreadsheet = sys.argv[1]
template = sys.argv[2]
sheetname = sys.argv[3]
rangee = sys.argv[4]
print(sys.argv)

# Handling if all rows in spreadsheet or just a defined amount
if rangee != 'all':
    rangestart = rangee.split('-')[0]
    rangeend = rangee.split('-')[1]
    rangee = 'defined'
else: 
    rangee = rangee

document = MailMerge(template)

# Initializes Excel source file and sheet name variables provided by user
wb = xlrd.open_workbook(spreadsheet)     # Name of your excel file
sheet = wb.sheet_by_name(sheetname)            # Name of the sheet within your workbook

# Defining loop range
if rangee == 'all':
    firstRow = 1
    lastRow = sheet.nrows
if rangee == 'defined':
    firstRow = int(rangestart)
    lastRow = int(rangeend)+1

#create output path if doesnt exist
if not os.path.exists(outputPath):
    try: 
        os.makedirs(outputPath)
    except OSError as e:
        if e.errno != errno.EEXIST:
            raise

for j in range(firstRow,lastRow):
    document = MailMerge(template)
    document.merge(C_LOT_FULL_ADDRESS= str(sheet.cell(j, 20).value),  C_LOT_ZIP= str(int(sheet.cell(j, 21).value)), CITY_LOT_LEGAL_DESCRIPTION= str(sheet.cell(j, 3).value), 
                   C_LOT_PIN_14= str(sheet.cell(j, 25).value))
    newDocName = 'PIN ' + str(sheet.cell(j, 25).value) + '.docx'
    document.write(outputPath+newDocName)


print("Success")
