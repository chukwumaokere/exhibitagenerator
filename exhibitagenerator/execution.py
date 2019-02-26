## copied from https://pbpython.com/python-word-template.html ##
from __future__ import print_function
from mailmerge import MailMerge
import os, xlrd
import sys

# Define the variables
dir_path = os.path.dirname(os.path.realpath(__file__))
spreadsheet = dir_path + '/../' + sys.argv[1]
template = dir_path + '/../' + sys.argv[2]
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

# Remove when done testing
#wb = xlrd.open_workbook('LLP2_Water_Certificate_List.xlsx')     # Name of your excel file
#sheet = wb.sheet_by_name('2018 Applicants')                     # Name of the sheet within your workbook
# Remove above when done testing

#for j in range(1,all_rows):
for j in range(firstRow,lastRow):
    document = MailMerge(template)
    document.merge(C_LOT_FULL_ADDRESS= str(sheet.cell(j, 20).value),  C_LOT_ZIP= str(int(sheet.cell(j, 21).value)), CITY_LOT_LEGAL_DESCRIPTION= str(sheet.cell(j, 3).value), 
                   C_LOT_PIN_14= str(sheet.cell(j, 25).value))
    newDocName = 'PIN ' + str(sheet.cell(j, 25).value) + '.docx'
    document.write('../'+newDocName)
    

#print("Your Exhibit A documents have been generated!")
#print("They are now in the same folder as this application.")
#print(" ")
#print(" ")
#print("<<<< Credits >>>>")
#print("Developer: Courtney Medina")
#print("Date: 2019-Feb-22")
#print("License: Creative Commons, Free to use")
#print("Compensation: Thank you messages via email or LinkedIn are appreciated :)")

print("Success")


# Excel cell reference info 
    
    #PIN = str(sheet.cell(i,2).value)            # PINs or whatever is in Column B
    #legalDesc = str(sheet.cell(i,4).value)      # Legal description/Exhibit A, or Col.D
    #address = str(sheet.cell(i,21).value)       # Lot's Address, column U
    #zipCode = str(sheet.cell(i,22).value)       # Lot's Zip code, column V
                                                

