import sys

#print("Hi, this is from python")
spreadsheet = sys.argv[1]
template = sys.argv[2]
sheetname = sys.argv[3]
rangee = sys.argv[4]
if rangee != 'all':
    rangestart = rangee.split('-')[0]
    rangeend = rangee.split('-')[1]
    rangee = 'defined'
else:
    rangee = rangee

if rangee == 'all':
    firstRow = 1
    lastRow = 200
if rangee == 'defined':
    firstRow = int(rangestart)-1
    lastRow = int(rangeend)+1

print(sys.argv)
print("Success")
#print(lastRow)
