import os,openpyxl
import sys,glob
from openpyxl import Workbook
from datetime import datetime

os.chdir("C:\Users\dell\Documents\Stock Picker")
book =  openpyxl.load_workbook("Scripts.xlsx")
scripts = book.get_sheet_by_name("Scripts")

os.chdir("C:\Users\dell\Documents\Stock Picker\DB")

start = datetime.now()
print "\n\nComputing 52 week low/high of the scripts..........\n\n"
book_52_week = Workbook()
page = book_52_week.active

page['A1'],page['B1'],page['C1'] = "Script", "52 week high", "52 week low"
row=2

'''
for script in glob.glob(os.path.join('.', '*.xlsx')):
    if script[2].isalpha() or script[2].isdigit():
        id=script[2:len(script)-5]
'''

for i in range(2,20):                       #scripts.max_row+1
    id = scripts['B'+str(i)].value
    print id
    wb=openpyxl.load_workbook(id+".xlsx")
    ws=wb.active
    low = float(ws['D2'].value)
    high = float(ws['C2'].value)
    for i in range(3,262):
        data1 = float(ws['D'+str(i)].value)
        data2 = float(ws['C'+str(i)].value)
        if data1 < low:
            low = data1
            if data2 > high:
                high = data2
                
    page['A'+str(row)] = id
    page['B'+str(row)] = high
    page['C'+str(row)] = low
    row+=1

end = datetime.now()

os.chdir("C:\Users\dell\Documents\Stock Picker\Results") 

os.remove("52 week low high.xlsx")  
book_52_week.save("52 week low high.xlsx")

print "\n\nSuccessfully computed 52 week low/high..."


# <------------------------Time Keeper Update------------------>

os.chdir("E:\Enthought Canopy\Stock Picker")

hours = round((end-start).seconds/3600,0)
minutes = (end-start).seconds%3600
minutes = round(minutes/60,0)
seconds = round((end-start).seconds%60,0)

sys.argv = [str(datetime.now().date()), "52 week low/high", str(start.time()), str(end.time()), int(hours), int(minutes), int(seconds)]
execfile("Datetime update.py")

