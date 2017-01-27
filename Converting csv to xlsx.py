import sys,glob
import csv,os
from xlsxwriter.workbook import Workbook

print "\n\nConverting CSV files to .xlsx files.........."
os.chdir("C:\Users\dell\Documents\Stock Picker\DB")

for csvfile in glob.glob(os.path.join('.', '*.csv')):
    workbook = Workbook(csvfile[2:len(csvfile)-4] + '.xlsx')
    worksheet = workbook.add_worksheet("Sheet1")
    with open(csvfile, 'rb') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
    workbook.close()  
    os.remove(csvfile)  

print "\nSuccessfully converted files.\n"
    
