import urllib,urllib2, os, openpyxl, sys
from bs4 import  BeautifulSoup
from mechanize import Browser
from datetime import datetime
from openpyxl import Workbook
import pandas as pd

start = datetime.now()

print "Downloading datebase.....\n\n"

os.chdir("C:\Users\dell\Documents\Stock Picker")
book = openpyxl.load_workbook("niftynext50.xlsx")
scripts = book.get_sheet_by_name("niftynext50")

br=Browser()
os.chdir("C:\Users\dell\Documents\Stock Picker\NiftyDB")


for i in range(2,scripts.max_row+1):
    script_id = scripts['C'+str(i)].value
        
    response=br.open("https://in.finance.yahoo.com/quote/"+script_id+".NS/history?period1=1486146600&period2=1497983400&interval=1d&filter=history&frequency=1d")
    soup = BeautifulSoup(response.read(),"html.parser")
    print br.title()
    right_table = soup.find( "table", { "class" : "W(100%) M(0) BdB Bdc($lightGray)" } )

    months = { "Feb":'02', "Mar":'03',"Apr":'04',"May":'05',"Jun":'06'}
    dates = []
    open = []
    close = []
    adj = []
    vol =[]
    high = []
    low = [] 
    
    for rows in right_table.findAll("tr"):
        cells = rows.findAll("td")
        if len(cells)==7:
            date = cells[0].text
            dd = date[:2]
            mm = months[date[3:6]]
            yy = date[7:]
            dates.append(yy+"-"+mm+"-"+dd)
            open.append(cells[1].text)
            high.append(cells[2].text)
            low.append(cells[3].text)
            close.append(cells[4].text)
            adj.append(cells[5].text)
            vol.append(cells[6].text)
        
    df = pd.DataFrame(dates,columns=['Date'])
    df['Open']=open
    df['High']=high
    df['Low']=low
    df['Close']=close
    df['Volume']=vol
    df['Adj Close']=adj
    
    writer = pd.ExcelWriter(script_id+".xlsx")
    #df = pandas.read_excel('your_xls_xlsx_filename', sheetname='Sheet 1')   
    df1 = pd.read_excel("C:\Users\dell\Documents\Stock Picker\DB\\"+script_id+".xlsx", sheetname='Sheet1')
    dataframe = pd.concat([df,df1])
    dataframe.to_excel(writer,index=False)
    writer.save()
    print script_id+".xlsx"+" is ready!"