import os # to navigate into the computer
import pandas as pd # to maniputale data
import datetime as dt # to get the year and quarter from invoice date
import xlwings as xw # to interact with Excel
import re #to match Excel cells column and row references in a function we will use later

# Navigate into the folder were the data set is. Create a variable with the Data and put it into a data frame
os.chdir(r'/Users/aymenbaraket/Downloads')
Data_Base='data.csv'
df = pd.read_csv(Data_Base,encoding= 'unicode_escape')

# Add column revenue 
df['revenue']=df['Quantity']*df['UnitPrice']
#Create a list month and isolate the first value of the invoice date to get the month
Month = df["InvoiceDate"].values
Month = [my_str.split("/")[0] for my_str in Month]
df["Month"] = Month
# Add column Year
df['year'] = pd.DatetimeIndex(df['InvoiceDate']).year
#Add column Quarter
df.date = pd.to_datetime(df.InvoiceDate) 
df['quarter'] = df.date.dt.quarter

# open the excel file, select the tab and the PivotTable to refresh
wb = xw.Book(r"/Users/aymenbaraket/Downloads/Dashboard.xlsb")
sheet = wb.sheets('DataBase')  #Name of sheet where to append df

#Function to dump the data frame in order to avoid a time out error.
def dumpLargeDf(sheet, df, startcell='A1', chunk_size=50000):
    # Dumps a large DataFrame in Excel via xlwings. Takes care of header.
    if len(df) <= (chunk_size + 1):
        sheet.range(startcell).options(index=False).value = df
    else:                                       # Chunk df and and dump each
        c = re.match(r"([a-z]+)([0-9]+)", startcell, re.I)      # A1
        row = c.group(1)                                        # A
        col = int(c.group(2))                                   # 1
        useHeader = True
        for chunk in (df[rw:rw + chunk_size] for rw in
                      range(0, len(df), chunk_size)):
            print("Dumping chunk in %s%s" % (row, col))
            sheet.range(row + str(col))                 .options(index=False, header=useHeader).value = chunk
            useHeader = False
            col += chunk_size


dumpLargeDf(sheet,df)
app = wb.app
macro_vba = app.macro("'Dashboard.XLSB'!refresh") 
macro_vba()
wb.save()
app = xw.apps.active 
app.quit()



