import pandas as pd
import dateparser
import datetime

#Read data from csv file
salesdata = pd.read_csv(r"C:\Users\rahul.pawar\Desktop\Proshop\Sales Export From Backend.csv")

#Select required columns
salesdata = salesdata[['Sale Order Status','Item SKU Code','Order Date as dd/mm/yyyy hh:MM:ss']]

#Filter data on the basis of dates

# Create mask to get other than cancelled data
GetOtherThanCnl = (salesdata['Sale Order Status'] != 'CANCELLED')

# Create mask to get skus of begins with 'con'
GetBeginCon = (salesdata['Item SKU Code'].str.startswith('CON'))

#Get Start and end date
StartDate = dateparser.parse('27-8-2019')

StartDate = datetime.date.strftime(StartDate, "%d/%m/%y")

EndDate = dateparser.parse('27-8-2019') 

EndDate = EndDate + datetime.timedelta(days=1)

EndDate = datetime.date.strftime(EndDate, "%d/%m/%y")

# create mask to filtter by date
BetweenDates = ((salesdata['Order Date as dd/mm/yyyy hh:MM:ss'] >= StartDate ) & (salesdata['Order Date as dd/mm/yyyy hh:MM:ss'] <= EndDate) )

# Run maskes and get data
salesdata = salesdata.loc[GetOtherThanCnl & GetBeginCon  & BetweenDates]


#Select required columns
salesdata = salesdata[['Sale Order Status','Item SKU Code']]


#Select data where Item SKU code begins with con
salesdata = salesdata.loc[salesdata['Item SKU Code'].str.startswith('CON')]


#Get count of data
SalesDataCount = salesdata.groupby(['Item SKU Code']).count().reset_index()
