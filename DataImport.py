import pandas as pd
import datetime
from pathlib import Path

class dataImport:
    def __init__(self):
        self
    
    def getGRNData(self, filepath,StartDate,EndDate):
        #Read data from csv file
        GRNdata = pd.read_csv(filepath,sep=';')
        #Select required columns
        GRNdata = GRNdata[['GRN Date','Item SkuCode','Quantity Received']]
        # Create mask to get skus of begins with 'con'
        GetBeginCon = (GRNdata['Item SkuCode'].str.startswith('CON'))

        StartDate = datetime.date.strftime(StartDate, "%d/%m/%Y")         

        EndDate = EndDate + datetime.timedelta(days=1)

        EndDate = datetime.date.strftime(EndDate, "%d/%m/%Y")

        # create mask to filtter by date
        BetweenDates = ((GRNdata['GRN Date'] >= StartDate ) & (GRNdata['GRN Date'] <= EndDate) )

        # Run maskes and get data
        GRNdata = GRNdata.loc[GetBeginCon & BetweenDates]

        #Get count of data
        GRNdataCount = GRNdata.groupby(['Item SkuCode']).count().reset_index()

        return GRNdataCount

    def getsalesData(self,filepath,StartDate,EndDate):
        
        #Read data from csv file
        salesdata = pd.read_csv(filepath,sep=',')

        #Select required columns
        salesdata = salesdata[['Sale Order Status','Item SKU Code','Order Date as dd/mm/yyyy hh:MM:ss']]

        #**Filter data on the basis of dates**
        # Create mask to get other than cancelled data
        GetOtherThanCnl = (salesdata['Sale Order Status'] != 'CANCELLED')

        # Create mask to get skus of begins with 'con'
        GetBeginCon = (salesdata['Item SKU Code'].str.startswith('CON'))                

        StartDate = datetime.date.strftime(StartDate, "%d/%m/%Y")         

        EndDate = EndDate + datetime.timedelta(days=1)

        EndDate = datetime.date.strftime(EndDate, "%d/%m/%Y")

        # create mask to filtter by date
        BetweenDates = ((salesdata['Order Date as dd/mm/yyyy hh:MM:ss'] >= StartDate ) & (salesdata['Order Date as dd/mm/yyyy hh:MM:ss'] <= EndDate) )

        # Run maskes and get data
        salesdata = salesdata.loc[GetOtherThanCnl & GetBeginCon  & BetweenDates]

        #Select required columns
        salesdata = salesdata[['Sale Order Status','Item SKU Code']]

        #Select data where Item SKU code begins with con
        #salesdata = salesdata.loc[salesdata['Item SKU Code'].str.startswith('CON')]

        #Get count of data
        SalesDataCount = salesdata.groupby(['Item SKU Code']).count().reset_index()
        return SalesDataCount
