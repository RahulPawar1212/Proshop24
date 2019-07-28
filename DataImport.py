import pandas as pd
import datetime
from pathlib import Path
from datetime import timedelta, date
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

        # convert column to date
        GRNdata['GRN Date'] = pd.to_datetime(GRNdata['GRN Date'])

        # String dates
        #StartDate = datetime.date.strftime(StartDate, "%d/%m/%Y")
        StartDate1 = datetime.datetime.strptime(StartDate, "%d/%m/%Y")         

        # End Dates
        #EndDate = datetime.date.strftime(EndDate, "%d/%m/%Y")
        EndDate1 = datetime.datetime.strptime(EndDate, "%d/%m/%Y")

        #Adjustment for end date
        EndDate1 = EndDate1 + datetime.timedelta(days=1)

        # create mask to filtter by date
        #BetweenDates = ((GRNdata['GRN Date'] >= StartDate ) & (GRNdata['GRN Date'] <= EndDate) )

        #Run maskes and get data
        GRNdata = GRNdata.loc[GetBeginCon]
        
        #Make GRN as index
        GRNdata = GRNdata.set_index(['GRN Date'])

        def daterange(self,StartDate, end_date):
            for n in range(int ((end_date - StartDate).days)):
                yield StartDate + timedelta(n)
      
        def Reverse(self,lst):
            return [ele for ele in reversed(list(enumerate(lst)))]

        for single_date in daterange(self,StartDate1, EndDate1):
            if (len(GRNdata.loc[single_date.strftime("%d-%b-%Y"):single_date.strftime("%d-%b-%Y")].index) > 0):
                StartDate = single_date.strftime("%d-%b-%Y")
                break

        for single_date in Reverse(self,daterange(self,StartDate1, EndDate1)):
            if (len(GRNdata.loc[single_date[1].strftime("%d-%b-%Y"):single_date[1].strftime("%d-%b-%Y")].index) > 0):
                EndDate = single_date[1].strftime("%d-%b-%Y")
                break

        #Select dates
        GRNdata = GRNdata.loc[StartDate:EndDate]
        
        GRNdata = GRNdata.reset_index()

        # Remove Time stamp from date
        GRNdata['GRN Date'] = GRNdata['GRN Date'].dt.date

        #Get count of data
        GRNdataCount = GRNdata.groupby(['GRN Date','Item SkuCode']).count().reset_index()

        return GRNdataCount

    def getsalesData(self,filepath,StartDate,EndDate):
        
        #Read data from csv file
        salesdata = pd.read_csv(filepath,sep=',')

        #Select required columns
        salesdata = salesdata[['Sale Order Status','Item SKU Code','Order Date as dd/mm/yyyy hh:MM:ss']]

        # convert column to date
        salesdata['Order Date as dd/mm/yyyy hh:MM:ss'] = pd.to_datetime(salesdata['Order Date as dd/mm/yyyy hh:MM:ss'])

        #**Filter data on the basis of dates**
        # Create mask to get other than cancelled data
        GetOtherThanCnl = (salesdata['Sale Order Status'] != 'CANCELLED')

        # Create mask to get skus of begins with 'con'
        GetBeginCon = (salesdata['Item SKU Code'].str.startswith('CON'))                

        # String dates
        StartDate1 = datetime.datetime.strptime(StartDate, "%d/%m/%Y")         

        # End Dates
        EndDate1 = datetime.datetime.strptime(EndDate, "%d/%m/%Y")

        #Adjustment for end date
        EndDate1 = EndDate1 + datetime.timedelta(days=1)

        # Run maskes and get data
        salesdata = salesdata.loc[GetOtherThanCnl & GetBeginCon]

        
        # create mask to filtter by date
        #BetweenDates = ((salesdata['Order Date as dd/mm/yyyy hh:MM:ss'] >= StartDate ) & (salesdata['Order Date as dd/mm/yyyy hh:MM:ss'] <= EndDate) )

        #Make GRN as index
        salesdata = salesdata.set_index(['Order Date as dd/mm/yyyy hh:MM:ss'])


        def daterange(self,StartDate, end_date):
            for n in range(int ((end_date - StartDate).days)):
                yield StartDate + timedelta(n)
      
        def Reverse(self,lst): 
            return [ele for ele in reversed(list(enumerate(lst)))]

        for single_date in daterange(self,StartDate1, EndDate1):
            if (len(salesdata.loc[single_date.strftime("%d-%b-%Y"):single_date.strftime("%d-%b-%Y")].index) > 0):
                StartDate = single_date.strftime("%d-%b-%Y")
                break

        for single_date in Reverse(self,daterange(self,StartDate1, EndDate1)):
            if (len(salesdata.loc[single_date[1].strftime("%d-%b-%Y"):single_date[1].strftime("%d-%b-%Y")].index) > 0):
                EndDate = single_date[1].strftime("%d-%b-%Y")
                break

        #Select dates
        salesdata = salesdata.loc[StartDate:EndDate]

        #Reset index
        salesdata = salesdata.reset_index()


        #Select required columns
        salesdata = salesdata[['Sale Order Status','Item SKU Code']]

        #Select data where Item SKU code begins with con
        #salesdata = salesdata.loc[salesdata['Item SKU Code'].str.startswith('CON')]

        #Get count of data
        SalesDataCount = salesdata.groupby(['Item SKU Code']).count().reset_index()
        
        return SalesDataCount
