import pandas as pd
import dateparser
import datetime
from pathlib import Path

class DataImport:
    def __init__(self):
        self
    
    def getSalesData(self, filepath):
        #Read data from csv file
        GRNdata = pd.read_csv(filepath,sep=",")
        #Select required columns
        GRNdata = GRNdata[['Item SkuCode','Quantity Received']]
        # Create mask to get skus of begins with 'con'
        GetBeginCon = (GRNdata['Item SkuCode'].str.startswith('CON'))
        # Run maskes and get data
        GRNdata = GRNdata.loc[GetBeginCon]
        return GRNdata