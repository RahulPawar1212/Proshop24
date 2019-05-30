import pandas as pd
import dateparser
import datetime

#Read data from csv file
GRNdata = pd.read_csv(r"C:\Users\rahul.pawar\Desktop\Proshop\Goods Received Notes (Stock Intake).csv")

#Select required columns
GRNdata = GRNdata[['Item SkuCode','Quantity Received']]

# Create mask to get skus of begins with 'con'
GetBeginCon = (GRNdata['Item SkuCode'].str.startswith('CON'))


# Run maskes and get data
GRNdata = GRNdata.loc[GetBeginCon]
