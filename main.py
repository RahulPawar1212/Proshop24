import DataImport
from pathlib import Path

data_folder = Path(r"C:\Users\rahul.pawar\Desktop\Proshop\Goods Received Notes (Stock Intake).csv")

p1 = DataImport.DataImport()
GRNdata = p1.getSalesData(data_folder)