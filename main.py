import DataImport
from pathlib import Path

GrnDataPath = Path(r"C:\Users\rahul.pawar\Desktop\Proshop\Goods Received Notes (Stock Intake).csv")
SalesDataPath = Path(r"C:\Users\rahul.pawar\Desktop\Proshop\Sales Export From Backend.csv")


p1 = DataImport.dataImport()
GRNdata = p1.getGRNData(GrnDataPath)
SalesData = p1.getsalesData(SalesDataPath)