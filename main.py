import DataImport
import DataProcessing
import openpyxl as pyxl
from pathlib import Path
import dateparser
import datetime
#from xlsxwriter.utility import xl_rowcol_to_cell,xl_col_to_name
#from openpyxl.styles import Fill,Font,Color
#from openpyxl.styles import colors

GrnDataPath = Path(r"C:\Users\rahul.pawar\Desktop\Proshop\Goods Received Notes (Stock Intake).csv")
SalesDataPath = Path(r"C:\Users\rahul.pawar\Desktop\Proshop\Sales Export From Backend.csv")

StartDate = dateparser.parse('27-8-2019')
EndDate = dateparser.parse('29-8-2019')

Di = DataImport.dataImport()
GRNdata = Di.getGRNData(GrnDataPath,StartDate,EndDate)
SalesData = Di.getsalesData(SalesDataPath,StartDate,EndDate)

SalesReporExcel = Path(r"C:\Users\rahul.pawar\Desktop\Proshop\Sample Report 2.xlsx")
wb = pyxl.load_workbook(filename= SalesReporExcel,read_only=False)


Dp = DataProcessing.processData()
Dp.SalesDataProcess(SalesData,wb,datetime.date.strftime(StartDate, "%d/%m/%Y") )
Dp.StockDataProcess(GRNdata,wb,datetime.date.strftime(StartDate, "%d/%m/%Y"))

#print(Sales_ReportsMaxCol)
#print(Sales_ReportsMaxRow)
#print(wsStock_UpdateMaxCol)
#print(wsStock_UpdateMaxRow)
wb.save("Test5.xlsx")