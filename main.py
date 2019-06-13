import DataImport
import DataProcessing
import openpyxl as pyxl
from pathlib import Path
import dateparser
import datetime
import os
import shutil
import glob
#from xlsxwriter.utility import xl_rowcol_to_cell,xl_col_to_name
#from openpyxl.styles import Fill,Font,Color
#from openpyxl.styles import colors

GrnDataPath = Path(r"C:\Users\rahul.pawar\Desktop\Proshop\Goods Received Notes (Stock Intake).csv")
SalesDataPath = Path(r"C:\Users\rahul.pawar\Desktop\Proshop\Sales Export From Backend.csv")
SalesReporExcel = Path(r"C:\Users\rahul.pawar\Desktop\Proshop")
strSalesReporExcel = os.path.abspath(str(SalesReporExcel))

StartDate = dateparser.parse('27-8-2019')
EndDate = dateparser.parse('29-8-2019')

#************************** Import Data *******************************
Di = DataImport.dataImport()
GRNdata = Di.getGRNData(GrnDataPath,StartDate,EndDate)
SalesData = Di.getsalesData(SalesDataPath,StartDate,EndDate)
#*****************************************************************

SalesRptXlintoOutPut = SalesReporExcel
SalesRptXlintoOutPut = SalesRptXlintoOutPut.joinpath('Output')

# Create folder if not exists / delete if exists and recreate folder
if not os.path.exists(SalesRptXlintoOutPut):
    os.makedirs(SalesRptXlintoOutPut)
elif os.path.exists(SalesRptXlintoOutPut):
    shutil.rmtree(SalesRptXlintoOutPut)
    os.makedirs(SalesRptXlintoOutPut)

import ntpath
strSalesReporExcel
# Get file names from path along with path
files = glob.glob(strSalesReporExcel + '\\' + '*.xlsx')
for entry in files:    
    # Sales report excel
    #wb = pyxl.load_workbook(filename= SalesReporExcel.joinpath ('Sample Report 2.xlsx') ,read_only=False)
    wb = pyxl.load_workbook(filename= entry ,read_only=False)

    Dp = DataProcessing.processData()
    Dp.SalesDataProcess(SalesData,wb,datetime.date.strftime(StartDate, "%d/%m/%Y") )
    Dp.StockDataProcess(GRNdata,wb,datetime.date.strftime(StartDate, "%d/%m/%Y"))   
    

    testpath =   ntpath.basename(entry)
    wb.save(SalesRptXlintoOutPut.joinpath(testpath))


#Testing
#print(Sales_ReportsMaxCol)
#print(Sales_ReportsMaxRow)
#print(wsStock_UpdateMaxCol)
#print(wsStock_UpdateMaxRow)