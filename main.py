import DataImport
import openpyxl as pyxl
from pathlib import Path
from xlsxwriter.utility import xl_rowcol_to_cell,xl_col_to_name

GrnDataPath = Path(r"C:\Users\rahul.pawar\Desktop\Proshop\Goods Received Notes (Stock Intake).csv")
SalesDataPath = Path(r"C:\Users\rahul.pawar\Desktop\Proshop\Sales Export From Backend.csv")


p1 = DataImport.dataImport()
GRNdata = p1.getGRNData(GrnDataPath)
SalesData = p1.getsalesData(SalesDataPath)

SalesReporExcel = Path(r"C:\Users\rahul.pawar\Desktop\Proshop\Sample Report 1.xlsx")
wb = pyxl.load_workbook(filename= SalesReporExcel,read_only=False)

wsSales_Reports  = wb['Sales Reports']
wsStock_Update  = wb['Stock Update']

def maxCol(ws,intStartRow,intStartCol):
    maxcol = intStartCol - 1 
    for i in range(intStartCol,16384):
        if ws.cell(intStartRow,i).value == None:
            break
        #print(ws.cell(intStartRow,i).value)
        maxcol = maxcol + 1    
    return maxcol

def maxRow(ws,intStartRow,intStartCol):
    maxrow = intStartRow -1
    for i in range(intStartRow,1048576):
        if ws.cell(i,intStartCol).value == None:            
            break
        #print(ws.cell(i,intStartCol).value)
        maxrow = maxrow + 1    
    return maxrow


Sales_ReportsMaxCol = maxCol(wsSales_Reports,9,1)
Sales_ReportsMaxRow = maxRow(wsSales_Reports,9,1)

wsStock_UpdateMaxCol = maxCol(wsStock_Update,1,1)
wsStock_UpdateMaxRow = maxRow(wsStock_Update,1,1)

# Add two columns 
wsSales_Reports.insert_cols(Sales_ReportsMaxCol)
wsSales_Reports.insert_cols(Sales_ReportsMaxCol)



for i in range(10,Sales_ReportsMaxRow + 1):
    if SalesData.loc[SalesData['Item SKU Code'] == wsSales_Reports.cell(i,2).value,'Sale Order Status'].count() == 1:
        wsSales_Reports.cell(i,Sales_ReportsMaxCol).value = SalesData.loc[SalesData['Item SKU Code'] == wsSales_Reports.cell(i,2).value,'Sale Order Status'].values[0]
    else:
        wsSales_Reports.cell(i,Sales_ReportsMaxCol).value = 0

formulaSumKcol =  '=SUM(' + xl_col_to_name(Sales_ReportsMaxCol -1,True) + str(10) + ':' + xl_col_to_name(Sales_ReportsMaxCol - 1,True) + str(Sales_ReportsMaxRow) + ")"
formulaSumLcol = '=SUM(' + xl_col_to_name(Sales_ReportsMaxCol ,True) + str(10) + ':' + xl_col_to_name(Sales_ReportsMaxCol,True) + str(Sales_ReportsMaxRow) + ")"

wsSales_Reports['K53'].value = formulaSumKcol
wsSales_Reports['L53'].value = formulaSumLcol


marginPercent = round((1 - wsSales_Reports['E10'].value),2)
strRange = xl_col_to_name(Sales_ReportsMaxCol)  + '10' + ':' + xl_col_to_name(Sales_ReportsMaxCol) + str(Sales_ReportsMaxRow)

formulaAP = '=(' + str(marginPercent) + '*$D{0})*' + xl_col_to_name(Sales_ReportsMaxCol -1 ,True) + '{0}'
for i, cellObj in enumerate(wsSales_Reports[strRange], 10):
    cellObj[0].value = formulaAP.format(i)

#'=IF($A{0}=$B{0}, "Match", "Mismatch")'



print(Sales_ReportsMaxCol)
print(Sales_ReportsMaxRow)
print(wsStock_UpdateMaxCol)
print(wsStock_UpdateMaxRow)
wb.save("Test5.xlsx")