#import DataImport
import datetime
from pathlib import Path
import pandas as pd
import openpyxl as pyxl
from openpyxl.styles import Alignment, Color, Font, PatternFill, colors
from openpyxl.styles.borders import Border, Side
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell
from py_linq import Enumerable

class processData:
    def __init__(self):
        self
    
    def SalesDataProcess(self,SalesData,wb,StartDate):
       
        wsSales_Reports  = wb['Sales Reports']
        
        Sales_ReportsMaxCol = maxCol(wsSales_Reports,9,1)
        Sales_ReportsMaxRow = maxRow(wsSales_Reports,9,1)

        # Add two columns 
        wsSales_Reports.insert_cols(Sales_ReportsMaxCol)
        wsSales_Reports.insert_cols(Sales_ReportsMaxCol)


        for i in range(10,Sales_ReportsMaxRow + 1):
            if SalesData.loc[SalesData['Item SKU Code'] == wsSales_Reports.cell(i,2).value,'Sale Order Status'].count() == 1:
                wsSales_Reports.cell(i,Sales_ReportsMaxCol).value = SalesData.loc[SalesData['Item SKU Code'] == wsSales_Reports.cell(i,2).value,'Sale Order Status'].values[0]
            else:
                wsSales_Reports.cell(i,Sales_ReportsMaxCol).value = 0

        formulaSumSQ =  '=SUM(' + xl_col_to_name(Sales_ReportsMaxCol -1,True) + str(10) + ':' + xl_col_to_name(Sales_ReportsMaxCol - 1,True) + str(Sales_ReportsMaxRow) + ")"
        formulaSumAP = '=SUM(' + xl_col_to_name(Sales_ReportsMaxCol ,True) + str(10) + ':' + xl_col_to_name(Sales_ReportsMaxCol,True) + str(Sales_ReportsMaxRow) + ")"
        FormulaBQ = '=SUM(' + xl_col_to_name(Sales_ReportsMaxCol + 1,True) + str(10) + ':' + xl_col_to_name(Sales_ReportsMaxCol + 1,True) + str(Sales_ReportsMaxRow) + ")"

        wsSales_Reports[xl_col_to_name(Sales_ReportsMaxCol -1) + str(Sales_ReportsMaxRow + 1)].value = formulaSumSQ
        wsSales_Reports[xl_col_to_name(Sales_ReportsMaxCol) + str(Sales_ReportsMaxRow + 1)].value = formulaSumAP
        wsSales_Reports[xl_col_to_name(Sales_ReportsMaxCol + 1) + str(Sales_ReportsMaxRow + 1)].value = FormulaBQ


        marginPercent = round((1 - wsSales_Reports['E10'].value),2)
        strRange = xl_col_to_name(Sales_ReportsMaxCol)  + '10' + ':' + xl_col_to_name(Sales_ReportsMaxCol) + str(Sales_ReportsMaxRow)

        formulaAP = '=(' + str(marginPercent) + '*$D{0})*' + xl_col_to_name(Sales_ReportsMaxCol -1 ,True) + '{0}'
        for i, cellObj in enumerate(wsSales_Reports[strRange], 10):
            cellObj[0].value = formulaAP.format(i)

        for i in range(10,Sales_ReportsMaxRow + 1):
            wsSales_Reports.cell(i,Sales_ReportsMaxCol + 2).value = "=" + GetFormuleBQ(Sales_ReportsMaxCol + 2,i)[:-1]
        
        blackFill = PatternFill(start_color='FF000000', end_color='FF000000', fill_type='solid')
        ft_White = Font(color=colors.WHITE)
        ft_WhiteBold = Font(bold=True,color=colors.WHITE)
        
        _cell1 = wsSales_Reports.cell(9,Sales_ReportsMaxCol)         
        _cell1.font = ft_White       
        _cell1.fill = blackFill
        _cell1.value = StartDate
        _cell1.alignment = Alignment(wrapText=True) 

        _cell2 = wsSales_Reports.cell(9,Sales_ReportsMaxCol + 1)        
        _cell2.font = ft_White
        _cell2.fill = blackFill
        _cell2.value = 'Amount Payable'
        _cell2.alignment = Alignment(wrapText=True)
        

        _cell3 = wsSales_Reports.cell(Sales_ReportsMaxRow + 1,Sales_ReportsMaxCol)
        _cell3.font = ft_WhiteBold
        _cell3.fill = blackFill
        
        _cell4 = wsSales_Reports.cell(Sales_ReportsMaxRow + 1,Sales_ReportsMaxCol + 1)
        _cell4.font = ft_WhiteBold
        _cell4.fill = blackFill

        strRange = xl_col_to_name(Sales_ReportsMaxCol - 1)  + '10' + ':' + xl_col_to_name(Sales_ReportsMaxCol) + str(Sales_ReportsMaxRow)

        set_border(wsSales_Reports,strRange)
        

    def StockDataProcess(self,StockData,wb,StartDate,end_date):
        wsStock_Update  = wb['Stock Update']

        wsStock_UpdateMaxCol = maxCol(wsStock_Update,1,1)
        wsStock_UpdateMaxRow = maxRow(wsStock_Update,1,1)

        from datetime import timedelta, date

        #StartDate =  datetime.datetime.strptime(StartDate, "%d/%m/%Y")
        #end_date = datetime.datetime.strptime(end_date, "%d/%m/%Y")

        StartDate = datetime.datetime.strptime(StartDate, "%d/%m/%Y")         

        end_date = datetime.datetime.strptime(end_date, "%d/%m/%Y")
        
        #Adjustment for end date
        end_date = end_date + datetime.timedelta(days=1)

        def daterange(self,StartDate, end_date):
            for n in range(int ((end_date - StartDate).days)):
                yield StartDate + timedelta(n)


        #marks = Enumerable(StockData)
        #passing = marks.where(lambda x: x['Item SkuCode'] == 'CONHSOFQRK001') # results in [50, 80, 90]

        #print(passing)

        col = 1
        col2 = 0
        for single_date in daterange(self,StartDate, end_date):
            for i in range(2,wsStock_UpdateMaxRow + 1): # Loop through excel data
                if StockData.loc[StockData['Item SkuCode'] == wsStock_Update.cell(i,2).value,'Quantity Received'].count() > 1:
                    #wsStock_Update.cell(i,wsStock_UpdateMaxCol + col).value = StockData.loc[StockData['Item SkuCode'] == wsStock_Update.cell(i,2).value].values[0]
                    #if datetime.datetime.strptime(StockData.loc[StockData['GRN Date']],'%d-%m-%Y').date() == single_date:
                    df = StockData.loc[(StockData['Item SkuCode'] == wsStock_Update.cell(i,2).value)]
                    df['GRN Date'] = pd.to_datetime(df['GRN Date'])    
                    df = df.set_index(['GRN Date'])                
                    #df['GRN Date'] = df['GRN Date'].dt.date
                    df = df.loc[single_date:single_date]
                    df2 = df.reset_index()
                    if df2.loc[df2['GRN Date'] == single_date,'Quantity Received'].count() == 1:
                                            wsStock_Update.cell(i,wsStock_UpdateMaxCol + col).value = df2.loc[(df2['GRN Date'] == single_date,'Quantity Received')].values[0]
                    else:
                        wsStock_Update.cell(i,wsStock_UpdateMaxCol + col).value = 0
                else:
                    wsStock_Update.cell(i,wsStock_UpdateMaxCol + col).value = 0
               

#print(df.loc[(df['GRN Date'] == '30-4-2019')])
            FormulaSQ = '=SUM(' + xl_col_to_name(wsStock_UpdateMaxCol + col2,True) + str(2) + ':' + xl_col_to_name(wsStock_UpdateMaxCol + col2,True) + str(wsStock_UpdateMaxRow) + ")"

            wsStock_Update[xl_col_to_name(wsStock_UpdateMaxCol + col2,True) + str(wsStock_UpdateMaxRow + 1)].value = FormulaSQ

            FormulaPB = '=SUM(' + xl_col_to_name(3,True) + str(2) + ':' + xl_col_to_name(3,True) + str(wsStock_UpdateMaxRow) + ")"

            wsStock_Update[xl_col_to_name(3,True) + str(wsStock_UpdateMaxRow + 1)].value = FormulaPB
            
            strRange = xl_col_to_name(3)  + '2' + ':' + xl_col_to_name(3) + str(wsStock_UpdateMaxRow)

            formulaAP = '=SUM(' + xl_col_to_name(4,True) + '{0}' + ':' + xl_col_to_name(wsStock_UpdateMaxCol + col2,True) + '{0}'+ ")"
            for i, cellObj in enumerate(wsStock_Update[strRange], 2):
                cellObj[0].value = formulaAP.format(i)



            ## Formating starts here 
            blackFill = PatternFill(start_color='FF000000', end_color='FF000000', fill_type='solid')
            ft_White = Font(color=colors.WHITE)
           #ft_WhiteBold = Font(bold=True,color=colors.WHITE)
            ft_BlackBold = Font(bold=True,color=colors.BLACK)

            _cell1 = wsStock_Update.cell(1,wsStock_UpdateMaxCol + col)        
            _cell1.font = ft_White
            _cell1.fill = blackFill
            _cell1.value = 'Qty Rcvd ' + str(single_date)
            _cell1.alignment = Alignment(wrapText=True)
            
            _cell2 = wsStock_Update.cell(wsStock_UpdateMaxRow + 1,wsStock_UpdateMaxCol + col)        
            _cell2.font = ft_BlackBold
            #_cell2.fill = blackFill
            
            strRange = xl_col_to_name(wsStock_UpdateMaxCol + col2,True)  + '2' + ':' + xl_col_to_name(wsStock_UpdateMaxCol + col2) + str(wsStock_UpdateMaxRow)

            set_border(wsStock_Update,strRange)
            col = col + 1
            col2 = col2 + 1

        for jj in range(2,wsStock_UpdateMaxRow + 1):
            mxColNow = maxCol(wsStock_Update,1,1)
            for j in range(5,mxColNow):
                formula = xl_col_to_name(j)
                #print(wsStock_Update.cell(1,j + 2).value)
                if j == 5 :
                    
                    if 'Returned' not in wsStock_Update.cell(1,j + 2).value:                
                        formula1 = 'E' + str(jj) + '+' + (formula + str(jj)) +  '+'
                    if 'Returned' in wsStock_Update.cell(1,j + 2).value:
                        formula1 = 'E' + str(jj) + '+' + (formula + str(jj)) +  '-'
                elif (j > 5) & (j <  mxColNow - 1):
                    if 'Returned' not in wsStock_Update.cell(1,j + 2).value:
                        formula1 = formula1 + (formula + str(jj)) +  '+'
                    if 'Returned' in wsStock_Update.cell(1,j + 2).value:
                        formula1 = formula1 + (formula + str(jj)) +  '-'
                elif (j ==  mxColNow -1):
                    formula1 = formula1 + (formula + str(jj))

            wsStock_Update.cell(jj,4).value = "=" + formula1        









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

def GetFormuleBQ(intMaxCol,inRow):    
    intcount = 1
    for i in range(7,intMaxCol):
        if intcount == 3:
            intcount = 1
        
        if intcount == 1:
            formula = xl_col_to_name(i - 1)
            if i == 7:
                formula1 = 'F' + str(inRow) + '-' + (formula + str(inRow)) +  '-' 
            elif i > 7 & i != intMaxCol:    
                formula1 = formula1 + (formula + str(inRow) + '-')        
            intcount = intcount + 1
        else:
            intcount = intcount + 1
    return formula1


def set_border(ws, cell_range):
    rows = ws[cell_range]
    side = Side(border_style='thin', color="FF000000")

    rows = list(rows)  # we convert iterator to list for simplicity, but it's not memory efficient solution
    max_y = len(rows) - 1  # index of the last row
    for pos_y, cells in enumerate(rows):
        max_x = len(cells) - 1  # index of the last cell
        for pos_x, cell in enumerate(cells):
            border = Border(
                left=cell.border.left,
                right=cell.border.right,
                top=cell.border.top,
                bottom=cell.border.bottom
            )
            border.left = side
            border.right = side
            border.top = side
            border.bottom = side
            #if pos_x == 0:
            #    border.left = side
            #if pos_x == max_x:
            #    border.right = side
            #if pos_y == 0:
            #    border.top = side
            #if pos_y == max_y:
            #    border.bottom = side

            # set new border only if it's one of the edge cells
            if pos_x == 0 or pos_x == max_x or pos_y == 0 or pos_y == max_y:
                cell.border = border
