import os
import win32com.client as win32
from win32com.client import Dispatch
from datetime import datetime


path_bazy = r'C:\Users\ROBERT\Desktop\IT\PYTHON\PYTHON 37 PROJEKTY\księgowość\skrypty osobno'


# """Sprawdza czy arkusz jest otwarty."""
# """Jeżeli arkusz jest zamknięty, otwiera go."""
try:
      ExcelApp = win32.GetActiveObject('Excel.Application')
      wb = ExcelApp.Workbooks("2014 BAZA MAGRO.xlsx")
      ws = wb.Worksheets("BAZA 2014")
      # workbook = ExcelApp.Workbooks("BAZA 2014.xlsx")

except:
      ExcelApp = Dispatch("Excel.Application")
      wb = ExcelApp.Workbooks.Open(path_bazy + "\\2014 BAZA MAGRO.xlsx")
      ws = wb.Worksheets("BAZA 2014")
      # wb.DisplayAlerts = False






tu = {'ALL': 'Allianz', 'AXA': 'AXA', 'COM': 'Compensa', 'EPZU': 'PZU', 'GEN': 'Generali', 'GOT': 'Gothaer',
      'HDI': 'HDI', 'HES': 'Ergo Hestia', 'IGS': 'IGS', 'INT': 'INTER', 'LIN': 'LINK 4', 'MTU': 'MTU',
      'PRO': 'Proama', 'PZU': 'PZU', 'RIS': 'InterRisk', 'TUW': 'TUW', 'TUZ': 'TUZ', 'UNI': 'Uniqa',
      'WAR': 'Warta', 'WIE': 'Wiener', 'YCD': 'You Can Drive'}




# column = ws.Range(f'AY1:AY{ws.UsedRange.Rows.Count}')

# for i, cash in enumerate(column):
#       # print(cash)
#       if cash == 'G':
#           print(cash)

data_wyst = ExcelApp.Cells(1097, 30).Value
tow_ub = ExcelApp.Cells(1097, 38).Value
nr_polisy = ExcelApp.Cells(1097, 40).Value
inkaso = ExcelApp.Cells(1097, 55).Value
forma_p = ExcelApp.Cells(1097, 51).Value
print(data_wyst, nr_polisy, inkaso, forma_p)





wb.Close(SaveChanges=False)



ExcelApp_cash = win32.gencache.EnsureDispatch('Excel.Application')
ExcelApp_cash.Visible = True
wb_cash = ExcelApp_cash.Workbooks.Add()
ws_cash = wb_cash.Worksheets.Add()
ws_cash.Name = 'Gotówka luty 2021 r.'


ws_cash.Cells(1, 1).Value = 'Data'
ws_cash.Cells(1, 2).Value = 'TU'
ws_cash.Cells(1, 3).Value = 'Nr polisy'
ws_cash.Cells(1, 4).Value = 'Kwota inkaso'



ws_cash.Range('A2:A32').Value = data_wyst.strftime('%Y.%m.%d')
ws_cash.Range('B2:B32').Value = tu[tow_ub]
ws_cash.Range('C2:C32').NumberFormat = 0
ws_cash.Range('C2:C32').Value = nr_polisy
ws_cash.Range('D2:D32').Value = inkaso






ws_cash.Columns.AutoFit()
ws_cash.Columns(1).ColumnWidth = 12
ws_cash.Columns(2).ColumnWidth = 10

wb.DisplayAlerts = False
# ws_cash.DisplayAlerts = False
path_do_zapisu_w = r'C:\Users\ROBERT\Desktop\IT\PYTHON\PYTHON 37 PROJEKTY\księgowość\skrypty osobno'

wb_cash.SaveAs(path_do_zapisu_w + "\\inkaso.xlsx")
wb_cash.Close()


ExcelApp.Application.Quit()
wb_cash.DisplayAlerts = True
