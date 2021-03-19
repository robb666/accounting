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

tu = {'ALL': 'Allianz', 'AXA': 'AXA', 'COM': 'Compensa', 'EPZU': 'PZU', 'GEN': 'Generali',
      'GOT': 'Gothaer', 'HDI': 'HDI', 'HES': 'Ergo Hestia', 'IGS': 'IGS', 'INT': 'INTER',
      'LIN': 'LINK 4', 'MTU': 'MTU', 'PRO': 'Proama', 'PZU': 'PZU', 'RIS': 'InterRisk', 'TUW': 'TUW',
      'TUZ': 'TUZ', 'UNI': 'Uniqa', 'WAR': 'Warta', 'WIE': 'Wiener', 'YCD': 'You Can Drive'}


# forma_płatności = ExcelApp.Cells(1096, 51).Value

for i in range(20):
      if ExcelApp.Cells(i, 51).Value == 'G':



            data_wyst = ExcelApp.Cells(i, 30).Value
            tow_ub = ExcelApp.Cells(i, 38).Value
            nr_polisy = ExcelApp.Cells(i, 40).Value
            inkaso = ExcelApp.Cells(i, 55).Value
            print(data_wyst, tu[tow_ub], nr_polisy, inkaso)

wb.Close(SaveChanges=False)





# """Rozpoznaje kolejny wiersz, który może zapisać."""
# row_to_write = wb.Worksheets(1).Cells(wb.Worksheets(1).Rows.Count, 30).End(-4162).Row + 1
ExcelApp = win32.gencache.EnsureDispatch('Excel.Application')
ExcelApp.Visible = True
wb = ExcelApp.Workbooks.Add()


cash_ws = wb.Worksheets.Add()

cash_ws.Name = 'Gotówka luty 2021 r.'


cash_ws.Cells(1, 1).Value = 'Data'
cash_ws.Cells(1, 2).Value = 'TU'
cash_ws.Cells(1, 3).Value = 'Nr polisy'
cash_ws.Cells(1, 4).Value = 'Kwota inkaso'


cash_ws.Range('A2:A32').Value = data_wyst.strftime('%Y.%m.%d')
cash_ws.Range('B2:B32').Value = tu[tow_ub]
cash_ws.Range('C2:C32').NumberFormat = 0
cash_ws.Range('C2:C32').Value = nr_polisy
cash_ws.Range('D2:D32').Value = inkaso

cash_ws.Columns.AutoFit()
cash_ws.Columns(1).ColumnWidth = 12
cash_ws.Columns(2).ColumnWidth = 10

wb.DisplayAlerts = False
# cash_ws.DisplayAlerts = False
path_do_zapisu_w = r'C:\Users\ROBERT\Desktop\IT\PYTHON\PYTHON 37 PROJEKTY\księgowość\skrypty osobno'

wb.SaveAs(path_do_zapisu_w + "\\inkaso.xlsx")
wb.Close()


ExcelApp.Application.Quit()
wb.DisplayAlerts = True
