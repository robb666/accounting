import os
import win32com.client as win32
from win32com.client import Dispatch
from win32com.client import constants
from datetime import datetime
from dateutil.relativedelta import relativedelta
import time

path_bazy = r'C:\Users\ROBERT\Desktop\IT\PYTHON\PYTHON 37 PROJEKTY\księgowość\skrypty osobno'


# """Sprawdza czy arkusz jest otwarty."""
# """Jeżeli arkusz jest zamknięty, otwiera go."""

try:
    ExcelApp = win32.GetActiveObject('Excel.Application')
    wb = ExcelApp.Workbooks("2014 BAZA MAGRO short.xlsx")
    ws = wb.Worksheets("BAZA 2014")

except:
    ExcelApp = Dispatch("Excel.Application")
    wb = ExcelApp.Workbooks.Open(path_bazy + "\\2014 BAZA MAGRO short.xlsx")
    ws = wb.Worksheets("BAZA 2014")

ExcelApp.Visible = True

tu = {'ALL': 'Allianz', 'AXA': 'AXA', 'COM': 'Compensa', 'EIN': 'Euroins', 'EPZU': 'PZU', 'GEN': 'Generali',
      'ŻGEN': 'Generali', 'GOT': 'Gothaer', 'HDI': 'HDI', 'HES': 'Ergo Hestia', 'IGS': 'IGS', 'INT': 'INTER',
      'LIN': 'LINK 4', 'MTU': 'MTU', 'PRO': 'Proama', 'PZU': 'PZU', 'RIS': 'InterRisk', 'TUW': 'TUW', 'TUZ': 'TUZ',
      'UNI': 'Uniqa', 'WAR': 'Warta', 'ŻWAR': 'Warta', 'WIE': 'Wiener', 'YCD': 'You Can Drive', 'None': ''}

m = (datetime.today() + relativedelta(months=-1)).strftime('%m')

ExcelApp_cash = win32.DispatchEx('Excel.Application')
ExcelApp_cash.Visible = True
wb_cash = ExcelApp_cash.Workbooks.Add()
ws_cash = wb_cash.Worksheets.Add()
ws_cash.Name = f'Inkaso {m}.2021r.'

ws_cash.Cells(1, 1).Value = 'Data'
ws_cash.Cells(1, 2).Value = 'TU'
ws_cash.Cells(1, 3).Value = 'Nr polisy'
ws_cash.Cells(1, 4).Value = 'Kwota inkaso'
ws_cash.Cells(1, 5).Value = 'Suma inkaso PLN:'
ws_cash.Cells(1, 5).Font.Bold = True


ws.Columns(1).AutoFilter(Field=2, Criteria1=f'21_{m}')
ws.Columns(1).AutoFilter(Field=51, Criteria1='G')


ws.Range(f'AD5:AD{ws.UsedRange.Rows.Count}').Copy()
time.sleep(.6)
ws_cash.Range(f'A2').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)

ws_cash.Range(f'A2:A300').HorizontalAlignment = constants.xlHAlignLeft
time.sleep(.6)


ws.Range(f'AL5:AL{ws.UsedRange.Rows.Count}').Copy()
time.sleep(.6)
ws_cash.Range(f'B2').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)
col_diff = wb.Worksheets(1).Cells(wb.Worksheets(1).Rows.Count, 2).End(-4162).Row
none_list = []
row = 2
for tow in ws_cash.Range(f'B2:B{ws.UsedRange.Rows.Count - col_diff}'):
    if none := str(tow) is None:
        none_list.append(none)
        row += 1
        if len(none_list) > 3:
            break
    ws_cash.Cells(row, 2).Value = tu[str(tow)]
    row += 1


ws.Range(f'AN5:AN{ws.UsedRange.Rows.Count}').Copy()
time.sleep(.6)
ws_cash.Columns(3).NumberFormat = 0
ws_cash.Range(f'C2').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)
time.sleep(.6)


ws.Range(f'BC5:BC{ws.UsedRange.Rows.Count}').Copy()
time.sleep(.6)
ws_cash.Range(f'D2').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)

for i, value in enumerate(ws_cash.Range(f'D2:D{ws.UsedRange.Rows.Count - col_diff}')):

    if str(value) in ('0.0', 'None', None, ''):
        ws_cash.Rows(i + 2).EntireRow.Delete()


ws_cash.Cells(1, 6).Value = '=SUM(D:D)'
ws_cash.Cells(1, 6).Font.Size = 15
ws_cash.Cells(1, 6).Font.Bold = True

xlAscending = 1
xlSortColumns = 1
ws_cash.Range(f"A2:D{ws.UsedRange.Rows.Count - col_diff}").Sort(Key1=ws_cash.Range("A1"),
                                                                Order1=xlAscending, Orientation=xlSortColumns)

ws_cash.Columns.AutoFit()
ws_cash.Columns(1).ColumnWidth = 11
ws_cash.Columns(2).ColumnWidth = 11

path_do_zapisu_w = r'C:\Users\ROBERT\Desktop\IT\PYTHON\PYTHON 37 PROJEKTY\księgowość\skrypty osobno'
wb_cash.DisplayAlerts = False
ExcelApp.Application.CutCopyMode = False

okres = (datetime.today() + relativedelta(months=-1)).strftime('%m.%Y')
wb_cash.SaveAs(path_do_zapisu_w + f"\\gotówka {okres}.xlsx")
wb.Close(SaveChanges=False)
wb_cash.Close()
ExcelApp.Application.Quit()
ExcelApp_cash.Application.Quit()

wb_cash.DisplayAlerts = True



