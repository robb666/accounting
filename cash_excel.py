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
    # workbook = ExcelApp.Workbooks("BAZA 2014.xlsx")

except:
    ExcelApp = Dispatch("Excel.Application")
    wb = ExcelApp.Workbooks.Open(path_bazy + "\\2014 BAZA MAGRO short.xlsx")
    ws = wb.Worksheets("BAZA 2014")
    # wb.DisplayAlerts = False

ExcelApp.Visible = True

tu = {'ALL': 'Allianz', 'AXA': 'AXA', 'COM': 'Compensa', 'EIN': 'Euroins', 'EPZU': 'PZU', 'GEN': 'Generali', 'ŻGEN': 'Generali',
      'GOT': 'Gothaer', 'HDI': 'HDI', 'HES': 'Ergo Hestia', 'IGS': 'IGS', 'INT': 'INTER', 'LIN': 'LINK 4', 'MTU': 'MTU',
      'PRO': 'Proama', 'PZU': 'PZU', 'RIS': 'InterRisk', 'TUW': 'TUW', 'TUZ': 'TUZ', 'UNI': 'Uniqa', 'WAR': 'Warta',
      'ŻWAR': 'Warta', 'WIE': 'Wiener', 'YCD': 'You Can Drive'}



ExcelApp_cash = win32.DispatchEx('Excel.Application')
ExcelApp_cash.Visible = True
wb_cash = ExcelApp_cash.Workbooks.Add()
ws_cash = wb_cash.Worksheets.Add()
m = (datetime.today() + relativedelta(months=-1)).strftime('%m')
ws_cash.Name = f'Gotówka {m}.2021r.'

ws_cash.Cells(1, 1).Value = 'Data'
ws_cash.Cells(1, 2).Value = 'TU'
ws_cash.Cells(1, 3).Value = 'Nr polisy'
ws_cash.Cells(1, 4).Value = 'Kwota inkaso'

print(str((datetime.today() + relativedelta(months=-1)).strftime('%Y-%m-%d')))
print( ExcelApp.Cells(800, 30).Value)

# ws.Columns(3).NumberFormat = "yyyy-MM"

# s = constants.Format(str((datetime.today() + relativedelta(months=-1)).strftime('%Y-%m-%d')), "yyyy-mm-dd")

ws.Columns(1).AutoFilter(Field=2, Criteria1="21_02")
ws.Columns(1).AutoFilter(Field=51, Criteria1='G')


ws.Range(f'AD5:AD{ws.UsedRange.Rows.Count}').Copy()
time.sleep(.6)
ws_cash.Range(f'A2').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)
time.sleep(.6)


t_ub = ws.Range(f'AL5:AL{ws.UsedRange.Rows.Count}')
time.sleep(.6)
j = 2
for tow in t_ub:
    if str(tow) in tu.keys():
        ws_cash.Cells(j, 2).Value = tu[str(tow)]
        j += 1
    ws.Range(f'AL5:AL{ws.UsedRange.Rows.Count}').copy()
    ws_cash.Range(f'B2').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)
time.sleep(.6)


ws.Range(f'AN5:AN{ws.UsedRange.Rows.Count}').Copy()
time.sleep(.6)
ws_cash.Columns(3).NumberFormat = 0
ws_cash.Range(f'C2').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)
time.sleep(.6)


ws.Range(f'BC5:BC{ws.UsedRange.Rows.Count}').Copy()
time.sleep(.6)
ws_cash.Range(f'D2').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)
time.sleep(.6)






ws_cash.Columns.AutoFit()
ws_cash.Columns(1).ColumnWidth = 12
ws_cash.Columns(2).ColumnWidth = 12




path_do_zapisu_w = r'C:\Users\ROBERT\Desktop\IT\PYTHON\PYTHON 37 PROJEKTY\księgowość\skrypty osobno'
wb_cash.DisplayAlerts = False
# wb_cash.SaveAs(path_do_zapisu_w + "\\inkaso.xlsx")
wb.Close(SaveChanges=False)
wb_cash.Close()
ExcelApp.Application.Quit()
ExcelApp_cash.Application.Quit()

wb_cash.DisplayAlerts = True





















# # ws.Columns(1).AutoFilter(Field=51, Criteria1='G')
# column = ws.Range(f'AY1:AY{ws.UsedRange.Rows.Count}')
# i = 1
# j = 2
# # print(column)
#
# # print(ExcelApp.Cells(i, 51).Value)
# for cash in column:
#     print(str(cash))
#     data_wyst = ExcelApp.Cells(i, 30).Value
#     tow_ub = ExcelApp.Cells(i, 38).Value
#     nr_polisy = ExcelApp.Cells(i, 40).Value
#     inkaso = ExcelApp.Cells(i, 55).Value
#
#     i += 1
#     # ExcelApp.Cells(i, 51).Value == 'G'
#     # if str(cash) == 'G' and str(cash) != 'P':
#     # str(cash) == 'G' and inkaso is not None and float(inkaso) > 0 and
#     # and isinstance(data_wyst, datetime)
#     if (str(cash) == 'G' or str(cash) == 'g') and isinstance(data_wyst, datetime) and inkaso is not None \
#             and float(inkaso) > 0 and (datetime.today() + relativedelta(months=-2)).strftime('%m') < \
#                                         datetime.date(data_wyst).strftime('%m') < datetime.today().strftime('%m'):
#
#
#         print(data_wyst.strftime('%Y.%m.%d'), nr_polisy, inkaso, cash)
#         ws_cash.Cells(j, 1).Value = data_wyst.strftime('%Y.%m.%d')
#         ws_cash.Cells(j, 2).Value = tu[tow_ub]
#         ws_cash.Columns(3).NumberFormat = 0
#         ws_cash.Cells(j, 3).Value = nr_polisy
#         ws_cash.Cells(j, 4).Value = inkaso
#         ws_cash.Cells(j, 5).Value = str(cash)
#
#         j += 1
#
#
#
# ws_cash.Columns.AutoFit()
# ws_cash.Columns(1).ColumnWidth = 12
# ws_cash.Columns(2).ColumnWidth = 12
#
# wb_cash.DisplayAlerts = False
# # ws_cash.DisplayAlerts = False
# path_do_zapisu_w = r'C:\Users\ROBERT\Desktop\IT\PYTHON\PYTHON 37 PROJEKTY\księgowość\skrypty osobno'
#
# wb_cash.SaveAs(path_do_zapisu_w + "\\inkaso.xlsx")
# wb.Close(SaveChanges=False)
# wb_cash.Close()
# ExcelApp.Application.Quit()
# ExcelApp_cash.Application.Quit()
#
# wb_cash.DisplayAlerts = True
