import win32com.client as win32
from win32com.client import constants
from datetime import datetime
from dateutil.relativedelta import relativedelta
import time


def gen_py():
    try:
        win32.gencache.EnsureDispatch('Excel.Application')
    except AttributeError:
        # Corner case dependencies.
        import os
        import re
        import sys
        import shutil

        # Remove cache and try again.
        MODULE_LIST = [m.__name__ for m in sys.modules.values()]
        for module in MODULE_LIST:
            if re.match(r'win32com\.gen_py\..+', module):
                del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))


def baza():
    path_bazy = r'\\Js\e\Agent baza'
    """Sprawdza czy arkusz jest otwarty. Jeżeli arkusz jest zamknięty, otwiera go."""
    try:
        ExcelApp = win32.GetActiveObject('Excel.Application')
        ExcelApp.DisplayAlerts = False
        wb = ExcelApp.Workbooks("\\2014 BAZA MAGRO.xlsm")
        ws = wb.Worksheets("BAZA 2014")
    except:
        ExcelApp = win32.gencache.EnsureDispatch("Excel.Application")
        wb = ExcelApp.Workbooks.OpenXML(path_bazy + "\\2014 BAZA MAGRO.xlsm")
        ws = wb.Worksheets("BAZA 2014")

    ExcelApp.Visible = True
    col_diff = wb.Worksheets(1).Cells(wb.Worksheets(1).Rows.Count, 2).End(-4162).Row

    return ExcelApp, wb, ws, col_diff


def filtr_tu(tow):
    tu = {'ALL': 'Allianz', 'AXA': 'AXA', 'COM': 'Compensa', 'EIN': 'Euroins', 'EPZU': 'PZU', 'GEN': 'Generali',
          'ŻGEN': 'Generali', 'GOT': 'Gothaer', 'HDI': 'HDI', 'HES': 'Ergo Hestia', 'IGS': 'IGS', 'INT': 'INTER',
          'LIN': 'LINK 4', 'MTU': 'MTU', 'PRO': 'Proama', 'PZU': 'PZU', 'RIS': 'InterRisk', 'TUW': 'TUW', 'TUZ': 'TUZ',
          'UNI': 'Uniqa', 'WAR': 'Warta', 'ŻWAR': 'Warta', 'WIE': 'Wiener', 'YCD': 'You Can Drive', 'None': ''}

    return tu[tow]


def okres(n):
    msc = (datetime.today() + relativedelta(months=n)).strftime('%m')
    msc_rok = (datetime.today() + relativedelta(months=n)).strftime('%m.%Y')
    rok_msc = (datetime.today() + relativedelta(months=n)).strftime('%y_%m')

    return msc, msc_rok, rok_msc


def arkusz_raportu(msc_rok):
    ExcelApp_cash = win32.DispatchEx('Excel.Application')
    ExcelApp_cash.Visible = True
    wb_cash = ExcelApp_cash.Workbooks.Add()
    ws_cash = wb_cash.Worksheets.Add()
    ws_cash.Name = f'Inkaso {msc_rok}r.'

    ws_cash.Cells(1, 1).Value = 'Data'
    ws_cash.Cells(1, 2).Value = 'TU'
    ws_cash.Cells(1, 3).Value = 'Nr polisy'
    ws_cash.Cells(1, 4).Value = 'Kwota inkaso'
    ws_cash.Cells(1, 5).Value = 'Suma inkaso w PLN:'
    ws_cash.Cells(1, 5).Font.Bold = True

    return ExcelApp_cash, wb_cash, ws_cash


def filtry_kolumn(ws, rok_msc):
    ws.Columns(1).AutoFilter(Field=2, Criteria1=f'{rok_msc}')
    ws.Columns(1).AutoFilter(Field=51, Criteria1='G')


def copy_paste_daty(ws, ws_cash):
    ws.Range(f'AD5:AD{ws.UsedRange.Rows.Count}').Copy()
    time.sleep(3)
    ws_cash.Range(f'A2:A{ws.UsedRange.Rows.Count}').NumberFormat = "@"
    ws_cash.Range(f'A2').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)  # 12
    ws_cash.Range(f'A2:A{ws.UsedRange.Rows.Count}').HorizontalAlignment = constants.xlHAlignLeft  # -4131

    time.sleep(.7)


def copy_paste_tu(ws, ws_cash, col_diff):
    ws.Range(f'AL5:AL{ws.UsedRange.Rows.Count}').Copy()
    time.sleep(1)
    ws_cash.Range(f'B2').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)
    none_list = []
    row = 2

    for tow in ws_cash.Range(f'B2:B{ws.UsedRange.Rows.Count - col_diff + 190}'):
        tow = str(tow)
        if none := tow is None:
            none_list.append(none)
            row += 1
            if len(none_list) > 3:
                break
        ws_cash.Cells(row, 2).Value = filtr_tu(tow)
        row += 1


def copy_paste_nr(ws, ws_cash):
    ws.Range(f'AN5:AN{ws.UsedRange.Rows.Count}').Copy()
    time.sleep(1)
    ws_cash.Columns(3).NumberFormat = 0
    ws_cash.Range(f'C2').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)
    ws_cash.Range(f'C2:C{ws.UsedRange.Rows.Count}').HorizontalAlignment = constants.xlHAlignLeft
    time.sleep(.7)


def copy_paste_inkaso(ws, ws_cash, col_diff):
    ws.Range(f'BC5:BC{ws.UsedRange.Rows.Count}').Copy()
    time.sleep(1)
    ws_cash.Range(f'D2').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)

    for i, value in enumerate(ws_cash.Range(f'D2:D{ws.UsedRange.Rows.Count - col_diff}')):
        if str(value) in ('0.0', 'None', None, ''):
            ws_cash.Rows(i + 2).EntireRow.Delete()

    ws_cash.Cells(1, 6).Value = '=SUM(D:D)'
    ws_cash.Cells(1, 6).Font.Size = 15
    ws_cash.Cells(1, 6).Font.Bold = True


def sortowanie(ws, ws_cash, col_diff):
    xlAscending = 1
    xlSortColumns = 1
    ws_cash.Range(f"A2:D{ws.UsedRange.Rows.Count - col_diff}").Sort(Key1=ws_cash.Range("A1"),
                                                                    Order1=xlAscending, Orientation=xlSortColumns)


def auto_fit(ws_cash):
    ws_cash.Columns.AutoFit()
    ws_cash.Columns(1).ColumnWidth = 11
    ws_cash.Columns(2).ColumnWidth = 11


def opcje_zapisu(ExcelApp, ExcelApp_cash, wb, wb_cash, msc_rok, next_month_path):
    path_do_zapisu_w = next_month_path
    wb_cash.DisplayAlerts = False
    ExcelApp.Application.CutCopyMode = False

    wb_cash.SaveAs(path_do_zapisu_w + f"Raport_kasowy_{msc_rok}.xlsx")
    wb.Close(SaveChanges=False)
    wb_cash.Close()
    ExcelApp.Application.Quit()
    ExcelApp_cash.Application.Quit()

    ExcelApp.DisplayAlerts = True
    wb_cash.DisplayAlerts = True


def raport_inkaso(*, za_okres, path):
    gen_py()
    try:
        print('Raport kasowy...')
        ExcelApp, wb, ws, col_diff = baza()

        msc, msc_rok, rok_msc = okres(za_okres)
        ExcelApp_cash, wb_cash, ws_cash = arkusz_raportu(msc_rok)

        filtry_kolumn(ws, rok_msc)
        copy_paste_daty(ws, ws_cash)
        copy_paste_tu(ws, ws_cash, col_diff)
        copy_paste_nr(ws, ws_cash)
        copy_paste_inkaso(ws, ws_cash, col_diff)
        sortowanie(ws, ws_cash, col_diff)
        auto_fit(ws_cash)
        opcje_zapisu(ExcelApp, ExcelApp_cash, wb, wb_cash, msc_rok, path)
        print('Raport kasowy ok')

    except Exception as e:
        with open(rf'{path}brak dokumentów.txt', 'a') as f:
            f.write('Brak raportu kasowego\n')
        print(f'Brak raportu kasowego: {e}')


# next_month_path = f'C:\\Users\\ROBERT\\Desktop\\Księgowość\\' \
#                       f'{(datetime.today() + relativedelta(months=-1)).strftime("%m.%Y")}\\'
# raport_inkaso(za_okres=-1, path=next_month_path)