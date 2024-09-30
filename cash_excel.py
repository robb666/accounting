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
          'UNI': 'Uniqa', 'WAR': 'Warta', 'ŻWAR': 'Warta', 'WIE': 'Wiener', 'YCD': 'You Can Drive', 'TRA': 'Trasti',
          'WEF': 'Wefox', 'BAL': 'Balcia', 'None': ''}

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

    ws_cash.Cells(1, 4).Value = f'Inkaso {msc_rok}r.'
    ws_cash.Cells(2, 1).Value = 'MAGRO UBEZPIECZENIA SP. Z O.O.'
    ws_cash.Cells(2, 1).Font.Bold = True
    ws_cash.Cells(3, 1).Value = '90-441 Łódź, Al. Kościuszki 123/307'
    ws_cash.Cells(4, 1).Value = 'NIP 7252160008'

    ws_cash.Cells(20, 1).Value = 'Data'
    ws_cash.Cells(20, 1).Font.Bold = True
    ws_cash.Cells(20, 2).Value = 'TU'
    ws_cash.Cells(20, 2).Font.Bold = True
    ws_cash.Cells(20, 3).Value = 'Nr polisy'
    ws_cash.Cells(20, 3).Font.Bold = True
    ws_cash.Cells(20, 4).Value = 'Kwota inkaso'
    ws_cash.Cells(20, 4).Font.Bold = True
    ws_cash.Cells(19, 3).Value = 'Razem inkaso:'
    ws_cash.Cells(19, 3).Font.Bold = True

    return ExcelApp_cash, wb_cash, ws_cash


def summary(ws_cash, start_row, end_row):
    insurers = {}

    for row in range(start_row, end_row + 1):
        insurer = ws_cash.Cells(row, 2).Value
        amount = ws_cash.Cells(row, 4).Value

        if insurer and amount is not None and isinstance(amount, (int, float)):
            if insurer in insurers:
                insurers[insurer] += amount
            else:
                insurers[insurer] = amount

    summary_start_row = 5 + 2  # Leave one row as a gap
    total_sum = 0

    for idx, (insurer, sum_value) in enumerate(insurers.items()):
        ws_cash.Cells(summary_start_row + idx, 3).Value = insurer
        ws_cash.Cells(summary_start_row + idx, 4).Value = sum_value
        ws_cash.Cells(summary_start_row + idx, 4).Font.Bold = False
        total_sum += sum_value

    # Write the grand total at the end
    grand_total_row = summary_start_row + len(insurers)
    ws_cash.Cells(grand_total_row - 1, 3).Borders(9).Weight = 2
    ws_cash.Cells(grand_total_row - 1, 4).Borders(9).Weight = 2
    ws_cash.Cells(grand_total_row, 3).Value = "Razem"
    ws_cash.Cells(grand_total_row, 3).Font.Bold = True
    ws_cash.Cells(grand_total_row, 4).Value = total_sum
    ws_cash.Cells(grand_total_row, 4).Font.Size = 12
    ws_cash.Cells(grand_total_row, 4).Font.Bold = True

    # Format the summary rows for clarity
    # ws_cash.Cells(grand_total_row, 4).NumberFormat = "#,##0.00"


def filtry_kolumn(ws, rok_msc):
    ws.Columns(1).AutoFilter(Field=2, Criteria1=f'{rok_msc}')
    ws.Columns(1).AutoFilter(Field=51, Criteria1='G')


def copy_paste_daty(ws, start_row, ws_cash):
    ws.Range(f'AD5:AD{ws.UsedRange.Rows.Count}').Copy()
    time.sleep(3)
    ws_cash.Range(f'A{start_row}:A{ws.UsedRange.Rows.Count}').NumberFormat = "@"
    ws_cash.Range(f'A{start_row}').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)  # 12
    ws_cash.Range(f'A{start_row}:A{ws.UsedRange.Rows.Count}').HorizontalAlignment = constants.xlHAlignLeft  # -4131

    time.sleep(.7)


def copy_paste_tu(ws, ws_cash, start_row, col_diff):
    ws.Range(f'AL5:AL{ws.UsedRange.Rows.Count}').Copy()
    time.sleep(1)
    ws_cash.Range(f'B{start_row}').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)
    none_list = []
    row = 21

    for tow in ws_cash.Range(f'B{start_row}:B{col_diff}'):
        tow = str(tow)
        if none := tow is None:
            none_list.append(none)
            row += 1
            if len(none_list) > 3:
                break
        ws_cash.Cells(row, 2).Value = filtr_tu(tow)
        row += 1


def copy_paste_nr(ws, ws_cash, start_row):
    ws.Range(f'AN5:AN{ws.UsedRange.Rows.Count}').Copy()
    time.sleep(1)
    ws_cash.Columns(3).NumberFormat = 0
    ws_cash.Range(f'C{start_row}').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)
    ws_cash.Range(f'C{start_row}:C{ws.UsedRange.Rows.Count}').HorizontalAlignment = constants.xlHAlignLeft
    time.sleep(.7)


def copy_paste_inkaso(ws, ws_cash, start_row, col_diff):
    ws.Range(f'BC5:BC{ws.UsedRange.Rows.Count}').Copy()
    time.sleep(1)
    ws_cash.Range(f'D{start_row}').PasteSpecial(Paste=constants.xlPasteValuesAndNumberFormats)

    for i, value in enumerate(ws_cash.Range(f'D{start_row}:D{ws.UsedRange.Rows.Count - col_diff}')):
        if str(value) in ('0.0', 'None', None, ''):
            ws_cash.Rows(i + 2).EntireRow.Delete()

    ws_cash.Cells(19, 4).Value = f'=SUM(D{start_row}:D2000)'
    ws_cash.Cells(19, 4).Font.Size = 15
    ws_cash.Cells(19, 4).Font.Bold = True


def sortowanie(ws, ws_cash, start_row, col_diff):
    xlAscending = 1
    xlSortColumns = 1
    ws_cash.Range(f"A{start_row}:A{ws_cash.UsedRange.Rows.Count}").Sort(Key1=ws_cash.Range("A1"),
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
    wb_cash.SaveAs(path_do_zapisu_w + f"Raport_kasowy_{msc_rok}.pdf", FileFormat=57)
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
        start_row = 21
        # end_row = ws_cash.Cells(ws_cash.Rows.Count, 2).End(-4162).Row  # Dynamically find the last row
        end_row = 300

        filtry_kolumn(ws, rok_msc)
        copy_paste_daty(ws, start_row, ws_cash)
        copy_paste_tu(ws, ws_cash, start_row, end_row)
        copy_paste_nr(ws, ws_cash, start_row)
        copy_paste_inkaso(ws, ws_cash, start_row, col_diff)
        sortowanie(ws, ws_cash, start_row, col_diff)
        summary(ws_cash, start_row, end_row)
        auto_fit(ws_cash)
        time.sleep(1)
        opcje_zapisu(ExcelApp, ExcelApp_cash, wb, wb_cash, msc_rok, path)
        print('Raport kasowy ok')

    except Exception as e:
        with open(rf'{path}brak dokumentów.txt', 'a') as f:
            f.write('Brak raportu kasowego\n')
        print(f'Brak raportu kasowego: {e}')


next_month_path = f'C:\\Users\\PipBoy3000\\Desktop\\'

raport_inkaso(za_okres=-1, path=next_month_path)