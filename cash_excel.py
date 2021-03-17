import os
import win32com.client
from win32com.client import Dispatch



"""Sprawdza czy arkusz jest otwarty."""
"""Jeżeli arkusz jest zamknięty, otwiera go."""
try:
    ExcelApp = win32com.client.GetActiveObject('Excel.Application')
    wb = ExcelApp.Workbooks("2014 BAZA MAGRO.xlsx")
    ws = wb.Worksheets("BAZA 2014")
    # workbook = ExcelApp.Workbooks("Baza.xlsx")

except:
    ExcelApp = Dispatch("Excel.Application")
    wb = ExcelApp.Workbooks.Open(path + "\\2014 BAZA MAGRO.xlsx")
    ws = wb.Worksheets("BAZA 2014")


    """Rozpoznaje kolejny wiersz, który może zapisać."""
    row_to_write = wb.Worksheets(1).Cells(wb.Worksheets(1).Rows.Count, 30).End(-4162).Row + 1

    # Rok_przypisu = ExcelApp.Cells(row_to_write, 1).Value =
    Rozlicz = ExcelApp.Cells(row_to_write, 7).Value =
    Podpis = ExcelApp.Cells(row_to_write, 10).Value =
    FIRMA = ExcelApp.Cells(row_to_write, 11).Value =
    Nazwisko = ExcelApp.Cells(row_to_write, 12).Value =
    Imie = ExcelApp.Cells(row_to_write, 13).Value =
    Pesel_Regon = ExcelApp.Cells(row_to_write, 14).Value =
    ExcelApp.Cells(row_to_write, 15).Value =
    ExcelApp.Cells(row_to_write, 16).Value =
    ExcelApp.Cells(row_to_write, 17).Value =
    ExcelApp.Cells(row_to_write, 18).Value =
    ExcelApp.Cells(row_to_write, 19).Value =
    ExcelApp.Cells(row_to_write, 20).Value =
    ExcelApp.Cells(row_to_write, 23).Value =
    ExcelApp.Cells(row_to_write, 24).Value =
    ExcelApp.Cells(row_to_write, 25).Value =
    ExcelApp.Cells(row_to_write, 26).Value =
    # ExcelApp.Cells(row_to_write, 29).Value =
    # ExcelApp.Cells(row_to_write, 30).NumberFormat =
    ExcelApp.Cells(row_to_write, 30).Value =
    # ExcelApp.Cells(row_to_write, 31).Value =
    ExcelApp.Cells(row_to_write, 32).Value =
    ExcelApp.Cells(row_to_write, 36).Value =
    tor = ExcelApp.Cells(row_to_write, 37).Value =
    ExcelApp.Cells(row_to_write, 38).Value =
    # ExcelApp.Cells(row_to_write, 39).Value =
    ExcelApp.Cells(row_to_write, 40).Value =
    # ExcelApp.Cells(row_to_write, 41).Value =
    # ExcelApp.Cells(row_to_write, 42).Value =
    # if wzn_idx:
    #     ExcelApp.Cells(row_to_write, 41).Value =
    #     ExcelApp.Cells(row_to_write, 42).Value =
    # else:
    #     ExcelApp.Cells(row_to_write, 41).Value =
    #     ExcelApp.Cells(row_to_write, 42).Value =
    # ryzyko = ExcelApp.Cells(row_to_write, 46).Value =
    ExcelApp.Cells(row_to_write, 48).Value =
    ExcelApp.Cells(row_to_write, 49).Value =
    # if I_rata_data:
    #     ExcelApp.Cells(row_to_write, 49).Value =
    if rata_I:
        ExcelApp.Cells(row_to_write, 50).Value =
    else:
        ExcelApp.Cells(row_to_write, 50).Value =
    ExcelApp.Cells(row_to_write, 51).Value =
    ExcelApp.Cells(row_to_write, 52).Value =
    ExcelApp.Cells(row_to_write, 53).Value =
    data_inkasa = ExcelApp.Cells(row_to_write, 54).Value =
    if rata_I:
        ExcelApp.Cells(row_to_write, 55).Value =
    else:
        ExcelApp.Cells(row_to_write, 55).Value =
    ExcelApp.Cells(row_to_write, 60).Value =



"""Opcje zapisania"""
ExcelApp.DisplayAlerts = False
wb.SaveAs("\\cash.xlsx")
wb.Close()
ExcelApp.DisplayAlerts = True
