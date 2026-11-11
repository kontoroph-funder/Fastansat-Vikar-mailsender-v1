"""Isolated Excel COM process; never attaches to an already open Excel window."""
import sys
from pathlib import Path


def export(source, sheet_name, output):
    import pythoncom
    import win32com.client
    import win32process
    pythoncom.CoInitialize()
    excel = workbook = None
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        Path(output).with_name('excel.pid').write_text(str(win32process.GetWindowThreadProcessId(excel.Hwnd)[1]))
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        excel.AskToUpdateLinks = False
        workbook = excel.Workbooks.Open(str(Path(source).resolve()), UpdateLinks=0, ReadOnly=True)
        sheet = workbook.Worksheets(sheet_name)
        page = sheet.PageSetup
        page.PrintArea = '$A$1:$K$75'
        page.Orientation, page.PaperSize = 2, 9
        page.Zoom, page.FitToPagesWide, page.FitToPagesTall = False, 1, 2
        page.CenterHorizontally = page.CenterVertically = False
        for name, inches in [('TopMargin', .3), ('BottomMargin', .3), ('LeftMargin', .2), ('RightMargin', .2), ('HeaderMargin', .1), ('FooterMargin', .1)]:
            setattr(page, name, excel.InchesToPoints(inches))
        sheet.ResetAllPageBreaks()
        sheet.HPageBreaks.Add(sheet.Rows(39))
        # Persist only a private sanitized copy, matching the original exporter's behavior.
        workbook.SaveAs(str(Path(output).with_suffix('.xlsx')), FileFormat=51)
        sheet.ExportAsFixedFormat(0, str(Path(output).resolve()))
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        finally:
            if excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()


if __name__ == '__main__':
    export(*sys.argv[1:])
