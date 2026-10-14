import win32com.client, pythoncom, os
try:
    import pymupdf
except ImportError:
    pymupdf = None

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(os.path.abspath('test_wide_rab.xlsx'))
ws = wb.Worksheets(1)

print('Initial VPageBreaks count:', ws.VPageBreaks.Count)
print('Initial Zoom:', ws.PageSetup.Zoom)
print('Initial FitToPagesWide:', ws.PageSetup.FitToPagesWide)
print('Initial PrintTitleColumns:', ws.PageSetup.PrintTitleColumns)

# Test 1: Forced FitToPagesWide = 1
ws.PageSetup.Zoom = False
ws.PageSetup.FitToPagesWide = 1
ws.PageSetup.FitToPagesTall = False
wb.ExportAsFixedFormat(0, os.path.abspath('pdf_fit_width.pdf'))

# Test 2: Natural multi-page layout (Smart Auto / Separate pages)
ws.PageSetup.Zoom = 100
ws.PageSetup.FitToPagesWide = False
ws.PageSetup.FitToPagesTall = False
ws.PageSetup.PrintTitleColumns = "$A:$C"
wb.ExportAsFixedFormat(0, os.path.abspath('pdf_multi_page.pdf'))

wb.Close(SaveChanges=False)
excel.Quit()
pythoncom.CoUninitialize()

if pymupdf:
    doc1 = pymupdf.open('pdf_fit_width.pdf')
    print('PDF Fit Width total pages:', len(doc1))
    doc1.close()

    doc2 = pymupdf.open('pdf_multi_page.pdf')
    print('PDF Multi Page total pages:', len(doc2))
    doc2.close()
