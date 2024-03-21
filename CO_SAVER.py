from win32com.client import Dispatch

xl = Dispatch("Excel.Application")
wb = xl.Workbooks.Open(r"C:\Users\osukhotin\PycharmProjects\CO_DB\document.xlsm")
xlOpenXMLWorkbookMacroEnabled = 52
filename = 'new_document.xlsm'
wb.SaveAs(rf"C:\Users\osukhotin\PycharmProjects\CO_DB\{filename}", FileFormat=xlOpenXMLWorkbookMacroEnabled)
wb.Close(SaveChanges=False)
xl.Quit()