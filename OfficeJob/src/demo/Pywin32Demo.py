import win32com.client

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = -1
myBook = excel.Workbooks.Open(r"C:\Users\m\Desktop\表格.xls")
sheet2 = myBook.Worksheets("sheet2")
sheet2.Range("A1:B5").value = "李丽萍"
sheet2.Columns.Autofit