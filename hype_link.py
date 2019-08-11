import win32com
from win32com.client import Dispatch


app = win32com.client.Dispatch("Excel.Application")
excel = app.WorkBooks.Open("E:\\Desktop\\1")
sheet = excel.Worksheets(1)
#sheet.active
max_rows = sheet.UsedRange.Count
print(max_rows)
for i in range(1,1001):
    print(i)
    
    sheet.Cells(i, 3).value = '''=concatenate("https:\\",hyperlink(B{}))s'''ss.format(i)
excel.Save()
del app
count.pyimport win32com
from win32com.client import Dispatch
import os


app = win32com.client.Dispatch("excel.application")
excel = app.Workbooks.Open(os.getcwd()+"\\b站麻省理工课程")
sheet = excel.Worksheets(1)
rows = sheet.usedrange.rows.count
count = 0
for i in range(1, rows + 1):
    string = sheet.cells(i, 1).value
    if "计算机" in string or "算法" in string or "leet" in string.lower():
        print(sheet.cells(i, 1).value)
        count += 1
print(count)