

import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # 可设为 True 观察 Excel 操作

wb = excel.Workbooks.Open(r"C:\Users\dsm\Desktop\20250304.xlsm")  # 修改为实际路径

try:
    result = excel.Application.Run("ThisWorkbook.GetFilesFromFolder", r"E:\文件存储卷_1\My Kindle Content",0)
    print(r"完成 E:\文件存储卷_1\My Kindle Content目录刷新")
finally:
    wb.Close(True)
    excel.Quit()

