
import win32com.client

xl =win32com.client.Dispatch("Excel.Application") #实例化Excel应用程序
wb =xl.Workbooks.Open(r'E:\新desktop\python.xlsm')
# xl.Application.Run('pythonrunvba.xlsm!模块1.mymacro("完美Excel")')
xl.Application.Run('python.xlsm!模块1.d2')
wb.Save()
xl.Application.Quit()