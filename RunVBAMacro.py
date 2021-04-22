# Run .xlsm Macro
import pandas as pd
import shutil
from win32com.client import Dispatch
import datetime

now_time = datetime.datetime.now().strftime('%Y%m%d')
shutil.copy('C:/Users/Desktop/**A**.xlsm','C:/Users/Desktop/old'+now_time+'.xlsm')
xlapp = Dispatch('Excel.Application')# 调用 excel程序
# xlapp.Visible = False  # 如果是True会打开 excel程序（界面）
# xlapp.DisplayAlerts = 0  # 不显示警告信息
excel = xlapp.Workbooks.Open(r'C:/Users/Desktop/**A**.xlsm')
xlapp.Run('MacroName')
excel.Close(True)  #  True关闭该文件并保存;False不保存并关闭工作簿
xlapp.Quit()  # 关闭 excel操作环境
copy = pd.read_excel('C:/Users/Desktop/**A**.xlsm')
shutil.copy('C:/Users/Desktop/**A**.xlsm','C:/Users/cn211183/Desktop/new'+now_time+'.xlsm')
