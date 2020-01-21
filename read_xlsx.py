# ！ /usr/bin/python # 第一行是特殊注释行，称之为组织行，用来告诉我们GUN/Linux系统应该使用哪个解释器来执行该程序
# -*- coding: utf-8 -*-
import os
from win32com.client import Dispatch

 # 打开excel应用程序
def openXlsx(filepath):
    # xl = Dispatch('Excel.Application')
    # file = excel.Documents.Open(filepath)
    xl = Dispatch('Excel.Application')
    # 后台运行,不显示
    xl.Visible = True
    # # 不显示提示弹窗
    # xl.DisplayAlerts = 0
    xl.Workbooks.Open(filepath)
    xlBook = xl.Workbooks(1)
    xlSheet = xl.Sheets(1)
    xlSheet.Cells(1,1).Value = 'What shall be the number of thy counting?'
    xlSheet.Cells(2,1).Value = 3
    # print(xlSheet.Cells(1,1).Value)
    # print(xlSheet.Cells(2,1).Value)

local = 'e:\\bigdata\\output.xlsx'
openXlsx(local)