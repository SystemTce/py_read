# encoding=utf-8
# Author：Tce
# Date:2019年6月2日 16:51:28
# doc文件转docx文件
import sys
import pickle
import re
import codecs
import string
import shutil
from win32com.client import Dispatch
import docx
import os
import read_docx as rd


# 单个doc文件转 docx
def doc2docx(path):
    # 打开word应用程序
    word = Dispatch('Word.Application')
    # 后台运行,不显示
    word.Visible = 0
    # 不显示提示弹窗
    word.DisplayAlerts = 0
    doc = word.Documents.Open(path)
    newpath = os.path.splitext(path)[0]+".docx"
    doc.SaveAs(newpath,12,False,"",True,"",False,False,False,False)
    doc.Close()

    word.Quit()
    # os.remove(path)

# 将目录下所有doc文件转成 docx
def alldoc2docx(path):
    # 获取本地文件路径
    list =findDocFileList(path)
    if len(list)>0 :
        # 打开word应用程序
        word = Dispatch('Word.Application')
        # 后台运行,不显示
        word.Visible = 0
        # 不显示提示弹窗
        word.DisplayAlerts = 0

        for p in list:
            print('--------------------------------------')
            print(p)
            # 重命名
            newpath = os.path.splitext(p)[0]+".docx"
            # 用word程序打开word文件
            doc = word.Documents.Open(p)
            # 另存为
            doc.SaveAs(newpath,12,False,"",True,"",False,False,False,False)
            # 关闭
            doc.Close()
            # 删除旧文件
            os.remove(p)
        # 退出word程序
        word.Quit()
           


# 读取目录下所有文件
def findDocFileList(path): 
    list = []
    for root,dirs,files in os.walk(path):
        # for dir in dirs:
            # print(os.path.join(root,dir))
        # print('--------------------------------------')
        for file in files:
            # 现有文件名 
            localFile = os.path.join(root,file)
           
            # 判断文件名是否存在空格
            if file.find(' ') >= 0:
                # 存在空格，将文件名去除所有空格符
                newFileName = file.replace(' ','')
                newFile = os.path.join(root,newFileName)
                os.rename(localFile,newFile)
                if localFile[-4:] == '.doc':
                    list.append(newFile) 
            else:
                 if localFile[-4:] == '.doc':
                    list.append(localFile)   
             
    print('--------------------',len(list),'------------------')
    return list


# filetest = "e:/test.doc"
# doc2docx(filetest)
# readdoclist(rd.path)  
# alldoc2docx(rd.path)
 