# !/usr/bin/env python
# -*- coding:utf-8 -*-
import os
import win32com
from win32com.client import Dispatch
import read_docx as rd

 
# 处理Word文档的类
class RemoteWord:
  def __init__(self, filename=None):
    # 此处使用的是Dispatch，原文中使用的DispatchEx会报错
    self.xlApp = win32com.client.Dispatch('Word.Application')  
    # 后台运行，不显示
    self.xlApp.Visible = 0 
    #不警告
    self.xlApp.DisplayAlerts = 0 
    if filename:
      self.filename = filename
      if os.path.exists(self.filename):
        self.doc = self.xlApp.Documents.Open(filename)
      else:
        # 创建新的文档
        self.doc = self.xlApp.Documents.Add()   
        self.doc.SaveAs(filename)
    else:
      self.doc = self.xlApp.Documents.Add()
      self.filename = ''
 
  def add_doc_end(self, string):
    '''在文档末尾添加内容'''
    rangee = self.doc.Range()
    rangee.InsertAfter('\n' + string)
 
  def add_doc_start(self, string):
    '''在文档开头添加内容'''
    rangee = self.doc.Range(0, 0)
    rangee.InsertBefore(string + '\n')
 
  def insert_doc(self, insertPos, string):
    '''在文档insertPos位置添加内容'''
    rangee = self.doc.Range(0, insertPos)
    if (insertPos == 0):
      rangee.InsertAfter(string)
    else:
      rangee.InsertAfter('\n' + string)
 
  def replace_doc(self, string, new_string):
    '''替换文字'''
    self.xlApp.Selection.Find.ClearFormatting()
    self.xlApp.Selection.Find.Replacement.ClearFormatting()
    #(string--搜索文本,
    # True--区分大小写,
    # True--完全匹配的单词，并非单词中的部分（全字匹配）,
    # True--使用通配符,
    # True--同音,
    # True--查找单词的各种形式,
    # True--向文档尾部搜索,
    # 1,
    # True--带格式的文本,
    # new_string--替换文本,
    # 2--替换个数（全部替换）
    self.xlApp.Selection.Find.Execute(string, False, False, False, False, False, True, 1, True, new_string, 2)
 
  def replace_docs(self, string, new_string):
    '''采用通配符匹配替换'''
    self.xlApp.Selection.Find.ClearFormatting()
    self.xlApp.Selection.Find.Replacement.ClearFormatting()
    self.xlApp.Selection.Find.Execute(string, False, False, True, False, False, False, 1, False, new_string, 2)
  def save(self):
    '''保存文档'''
    self.doc.Save()
 
  def save_as(self, filename):
    '''文档另存为'''
    self.doc.SaveAs(filename,2)
 
  def close(self):
    '''保存文件、关闭文件'''
    self.save()
    self.xlApp.Documents.Close()
    self.xlApp.Quit()
 
 
if __name__ == '__main__':
    out = "e:\\bigdata\out.txt"
    # path = 'E:\\XXX.docx'
    error1  = rd.path3+'01中国共产党基层组织选举工作暂行条例题库（单选10）.docx'
    doc = RemoteWord(error1)   # 初始化一个doc对象
    # 这里演示替换内容，其他功能自己按照上面类的功能按需使用
    
    doc.replace_doc(' ', '')   # 替换文本内容
    doc.replace_doc('．', '.')  # 替换．为.
    doc.replace_doc('\n', '')   # 去除空行
    doc.replace_doc('(','（')   
    doc.replace_doc(')','）')   
    # doc.replace_docs('([0-9])@[、,，]([0-9])@', '\1.\2')   使用@不能识别改用{1,}，\需要使用反斜杠转义
    #   doc.replace_docs('([0-9]){1,}[、,，．]([0-9]){1,}', '\\1.\\2')   # 将数字中间的，,、．替换成.
    #   doc.replace_docs('([0-9]){1,}[旧]([0-9]){1,}', '\\101\\2')    # 将数字中间的“旧”替换成“01”
    
    doc.save_as(out)
    
    doc.close()
  







