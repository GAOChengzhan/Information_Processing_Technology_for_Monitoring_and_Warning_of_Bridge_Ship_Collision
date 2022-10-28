# -*- coding: utf-8 -*-
"""
Created on Mon Aug 20 20:44:33 2018

@author: DELL
"""
import win32com.client
import os
import time
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import RGBColor
import xlrd
import DataBase
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Inches, Pt
class RemoteWord():
    def __init__(self,filename=None):
        self.xlApp=win32com.client.DispatchEx('Word.Application')
        self.xlApp.Visible=0
        self.xlApp.DisplayAlerts=0    #后台运行，不显示，不警告
        if filename:
            self.filename=filename
            if os.path.exists(self.filename):
                self.doc=self.xlApp.Documents.Open(filename)
            else:
                self.doc = self.xlApp.Documents.Add()    #创建新的文档
                self.doc.SaveAs(filename)
        else:
            self.doc=self.xlApp.Documents.Add()
            self.filename=''
    def add_doc_end(self, string):
        '''在文档末尾添加内容'''
        rangee = self.doc.Range()
        rangee.InsertAfter('\n'+string)

    def add_doc_start(self, string):
        '''在文档开头添加内容'''
        rangee = self.doc.Range(0, 0)
        rangee.InsertBefore(string+'\n')

    def insert_doc(self, insertPos, string):
        '''在文档insertPos位置添加内容'''
        rangee = self.doc.Range(0, insertPos)
        if (insertPos == 0):
            rangee.InsertAfter(string)
        else:
            rangee.InsertAfter('\n'+string)
    def copy_pagenumber(self,filename=None):
        self.filename=filename
        self.doc =self.xlApp.Documents.Open('d:\\biaozhun\\biaozhun.docx') 
        self.xlApp.ActiveDocument.Sections[0].Headers[0].Range.Copy()
        self.wc = win32com.client.constants 
        self.doc.Close() 
        self.doc2= self.xlApp.Documents.Open(self.filename) 
        self.xlApp.ActiveDocument.Sections[0].Headers[0].Range.Paste()
        self.xlApp.ActiveDocument.SaveAs(self.filename)
        self.doc2.Close()
        self.doc3 = self.xlApp.Documents.Open('d:\\biaozhun\\biaozhun.docx')
        self.xlApp.ActiveDocument.Sections[0].Footers[0].Range.Copy()
        self.doc3.Close()
        self.doc4= self.xlApp.Documents.Open(self.filename)
        self.xlApp.ActiveDocument.Sections[0].Footers[0].Range.Paste()
        self.xlApp.ActiveDocument.SaveAs(self.filename)
        #设置页眉文字，如果要设置页脚值需要把SeekView由9改为10就可以了。。。
    def add_headers(self,string):
        self.string=string
        self.xlApp.ActiveWindow.ActivePane.View.SeekView = 9 #9 - 页眉； 10 - 页脚
        self.xlApp.Selection.ParagraphFormat.Alignment = 0
        self.xlApp.Selection.Text = self.string
        self.xlApp.ActiveWindow.ActivePane.View.SeekView = 0 # 释放焦点，返回主文档
    def add_footers(self,string):
        self.string=string
        self.xlApp.ActiveWindow.ActivePane.View.SeekView = 10 #9 - 页眉； 10 - 页脚
        self.xlApp.Selection.ParagraphFormat.Alignment = 0
        self.xlApp.Selection.Text = self.string
        self.xlApp.ActiveWindow.ActivePane.View.SeekView = 0 # 释放焦点，返回主文档

    def replace_headers(self,string1,string2):        
        # 页眉文字替换
        self.string1=string1#被替换的文字
        self.string2=string2#新的文字
        self.xlApp.ActiveDocument.Sections[0].Headers[0].Range.Find.ClearFormatting()
        self.xlApp.ActiveDocument.Sections[0].Headers[0].Range.Find.Replacement.ClearFormatting()
        self.xlApp.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute(self.string1, False, False, False, False, False, True, 1, False, self.string2, 2)
     
    def save(self):
        '''保存文档'''
        self.doc.Save()

    def save_as(self, filename):
        '''文档另存为'''
        self.doc.SaveAs(filename)

    def close(self):
        '''保存文件、关闭文件'''
        self.save()
        self.xlApp.Documents.Close()
        self.xlApp.Quit()

#方法样例如下
#if __name__=='__main__':
    
    #path=r"C:\Users\DELL\Desktop\自动报告尝试1\docsample\$.docx"
    #doc = RemoteWord(path)
    #doc.insert_doc(0, '0123456789')
    #doc.add_doc_end('9876543210')
    #doc.add_doc_start('asdfghjklm')
    #doc.insert_doc(10, 'qwertyuiop')
    #doc.copy_pagenumber(path)
    #doc.add_headers('tongji university')
    #doc.add_footers('TJ university')
    #doc.replace_headers("tongji university","TJU")

