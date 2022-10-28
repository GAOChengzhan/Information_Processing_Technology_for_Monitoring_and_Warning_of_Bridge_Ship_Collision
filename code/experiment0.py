# -*- coding: utf-8 -*-
"""
Created on Wed Oct 31 07:51:01 2018

@author: DELL
"""
from package_pydocx import GENERATEWord
from package import  RemoteWord
from DataBase import DataBase
import numpy as np
address = r'C:\Users\DELL\Desktop\自动报告尝试1\docsample\testtest.db'
df = DataBase(address)
tblName = '预警报告模板'
tblName1 = '桥梁信息'
tblName2 = '船只信息'
#values=['2','Mom','30','Guangzhou','30000']
#df.insertData(tblName,values)
result = df.getAllData(tblName)
colname1 = df.getColumns(tblName1)
colname2 = df.getColumns(tblName2)
path=r"C:\Users\DELL\Desktop\自动报告尝试1\docsample\预警报告.docx"
docx = GENERATEWord(path)
docx.add_text_with_style('预警报告')
result3=""
k=1
for i in list(range(0,33)):
    result1=result[i]
    result2=result1[1]
    result21=result2[:7]
    result22=result2[:5]
    result23=result2[:9]
    if(result21=='_fixed_'):
        result3=result3+str(result2[7:])
    if(result22=='_var_'):
        result4=result2[:9]
        if(result4=='_var_桥梁信息'):
            data=df.selectData(colname1[1:2],tblName1,'id=1')
            data=data[0][0]
            result3=result3+str(data)
        if(result4=='_var_船只信息'):
            data=df.selectData(colname2[k:k+1],tblName2,'id=1')
            data=data[0][0]
            if(result1[-1]=='T'):
                docx.add_picture(str(data),5,0)
            else:
                result3=result3+str(data)
            k=k+1
    if(result23=='_newline_'):
        docx.add_text(str(result3))
        result3=""
        docx.save_doc(path)
docx.change_all_word_style()
docx.save_doc(path)
df.closeDataBase()
        



