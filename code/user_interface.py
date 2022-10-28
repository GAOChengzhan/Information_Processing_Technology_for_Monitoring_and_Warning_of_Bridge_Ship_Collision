# -*- coding: utf-8 -*-
"""
Created on Wed Feb 13 14:14:22 2019

@author: DELL
"""

# -*- coding: utf-8 -*-
"""
Created on Fri Feb  8 15:35:45 2019

@author: Hwanghokshun
"""
import itchat
from sendbymail import send_email

from PyQt5 import QtWidgets
from srs1 import Ui_Dialog

from package_pydocx import GENERATEWord
from package import  RemoteWord
from DataBase import DataBase

import comtypes.client
import os

class mywindow(QtWidgets.QDialog, Ui_Dialog):
    def  __init__ (self):
        super(mywindow, self).__init__()
        self.setupUi(self)
        self.GenerateReport.clicked.connect(self.gereport)
        self.ViewReport.clicked.connect(self.showreport)
        self.SendByWechat.clicked.connect(self.sendbywechat)
        self.SendByMail.clicked.connect(self.sendbymail)
        self.SendByBoth.clicked.connect(self.sendbyboth)
    def gereport(self):
        address = r'C:\Users\DELL\Desktop\自动报告尝试1\docsample\SAMPLE.db'
        df = DataBase(address)
        tblName = '检测评估报告模板'
        tblName1 = '桥梁信息'
        tblName2 = '船只信息'
        tblName3='委托内容'
        tblName4='事故记录'
        #values=['2','Mom','30','Guangzhou','30000']
        #df.insertData(tblName,values)
        result = df.getAllData(tblName)
        colname1 = df.getColumns(tblName1)
        colname2 = df.getColumns(tblName2)
        colname3 = df.getColumns(tblName3)
        colname4 = df.getColumns(tblName4)
        path=r"C:\Users\DELL\Desktop\test.docx"
        docx = GENERATEWord(path)
        docx.add_text_with_style('杭州湾跨海大桥船撞事故',1,0,0,22)
        docx.add_prargraph_with_style('   ',5)
        docx.add_text_with_style('结构检测评估报告',1,1,0,30)
        docx.add_prargraph_with_style('   ',460)
        docx.add_text_with_style('上海同济建设工程质量监测站',1,0,0,8)
        docx.add_text_with_style('二零零六年十月三十日',1,0,0,8)
        docx.add_page_break()#封面
        result3=""
        k=1
        j=1
        l=1
        maxID=df.getMaxID(tblName) 
        maxID1=df.getMaxID(tblName1) 
        #maxID2=df.getMaxID(tblName2) 
        maxID3=df.getMaxID(tblName3) 
        maxID4=df.getMaxID(tblName4) #得到最新的id的数值
        for i in list(range(0,maxID)):
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
                    data=df.selectData(colname1[k:k+1],tblName1,'id='+str(maxID1))
                    data=data[0][0]
                    result3=result3+str(data)
                    k=k+1   
                if(result4=='_var_委托内容'):
                    data=df.selectData(colname3[j:j+1],tblName3,'id='+str(maxID3))
                    data=data[0][0]
                    result3=result3+str(data)
                    j=j+1
                if(result4=='_var_事故记录'):
                    data=df.selectData(colname4[l:l+1],tblName4,'id='+str(maxID4))
                    data=data[0][0]
                    result3=result3+str(data)
                    l=l+1
            if(result23=='_newline_'):
                docx.add_text(str(result3))
                result3=""
                docx.save_doc(path)
            if(result23=='_newline+'):
                docx.add_prargraph_with_style('   ',5)
                docx.add_text_with_style(str(result3),0,1,0,16)
                result3=""
                docx.save_doc(path)
        print(maxID)
        docx.change_all_word_style()
        docx.save_doc(path)
        df.closeDataBase()
        
    def showreport(self):
        wdFormatPDF = 17
        path=r"C:\Users\DELL\Desktop\自动报告尝试1\docsample\test.docx"
        in_file = path
        out_file = r"C:\Users\DELL\Desktop\自动报告尝试1\docsample\test.pdf"
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        os.popen(r'"C:\Program Files (x86)\SmartPDF阅读器\smartpdf.exe" test.pdf')
 
    def sendbywechat(self):
        itchat.auto_login()
        users = itchat.search_friends(name="高程展")
        print(users)
        userName = users[0]['UserName']
        itchat.send("桥梁监测预警报告已发送，请查收",toUserName =userName)
        
    def sendbymail(self):
        if __name__ == '__main__':
            receiver = ['1812535185@qq.com']
            topic = '邮件测试'
            content = '邮件测试，请勿回复'
            send_email(receiver, topic, content)
            
    def sendbyboth(self):
        itchat.auto_login()
        users = itchat.search_friends(name="黄学纯")
        print(users)
        userName = users[0]['UserName']
        itchat.send("桥梁监测预警报告已发送，请查收",toUserName =userName)
        if __name__ == '__main__':
            receiver = ['1812535185@qq.com']
            topic = '邮件测试'
            content = '邮件测试，请勿回复'
            send_email(receiver, topic, content) 
            
if __name__=="__main__":
    import sys
    app=QtWidgets.QApplication(sys.argv)
    ui = mywindow()    
    ui.show()
    sys.exit(app.exec_())





