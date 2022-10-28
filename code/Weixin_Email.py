# -*- coding: utf-8 -*-
"""
Created on Fri Feb  8 15:35:45 2019

@author: Hwanghokshun
"""

from PyQt5 import QtWidgets
from srs1 import Ui_Dialog
import itchat
from sendbymail import send_email


class mywindow(QtWidgets.QDialog, Ui_Dialog):
    def  __init__ (self):
        super(mywindow, self).__init__()
        self.setupUi(self)
        self.SendByWechat.clicked.connect(self.sendbywechat)
        self.SendByMail.clicked.connect(self.sendbymail)
        self.SendByBoth.clicked.connect(self.sendbyboth)
        
    def sendbywechat(self):
        itchat.auto_login()
        users = itchat.search_friends(name="高程展")
        print(users)
        userName = users[0]['UserName']
        itchat.send("桥梁监测预警报告",toUserName =userName)
        
    def sendbymail(self):
        if __name__ == '__main__':
            receiver = ['302630314@qq.com']
            topic = '邮件测试'
            content = '邮件测试，请勿回复'
            send_email(receiver, topic, content)
            
    def sendbyboth(self):
        itchat.auto_login()
        users = itchat.search_friends(name="高程展")
        print(users)
        userName = users[0]['UserName']
        itchat.send("桥梁监测预警报告",toUserName =userName)
        if __name__ == '__main__':
            receiver = ['302630314@qq.com']
            topic = '邮件测试'
            content = '邮件测试，请勿回复'
            send_email(receiver, topic, content)

if __name__=="__main__":
    import sys
    app=QtWidgets.QApplication(sys.argv)
    ui = mywindow()    
    ui.show()
    sys.exit(app.exec_())
