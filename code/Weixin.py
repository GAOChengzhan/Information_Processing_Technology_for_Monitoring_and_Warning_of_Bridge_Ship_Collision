# -*- coding: utf-8 -*-
"""
Created on Wed Jul 11 18:39:26 2018

@author: Hwanghokshun
"""

import itchat
# 登录，执行本函数，itchat自动把二维码下载到本地并打开，手机微信扫描即可。
class MessagebyWechat:
    def sendmessage():
        itchat.auto_login()
        #想给谁发信息，先查找到这个朋友,name后填微信备注即可,deepin测试成功
        users = itchat.search_friends(name=input("发送对象的备注名："))
        userName = users[0]['UserName']
        itchat.send(input("发送信息内容："),toUserName =userName)