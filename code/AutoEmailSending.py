# -*- coding: utf-8 -*-
"""
Created on Wed Jul 11 17:22:21 2018

@author: Hwanghokshun
"""

from email.header import Header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import parseaddr, formataddr
import smtplib
 
class MessagebyMail:
    
    def sendmessage():
        
        def _format_addr(s):
            name, addr=parseaddr(s)
            return formataddr((Header(name, 'utf-8').encode(), addr))
        
        from_addr = input('发件人邮箱： ')
        password = input('登陆密码(或授权码)： ')
        to_addr = input('收件人邮箱： ')
        smtp_server = input('SMTP服务器(例：smtp.qq.com)： ')
        msg = MIMEMultipart()
        msg['From'] = _format_addr("检测中心 <%s>" % from_addr)
        msg['To'] = _format_addr("管理员 <%s>" % to_addr)
        msg['Subject'] = Header(input("邮件标题："), 'utf-8').encode()
        msg.attach(MIMEText(input("正文部分： "), 'plain', 'utf-8'))
        mime = MIMEApplication(open('监测报告.pdf','rb').read())
        mime.add_header('Content-Disposition', 'attachment', filename="监测报告.pdf")
        msg.attach(mime)
        server = smtplib.SMTP(smtp_server, 25)
        server.set_debuglevel(1)
        server.login(from_addr, password)
        server.sendmail(from_addr, [to_addr], msg.as_string())
        server.quit()
    
a=MessagebyMail
a.sendmessage()