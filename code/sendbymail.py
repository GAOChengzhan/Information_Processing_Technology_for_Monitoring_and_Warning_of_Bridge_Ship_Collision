# -*- coding: utf-8 -*-
"""
Created on Sat Feb 16 13:09:02 2019

@author: Hwanghokshun
"""

from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from smtplib import SMTP


def send_email(receivers, topic, content, sender='302630314@qq.com', password='dtvwxlghmypcbifc'):
    # 自己填好相关信息
    for receiver in receivers:
        try:
            msg = MIMEMultipart()
            msg['From'] = Header(sender, 'utf-8')  # 编辑邮件头
            msg['To'] = Header(receiver, 'utf-8')
            msg['Subject'] = Header(topic, 'utf-8')
            msg.attach(MIMEText(content, 'plain', 'utf-8'))  # 把正文附在邮件上

            with open('检测评估报告.pdf', 'rb') as f:
                mime = MIMEBase('file', 'pdf', filename='检测评估报告.pdf') 
                mime.add_header('Content-Disposition', 'attachment', filename='检测评估报告.pdf')
                mime.set_payload(f.read())  # 读取附件内容
                encoders.encode_base64(mime)  # 对附件Base64编码
                msg.attach(mime)  # 把附件附在邮件上
                server = SMTP('smtp.qq.com', 25)
                server.login(sender, password)
                server.sendmail(sender, receiver, msg.as_string())
                print('发送成功！')
        except Exception as error:
            print(error)
            continue


if __name__ == '__main__':
    receiver = ['1812535185@qq.com']
    topic = '邮件测试'
    content = '邮件测试，请勿回复'
    send_email(receiver, topic, content)
