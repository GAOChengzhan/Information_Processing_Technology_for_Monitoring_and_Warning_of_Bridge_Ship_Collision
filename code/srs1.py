# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled1.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(552, 368)
        self.layoutWidget = QtWidgets.QWidget(Dialog)
        self.layoutWidget.setGeometry(QtCore.QRect(30, 50, 481, 261))
        self.layoutWidget.setObjectName("layoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.layoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.GenerateReport = QtWidgets.QPushButton(self.layoutWidget)
        self.GenerateReport.setObjectName("GenerateReport")
        self.gridLayout.addWidget(self.GenerateReport, 0, 0, 1, 1)
        self.ViewReport = QtWidgets.QPushButton(self.layoutWidget)
        self.ViewReport.setObjectName("ViewReport")
        self.gridLayout.addWidget(self.ViewReport, 0, 1, 1, 1)
        self.SendByMail = QtWidgets.QPushButton(self.layoutWidget)
        self.SendByMail.setObjectName("SendByMail")
        self.gridLayout.addWidget(self.SendByMail, 1, 0, 1, 1)
        self.SendByWechat = QtWidgets.QPushButton(self.layoutWidget)
        self.SendByWechat.setObjectName("SendByWechat")
        self.gridLayout.addWidget(self.SendByWechat, 1, 1, 1, 1)
        self.SendByBoth = QtWidgets.QPushButton(self.layoutWidget)
        self.SendByBoth.setObjectName("SendByBoth")
        self.gridLayout.addWidget(self.SendByBoth, 1, 2, 1, 1)
        self.ok = QtWidgets.QDialogButtonBox(self.layoutWidget)
        self.ok.setOrientation(QtCore.Qt.Horizontal)
        self.ok.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Ok)
        self.ok.setObjectName("ok")
        self.gridLayout.addWidget(self.ok, 2, 0, 1, 3)

        self.retranslateUi(Dialog)
        self.ok.accepted.connect(Dialog.accept)
        self.ok.rejected.connect(Dialog.reject)
        self.ok.clicked['QAbstractButton*'].connect(Dialog.close)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.GenerateReport.setText(_translate("Dialog", "生成报告"))
        self.ViewReport.setText(_translate("Dialog", "报告预览"))
        self.SendByMail.setText(_translate("Dialog", "邮件发送"))
        self.SendByWechat.setText(_translate("Dialog", "微信发送"))
        self.SendByBoth.setText(_translate("Dialog", "同步发送"))

