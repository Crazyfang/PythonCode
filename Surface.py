# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Surface.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setEnabled(True)
        MainWindow.resize(482, 346)
        MainWindow.setMinimumSize(QtCore.QSize(482, 346))
        MainWindow.setMaximumSize(QtCore.QSize(482, 346))
        MainWindow.setIconSize(QtCore.QSize(16, 16))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.groupBox_Operate = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_Operate.setGeometry(QtCore.QRect(10, 10, 461, 141))
        self.groupBox_Operate.setObjectName("groupBox_Operate")
        self.label_SelectZipFile = QtWidgets.QLabel(self.groupBox_Operate)
        self.label_SelectZipFile.setGeometry(QtCore.QRect(10, 30, 81, 21))
        self.label_SelectZipFile.setObjectName("label_SelectZipFile")
        self.lineEdit_SelectZipFile = QtWidgets.QLineEdit(self.groupBox_Operate)
        self.lineEdit_SelectZipFile.setGeometry(QtCore.QRect(110, 30, 241, 20))
        self.lineEdit_SelectZipFile.setObjectName("lineEdit_SelectZipFile")
        self.Button_SelectZipFile = QtWidgets.QPushButton(self.groupBox_Operate)
        self.Button_SelectZipFile.setGeometry(QtCore.QRect(370, 30, 75, 23))
        self.Button_SelectZipFile.setObjectName("Button_SelectZipFile")
        self.label_SelectExcelFile = QtWidgets.QLabel(self.groupBox_Operate)
        self.label_SelectExcelFile.setGeometry(QtCore.QRect(10, 70, 81, 21))
        self.label_SelectExcelFile.setObjectName("label_SelectExcelFile")
        self.lineEdit_SelectExcelFile = QtWidgets.QLineEdit(self.groupBox_Operate)
        self.lineEdit_SelectExcelFile.setGeometry(QtCore.QRect(110, 70, 241, 20))
        self.lineEdit_SelectExcelFile.setObjectName("lineEdit_SelectExcelFile")
        self.Button_SelectExcelFile = QtWidgets.QPushButton(self.groupBox_Operate)
        self.Button_SelectExcelFile.setGeometry(QtCore.QRect(370, 70, 75, 23))
        self.Button_SelectExcelFile.setObjectName("Button_SelectExcelFile")
        self.progressBar_Progress = QtWidgets.QProgressBar(self.groupBox_Operate)
        self.progressBar_Progress.setGeometry(QtCore.QRect(110, 110, 241, 23))
        self.progressBar_Progress.setProperty("value", 24)
        self.progressBar_Progress.setObjectName("progressBar_Progress")
        self.label_Progress = QtWidgets.QLabel(self.groupBox_Operate)
        self.label_Progress.setGeometry(QtCore.QRect(10, 110, 81, 21))
        self.label_Progress.setObjectName("label_Progress")
        self.Button_Start = QtWidgets.QPushButton(self.groupBox_Operate)
        self.Button_Start.setGeometry(QtCore.QRect(370, 110, 75, 23))
        self.Button_Start.setObjectName("Button_Start")
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setGeometry(QtCore.QRect(10, 160, 461, 16))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.groupBox_Info = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_Info.setGeometry(QtCore.QRect(10, 180, 461, 141))
        self.groupBox_Info.setObjectName("groupBox_Info")
        self.listView_Info = QtWidgets.QListView(self.groupBox_Info)
        self.listView_Info.setGeometry(QtCore.QRect(10, 20, 441, 111))
        self.listView_Info.setObjectName("listView_Info")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "二维码转换"))
        self.groupBox_Operate.setTitle(_translate("MainWindow", "操作"))
        self.label_SelectZipFile.setText(_translate("MainWindow", "选择压缩文件："))
        self.Button_SelectZipFile.setText(_translate("MainWindow", "选择文件"))
        self.label_SelectExcelFile.setText(_translate("MainWindow", "选择名称文件："))
        self.Button_SelectExcelFile.setText(_translate("MainWindow", "选择文件"))
        self.label_Progress.setText(_translate("MainWindow", "    处理进度："))
        self.Button_Start.setText(_translate("MainWindow", "开始处理"))
        self.groupBox_Info.setTitle(_translate("MainWindow", "处理信息"))

