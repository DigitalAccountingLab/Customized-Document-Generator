# -*- coding: utf-8 -*-



from __future__ import print_function 
from mailmerge import MailMerge
import os
import sys
import comtypes.client
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
import pythoncom

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(930, 627)
        Dialog.setMouseTracking(False)
        Dialog.setFocusPolicy(QtCore.Qt.WheelFocus)
        Dialog.setToolTip("")
        Dialog.setLayoutDirection(QtCore.Qt.LeftToRight)
        Dialog.setStyleSheet("")
        self.widget = QtWidgets.QWidget(Dialog)
        self.widget.setGeometry(QtCore.QRect(40, 40, 861, 531))
        self.widget.setMouseTracking(True)
        self.widget.setFocusPolicy(QtCore.Qt.TabFocus)
        self.widget.setAutoFillBackground(False)
        self.widget.setStyleSheet("QWidget #widget {border-image: url(:/Image/image/XJTLU.jpg);}\n"
"QWidget #widget {border-top-right-radius:30px;}\n"
"QWidget #widget {border-bottom-left-radius:30px;}")
        self.widget.setObjectName("widget")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.widget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(620, 90, 121, 152))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.add_word = QtWidgets.QPushButton(self.verticalLayoutWidget)
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.PlaceholderText, brush)
        self.add_word.setPalette(palette)
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(24)
        self.add_word.setFont(font)
        self.add_word.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.add_word.setStyleSheet("#add_word{\n"
"    \n"
"    color: rgb(255, 255, 255);\n"
"    border:3px solid rgb(255,255,255);\n"
"    border-radius:8px;\n"
"}\n"
"#add_word:hover{\n"
"    \n"
"    border-color: rgb(44, 9, 103);\n"
"}\n"
"#add_word:pressed{\n"
"    padding-top:5px;\n"
"    padding-left:5px;\n"
"    \n"
"}")
        self.add_word.setObjectName("add_word")
        self.verticalLayout.addWidget(self.add_word)
        self.add_excel = QtWidgets.QPushButton(self.verticalLayoutWidget)
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.PlaceholderText, brush)
        self.add_excel.setPalette(palette)
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(24)
        self.add_excel.setFont(font)
        self.add_excel.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.add_excel.setStyleSheet("#add_excel{\n"
"    \n"
"    color: rgb(255, 255, 255);\n"
"    border:3px solid rgb(255,255,255);\n"
"    border-radius:8px;\n"
"}\n"
"#add_excel:hover{\n"
"     border-color: rgb(44, 9, 103);\n"
"\n"
"}\n"
"#add_excel:pressed{\n"
"    padding-top:5px;\n"
"    padding-left:5px;\n"
"}")
        self.add_excel.setObjectName("add_excel")
        self.verticalLayout.addWidget(self.add_excel)
        self.choose_output = QtWidgets.QPushButton(self.verticalLayoutWidget)
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.PlaceholderText, brush)
        self.choose_output.setPalette(palette)
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(24)
        self.choose_output.setFont(font)
        self.choose_output.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.choose_output.setFocusPolicy(QtCore.Qt.StrongFocus)
        self.choose_output.setStyleSheet("#choose_output{\n"
"    \n"
"    color: rgb(255, 255, 255);\n"
"    border:3px solid rgb(255,255,255);\n"
"    border-radius:8px;\n"
"}\n"
"#choose_output:hover{\n"
"    border-color: rgb(44, 9, 103);\n"
"\n"
"}\n"
"#choose_output:pressed{\n"
"    padding-top:5px;\n"
"    padding-left:5px;\n"
"}")
        self.choose_output.setObjectName("choose_output")
        self.verticalLayout.addWidget(self.choose_output)
        self.verticalLayoutWidget_2 = QtWidgets.QWidget(self.widget)
        self.verticalLayoutWidget_2.setGeometry(QtCore.QRect(110, 90, 501, 151))
        self.verticalLayoutWidget_2.setObjectName("verticalLayoutWidget_2")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_2)
        self.verticalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.lineEdit_1 = QtWidgets.QLineEdit(self.verticalLayoutWidget_2)
        self.lineEdit_1.setMaximumSize(QtCore.QSize(481, 41))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(16)
        font.setItalic(True)
        self.lineEdit_1.setFont(font)
        self.lineEdit_1.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        self.lineEdit_1.setFocusPolicy(QtCore.Qt.NoFocus)
        self.lineEdit_1.setContextMenuPolicy(QtCore.Qt.DefaultContextMenu)
        self.lineEdit_1.setAcceptDrops(False)
        self.lineEdit_1.setStyleSheet("color:rgb(0,0,0);\n"
"\n"
"border-style:outset;\n"
"\n"
"border-radius:8px;\n"
"")
        self.lineEdit_1.setInputMask("")
        self.lineEdit_1.setFrame(False)
        self.lineEdit_1.setDragEnabled(True)
        self.lineEdit_1.setCursorMoveStyle(QtCore.Qt.LogicalMoveStyle)
        self.lineEdit_1.setClearButtonEnabled(False)
        self.lineEdit_1.setObjectName("lineEdit_1")
        self.verticalLayout_2.addWidget(self.lineEdit_1)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.verticalLayoutWidget_2)
        self.lineEdit_2.setMaximumSize(QtCore.QSize(481, 41))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(16)
        font.setItalic(True)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        self.lineEdit_2.setFocusPolicy(QtCore.Qt.NoFocus)
        self.lineEdit_2.setAcceptDrops(True)
        self.lineEdit_2.setStyleSheet("border-radius:8px;\n"
"border-style:outset;")
        self.lineEdit_2.setFrame(False)
        self.lineEdit_2.setDragEnabled(False)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.verticalLayout_2.addWidget(self.lineEdit_2)
        self.lineEdit_3 = QtWidgets.QLineEdit(self.verticalLayoutWidget_2)
        self.lineEdit_3.setMaximumSize(QtCore.QSize(481, 41))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(16)
        font.setItalic(True)
        self.lineEdit_3.setFont(font)
        self.lineEdit_3.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        self.lineEdit_3.setFocusPolicy(QtCore.Qt.NoFocus)
        self.lineEdit_3.setAcceptDrops(True)
        self.lineEdit_3.setStyleSheet("border-radius:8px;\n"
"")
        self.lineEdit_3.setText("")
        self.lineEdit_3.setFrame(False)
        self.lineEdit_3.setDragEnabled(False)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.verticalLayout_2.addWidget(self.lineEdit_3)
        self.progressBar = QtWidgets.QProgressBar(self.widget)
        self.progressBar.setGeometry(QtCore.QRect(30, 490, 118, 23))
        self.progressBar.setStyleSheet("QProgressBar {\n"
"        background-color: rgb(98,114,164);\n"
"        color: rgb(200,200,200);\n"
"        border-style:none;\n"
"        border-radius: 10px;\n"
"}\n"
"QProgressBar::chunk {\n"
"    background-color: qlineargradient(spread:pad, x1:0, y1:0.511364, x2:1, y2:0.523, stop:0 rgba(0, 0, 0, 255), stop:1 rgba(100, 0, 223, 255));\n"
"        border-radius: 10px;\n"
"}")
        self.progressBar.setProperty("value", 0)
        self.progressBar.setTextVisible(False)
        self.progressBar.setVisible(False)
        self.progressBar.setObjectName("progressBar")
        self.result = QtWidgets.QLabel(self.widget)
        self.result.setGeometry(QtCore.QRect(100, 250, 661, 81))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(17)
        font.setBold(True)
        font.setWeight(75)
        self.result.setFont(font)
        self.result.setAutoFillBackground(False)
        self.result.setStyleSheet("color: rgb(255, 255, 255);\n"
"\n"
"border-radius:8px;\n"
"\n"
"background-color:  rgba(0, 0, 0, 50)\n"
"")
        self.result.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.result.setFrameShadow(QtWidgets.QFrame.Plain)
        self.result.setTextFormat(QtCore.Qt.AutoText)
        self.result.setScaledContents(True)
        self.result.setAlignment(QtCore.Qt.AlignCenter)
        self.result.setWordWrap(True)
        self.result.setObjectName("result")
        self.label_2 = QtWidgets.QLabel(self.widget)
        self.label_2.setGeometry(QtCore.QRect(810, 500, 54, 31))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(16)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setStyleSheet("color: rgb(255,255,255)\n"
"")
        self.label_2.setObjectName("label_2")
        self.heading = QtWidgets.QLabel(self.widget)
        self.heading.setGeometry(QtCore.QRect(0, 0, 861, 41))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(17)
        font.setBold(True)
        font.setWeight(75)
        self.heading.setFont(font)
        self.heading.setAutoFillBackground(False)
        self.heading.setStyleSheet("color: rgb(255, 255, 255);\n"
"border-top-right-radius:30px;\n"
"background-color:  rgba(0, 0, 0, 110)\n"
"")
        self.heading.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.heading.setFrameShadow(QtWidgets.QFrame.Plain)
        self.heading.setTextFormat(QtCore.Qt.AutoText)
        self.heading.setScaledContents(True)
        self.heading.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.heading.setWordWrap(True)
        self.heading.setIndent(10)
        self.heading.setObjectName("heading")
        self.pushButton_1 = QtWidgets.QPushButton(self.widget)
        self.pushButton_1.setGeometry(QtCore.QRect(810, 10, 41, 31))
        self.pushButton_1.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton_1.setStyleSheet("background-color:  rgba(0, 0, 0, 110)\n"
"")
        self.pushButton_1.setText("")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/Image/image/Picture1.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton_1.setIcon(icon)
        self.pushButton_1.setIconSize(QtCore.QSize(20, 20))
        self.pushButton_1.setObjectName("pushButton_1")
        self.pushButton_2 = QtWidgets.QPushButton(self.widget)
        self.pushButton_2.setGeometry(QtCore.QRect(760, 10, 41, 31))
        self.pushButton_2.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton_2.setStyleSheet("background-color:  rgba(0, 0, 0, 110)")
        self.pushButton_2.setText("")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/Image/image/Picture2.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton_2.setIcon(icon1)
        self.pushButton_2.setIconSize(QtCore.QSize(20, 20))
        self.pushButton_2.setObjectName("pushButton_2")
        self.save = QtWidgets.QPushButton(self.widget)
        self.save.setGeometry(QtCore.QRect(260, 420, 161, 51))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.PlaceholderText, brush)
        self.save.setPalette(palette)
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(24)
        self.save.setFont(font)
        self.save.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.save.setStyleSheet("#save{\n"
"    background-color: rgb(0, 0, 0);\n"
"    color: rgb(255, 255, 255);\n"
"    border:3px solid rgb(255,255,255);\n"
"    border-radius:8px;\n"
"}\n"
"#save:hover{\n"
"    background-color: rgb(255, 255, 255);\n"
"    border-color: rgb(44, 9, 103);\n"
"    color: rgb(0, 0, 0);\n"
"}\n"
"#save:pressed{\n"
"    padding-top:5px;\n"
"    padding-left:5px;\n"
"}")
        self.save.setObjectName("save")
        self.quit = QtWidgets.QPushButton(self.widget)
        self.quit.setGeometry(QtCore.QRect(450, 420, 161, 51))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.PlaceholderText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255, 128))
        brush.setStyle(QtCore.Qt.NoBrush)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.PlaceholderText, brush)
        self.quit.setPalette(palette)
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(24)
        self.quit.setFont(font)
        self.quit.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.quit.setStyleSheet("#quit{\n"
"    background-color: rgb(0, 0, 0);\n"
"    color: rgb(255, 255, 255);\n"
"    border:3px solid rgb(255,255,255);\n"
"    border-radius:8px;\n"
"}\n"
"#quit:hover{\n"
"    background-color: rgb(255, 255, 255);\n"
"    border-color: rgb(44, 9, 103);\n"
"    color: rgb(0, 0, 0);\n"
"}\n"
"#quit:pressed{\n"
"    padding-top:5px;\n"
"    padding-left:5px;\n"
"}")
        self.quit.setObjectName("quit")
        self.PDF = QtWidgets.QCheckBox(self.widget)
        self.PDF.setGeometry(QtCore.QRect(620, 350, 104, 41))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(20)
        self.PDF.setFont(font)
        self.PDF.setStyleSheet("color: rgb(255, 255, 255);")
        self.PDF.setAutoExclusive(False)
        self.PDF.setObjectName("PDF")
        self.Docx = QtWidgets.QCheckBox(self.widget)
        self.Docx.setGeometry(QtCore.QRect(510, 350, 101, 41))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(20)
        self.Docx.setFont(font)
        self.Docx.setAcceptDrops(False)
        self.Docx.setStyleSheet("color: rgb(255, 255, 255);")
        self.Docx.setCheckable(True)
        self.Docx.setChecked(False)
        self.Docx.setAutoExclusive(False)
        self.Docx.setObjectName("Docx")
        self.FileName = QtWidgets.QComboBox(self.widget)
        self.FileName.setGeometry(QtCore.QRect(190, 350, 281, 41))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(16)
        self.FileName.setFont(font)
        self.FileName.setStyleSheet("\n"
"border-style:outset;\n"
"\n"
"border-width:1px;\n"
"")
        self.FileName.setObjectName("FileName")
        self.save.raise_()
        self.quit.raise_()
        self.result.raise_()
        self.verticalLayoutWidget.raise_()
        self.verticalLayoutWidget_2.raise_()
        self.progressBar.raise_()
        self.label_2.raise_()
        self.heading.raise_()
        self.pushButton_1.raise_()
        self.pushButton_2.raise_()
        self.PDF.raise_()
        self.Docx.raise_()
        self.FileName.raise_()

        self.retranslateUi(Dialog)
        self.quit.clicked.connect(Dialog.close)
        self.save.clicked.connect(self.lineEdit_1.clear)
        self.save.clicked.connect(self.lineEdit_2.clear)
        self.save.clicked.connect(self.lineEdit_3.clear)
        self.save.clicked.connect(self.progressBar.show)
        self.pushButton_1.clicked.connect(Dialog.close)
        self.pushButton_2.clicked.connect(Dialog.showMinimized)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.add_word.setText(_translate("Dialog", "Browse"))
        self.add_excel.setText(_translate("Dialog", "Browse"))
        self.choose_output.setText(_translate("Dialog", "Select"))
        self.lineEdit_1.setPlaceholderText(_translate("Dialog", " Word Template"))
        self.lineEdit_2.setPlaceholderText(_translate("Dialog", " Excel"))
        self.lineEdit_3.setPlaceholderText(_translate("Dialog", " Output Path"))
        self.result.setText(_translate("Dialog", "Welcome! \
                                       \
                                           Choose the filename field below"))
        self.label_2.setText(_translate("Dialog", "v1.0"))
        self.heading.setText(_translate("Dialog", "Customized Document Generator by DALab "))
        self.save.setText(_translate("Dialog", "Run"))
        self.quit.setText(_translate("Dialog", "Quit"))
        self.PDF.setText(_translate("Dialog", ".pdf"))
        self.Docx.setText(_translate("Dialog", ".docx"))
import resource



from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QComboBox, QCheckBox, QLineEdit,\
      QWidget
from CDG import Ui_Dialog

        
class Mywindow(QMainWindow, Ui_Dialog, QWidget):
    def __init__(self, parent=None):
        super(Mywindow, self).__init__(parent)
        self.setupUi(self)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setWindowTitle("Customized Document Generator v1.0")
        self.add_word.clicked.connect(self.read_word)
        self.add_excel.clicked.connect(self.read_excel)
        self.choose_output.clicked.connect(self.write_folder)    
        self.save.clicked.connect(self.process)


    def read_word(self):
        global word
        word = QFileDialog.getOpenFileName(self,'选择文件','','word files(*.doc , *.docx)')
        self.lineEdit_1.setText(word[0])
        self.progressBar.setVisible(False)
        
    def read_excel(self):
        global excel
        excel = QFileDialog.getOpenFileName(self,'选择文件','','Excel files(*.xlsx , *.xls)')
        self.lineEdit_2.setText(excel[0])
        Info = pd.read_excel(excel[0])
        excel_field = Info.columns
        self.FileName.clear()
        self.FileName.addItems(excel_field)
        
    def write_folder(self):
        global Path
        Path = QFileDialog.getExistingDirectory(self,'选择文件','')
        self.lineEdit_3.setText(Path)
                 

    def process(self):
        QApplication.processEvents()
        try:
            template = word[0]
            Info = pd.read_excel(excel[0])
            foldername = Path

            
        except (NameError, FileNotFoundError, AssertionError):
               self.result.setText("Please check your inputs")
               self.progressBar.setProperty("value",100)

        else:
            pythoncom.CoInitialize()
            Word = comtypes.client.CreateObject('Word.Application') 

            document = MailMerge(template)
    
            word_field = document.get_merge_fields()
            excel_field = Info.columns
            tmp=list(word_field.difference(excel_field))
        
    
            self.progressBar.setRange(0,len(Info)-1)
   
            if len(tmp)==0 and (self.Docx.isChecked() or self.PDF.isChecked()):
                for i in range(len(Info)):
                     QApplication.processEvents();
                     document = MailMerge(template)
                     d={}
                     for j in range(len(word_field)):
                         x = list(word_field)[j]
                         y = Info[x][i]
                         y = str(y)
                         d[x] = y
                     document.merge(**d)
                     Title = self.FileName.currentText()
                     filename = f"{Info[Title][i]}"



                     if self.Docx.isChecked() and self.PDF.isChecked():
                         document.write(foldername+'/'+filename+".docx")
                         document.close()
                         pdfdoc = Word.Documents.Open(os.path.abspath(foldername+"/"+filename+".docx"))     
                         pdfdoc.SaveAs(os.path.abspath(foldername+'/'+filename+'.pdf'),17)
                         pdfdoc.Close() 
                         self.progressBar.setProperty("value",i)
                         self.result.setText("Success!")
                     elif self.Docx.isChecked():
                         document.write(foldername+'/'+filename+".docx")
                         document.close()
                         self.progressBar.setProperty("value",i)
                         self.result.setText("Success!")
                     else:
                         document.write(foldername+'/'+filename+".docx")
                         document.close()
                         pdfdoc = Word.Documents.Open(os.path.abspath(foldername+"/"+filename+".docx"))     
                         pdfdoc.SaveAs(os.path.abspath(foldername+'/'+filename+'.pdf'),17)
                         pdfdoc.Close()
                         os.remove(foldername+'/'+filename+".docx")
                         self.progressBar.setProperty("value",i)
                         self.result.setText("Success!")        
                

            elif len(tmp) != 0:
                self.progressBar.setProperty("value",len(Info)-1)
                self.result.setText('No field(s) called '+'"'+'","'.join(tmp)+'"'+' !')
            else:
                self.progressBar.setProperty("value",len(Info)-1)
                self.result.setText("Please choose the output format")
                
            Word.Quit()
            pythoncom.CoUninitialize()
                
            del globals()['word']
            del globals()['excel']
            del globals()['Path']
            
    def mouseMoveEvent(self, e: QtGui.QMouseEvent):  
        if e.y()<100:
            self._endPos = e.pos() - self._startPos
            self.move(self.pos() + self._endPos)


    def mousePressEvent(self, e: QtGui.QMouseEvent):
            if e.button() == QtCore.Qt.LeftButton:
                self._isTracking = True
                self._startPos = QtCore.QPoint(e.x(), e.y())

    def mouseReleaseEvent(self, e: QtGui.QMouseEvent):
            if e.button() == QtCore.Qt.LeftButton:
                self._isTracking = False
                self._startPos = None
                self._endPos = None
   
if __name__ == '__main__':
    app = QApplication(sys.argv) 
    ui = Mywindow()
    ui.show()
    sys.exit(app.exec_())    
    
