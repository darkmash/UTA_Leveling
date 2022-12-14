from datetime import datetime
import glob
import logging
from logging.handlers import RotatingFileHandler
import os
import re
import math
import time
import numpy as np
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import Qt, QCoreApplication
from PyQt5.QtGui import QDoubleValidator, QStandardItemModel, QIcon, QStandardItem, QIntValidator, QFont
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QProgressBar, QPlainTextEdit, QWidget, QGridLayout, QGroupBox, QLineEdit, QSizePolicy, QToolButton, QLabel, QFrame, QListView, QMenuBar, QStatusBar, QPushButton, QApplication, QCalendarWidget, QVBoxLayout, QFileDialog, QCheckBox
from PyQt5.QtCore import pyqtSlot, pyqtSignal, QObject, QThread, QRect, QSize, QDate
import pandas as pd


class Worker(QObject):
    progressChanged = pyqtSignal(int)

    def work(self):
        progressbar_value = 0
        while progressbar_value < 100:
            self.progressChanged.emit(progressbar_value)
            time.sleep(0.1)


class CustomFormatter(logging.Formatter):
    FORMATS = {
        logging.ERROR: ('[%(asctime)s] %(levelname)s:%(message)s', 'white'),
        logging.DEBUG: ('[%(asctime)s] %(levelname)s:%(message)s', 'white'),
        logging.INFO: ('[%(asctime)s] %(levelname)s:%(message)s', 'white'),
        logging.WARNING: ('[%(asctime)s] %(levelname)s:%(message)s', 'yellow')
    }

    def format(self, record):
        last_fmt = self._style._fmt
        opt = CustomFormatter.FORMATS.get(record.levelno)
        if opt:
            fmt, color = opt
            self._style._fmt = "<font color=\"{}\">{}</font>".format(QtGui.QColor(color).name(), fmt)
        res = logging.Formatter.format(self, record)
        self._style._fmt = last_fmt
        return res


class QPlainTextEditLogger(logging.Handler):
    def __init__(self, parent=None):
        super().__init__()
        self.widget = QPlainTextEdit(parent)
        self.widget.setReadOnly(True)
        self.widget.setGeometry(QRect(10, 260, 661, 161))
        self.widget.setStyleSheet('background-color: rgb(53, 53, 53);\ncolor: rgb(255, 255, 255);')
        self.widget.setObjectName('logBrowser')
        font = QFont()
        font.setFamily('Nanum Gothic')
        font.setBold(False)
        font.setPointSize(9)
        self.widget.setFont(font)

    def emit(self, record):
        msg = self.format(record)
        self.widget.appendHtml(msg)
        scrollbar = self.widget.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())


class CalendarWindow(QWidget):
    submitClicked = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        cal = QCalendarWidget(self)
        cal.setGridVisible(True)
        cal.clicked[QDate].connect(self.showDate)
        self.lb = QLabel(self)
        date = cal.selectedDate()
        self.lb.setText(date.toString("yyyy-MM-dd"))
        vbox = QVBoxLayout()
        vbox.addWidget(cal)
        vbox.addWidget(self.lb)
        self.submitBtn = QToolButton(self)
        sizePolicy = QSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        self.submitBtn.setSizePolicy(sizePolicy)
        self.submitBtn.setMinimumSize(QSize(0, 35))
        self.submitBtn.setStyleSheet('background-color: rgb(63, 63, 63);\ncolor: rgb(255, 255, 255);')
        self.submitBtn.setObjectName('submitBtn')
        self.submitBtn.setText('??????????????? ??????')
        self.submitBtn.clicked.connect(self.confirm)
        vbox.addWidget(self.submitBtn)

        self.setLayout(vbox)
        self.setWindowTitle('?????????')
        self.setGeometry(500, 500, 500, 400)
        self.show()

    def showDate(self, date):
        self.lb.setText(date.toString("yyyy-MM-dd"))

    @pyqtSlot()
    def confirm(self):
        self.submitClicked.emit(self.lb.text())
        self.close()


class UISubWindow(QMainWindow):
    submitClicked = pyqtSignal(list)
    status = ''

    def __init__(self):
        super().__init__()
        self.setupUi()

    def setupUi(self):
        self.setObjectName('SubWindow')
        self.resize(600, 600)
        self.setStyleSheet('background-color: rgb(252, 252, 252);')
        self.centralwidget = QWidget(self)
        self.centralwidget.setObjectName('centralwidget')
        self.gridLayout2 = QGridLayout(self.centralwidget)
        self.gridLayout2.setObjectName('gridLayout2')
        self.gridLayout = QGridLayout()
        self.gridLayout.setObjectName('gridLayout')
        self.groupBox = QGroupBox(self.centralwidget)
        self.groupBox.setTitle('')
        self.groupBox.setObjectName('groupBox')
        self.gridLayout4 = QGridLayout(self.groupBox)
        self.gridLayout4.setObjectName('gridLayout4')
        self.gridLayout3 = QGridLayout()
        self.gridLayout3.setObjectName('gridLayout3')
        self.linkageInput = QLineEdit(self.groupBox)
        self.linkageInput.setMinimumSize(QSize(0, 25))
        self.linkageInput.setObjectName('linkageInput')
        self.linkageInput.setValidator(QDoubleValidator(self))
        self.gridLayout3.addWidget(self.linkageInput, 0, 1, 1, 3)
        self.linkageInputBtn = QPushButton(self.groupBox)
        self.linkageInputBtn.setMinimumSize(QSize(0, 25))
        self.gridLayout3.addWidget(self.linkageInputBtn, 0, 4, 1, 2)
        self.linkageAddExcelBtn = QPushButton(self.groupBox)
        self.linkageAddExcelBtn.setMinimumSize(QSize(0, 25))
        self.gridLayout3.addWidget(self.linkageAddExcelBtn, 0, 6, 1, 2)
        self.mscodeInput = QLineEdit(self.groupBox)
        self.mscodeInput.setMinimumSize(QSize(0, 25))
        self.mscodeInput.setObjectName('mscodeInput')
        self.mscodeInputBtn = QPushButton(self.groupBox)
        self.mscodeInputBtn.setMinimumSize(QSize(0, 25))
        self.gridLayout3.addWidget(self.mscodeInput, 1, 1, 1, 3)
        self.gridLayout3.addWidget(self.mscodeInputBtn, 1, 4, 1, 2)
        self.mscodeAddExcelBtn = QPushButton(self.groupBox)
        self.mscodeAddExcelBtn.setMinimumSize(QSize(0, 25))
        self.gridLayout3.addWidget(self.mscodeAddExcelBtn, 1, 6, 1, 2)
        sizePolicy = QSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        self.submitBtn = QToolButton(self.groupBox)
        sizePolicy.setHeightForWidth(self.submitBtn.sizePolicy().hasHeightForWidth())
        self.submitBtn.setSizePolicy(sizePolicy)
        self.submitBtn.setMinimumSize(QSize(100, 35))
        self.submitBtn.setStyleSheet('background-color: rgb(63, 63, 63);\ncolor: rgb(255, 255, 255);')
        self.submitBtn.setObjectName('submitBtn')
        self.gridLayout3.addWidget(self.submitBtn, 3, 5, 1, 2)
        self.label = QLabel(self.groupBox)
        self.label.setAlignment(Qt.AlignRight | Qt.AlignTrailing | Qt.AlignVCenter)
        self.label.setObjectName('label')
        self.gridLayout3.addWidget(self.label, 0, 0, 1, 1)
        self.label2 = QLabel(self.groupBox)
        self.label2.setAlignment(Qt.AlignRight | Qt.AlignTrailing | Qt.AlignVCenter)
        self.label2.setObjectName('label2')
        self.gridLayout3.addWidget(self.label2, 1, 0, 1, 1)
        self.line = QFrame(self.groupBox)
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)
        self.line.setObjectName('line')
        self.gridLayout3.addWidget(self.line, 2, 0, 1, 10)
        self.gridLayout4.addLayout(self.gridLayout3, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox, 0, 0, 1, 1)
        self.groupBox2 = QGroupBox(self.centralwidget)
        self.groupBox2.setTitle('')
        self.groupBox2.setObjectName('groupBox2')
        self.gridLayout6 = QGridLayout(self.groupBox2)
        self.gridLayout6.setObjectName('gridLayout6')
        self.gridLayout5 = QGridLayout()
        self.gridLayout5.setObjectName('gridLayout5')
        listViewModelLinkage = QStandardItemModel()
        self.listViewLinkage = QListView(self.groupBox2)
        self.listViewLinkage.setModel(listViewModelLinkage)
        self.gridLayout5.addWidget(self.listViewLinkage, 1, 0, 1, 1)
        self.label3 = QLabel(self.groupBox2)
        self.label3.setAlignment(Qt.AlignLeft | Qt.AlignTrailing | Qt.AlignVCenter)
        self.label3.setObjectName('label3')
        self.gridLayout5.addWidget(self.label3, 0, 0, 1, 1)
        self.vline = QFrame(self.groupBox2)
        self.vline.setFrameShape(QFrame.VLine)
        self.vline.setFrameShadow(QFrame.Sunken)
        self.vline.setObjectName('vline')
        self.gridLayout5.addWidget(self.vline, 1, 1, 1, 1)
        listViewModelmscode = QStandardItemModel()
        self.listViewmscode = QListView(self.groupBox2)
        self.listViewmscode.setModel(listViewModelmscode)
        self.gridLayout5.addWidget(self.listViewmscode, 1, 2, 1, 1)
        self.label4 = QLabel(self.groupBox2)
        self.label4.setAlignment(Qt.AlignLeft | Qt.AlignTrailing | Qt.AlignVCenter)
        self.label4.setObjectName('label4')
        self.gridLayout5.addWidget(self.label4, 0, 2, 1, 1)
        self.label5 = QLabel(self.groupBox2)
        self.label5.setAlignment(Qt.AlignLeft | Qt.AlignTrailing | Qt.AlignVCenter)
        self.label5.setObjectName('label5')
        self.gridLayout5.addWidget(self.label5, 0, 3, 1, 1)
        self.linkageDelBtn = QPushButton(self.groupBox2)
        self.linkageDelBtn.setMinimumSize(QSize(0, 25))
        self.gridLayout5.addWidget(self.linkageDelBtn, 2, 0, 1, 1)
        self.mscodeDelBtn = QPushButton(self.groupBox2)
        self.mscodeDelBtn.setMinimumSize(QSize(0, 25))
        self.gridLayout5.addWidget(self.mscodeDelBtn, 2, 2, 1, 1)
        self.gridLayout6.addLayout(self.gridLayout5, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox2, 1, 0, 1, 1)
        self.gridLayout2.addLayout(self.gridLayout, 0, 0, 1, 1)
        self.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(self)
        self.menubar.setGeometry(QRect(0, 0, 653, 21))
        self.menubar.setObjectName('menubar')
        self.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(self)
        self.statusbar.setObjectName('statusbar')
        self.setStatusBar(self.statusbar)
        self.retranslateUi(self)
        self.mscodeInput.returnPressed.connect(self.addmscode)
        self.linkageInput.returnPressed.connect(self.addLinkage)
        self.linkageInputBtn.clicked.connect(self.addLinkage)
        self.mscodeInputBtn.clicked.connect(self.addmscode)
        self.linkageDelBtn.clicked.connect(self.delLinkage)
        self.mscodeDelBtn.clicked.connect(self.delmscode)
        self.submitBtn.clicked.connect(self.confirm)
        self.linkageAddExcelBtn.clicked.connect(self.addLinkageExcel)
        self.mscodeAddExcelBtn.clicked.connect(self.addmscodeExcel)
        self.retranslateUi(self)
        self.show()

    def retranslateUi(self, MainWindow):
        _translate = QCoreApplication.translate
        MainWindow.setWindowTitle(_translate('SubWindow', '??????/???????????? ??????'))
        MainWindow.setWindowIcon(QIcon('.\\Logo\\logo.png'))
        self.label.setText(_translate('SubWindow', 'Linkage No ?????? :'))
        self.linkageInputBtn.setText(_translate('SubWindow', '??????'))
        self.label2.setText(_translate('SubWindow', 'MS-CODE ?????? :'))
        self.mscodeInputBtn.setText(_translate('SubWindow', '??????'))
        self.submitBtn.setText(_translate('SubWindow', '?????? ??????'))
        self.label3.setText(_translate('SubWindow', 'Linkage No List'))
        self.label4.setText(_translate('SubWindow', 'MS-Code List'))
        self.linkageDelBtn.setText(_translate('SubWindow', '??????'))
        self.mscodeDelBtn.setText(_translate('SubWindow', '??????'))
        self.linkageAddExcelBtn.setText(_translate('SubWindow', '?????? ??????'))
        self.mscodeAddExcelBtn.setText(_translate('SubWindow', '?????? ??????'))

    @pyqtSlot()
    def addLinkage(self):
        linkageNo = self.linkageInput.text()
        if len(linkageNo) == 16:
            if linkageNo.isdigit():
                model = self.listViewLinkage.model()
                linkageItem = QStandardItem()
                linkageItemModel = QStandardItemModel()
                dupFlag = False
                for i in range(model.rowCount()):
                    index = model.index(i, 0)
                    item = model.data(index)
                    if item == linkageNo:
                        dupFlag = True
                    linkageItem = QStandardItem(item)
                    linkageItemModel.appendRow(linkageItem)
                if not dupFlag:
                    linkageItem = QStandardItem(linkageNo)
                    linkageItemModel.appendRow(linkageItem)
                    self.listViewLinkage.setModel(linkageItemModel)
                else:
                    QMessageBox.information(self, 'Error', '????????? ???????????? ????????????.')
            else:
                QMessageBox.information(self, 'Error', '????????? ??????????????????.')
        elif len(linkageNo) == 0:
            QMessageBox.information(self, 'Error', 'Linkage Number ???????????? ???????????? ???????????????.')
        else:
            QMessageBox.information(self, 'Error', '16????????? Linkage Number??? ??????????????????.')

    @pyqtSlot()
    def delLinkage(self):
        model = self.listViewLinkage.model()
        linkageItem = QStandardItem()
        linkageItemModel = QStandardItemModel()
        for index in self.listViewLinkage.selectedIndexes():
            selected_item = self.listViewLinkage.model().data(index)
            for i in range(model.rowCount()):
                index = model.index(i, 0)
                item = model.data(index)
                linkageItem = QStandardItem(item)
                if selected_item != item:
                    linkageItemModel.appendRow(linkageItem)
            self.listViewLinkage.setModel(linkageItemModel)

    @pyqtSlot()
    def addmscode(self):
        mscode = self.mscodeInput.text()
        if len(mscode) > 0:
            model = self.listViewmscode.model()
            mscodeItem = QStandardItem()
            mscodeItemModel = QStandardItemModel()
            dupFlag = False
            for i in range(model.rowCount()):
                index = model.index(i, 0)
                item = model.data(index)
                if item == mscode:
                    dupFlag = True
                mscodeItem = QStandardItem(item)
                mscodeItemModel.appendRow(mscodeItem)
            if not dupFlag:
                mscodeItem = QStandardItem(mscode)
                mscodeItemModel.appendRow(mscodeItem)
                self.listViewmscode.setModel(mscodeItemModel)
            else:
                QMessageBox.information(self, 'Error', '????????? ???????????? ????????????.')
        else:
            QMessageBox.information(self, 'Error', 'MS-CODE ???????????? ???????????? ???????????????.')

    @pyqtSlot()
    def delmscode(self):
        model = self.listViewmscode.model()
        mscodeItem = QStandardItem()
        mscodeItemModel = QStandardItemModel()
        for index in self.listViewmscode.selectedIndexes():
            selected_item = self.listViewmscode.model().data(index)
            for i in range(model.rowCount()):
                index = model.index(i, 0)
                item = model.data(index)
                mscodeItem = QStandardItem(item)
                if selected_item != item:
                    mscodeItemModel.appendRow(mscodeItem)
            self.listViewmscode.setModel(mscodeItemModel)

    @pyqtSlot()
    def addLinkageExcel(self):
        try:
            fileName = QFileDialog.getOpenFileName(self, 'Open File', './', 'Excel Files (*.xlsx)')[0]
            if fileName != "":
                df = pd.read_excel(fileName)
                for i in df.index:
                    linkageNo = str(df[df.columns[0]][i])
                    if len(linkageNo) == 16:
                        if linkageNo.isdigit():
                            model = self.listViewLinkage.model()
                            linkageItem = QStandardItem()
                            linkageItemModel = QStandardItemModel()
                            dupFlag = False
                            for i in range(model.rowCount()):
                                index = model.index(i, 0)
                                item = model.data(index)
                                if item == linkageNo:
                                    dupFlag = True
                                linkageItem = QStandardItem(item)
                                linkageItemModel.appendRow(linkageItem)
                            if not dupFlag:
                                linkageItem = QStandardItem(linkageNo)
                                linkageItemModel.appendRow(linkageItem)
                                self.listViewLinkage.setModel(linkageItemModel)
                            else:
                                QMessageBox.information(self, 'Error', '????????? ???????????? ????????????.')
                        else:
                            QMessageBox.information(self, 'Error', '????????? ??????????????????.')
                    elif len(linkageNo) == 0:
                        QMessageBox.information(self, 'Error', 'Linkage Number ???????????? ???????????? ???????????????.')
                    else:
                        QMessageBox.information(self, 'Error', '16????????? Linkage Number??? ??????????????????.')
        except Exception as e:
            QMessageBox.information(self, 'Error', '???????????? : ' + e)

    @pyqtSlot()
    def addmscodeExcel(self):
        try:
            fileName = QFileDialog.getOpenFileName(self, 'Open File', './', 'Excel Files (*.xlsx)')[0]
            if fileName != "":
                df = pd.read_excel(fileName)
                for i in df.index:
                    mscode = str(df[df.columns[0]][i])
                    if len(mscode) > 0:
                        model = self.listViewmscode.model()
                        mscodeItem = QStandardItem()
                        mscodeItemModel = QStandardItemModel()
                        dupFlag = False
                        for i in range(model.rowCount()):
                            index = model.index(i, 0)
                            item = model.data(index)
                            if item == mscode:
                                dupFlag = True
                            mscodeItem = QStandardItem(item)
                            mscodeItemModel.appendRow(mscodeItem)
                        if not dupFlag:
                            mscodeItem = QStandardItem(mscode)
                            mscodeItemModel.appendRow(mscodeItem)
                            self.listViewmscode.setModel(mscodeItemModel)
                        else:
                            QMessageBox.information(self, 'Error', '????????? ???????????? ????????????.')
                    else:
                        QMessageBox.information(self, 'Error', 'MS-CODE ???????????? ???????????? ???????????????.')
        except Exception as e:
            QMessageBox.information(self, 'Error', '???????????? : ' + e)

    @pyqtSlot()
    def confirm(self):
        self.submitClicked.emit([self.listViewLinkage.model(), self.listViewmscode.model()])
        self.close()


class Ui_MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi()

    def setupUi(self):
        rfh = RotatingFileHandler(filename='./Log.log',
                                    mode='a',
                                    maxBytes=5 * 1024 * 1024,
                                    backupCount=2,
                                    encoding=None,
                                    delay=0)
        logging.basicConfig(level=logging.DEBUG,
                            format='%(asctime)s:%(levelname)s:%(message)s',
                            datefmt='%m/%d/%Y %H:%M:%S',
                            handlers=[rfh])
        self.setObjectName('MainWindow')
        self.resize(800, 800)
        self.setStyleSheet('background-color: rgb(252, 252, 252);')
        self.centralwidget = QWidget(self)
        self.centralwidget.setObjectName('centralwidget')
        self.gridLayout2 = QGridLayout(self.centralwidget)
        self.gridLayout2.setObjectName('gridLayout2')
        self.gridLayout = QGridLayout()
        self.gridLayout.setObjectName('gridLayout')
        self.groupBox = QGroupBox(self.centralwidget)
        self.groupBox.setTitle('')
        self.groupBox.setObjectName('groupBox')
        self.gridLayout4 = QGridLayout(self.groupBox)
        self.gridLayout4.setObjectName('gridLayout4')
        self.gridLayout3 = QGridLayout()
        self.gridLayout3.setObjectName('gridLayout3')
        self.maxOrderinput = QLineEdit(self.groupBox)
        self.maxOrderinput.setMinimumSize(QSize(0, 25))
        self.maxOrderinput.setObjectName('maxOrderinput')
        self.maxOrderinput.setValidator(QIntValidator(self))
        self.gridLayout3.addWidget(self.maxOrderinput, 0, 1, 1, 3)
        self.dateBtn = QToolButton(self.groupBox)
        self.dateBtn.setMinimumSize(QSize(0, 25))
        self.dateBtn.setObjectName('dateBtn')
        self.gridLayout3.addWidget(self.dateBtn, 1, 1, 1, 1)
        self.emgFileInputBtn = QPushButton(self.groupBox)
        self.emgFileInputBtn.setMinimumSize(QSize(0, 25))
        self.gridLayout3.addWidget(self.emgFileInputBtn, 2, 1, 1, 1)
        self.holdFileInputBtn = QPushButton(self.groupBox)
        self.holdFileInputBtn.setMinimumSize(QSize(0, 25))
        self.gridLayout3.addWidget(self.holdFileInputBtn, 5, 1, 1, 1)
        self.label4 = QLabel(self.groupBox)
        self.label4.setAlignment(Qt.AlignRight | Qt.AlignTrailing | Qt.AlignVCenter)
        self.label4.setObjectName('label4')
        self.gridLayout3.addWidget(self.label4, 3, 1, 1, 1)
        self.label5 = QLabel(self.groupBox)
        self.label5.setAlignment(Qt.AlignRight | Qt.AlignTrailing | Qt.AlignVCenter)
        self.label5.setObjectName('label5')
        self.gridLayout3.addWidget(self.label5, 3, 2, 1, 1)
        self.label6 = QLabel(self.groupBox)
        self.label6.setAlignment(Qt.AlignRight | Qt.AlignTrailing | Qt.AlignVCenter)
        self.label6.setObjectName('label6')
        self.gridLayout3.addWidget(self.label6, 6, 1, 1, 1)
        self.label7 = QLabel(self.groupBox)
        self.label7.setAlignment(Qt.AlignRight | Qt.AlignTrailing | Qt.AlignVCenter)
        self.label7.setObjectName('label7')
        self.gridLayout3.addWidget(self.label7, 6, 2, 1, 1)
        listViewModelEmgLinkage = QStandardItemModel()
        self.listViewEmgLinkage = QListView(self.groupBox)
        self.listViewEmgLinkage.setModel(listViewModelEmgLinkage)
        self.gridLayout3.addWidget(self.listViewEmgLinkage, 4, 1, 1, 1)
        listViewModelEmgmscode = QStandardItemModel()
        self.listViewEmgmscode = QListView(self.groupBox)
        self.listViewEmgmscode.setModel(listViewModelEmgmscode)
        self.gridLayout3.addWidget(self.listViewEmgmscode, 4, 2, 1, 1)
        listViewModelHoldLinkage = QStandardItemModel()
        self.listViewHoldLinkage = QListView(self.groupBox)
        self.listViewHoldLinkage.setModel(listViewModelHoldLinkage)
        self.gridLayout3.addWidget(self.listViewHoldLinkage, 7, 1, 1, 1)
        listViewModelHoldmscode = QStandardItemModel()
        self.listViewHoldmscode = QListView(self.groupBox)
        self.listViewHoldmscode.setModel(listViewModelHoldmscode)
        self.gridLayout3.addWidget(self.listViewHoldmscode, 7, 2, 1, 1)
        self.labelBlank = QLabel(self.groupBox)
        self.labelBlank.setObjectName('labelBlank')
        self.gridLayout3.addWidget(self.labelBlank, 2, 4, 1, 1)
        self.progressbar = QProgressBar(self.groupBox)
        self.progressbar.setObjectName('progressbar')
        self.gridLayout3.addWidget(self.progressbar, 10, 1, 1, 2)
        thread = QThread()
        worker = Worker()
        worker.moveToThread(thread)
        worker.progressChanged.connect(self.progressbar.setValue)
        self.runBtn = QToolButton(self.groupBox)
        sizePolicy = QSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.runBtn.sizePolicy().hasHeightForWidth())
        self.runBtn.setSizePolicy(sizePolicy)
        self.runBtn.setMinimumSize(QSize(0, 35))
        self.runBtn.setStyleSheet('background-color: rgb(63, 63, 63);\ncolor: rgb(255, 255, 255);')
        self.runBtn.setObjectName('runBtn')
        self.gridLayout3.addWidget(self.runBtn, 10, 3, 1, 2)
        self.label = QLabel(self.groupBox)
        self.label.setAlignment(Qt.AlignRight | Qt.AlignTrailing | Qt.AlignVCenter)
        self.label.setObjectName('label')
        self.gridLayout3.addWidget(self.label, 0, 0, 1, 1)
        self.label8 = QLabel(self.groupBox)
        self.label8.setAlignment(Qt.AlignRight | Qt.AlignTrailing | Qt.AlignVCenter)
        self.label8.setObjectName('label8')
        self.gridLayout3.addWidget(self.label8, 1, 0, 1, 1)
        self.labelDate = QLabel(self.groupBox)
        self.labelDate.setAlignment(Qt.AlignRight | Qt.AlignTrailing | Qt.AlignVCenter)
        self.labelDate.setObjectName('labelDate')
        self.gridLayout3.addWidget(self.labelDate, 1, 2, 1, 1)
        self.label2 = QLabel(self.groupBox)
        self.label2.setAlignment(Qt.AlignRight | Qt.AlignTrailing | Qt.AlignVCenter)
        self.label2.setObjectName('label2')
        self.gridLayout3.addWidget(self.label2, 2, 0, 1, 1)
        self.label3 = QLabel(self.groupBox)
        self.label3.setAlignment(Qt.AlignRight | Qt.AlignTrailing | Qt.AlignVCenter)
        self.label3.setObjectName('label3')
        self.gridLayout3.addWidget(self.label3, 5, 0, 1, 1)
        self.cbLimit = QCheckBox('UT32A, UT52A, UM33A ????????? 50%??????', self)
        self.cbLimit.setObjectName('cbLimit')
        self.gridLayout3.addWidget(self.cbLimit, 8, 1, 1, 2)
        self.cbLimit.setChecked(True)
        self.line = QFrame(self.groupBox)
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)
        self.line.setObjectName('line')
        self.gridLayout3.addWidget(self.line, 9, 0, 1, 10)
        self.gridLayout4.addLayout(self.gridLayout3, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox, 0, 0, 1, 1)
        self.groupBox2 = QGroupBox(self.centralwidget)
        self.groupBox2.setTitle('')
        self.groupBox2.setObjectName('groupBox2')
        self.gridLayout6 = QGridLayout(self.groupBox2)
        self.gridLayout6.setObjectName('gridLayout6')
        self.gridLayout5 = QGridLayout()
        self.gridLayout5.setObjectName('gridLayout5')
        self.logBrowser = QPlainTextEditLogger(self.groupBox2)
        self.logBrowser.setFormatter(CustomFormatter())
        logging.getLogger().addHandler(self.logBrowser)
        logging.getLogger().setLevel(logging.INFO)
        self.gridLayout5.addWidget(self.logBrowser.widget, 0, 0, 1, 1)
        self.gridLayout6.addLayout(self.gridLayout5, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox2, 1, 0, 1, 1)
        self.gridLayout2.addLayout(self.gridLayout, 0, 0, 1, 1)
        self.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(self)
        self.menubar.setGeometry(QRect(0, 0, 653, 21))
        self.menubar.setObjectName('menubar')
        self.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(self)
        self.statusbar.setObjectName('statusbar')
        self.setStatusBar(self.statusbar)
        self.retranslateUi(self)
        self.dateBtn.clicked.connect(self.selectStartDate)
        self.emgFileInputBtn.clicked.connect(self.emgWindow)
        self.holdFileInputBtn.clicked.connect(self.holdWindow)
        self.runBtn.clicked.connect(self.startLeveling)
        self.cbLimit.stateChanged.connect(self.changeCbLimit)
        # ???????????? ?????????
        self.isDebug = False
        if self.isDebug:
            self.debugDate = QLineEdit(self.groupBox)
            self.debugDate.setObjectName('debugDate')
            self.gridLayout3.addWidget(self.debugDate, 10, 0, 1, 1)
            self.debugDate.setPlaceholderText('???????????? ????????????')
        self.show()

    def retranslateUi(self, MainWindow):
        _translate = QCoreApplication.translate
        MainWindow.setWindowTitle(_translate('MainWindow', 'UTA ????????? ????????? ???????????? Rev1.02'))
        MainWindow.setWindowIcon(QIcon('.\\Logo\\logo.png'))
        self.label.setText(_translate('MainWindow', '???????????? (?????? ?????????):'))
        self.runBtn.setText(_translate('MainWindow', '??????'))
        self.label2.setText(_translate('MainWindow', '???????????? ?????? :'))
        self.label3.setText(_translate('MainWindow', '???????????? ?????? :'))
        self.label4.setText(_translate('MainWindow', 'Linkage No List'))
        self.label5.setText(_translate('MainWindow', 'MSCode List'))
        self.label6.setText(_translate('MainWindow', 'Linkage No List'))
        self.label7.setText(_translate('MainWindow', 'MSCode List'))
        self.label8.setText(_translate('MainWndow', '??????????????? ?????? :'))
        self.labelDate.setText(_translate('MainWndow', '?????????'))
        self.dateBtn.setText(_translate('MainWindow', ' ??????????????? ?????? '))
        self.emgFileInputBtn.setText(_translate('MainWindow', '????????? ??????'))
        self.holdFileInputBtn.setText(_translate('MainWindow', '????????? ??????'))
        self.labelBlank.setText(_translate('MainWindow', '            '))
        logging.info('??????????????? ?????? ??????????????????')

    @pyqtSlot()
    def emgWindow(self):
        self.w = UISubWindow()
        self.w.submitClicked.connect(self.getEmgListview)
        self.w.show()

    @pyqtSlot()
    def holdWindow(self):
        self.w = UISubWindow()
        self.w.submitClicked.connect(self.getHoldListview)
        self.w.show()

    def getEmgListview(self, list):
        if len(list) > 0:
            self.listViewEmgLinkage.setModel(list[0])
            self.listViewEmgmscode.setModel(list[1])
            logging.info('???????????? ???????????? ??????????????? ??????????????????.')
        else:
            logging.error('???????????? ???????????? ????????????. ?????? ?????? ??????????????????')

    def getHoldListview(self, list):
        if len(list) > 0:
            self.listViewHoldLinkage.setModel(list[0])
            self.listViewHoldmscode.setModel(list[1])
            logging.info('???????????? ???????????? ??????????????? ??????????????????.')
        else:
            logging.error('???????????? ???????????? ????????????. ?????? ?????? ??????????????????')

    def updateProgressbar(self, val):
        self.progressbar.setValue(val)

    def selectStartDate(self):
        self.w = CalendarWindow()
        self.w.submitClicked.connect(self.getStartDate)
        self.w.show()

    def getStartDate(self, date):
        if len(date) > 0:
            self.labelDate.setText(date)
            logging.info('?????????????????? %s ??? ??????????????? ?????????????????????.', date)
        else:
            logging.error('?????????????????? ???????????? ???????????????.')

    @pyqtSlot()
    def changeCbLimit(self):
        if self.cbLimit.isChecked():
            logging.info('UT32A, UT52A, UM33A ???????????? ?????? ???????????? 50%??? ???????????????.')
        else:
            logging.info('UT32A, UT52A, UM33A ????????? ????????? ?????????????????????.')

    @pyqtSlot()
    def startLeveling(self):
        # ??????????????? ??????????????? ????????????
        def loadMasterFile():
            masterFileList = []
            date = datetime.today().strftime('%Y%m%d')
            if self.isDebug:
                date = self.debugDate.text()
            utaOrderFilePath = r'.\\input\\Master_File\\' + date + r'\\UTA ?????? ' + date[4:] + r' (?????? ?????? DATA).xlsx'
            levelingListFilePath = r'.\\input\\Master_File\\' + date + r'\\5400_A0100A81_' + date + r'_Leveling_List.xlsx'
            conditionFilePath = r'.\\Input\\mscODE_Table\\UTA_????????????_?????????.xlsx'
            calendarFilePath = r'.\\Input\\Calendar_File\\FY' + date[2:4] + '_Calendar.xlsx'
            if os.path.exists(utaOrderFilePath):
                if os.path.exists(levelingListFilePath):
                    if os.path.exists(conditionFilePath):
                        if os.path.exists(calendarFilePath):
                            utaOrderFile = glob.glob(utaOrderFilePath)[0]
                            levelingListFile = glob.glob(levelingListFilePath)[0]
                            conditionFile = glob.glob(conditionFilePath)[0]
                            caledarFile = glob.glob(calendarFilePath)[0]
                            masterFileList.append(utaOrderFile)
                            masterFileList.append(levelingListFile)
                            masterFileList.append(conditionFile)
                            masterFileList.append(caledarFile)
                            logging.info('????????? ?????? ??? ???????????? ???????????? ??????????????? ??????????????????.')
                        else:
                            logging.error('%s ????????? ????????????. ??????????????????.', calendarFilePath)
                            self.runBtn.setEnabled(True)
                    else:
                        logging.error('%s ????????? ????????????. ??????????????????.', conditionFilePath)
                        self.runBtn.setEnabled(True)
                else:
                    logging.error('%s ????????? ????????????. ??????????????????.', levelingListFilePath)
                    self.runBtn.setEnabled(True)
            else:
                logging.error('%s ????????? ????????????. ??????????????????.', utaOrderFilePath)
                self.runBtn.setEnabled(True)
            return masterFileList

        def checkWorkDay(df, today, compDate):
            dtToday = pd.to_datetime(datetime.strptime(today, '%Y%m%d'))
            dtComp = pd.to_datetime(compDate, unit='s')
            workDay = 0
            for i in df.index:
                dt = pd.to_datetime(df['Date'][i], unit='s')
                if dtToday < dt and dt <= dtComp:
                    if df['WorkingDay'][i] == 1:
                        workDay += 1
            return workDay

        self.runBtn.setEnabled(False)

        # 50%?????? ?????????
        limitFlag = False
        if self.cbLimit.isChecked():
            limitFlag = True

        # pandas ?????? ?????? OFF
        pd.set_option('mode.chained_assignment', None)
        try:
            listMasterFile = loadMasterFile()
            # ??? ????????? ?????? ?????? ???, ???????????? ??????
            if len(self.maxOrderinput.text()) > 0 and len(listMasterFile) > 0 and int(self.maxOrderinput.text()) != 0:
                maxConstructCnt = float(self.maxOrderinput.text())
                # ??????????????? ????????????, ???????????? ????????? ????????????
                emgLinkage = [str(self.listViewEmgLinkage.model().data(self.listViewEmgLinkage.model().index(x, 0))) for x in range(self.listViewEmgLinkage.model().rowCount())]
                emgmscode = [self.listViewEmgmscode.model().data(self.listViewEmgmscode.model().index(x, 0)) for x in range(self.listViewEmgmscode.model().rowCount())]
                holdLinkage = [str(self.listViewHoldLinkage.model().data(self.listViewHoldLinkage.model().index(x, 0))) for x in range(self.listViewHoldLinkage.model().rowCount())]
                holdmscode = [self.listViewHoldmscode.model().data(self.listViewHoldmscode.model().index(x, 0)) for x in range(self.listViewHoldmscode.model().rowCount())]
                # ????????????, ???????????? ???????????? DataFrame?????? ??????
                dfEmgLinkage = pd.DataFrame({'Linkage Number': emgLinkage})
                dfEmgmscode = pd.DataFrame({'MS Code': emgmscode})
                dfHoldLinkage = pd.DataFrame({'Linkage Number': holdLinkage})
                dfHoldmscode = pd.DataFrame({'MS Code': holdmscode})
                # ??? Linkage Number Column??? int64 Type?????? ??????
                dfEmgLinkage['Linkage Number'] = dfEmgLinkage['Linkage Number'].astype(np.int64)
                dfHoldLinkage['Linkage Number'] = dfHoldLinkage['Linkage Number'].astype(np.int64)
                # Merge ???, ????????? ??????, ?????? ?????? ????????? ??????
                dfEmgLinkage['????????????'] = '??????'
                dfEmgmscode['????????????'] = '??????'
                dfHoldLinkage['????????????'] = '??????'
                dfHoldmscode['????????????'] = '??????'
                # Leveling Data ????????????
                dfLeveling = pd.read_excel(listMasterFile[1], converters={'Component Number': str})
                if self.isDebug:
                    dfLeveling.to_excel('.\\debug\\flow1.xlsx')
                # ????????? ?????? ?????? ???, ??? ?????? ??????
                dfLevelingDropSEQ = dfLeveling[dfLeveling['Sequence No'].isnull()]
                dfLevelingUndepSeq = dfLeveling[dfLeveling['Sequence No'] == 'Undep']
                dfLevelingUncorSeq = dfLeveling[dfLeveling['Sequence No'] == 'Uncor']
                dfLevelingResult = pd.concat([dfLevelingDropSEQ, dfLevelingUndepSeq, dfLevelingUncorSeq])
                dfLevelingResult = dfLevelingResult.reset_index(level=None, drop=False, inplace=False)
                for i in dfLevelingResult.index:
                    if len(str(dfLevelingResult['Component Number'][i])) != 4:
                        logging.warning(f'???Planned Order : {str(dfLevelingResult["Planned Order"][i]).zfill(10)}?????? Component Number??? ??????????????????. (????????? ??????)')

                dfLevelingResult['???????????????'] = dfLevelingResult.groupby('Linkage Number')['Linkage Number'].transform('size')
                dfUtaOrder = pd.read_excel(listMasterFile[0])
                if self.isDebug:
                    dfUtaOrder.to_excel('.\\debug\\flow2.xlsx')
                dfMerge = pd.merge(dfUtaOrder, dfLevelingResult, on='Linkage Number').drop_duplicates(['Linkage Number'])
                if self.isDebug:
                    dfMerge.to_excel('.\\debug\\flow3.xlsx')
                dfMergeLink = pd.merge(dfMerge, dfEmgLinkage, on='Linkage Number', how='left')
                dfMergemscode = pd.merge(dfMerge, dfEmgmscode, on='MS Code', how='left')
                dfMergeLink = pd.merge(dfMergeLink, dfHoldLinkage, on='Linkage Number', how='left')
                if self.isDebug:
                    dfMergeLink.to_excel('.\\debug\\flow4.xlsx')
                dfMergemscode = pd.merge(dfMergemscode, dfHoldmscode, on='MS Code', how='left')
                # ???????????? ????????????
                if self.isDebug:
                    dfMergemscode.to_excel('.\\debug\\flow5.xlsx')
                dfMergeLink['????????????'] = dfMergeLink['????????????'].combine_first(dfMergemscode['????????????'])
                dfMergeLink['????????????'] = dfMergeLink['????????????'].combine_first(dfMergemscode['????????????'])
                dfMergeLink['Linkage Number'] = dfMergeLink['Linkage Number'].astype(str)
                # ???????????? ????????????
                if self.isDebug:
                    dfMergeLink.to_excel('.\\debug\\flow6.xlsx')
                # ????????? Column??? ??????
                dfMergeResult = dfMergeLink[['Country: Ship-to Party',
                                                'Linkage Number',
                                                'Material',
                                                'MS Code',
                                                'Status Category',
                                                'Order Quantity',
                                                '???????????????',
                                                'Planned Prod. Completion date',
                                                'Planned Shipping date',
                                                '????????????',
                                                '????????????']]
                # ???????????? Copy??? ??? DataFrame ??????
                dfCopy = pd.DataFrame(columns=['Country: Ship-to Party',
                                                'Linkage Number',
                                                'Material',
                                                'MS Code',
                                                'Status Category',
                                                'Order Quantity',
                                                '???????????????',
                                                'Planned Prod. Completion date',
                                                'Planned Shipping date',
                                                '????????????',
                                                '????????????',
                                                '????????????',
                                                '????????????'])
                # ???????????? ????????????
                if self.isDebug:
                    dfMergeResult.to_excel('.\\debug\\flow7.xlsx')
                # ?????????????????? ?????????????????? ?????? ????????? ??????
                dfMergeResultS = dfMergeResult.sort_values(by=['Planned Prod. Completion date', 'Planned Shipping date'], ascending=[True, True])
                # ????????? ?????? ????????? ?????????
                dfMergeResultS = dfMergeResultS.reset_index(level=None, drop=False, inplace=False)
                # ?????? ????????? ????????? Column??????
                dfMergeResultS['????????????'] = 0
                dfMergeResultS['????????????'] = ""
                dfMergeResultS['????????????'] = ""
                dfMergeResultS['UT32A, UT52A, UM33A????????????'] = ""
                dfMergeResultS['Cycle??????'] = np.nan
                dfMergeResultS['MAX ?????? ??????'] = ""
                dfMergeResultS['????????????(%)'] = ""
                dfMergeResultS['?????? ????????????'] = np.nan
                # ?????? ?????? ????????? ????????????
                dfCondition = pd.read_excel(listMasterFile[2])
                # ??? ????????? ?????? ?????? ?????? ????????????
                dfCondition['No'] = dfCondition['No'].fillna(method='ffill')
                dfCondition['???(LINE)????????????'] = dfCondition['???(LINE)????????????'].fillna(method='ffill')
                dfCondition['????????????(%)'] = dfCondition['????????????(%)'].fillna(method='ffill')
                dfCondition['????????????(??????)'] = dfCondition['????????????(??????)'].fillna(method='ffill')
                dfCondition['Cycle ?????? ??????'] = dfCondition['Cycle ?????? ??????'].fillna(method='ffill')
                dfCondition['?????? ??????'] = dfCondition['?????? ??????'].fillna(method='ffill')
                dfCondition['?????? ?????? (?????? ??????)'] = dfCondition['?????? ?????? (?????? ??????)'].fillna(method='ffill')
                dfCondition['MAX ?????? ??????'] = dfCondition['MAX ?????? ??????'].fillna(method='ffill')
                # ????????????????????? ??? ????????? ?????????, UT??? ????????? ????????? ?????? ?????? ??????
                constructTempCnt = float(maxConstructCnt)
                capableCntDic = {}
                preOrderDic = {}
                # ?????? ?????? ???????????? ?????? ?????? Data??? ????????????
                for i in dfCondition.index:
                    # ??? ???????????? ?????? ??????, ?????????????????? ???????????? ??????
                    if dfCondition['???(LINE)????????????'][i] == '-':
                        capableCntDic[float(dfCondition['No'][i])] = maxConstructCnt
                    else:
                        capableCntDic[float(dfCondition['No'][i])] = dfCondition['???(LINE)????????????'][i]
                        preOrderDic[float(dfCondition['No'][i])] = math.ceil(float(dfCondition['???(LINE)????????????'][i]) * float(dfCondition['????????????(%)'][i]))
                for i in dfMergeResultS.index:
                    # ?????? ?????? 4850U1-???, UTAP????????? ??????
                    if str(dfMergeResultS['MS Code'][i]).find('4850U1') > -1 or str(dfMergeResultS['MS Code'][i]).find('UTAP') > -1:
                        dfMergeResultS.drop(i, inplace=True)
                    # Status 100???????????? 0??? ?????? ??????
                    if int(dfMergeResultS['Status Category'][i]) >= 100 or int(dfMergeResultS['Status Category'][i]) == 0:
                        dfMergeResultS.drop(i, inplace=True)
                    # UT??? ??????
                    if str(dfMergeResultS['MS Code'][i]).find('UT32A') > -1 or str(dfMergeResultS['MS Code'][i]).find('UT52A') > -1 or str(dfMergeResultS['MS Code'][i]).find('UM33A') > -1:
                        dfMergeResultS['UT32A, UT52A, UM33A????????????'][i] = 'UT32A, UT52A, UM33A'
                    # ?????? ???????????? ???????????? ?????? ??????, ??????????????? ??????
                    dt = pd.to_datetime(dfMergeResultS['Planned Prod. Completion date'][i], unit='s')
                    if self.labelDate.text() != '?????????':
                        dtInput = pd.to_datetime(datetime.strptime(self.labelDate.text(), '%Y-%m-%d'))
                    if self.labelDate.text() == '?????????':
                        if dt < pd.Timestamp.now():
                            dfMergeResultS['????????????'][i] = '??????'
                    else:
                        if dt < dtInput:
                            dfMergeResultS['????????????'][i] = '??????'
                    # ???????????? ????????? ????????????
                    for j in dfCondition.index:
                        optionCode = ""
                        mscode = ''
                        if str(dfCondition['Model'][j]).strip() != 'nan' and str(dfCondition['Model'][j]).strip() != '':
                            mscode = str(dfCondition['Model'][j]).strip()
                        else:
                            mscode = '.....'
                        if mscode not in ('LL50A', 'MDUT'):
                            for k in range(1, 8):
                                if str(dfCondition['Suffix Code' + str(k)][j]).strip() != 'nan' and str(dfCondition['Suffix Code' + str(k)][j]).strip() != '':
                                    if 'UM33A' not in str(dfMergeResultS['MS Code'][i]).strip():
                                        if k in (1, 4, 6):
                                            mscode = mscode + '-' + str(dfCondition['Suffix Code' + str(k)][j]).strip()
                                        else:
                                            mscode = mscode + str(dfCondition['Suffix Code' + str(k)][j]).strip()
                                    else:
                                        if k in (1, 4):
                                            mscode = mscode + '-' + str(dfCondition['Suffix Code' + str(k)][j]).strip()
                                        elif k < 6:
                                            mscode = mscode + str(dfCondition['Suffix Code' + str(k)][j]).strip()
                                else:
                                    if 'UM33A' not in str(dfMergeResultS['MS Code'][i]).strip():
                                        if k in (1, 4, 6):
                                            mscode = mscode + '-' + '.'
                                        else:
                                            mscode = mscode + '.'
                                    else:
                                        if k in (1, 4):
                                            mscode = mscode + '-' + '.'
                                        elif k < 6:
                                            mscode = mscode + '.'
                            if str(dfCondition['Option Code'][j]).strip() != 'nan' and str(dfCondition['Option Code'][j]).strip() != '':

                                optionCode = str(dfCondition['Option Code'][j]).strip()
                            else:
                                optionCode = ""
                        # if optionCode == '/DC' and str(dfMergeResultS['MS Code'][i]) == 'UM33A-020-00/DC':
                        #     print()
                        if bool(re.search(mscode, str(dfMergeResultS['MS Code'][i]))):
                            if optionCode != "":
                                if str(dfMergeResultS['MS Code'][i]).find(optionCode) > -1:
                                    if dfCondition['?????? ?????? (?????? ??????)'][j] == '??????':
                                        if dfMergeResultS['????????????'][i] == "":
                                            dfMergeResultS['????????????'][i] = str(dfCondition['No'][j])
                                        else:
                                            dfMergeResultS['????????????'][i] = dfMergeResultS['????????????'][i] + ',' + str(dfCondition['No'][j])
                                    else:
                                        if dfMergeResultS['????????????'][i] == "":
                                            dfMergeResultS['????????????'][i] = str(dfCondition['No'][j])
                                        else:
                                            dfMergeResultS['????????????'][i] = dfMergeResultS['????????????'][i] + ',' + str(dfCondition['No'][j])
                                    if dfCondition['Cycle ?????? ??????'][j] != '-':
                                        dfMergeResultS['Cycle??????'][i] = int(dfCondition['Cycle ?????? ??????'][j])
                                    if dfMergeResultS['????????????'][i] == "" or dfMergeResultS['????????????'][i] < dfCondition['?????? ??????'][j]:
                                        dfMergeResultS['????????????'][i] = dfCondition['?????? ??????'][j]
                                    dfMergeResultS['????????????(%)'][i] = dfCondition['????????????(%)'][j]
                                    if str(dfCondition['MAX ?????? ??????'][j]) != 'nan' and str(dfCondition['MAX ?????? ??????'][j]) != '' and str(dfCondition['MAX ?????? ??????'][j]) != '-':
                                        dfMergeResultS['MAX ?????? ??????'][i] = '??????'
                                        dfMergeResultS['????????????(%)'][i] = 1.0
                            elif mscode != '.....-...-..':
                                if dfCondition['?????? ?????? (?????? ??????)'][j] == '??????':
                                    if dfMergeResultS['????????????'][i] == "":
                                        dfMergeResultS['????????????'][i] = str(dfCondition['No'][j])
                                    else:
                                        dfMergeResultS['????????????'][i] = dfMergeResultS['????????????'][i] + ',' + str(dfCondition['No'][j])
                                else:
                                    if dfMergeResultS['????????????'][i] == "":
                                        dfMergeResultS['????????????'][i] = str(dfCondition['No'][j])
                                    else:
                                        dfMergeResultS['????????????'][i] = dfMergeResultS['????????????'][i] + ',' + str(dfCondition['No'][j])
                                if dfCondition['Cycle ?????? ??????'][j] != '-':
                                    dfMergeResultS['Cycle??????'][i] = int(dfCondition['Cycle ?????? ??????'][j])
                                if dfMergeResultS['????????????'][i] == "" or dfMergeResultS['????????????'][i] < dfCondition['?????? ??????'][j]:
                                    dfMergeResultS['????????????'][i] = dfCondition['?????? ??????'][j]
                                    dfMergeResultS['????????????(%)'][i] = dfCondition['????????????(%)'][j]
                                if str(dfCondition['MAX ?????? ??????'][j]) != 'nan' and str(dfCondition['MAX ?????? ??????'][j]) != '' and str(dfCondition['MAX ?????? ??????'][j]) != '-':
                                    dfMergeResultS['MAX ?????? ??????'][i] = '??????'
                                    dfMergeResultS['????????????(%)'][i] = 1.0
                # ????????? ???????????? ????????? ???????????? 1??? ??????, ??????????????? ????????? ?????? ???????????????*1??? ??????
                dfMergeResultS.loc[dfMergeResultS['????????????'] == 0, '????????????'] = 1
                # ???????????? ?????? ??????
                if self.isDebug:
                    dfMergeResultS.to_excel('.\\debug\\flow8.xlsx')
                # ??????????????? ??????????????? ?????? ????????? ?????? ????????? ??????
                maxSpCommaCnt = 0
                tempSpCommaCnt = 0
                maxNormalCommaCnt = 0
                tempNormalCommaCnt = 0

                # ???????????? ????????? ????????????
                dfCalendar = pd.read_excel(listMasterFile[3])
                today = datetime.today().strftime('%Y%m%d')
                if self.isDebug:
                    today = self.debugDate.text()

                # ??????????????? ??????????????? ?????? ????????? ?????? + ???????????? ??????
                for i in dfMergeResultS.index:
                    tempSpCommaCnt = str(dfMergeResultS['????????????'][i]).count(',')
                    if int(tempSpCommaCnt) > int(maxSpCommaCnt):
                        maxSpCommaCnt = tempSpCommaCnt
                    tempNormalCommaCnt = str(dfMergeResultS['????????????'][i]).count(',')
                    if int(tempNormalCommaCnt) > int(maxNormalCommaCnt):
                        maxNormalCommaCnt = tempNormalCommaCnt
                    dfMergeResultS['?????? ????????????'][i] = checkWorkDay(dfCalendar, today, dfMergeResultS['Planned Prod. Completion date'][i])

                # ???????????? Column??? ??????
                for i in range(0, maxSpCommaCnt + 1):
                    dfMergeResultS['????????????' + str(i + 1)] = dfMergeResultS['????????????'].str.split(',').str[i]
                # ???????????? Column??? ??????
                for i in range(0, maxNormalCommaCnt + 1):
                    dfMergeResultS['????????????' + str(i + 1)] = dfMergeResultS['????????????'].str.split(',').str[i]
                dfMergeResultS.fillna(0)

                # ???????????? ?????? ??????
                if self.isDebug:
                    dfMergeResultS.to_excel('.\\debug\\flow9.xlsx')
                # ??????????????? -> ??????????????? -> ???????????? -> ???????????? ????????? ???????????????????????? ????????? ?????? ??????
                dfMergeResultSf = dfMergeResultS.sort_values(by=['Planned Prod. Completion date', '????????????', 'MAX ?????? ??????', 'Planned Shipping date', '????????????', '????????????'],
                                                                ascending=[True, False, False, True, False, False])
                # ????????? ???????????? ?????? ?????? ????????? ??????
                dfMergeResultSf.drop('index', axis=1, inplace=True)
                # ????????? ?????????
                dfMergeResultSfReset = dfMergeResultSf.reset_index(level=None, drop=False, inplace=False)
                # ?????? ????????? ??????
                dfMergeResultSfReset.drop('index', axis=1, inplace=True)
                # ???????????? ?????? ??????
                if self.isDebug:
                    dfMergeResultSfReset.to_excel('.\\debug\\flow10.xlsx')
                dfMergeResultSfReset['????????????'] = dfMergeResultSfReset['???????????????']

                dfMergeResultSfReset['??????????????????'] = ''
                integCntDic = {}
                copyCntDic = capableCntDic.copy()
                minContCntDic = {}
                for i in dfMergeResultSfReset.index:
                    for key, value in capableCntDic.items():
                        for x in range(0, maxSpCommaCnt + 1):
                            if str(dfMergeResultSfReset['????????????' + str(x + 1)][i]) != '' and str(dfMergeResultSfReset['????????????' + str(x + 1)][i]) != 'nan':
                                if key == float(dfMergeResultSfReset['????????????' + str(x + 1)][i]):
                                    if key in integCntDic:
                                        integCntDic[key] += dfMergeResultSfReset['???????????????'][i]
                                    else:
                                        integCntDic[key] = dfMergeResultSfReset['???????????????'][i]
                        for y in range(0, maxNormalCommaCnt + 1):
                            if str(dfMergeResultSfReset['????????????' + str(y + 1)][i]) != '' and str(dfMergeResultSfReset['????????????' + str(y + 1)][i]) != 'nan':
                                if key == float(dfMergeResultSfReset['????????????' + str(y + 1)][i]):
                                    if key in integCntDic:
                                        integCntDic[key] += dfMergeResultSfReset['???????????????'][i]
                                    else:
                                        integCntDic[key] = dfMergeResultSfReset['???????????????'][i]
                    if dfMergeResultSfReset['?????? ????????????'][i] == 0:
                        workDay = 1
                    else:
                        workDay = dfMergeResultSfReset['?????? ????????????'][i]
                    for x in range(0, maxSpCommaCnt + 1):
                        if str(dfMergeResultSfReset['????????????' + str(x + 1)][i]) != '' and str(dfMergeResultSfReset['????????????' + str(x + 1)][i]) != 'nan':
                            limitCopyCnt = capableCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])]
                            if (math.ceil(limitCopyCnt) <= (integCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] / workDay)) and copyCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] > 0:
                                if len(minContCntDic) > 0:
                                    if dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(x + 1)][i])]['?????????'].values[0] in minContCntDic:
                                        if minContCntDic[dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(x + 1)][i])]['?????????'].values[0]][0] < (integCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] / workDay):
                                            minContCntDic[dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(x + 1)][i])]['?????????'].values[0]][0] = (integCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] / workDay)
                                            minContCntDic[dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(x + 1)][i])]['?????????'].values[0]][1] = int(dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(x + 1)][i])]['???(LINE)????????????'].values[0])
                                            minContCntDic[dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(x + 1)][i])]['?????????'].values[0]][2] = dfMergeResultSfReset['Planned Prod. Completion date'][i]
                                else:
                                    minContCntDic[dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(x + 1)][i])]['?????????'].values[0]] = [(integCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] / workDay),
                                                                                                                                                                int(dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(x + 1)][i])]['???(LINE)????????????'].values[0]),
                                                                                                                                                                dfMergeResultSfReset['Planned Prod. Completion date'][i]]
                            limitCnt = capableCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] * float(dfMergeResultSfReset['????????????(%)'][i])
                            if (math.ceil(limitCnt) <= (integCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] / workDay)) and preOrderDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] > 0:
                                for j in dfMergeResultSfReset.index:
                                    if dfMergeResultSfReset['????????????' + str(x + 1)][i] in dfMergeResultSfReset['????????????'][j]:
                                        if preOrderDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] > 0:
                                            dfMergeResultSfReset['??????????????????'][j] = '??????'
                                            preOrderDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] -= dfMergeResultSfReset['????????????'][j]
                                        else:
                                            break
                    for y in range(0, maxNormalCommaCnt + 1):
                        if str(dfMergeResultSfReset['????????????' + str(y + 1)][i]) != '' and str(dfMergeResultSfReset['????????????' + str(y + 1)][i]) != 'nan':
                            limitCopyCnt = capableCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])]
                            if (math.ceil(limitCopyCnt) <= (integCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] / workDay)) and copyCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] > 0:
                                if len(minContCntDic) > 0:
                                    if dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(y + 1)][i])]['?????????'].values[0] in minContCntDic:
                                        if minContCntDic[dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(y + 1)][i])]['?????????'].values[0]][0] < (integCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] / workDay):
                                            minContCntDic[dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(y + 1)][i])]['?????????'].values[0]][0] = (integCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] / workDay)
                                            minContCntDic[dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(y + 1)][i])]['?????????'].values[0]][1] = int(dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(y + 1)][i])]['???(LINE)????????????'].values[0])
                                            minContCntDic[dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(y + 1)][i])]['?????????'].values[0]][2] = dfMergeResultSfReset['Planned Prod. Completion date'][i]
                                else:
                                    minContCntDic[dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(y + 1)][i])]['?????????'].values[0]] = [(integCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] / workDay),
                                                                                                                                                                int(dfCondition[dfCondition['No'] == float(dfMergeResultSfReset['????????????' + str(y + 1)][i])]['???(LINE)????????????'].values[0]),
                                                                                                                                                                dfMergeResultSfReset['Planned Prod. Completion date'][i]]
                            limitCnt = capableCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] * float(dfMergeResultSfReset['????????????(%)'][i])
                            if (math.ceil(limitCnt) <= (integCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] / workDay)) and preOrderDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] > 0:
                                for j in dfMergeResultSfReset.index:
                                    if str(dfMergeResultSfReset['????????????' + str(y + 1)][i]) in str(dfMergeResultSfReset['????????????'][j]):
                                        if preOrderDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] > 0:
                                            dfMergeResultSfReset['??????????????????'][j] = '??????'
                                            preOrderDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] -= dfMergeResultSfReset['????????????'][j]
                                        else:
                                            break
                if len(minContCntDic) > 0:
                    for key, value in minContCntDic.items():
                        logging.warning('???%s??? ????????? ??????????????????: %s??? ?????? ?????? ?????? ????????????: %i ???????????? ????????? ????????? ???????????????. ?????? ?????? ???????????? ???%i ?????? ?????????.',
                                                key,
                                                str(value[2]),
                                                value[1],
                                                math.ceil(value[0]))
                if self.isDebug:
                    dfMergeResultSfReset.to_excel('.\\debug\\flow10-2.xlsx')

                dfMergeResultSfReset = dfMergeResultSfReset.sort_values(by=['????????????', 'MAX ?????? ??????', '??????????????????', 'Planned Prod. Completion date', 'Planned Shipping date', '????????????', '????????????'],
                                                                            ascending=[False, False, False, True, True, True, True])
                # ????????? ?????????
                dfMergeResultSfReset = dfMergeResultSfReset.reset_index(level=None, drop=False, inplace=False)
                # ?????? ????????? ??????
                dfMergeResultSfReset.drop('index', axis=1, inplace=True)

                if self.isDebug:
                    dfMergeResultSfReset.to_excel('.\\debug\\flow10-3.xlsx')
                QApplication.processEvents()
                self.progressbar.setRange(0, dfMergeResultSfReset.index[-1])

                limitCnt = constructTempCnt / 2
                # ????????? ?????? ??????
                for i in dfMergeResultSfReset.index:
                    self.progressbar.setValue(i / 2)
                    if i == 157:
                        print()
                    if constructTempCnt > 0:
                        if dfMergeResultSfReset['????????????'][i] != '??????':
                            if dfMergeResultSfReset['????????????'][i] != '??????':
                                tempMinCnt = constructTempCnt
                                # ??????, ????????? ????????? ????????? ???????????? ?????? ?????????
                                normalFlag = True
                                # UT32A, UT52A, UM33A ????????? ????????? ??????
                                if limitFlag and (str(dfMergeResultSfReset['UT32A, UT52A, UM33A????????????'][i]) != '' and str(dfMergeResultSfReset['UT32A, UT52A, UM33A????????????'][i]) != 'nan'):
                                    if limitCnt <= tempMinCnt:
                                        tempMinCnt = limitCnt
                                # ????????? ?????????????????? ??????
                                if float(dfMergeResultSfReset['???????????????'][i]) * float(dfMergeResultSfReset['????????????'][i]) <= tempMinCnt:
                                    tempMinCnt = float(dfMergeResultSfReset['???????????????'][i]) * float(dfMergeResultSfReset['????????????'][i])
                                # ??????????????? ?????? ??? ???????????? ??????
                                for x in range(0, maxSpCommaCnt + 1):
                                    if str(dfMergeResultSfReset['????????????' + str(x + 1)][i]) != '' and str(dfMergeResultSfReset['????????????' + str(x + 1)][i]) != 'nan':
                                        normalFlag = False
                                        if capableCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] * float(dfMergeResultSfReset['????????????'][i]) <= tempMinCnt:
                                            tempMinCnt = capableCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] * float(dfMergeResultSfReset['????????????'][i])
                                # ??????????????? ?????? ??? ???????????? ??????
                                for y in range(0, maxNormalCommaCnt + 1):
                                    if str(dfMergeResultSfReset['????????????' + str(y + 1)][i]) != '' and str(dfMergeResultSfReset['????????????' + str(y + 1)][i]) != 'nan':
                                        normalFlag = False
                                        if capableCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] * float(dfMergeResultSfReset['????????????'][i]) <= tempMinCnt:
                                            tempMinCnt = capableCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] * float(dfMergeResultSfReset['????????????'][i])
                                # ?????? ???????????? ?????? ?????? ??????
                                if tempMinCnt > 0:
                                    dfMergeResultSfReset['????????????'][i] = math.ceil(tempMinCnt / float(dfMergeResultSfReset['????????????'][i]))
                                    # ?????? or ?????? ????????? ??????
                                    if not normalFlag:
                                        for x in range(0, maxSpCommaCnt + 1):
                                            if str(dfMergeResultSfReset['????????????' + str(x + 1)][i]) != '' and str(dfMergeResultSfReset['????????????' + str(x + 1)][i]) != 'nan':
                                                if capableCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] - dfMergeResultSfReset['????????????'][i] >= 0:
                                                    capableCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] -= dfMergeResultSfReset['????????????'][i]
                                                else:
                                                    capableCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] = 0
                                        for y in range(0, maxNormalCommaCnt + 1):
                                            if str(dfMergeResultSfReset['????????????' + str(y + 1)][i]) != '' and str(dfMergeResultSfReset['????????????' + str(y + 1)][i]) != 'nan':
                                                if capableCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] - dfMergeResultSfReset['????????????'][i] >= 0:
                                                    capableCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] -= dfMergeResultSfReset['????????????'][i]
                                                else:
                                                    capableCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] = 0
                                    # UT32A, UT52A, UM33A ????????? ????????? ??????Cnt????????? ??????
                                    if (str(dfMergeResultSfReset['UT32A, UT52A, UM33A????????????'][i]) != '' and str(dfMergeResultSfReset['UT32A, UT52A, UM33A????????????'][i]) != 'nan'):
                                        if limitCnt > 0:
                                            limitCnt -= dfMergeResultSfReset['????????????'][i] * float(dfMergeResultSfReset['????????????'][i])
                                            constructTempCnt -= dfMergeResultSfReset['????????????'][i] * float(dfMergeResultSfReset['????????????'][i])
                                            dfCopy = dfCopy.append(dfMergeResultSfReset.iloc[i])
                                    else:
                                        constructTempCnt -= dfMergeResultSfReset['????????????'][i] * float(dfMergeResultSfReset['????????????'][i])
                                        dfCopy = dfCopy.append(dfMergeResultSfReset.iloc[i])
                            # ?????? ?????? ??? ??????
                            else:
                                dfMergeResultSfReset['????????????'][i] = dfMergeResultSfReset['???????????????'][i]
                                # ??????????????? ?????? ??? ???????????? ?????? (?????? ??? ????????? ??? ???, ???????????? ???????????? ????????????)
                                for x in range(0, maxSpCommaCnt + 1):
                                    if str(dfMergeResultSfReset['????????????' + str(x + 1)][i]) != '' and str(dfMergeResultSfReset['????????????' + str(x + 1)][i]) != 'nan':
                                        if capableCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] - float(dfMergeResultSfReset['????????????'][i]) >= 0:
                                            capableCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] -= float(dfMergeResultSfReset['????????????'][i])
                                        else:
                                            capableCntDic[float(dfMergeResultSfReset['????????????' + str(x + 1)][i])] = 0
                                # ??????????????? ?????? ??? ???????????? ?????? (?????? ??? ????????? ??? ???, ???????????? ???????????? ????????????)
                                for y in range(0, maxNormalCommaCnt + 1):
                                    if str(dfMergeResultSfReset['????????????' + str(y + 1)][i]) != '' and str(dfMergeResultSfReset['????????????' + str(y + 1)][i]) != 'nan':
                                        if capableCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] - float(dfMergeResultSfReset['????????????'][i]) >= 0:
                                            capableCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] -= float(dfMergeResultSfReset['????????????'][i])
                                        else:
                                            capableCntDic[float(dfMergeResultSfReset['????????????' + str(y + 1)][i])] = 0
                                # UT32A, UT52A, UM33A ????????? ??????, ??????Cnt?????? ?????? (?????? ??? ????????? ??? ???, ???????????? ???????????? ????????????)
                                if (str(dfMergeResultSfReset['UT32A, UT52A, UM33A????????????'][i]) != '' and str(dfMergeResultSfReset['UT32A, UT52A, UM33A????????????'][i]) != 'nan'):
                                    limitCnt -= float(dfMergeResultSfReset['????????????'][i]) * float(dfMergeResultSfReset['????????????'][i])
                                constructTempCnt -= float(dfMergeResultSfReset['????????????'][i]) * float(dfMergeResultSfReset['????????????'][i])
                                dfCopy = dfCopy.append(dfMergeResultSfReset.iloc[i])
                    else:
                        self.progressbar.setValue(dfMergeResultSfReset.index[-1] / 2)
                        break
                # ???????????? ?????? ??????
                if self.isDebug:
                    dfCopy.to_excel('.\\debug\\flow11.xlsx')
                dfCopy['????????????*??????'] = dfCopy['????????????'] * dfCopy['????????????']
                # ???????????? ?????? ??????
                if self.isDebug:
                    dfCopy.to_excel('.\\debug\\flow12.xlsx')
                # ?????? ????????? = 4
                cycleStd = 4
                totalOrderCntDic = {}
                cycleStdDic = {}
                # ????????? ????????? ???????????? ?????? ??????
                for i in dfCopy.index:
                    if str(dfCopy['????????????'][i]) != '' and str(dfCopy['????????????'][i]) != 'nan':
                        for j in range(1, maxSpCommaCnt + 2):
                            # ??????????????? ?????? ??????
                            if str(dfCopy['????????????' + str(j)][i]) != 'nan' and str(dfCopy['????????????' + str(j)][i]) != '':
                                if dfCopy['????????????' + str(j)][i] in totalOrderCntDic:
                                    totalOrderCntDic[dfCopy['????????????' + str(j)][i]] += dfCopy['????????????'][i]
                                else:
                                    totalOrderCntDic[dfCopy['????????????' + str(j)][i]] = dfCopy['????????????'][i]
                                    cycleStdDic[dfCopy['????????????' + str(j)][i]] = dfCopy['Cycle??????'][i]
                                # normalSpec_flag = False
                        for j in range(1, maxNormalCommaCnt + 2):
                            # ???????????? + ???????????? ?????? ??????
                            if str(dfCopy['????????????' + str(j)][i]) != 'nan' and str(dfCopy['????????????' + str(j)][i]) != '':
                                # normalSpec_flag = False
                                if j == 1:
                                    if dfCopy['????????????' + str(j)][i] in totalOrderCntDic:
                                        totalOrderCntDic[dfCopy['????????????' + str(j)][i]] += dfCopy['????????????'][i]
                                    else:
                                        totalOrderCntDic[dfCopy['????????????' + str(j)][i]] = dfCopy['????????????'][i]
                                        cycleStdDic[dfCopy['????????????' + str(j)][i]] = dfCopy['Cycle??????'][i]
                    # ?????? ????????? ?????? ??????,
                    elif str(dfCopy['????????????'][i]) != '' and str(dfCopy['????????????'][i]) != 'nan':
                        for j in range(1, maxNormalCommaCnt + 2):
                            if str(dfCopy['????????????' + str(j)][i]) != 'nan' and str(dfCopy['????????????' + str(j)][i]) != '':
                                # normalSpec_flag = False
                                if j == 1:
                                    if dfCopy['????????????' + str(j)][i] in totalOrderCntDic:
                                        totalOrderCntDic[dfCopy['????????????' + str(j)][i]] += dfCopy['????????????'][i]
                                    else:
                                        totalOrderCntDic[dfCopy['????????????' + str(j)][i]] = dfCopy['????????????'][i]
                                        cycleStdDic[dfCopy['????????????' + str(j)][i]] = dfCopy['Cycle??????'][i]
                    # ????????????, ???????????? ?????? ?????? ?????? ??????
                    elif str(dfCopy['????????????'][i]) == '' or str(dfCopy['????????????'][i]) == 'nan':
                        if '??????' in totalOrderCntDic:
                            totalOrderCntDic['??????'] += dfCopy['????????????'][i]
                        else:
                            totalOrderCntDic['??????'] = dfCopy['????????????'][i]
                            cycleStdDic['??????'] = dfCopy['Cycle??????'][i]
                for key_i, value_i in cycleStdDic.items():
                    if value_i != 0 and str(value_i) != 'nan':
                        for key_j, value_j in totalOrderCntDic.items():
                            if key_i == key_j:
                                if math.ceil(value_j / value_i) > cycleStd:
                                    cycleStd = math.ceil(value_j / value_i)
                # Linkage Number ???????????? ??????
                dfCopy = dfCopy.astype({'Linkage Number': 'str'})
                dfLevelingResult = dfLevelingResult.astype({'Linkage Number': 'str'})
                dfMergeOrder = pd.merge(dfCopy, dfLevelingResult, on='Linkage Number', how='left')
                # ???????????? ?????? ??????
                if self.isDebug:
                    dfMergeOrder.to_excel('.\\debug\\flow13.xlsx')
                # ????????? ?????????
                dfMergeOrderResult = pd.DataFrame().reindex_like(dfMergeOrder)
                dfMergeOrderResult = dfMergeOrderResult[0:0]
                self.progressbar.setRange(0, (dfMergeResultSfReset.index[-1] / 2) + dfCopy.index[-1])
                # ?????? ???????????? ??????
                for i in dfCopy.index:
                    self.progressbar.setValue((dfMergeResultSfReset.index[-1] / 2) + i)
                    for j in dfMergeOrder.index:
                        if dfCopy['Linkage Number'][i] == dfMergeOrder['Linkage Number'][j]:
                            if j > 0:
                                if dfMergeOrder['Linkage Number'][j] != dfMergeOrder['Linkage Number'][j - 1]:
                                    orderCnt = int(dfCopy['????????????'][i])
                            else:
                                orderCnt = int(dfCopy['????????????'][i])
                            if orderCnt > 0:
                                dfMergeOrderResult = dfMergeOrderResult.append(dfMergeOrder.iloc[j])
                                orderCnt -= 1
                            # total_orderCnt -= 1
                # ????????? ?????????
                dfMergeOrderResult.reset_index(drop=True)
                # ???????????? ?????? ??????
                if self.isDebug:
                    dfMergeOrderResult.to_excel('.\\debug\\flow14.xlsx')
                # ????????? ????????? ?????? ?????? ??????
                cycleList = []
                cycleOrderCntDic = {}
                dfMergeOrderResult['???????????????'] = 0
                for key, value in totalOrderCntDic.items():
                    cycleOrderCntDic[key] = math.ceil(value / cycleStd)
                for i in range(0, cycleStd):
                    cycleList.append(cycleOrderCntDic.copy())
                for i in dfMergeOrderResult.index:
                    for j in range(0, cycleStd):
                        normalSpec = ""
                        # print(str(dfMergeOrderResult['????????????'][i]))
                        if str(dfMergeOrderResult['????????????'][i]) == '??????' or len(str(dfMergeOrderResult['????????????'][i])) != 0:
                            dfMergeOrderResult['???????????????'][i] = 0
                        else:
                            if str(dfMergeOrderResult['????????????1'][i]) != 'nan' and str(dfMergeOrderResult['????????????1'][i]) != '':
                                normalSpec = dfMergeOrderResult['????????????1'][i]
                            elif str(dfMergeOrderResult['????????????1'][i]) == 'nan' and str(dfMergeOrderResult['????????????1'][i]) == 'nan':
                                normalSpec = '??????'
                            elif str(dfMergeOrderResult['????????????1'][i]) == '' and str(dfMergeOrderResult['????????????1'][i]) == '':
                                normalSpec = '??????'
                            if normalSpec != "" and cycleList[j][normalSpec] != 0:
                                dfMergeOrderResult['???????????????'][i] = j + 1
                                cycleList[j][normalSpec] -= 1
                                break
                dfMergeOrderResult = dfMergeOrderResult.sort_values(by=['????????????', '???????????????'], ascending=[False, True])
                dfMergeOrderResult = dfMergeOrderResult.sort_values(by=['????????????', '???????????????', '????????????', 'MS-CODE', 'Linkage Number', 'Planned Prod. Completion date', '????????????'],
                                                                        ascending=[False, True, True, True, True, True, True])
                dfMergeOrderResult = dfMergeOrderResult.reset_index(drop=True)
                dfMergeOrderResult['No (*)'] = (dfMergeOrderResult.index.astype(int) + 1) * 10
                dfMergeOrderResult['Scheduled Start Date (*)'] = self.labelDate.text()
                dfMergeOrderResult['Planned Order'] = dfMergeOrderResult['Planned Order'].astype(int).astype(str).str.zfill(10)
                dfMergeOrderResult['Scheduled End Date'] = dfMergeOrderResult['Scheduled End Date'].astype(str).str.zfill(10)
                dfMergeOrderResult['Specified Start Date'] = dfMergeOrderResult['Specified Start Date'].astype(str).str.zfill(10)
                dfMergeOrderResult['Specified End Date'] = dfMergeOrderResult['Specified End Date'].astype(str).str.zfill(10)
                dfMergeOrderResult['Spec Freeze Date'] = dfMergeOrderResult['Spec Freeze Date'].astype(str).str.zfill(10)
                dfMergeOrderResult['Component Number'] = dfMergeOrderResult['Component Number'].astype(int).astype(str).str.zfill(4)

                # ???????????? ?????? ??????
                if self.isDebug:
                    dfMergeOrderResult.to_excel('.\\debug\\flow15.xlsx')
                dfMergeOrderResult = dfMergeOrderResult[['No (*)',
                                                        'Sequence No',
                                                        'Production Order',
                                                        'Planned Order',
                                                        'Manual',
                                                        'Scheduled Start Date (*)',
                                                        'Scheduled End Date',
                                                        'Specified Start Date',
                                                        'Specified End Date',
                                                        'Demand destination country',
                                                        'MS-CODE',
                                                        'Allocate',
                                                        'Spec Freeze Date',
                                                        'Linkage Number',
                                                        'Order Number',
                                                        'Order Item',
                                                        'Combination flag',
                                                        'Project Definition',
                                                        'Error message',
                                                        'Leveling Group',
                                                        'Leveling Class',
                                                        'Planning Plant',
                                                        'Component Number',
                                                        'Serial Number']]
                # ???????????? ?????? ??????
                if self.isDebug:
                    dfMergeOrderResult.to_excel('.\\debug\\flow16.xlsx')
                # ??????????????? ??????
                outputFile = '.\\result\\5400_A0100A81_' + today + '_Leveling_List.xlsx'
                dfMergeOrderResult.to_excel(outputFile, index=False)
                logging.info('???????????? %s ????????? ??????????????? ?????????????????????.', outputFile)
                self.runBtn.setEnabled(True)
            elif int(self.maxOrderinput.text()) == 0:
                logging.warning('?????? ???????????? 0?????????. ?????? ??????????????????.')
                self.runBtn.setEnabled(True)
            elif len(self.maxOrderinput.text()) == 0:
                logging.warning('?????? ???????????? ???????????? ???????????????.')
                self.runBtn.setEnabled(True)
        except Exception as e:
            logging.exception(e, exc_info=True)
            self.runBtn.setEnabled(True)


if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ui = Ui_MainWindow()
    sys.exit(app.exec_())
