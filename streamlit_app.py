import streamlit as st
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog,QMessageBox,QTableWidgetItem,QApplication,QProgressBar,QVBoxLayout
from PyQt5.QtCore import QThread,pyqtSignal,QObject
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as chrome_service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchWindowException,SessionNotCreatedException,TimeoutException,WebDriverException,UnexpectedAlertPresentException
from subprocess import CREATE_NO_WINDOW
import time
import logging
import pandas as pd
from types import SimpleNamespace
import re
from datetime import date
import socket

# Show title and description.
st.title("💬 Chatbot")
st.write(
    "tes"
)

logging.basicConfig(filename="log_file.log",filemode="w",format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger()
logger.setLevel(logging.INFO)

class Ui_MainWindow(QObject):
    global data_crm,element
    data_crm_ready = pyqtSignal(list)
    hide_browser_sinyal = pyqtSignal(bool)
    data_crm=[]
    element = SimpleNamespace(
        sales = "",
        transaksi = "",
        pekerjaan = "",
        opportunity = "",
        followUp = "",
        tgl = "",
        note = "",
        user = "",
        password = ""
	)
    def __init__(self):
        super().__init__()
        self.is_running = False
        self.proses_data = None
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(965, 635)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/icon/crm.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setAutoFillBackground(False)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(0, 0, 971, 61))
        self.frame.setStyleSheet("\n"
"background-color: rgb(3, 221, 138);")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.label_7 = QtWidgets.QLabel(self.frame)
        self.label_7.setGeometry(QtCore.QRect(90, 10, 231, 31))
        self.label_7.setObjectName("label_7")
        self.pushButton = QtWidgets.QPushButton(self.frame)
        self.pushButton.setGeometry(QtCore.QRect(20, 0, 75, 61))
        self.pushButton.setText("")
        self.pushButton.setIcon(icon)
        self.pushButton.setIconSize(QtCore.QSize(40, 40))
        self.pushButton.setFlat(True)
        self.pushButton.setObjectName("pushButton")
        self.label_8 = QtWidgets.QLabel(self.frame)
        self.label_8.setGeometry(QtCore.QRect(870, 20, 131, 20))
        self.label_8.setObjectName("label_8")
        self.pushButton_2 = QtWidgets.QPushButton(self.frame)
        self.pushButton_2.setGeometry(QtCore.QRect(910, 0, 75, 61))
        self.pushButton_2.setText("")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/icon/no-signal.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton_2.setIcon(icon1)
        self.pushButton_2.setIconSize(QtCore.QSize(20, 20))
        self.pushButton_2.setFlat(True)
        self.pushButton_2.setObjectName("pushButton_2")
        self.lineEdit_crm = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_crm.setGeometry(QtCore.QRect(10, 80, 521, 25))
        self.lineEdit_crm.setMinimumSize(QtCore.QSize(521, 0))
        self.lineEdit_crm.setStyleSheet("QLineEdit {\n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 5px;\n"
"    color:rgb(100, 108, 122)\n"
"}\n"
"")
        self.lineEdit_crm.setReadOnly(True)
        self.lineEdit_crm.setClearButtonEnabled(False)
        self.lineEdit_crm.setObjectName("lineEdit_crm")
        self.btn_crm = QtWidgets.QPushButton(self.centralwidget)
        self.btn_crm.setGeometry(QtCore.QRect(540, 80, 91, 25))
        self.btn_crm.setStyleSheet("QPushButton {\n"
"    color: rgb(255, 255, 255);\n"
"    background-color: rgb(3, 221, 138);\n"
"    border-radius: 5px\n"
"}\n"
"QPushButton:hover{\n"
"    background-color: rgb(87, 238, 183);\n"
"}")
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("icons/cil-magnifying-glass.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_crm.setIcon(icon2)
        self.btn_crm.setObjectName("btn_crm")
        self.btn_crm.clicked.connect(self.import_crm)
        self.lineEdit_wa = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_wa.setGeometry(QtCore.QRect(10, 110, 521, 25))
        self.lineEdit_wa.setMinimumSize(QtCore.QSize(521, 0))
        self.lineEdit_wa.setStyleSheet("QLineEdit {\n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 5px;\n"
"    color:rgb(100, 108, 122)\n"
"}\n"
"")
        self.lineEdit_wa.setReadOnly(True)
        self.lineEdit_wa.setClearButtonEnabled(False)
        self.lineEdit_wa.setObjectName("lineEdit_wa")
        self.btn_wa = QtWidgets.QPushButton(self.centralwidget)
        self.btn_wa.setGeometry(QtCore.QRect(540, 110, 91, 25))
        self.btn_wa.setStyleSheet("QPushButton {\n"
"    color: rgb(255, 255, 255);\n"
"    background-color: rgb(3, 221, 138);\n"
"    border-radius: 5px\n"
"}\n"
"QPushButton:hover{\n"
"    background-color: rgb(87, 238, 183);\n"
"}")
        self.btn_wa.setIcon(icon2)
        self.btn_wa.setObjectName("btn_wa")
        self.btn_wa.clicked.connect(self.import_wa)
        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(10, 140, 621, 431))
        self.tableWidget.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"border-radius:5px;")
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(10, 580, 621, 23))
        self.progressBar.setStyleSheet("border: 1px solid #03DD8A;\n"
"border-radius: 5px;\n"
"text-align: center;\n"
"background-color: rgb(255, 255, 255);\n"
"width: 10px;\n"
"color: rgb(100, 108, 122);")
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.groupBox_source = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_source.setGeometry(QtCore.QRect(640, 140, 271, 271))
        self.groupBox_source.setStyleSheet("color: rgb(100, 108, 122)")
        self.groupBox_source.setObjectName("groupBox_source")
        self.combo_sales = QtWidgets.QComboBox(self.groupBox_source)
        self.combo_sales.setGeometry(QtCore.QRect(90, 20, 171, 25))
        self.combo_sales.setStyleSheet("color: rgb(100, 108, 122);\n"
"background-color: rgb(255, 255, 255);\n"
"border-radius: 5px;")
        self.combo_sales.setObjectName("combo_sales")
        self.combo_type_transaksi = QtWidgets.QComboBox(self.groupBox_source)
        self.combo_type_transaksi.setGeometry(QtCore.QRect(90, 50, 171, 25))
        self.combo_type_transaksi.setStyleSheet("color: rgb(100, 108, 122);\n"
"background-color: rgb(255, 255, 255);\n"
"border-radius: 5px;")
        self.combo_type_transaksi.setObjectName("combo_type_transaksi")
        self.combo_pekerjaan = QtWidgets.QComboBox(self.groupBox_source)
        self.combo_pekerjaan.setGeometry(QtCore.QRect(90, 80, 171, 25))
        self.combo_pekerjaan.setStyleSheet("color: rgb(100, 108, 122);\n"
"background-color: rgb(255, 255, 255);\n"
"border-radius: 5px;")
        self.combo_pekerjaan.setObjectName("combo_pekerjaan")
        self.combo_oppotunity = QtWidgets.QComboBox(self.groupBox_source)
        self.combo_oppotunity.setGeometry(QtCore.QRect(90, 110, 171, 25))
        self.combo_oppotunity.setStyleSheet("color: rgb(100, 108, 122);\n"
"background-color: rgb(255, 255, 255);\n"
"border-radius: 5px;")
        self.combo_oppotunity.setObjectName("combo_oppotunity")

        self.combo_followUp = QtWidgets.QComboBox(self.groupBox_source)
        self.combo_followUp.setGeometry(QtCore.QRect(90, 140, 171, 25))
        self.combo_followUp.setStyleSheet("color: rgb(100, 108, 122);\n"
"background-color: rgb(255, 255, 255);\n"
"border-radius: 5px;")
        self.combo_followUp.setObjectName("combo_followUp")

        self.label_2 = QtWidgets.QLabel(self.groupBox_source)
        self.label_2.setGeometry(QtCore.QRect(10, 20, 71, 16))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.groupBox_source)
        self.label_3.setGeometry(QtCore.QRect(10, 50, 71, 16))
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.groupBox_source)
        self.label_4.setGeometry(QtCore.QRect(10, 80, 71, 16))
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.groupBox_source)
        self.label_5.setGeometry(QtCore.QRect(10, 110, 71, 16))
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(650, 310, 71, 16))
        self.label_6.setStyleSheet("color: rgb(100, 108, 122)")
        self.label_6.setObjectName("label_6")
        self.label_9 = QtWidgets.QLabel(self.centralwidget)
        self.label_9.setGeometry(QtCore.QRect(650, 275, 71, 22))
        self.label_9.setStyleSheet("color: rgb(100, 108, 122)")
        self.label_9.setObjectName("label_6")

        self.dateEdit = QtWidgets.QDateEdit(self.centralwidget)
        self.dateEdit.setGeometry(QtCore.QRect(730, 310, 171, 25))
        self.dateEdit.setStyleSheet("color: rgb(100, 108, 122);\n"
"background-color: rgb(255, 255, 255);\n"
"border-radius: 5px;\n"
"")
        self.dateEdit.setMinimumDate(QtCore.QDate(2024, 3, 14))
        self.dateEdit.setCalendarPopup(True)
        self.dateEdit.setObjectName("dateEdit")
        self.dateEdit.setDate(date.today())
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit.setGeometry(QtCore.QRect(650, 350, 251, 50))
        self.plainTextEdit.setStyleSheet("color: rgb(100, 108, 122);\n"
"background-color: rgb(255, 255, 255);\n"
"border-radius: 5px")
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.groupBox_user = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_user.setGeometry(QtCore.QRect(640, 420, 271, 91))
        self.groupBox_user.setStyleSheet("color: rgb(100, 108, 122)")
        self.groupBox_user.setObjectName("groupBox_user")
        self.lineEdit_user = QtWidgets.QLineEdit(self.groupBox_user)
        self.lineEdit_user.setGeometry(QtCore.QRect(10, 20, 251, 25))
        self.lineEdit_user.setStyleSheet("QLineEdit {\n"
"    background-color: rgb(255, 255, 255);\n"
"    color: rgb(100, 108, 122);\n"
"    border-radius: 5px;    \n"
"}\n"
"")
        self.lineEdit_user.setReadOnly(False)
        self.lineEdit_user.setClearButtonEnabled(True)
        self.lineEdit_user.setObjectName("lineEdit_user")
        self.lineEdit_password = QtWidgets.QLineEdit(self.groupBox_user)
        self.lineEdit_password.setGeometry(QtCore.QRect(10, 50, 251, 25))
        self.lineEdit_password.setStyleSheet("QLineEdit {\n"
"    background-color: rgb(255, 255, 255);\n"
"    color: rgb(100, 108, 122);\n"
"    border-radius: 5px;    \n"
"}\n"
"")
        self.lineEdit_password.setReadOnly(False)
        self.lineEdit_password.setClearButtonEnabled(True)
        self.lineEdit_password.setObjectName("lineEdit_password")
        self.btn_stop = QtWidgets.QPushButton(self.centralwidget)
        self.btn_stop.setGeometry(QtCore.QRect(790, 520, 121, 25))
        self.btn_stop.setStyleSheet("QPushButton {\n"
"    color: rgb(255, 255, 255);\n"
"    background-color: rgb(3, 221, 138);\n"
"    border-radius: 5px\n"
"}\n"
"QPushButton:hover{\n"
"    background-color: rgb(87, 238, 183);\n"
"}")
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("icons/cil-media-stop.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_stop.setIcon(icon3)
        self.btn_stop.setObjectName("btn_stop")
        self.btn_run = QtWidgets.QPushButton(self.centralwidget)
        self.btn_run.setGeometry(QtCore.QRect(640, 520, 131, 25))
        self.btn_run.setStyleSheet("QPushButton {\n"
"    color: rgb(255, 255, 255);\n"
"    background-color: rgb(3, 221, 138);\n"
"    border-radius: 5px\n"
"}\n"
"QPushButton:hover{\n"
"    background-color: rgb(87, 238, 183);\n"
"}")
        icon4 = QtGui.QIcon()
        icon4.addPixmap(QtGui.QPixmap("icons/cil-media-play.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_run.setIcon(icon4)
        self.btn_run.setObjectName("btn_run")
        self.btn_run.clicked.connect(self.toggle)
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(20, 610, 300, 16))
        self.label.setStyleSheet("color: rgb(100, 108, 122);")
        self.label.setObjectName("label")

        self.checkBox = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox.setGeometry(QtCore.QRect(645, 560, 91, 17))
        self.checkBox.setStyleSheet("color: rgb(100, 108, 122);")
        self.checkBox.setObjectName("checkBox")
        self.checkBox.stateChanged.connect(self.hide_browser)

        MainWindow.setCentralWidget(self.centralwidget)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
    def hide_browser(self):
        pass
    def message(self,judul,pesan):
        msg = QMessageBox()
        msg.setWindowTitle(judul)
        msg.setText(pesan)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setDefaultButton(QMessageBox.Ok)
        msg.exec_()
    def validasi_hp(self,phone):
        if str(phone).startswith("0"):
            return str(phone).replace("0","62",1)
        elif str(phone)[:1] != "0" or str(phone)[:2] != "62" :
            return "62" + str(phone)
        else:
            return str(phone)
    def import_crm(self):
        data_crm = []
        path = QFileDialog.getOpenFileName(None,'Open File','','Excel File(*.xls *.xlsx)')
        try :
            if path :
                self.lineEdit_crm.setText(str(path[0]).replace('/','\\'))
                dt = pd.read_excel(self.lineEdit_crm.text(),skiprows=4)
                crm = dt[['Subject','Customer','Phone']].dropna(how="all")
                crm['Phone'] = crm['Phone'].apply(lambda x:self.validasi_hp(x))
                self.tableWidget.setRowCount(len(crm))
                self.tableWidget.setColumnCount(5)
                lebar_kolom = [150, 170, 100, 80, 90]
                header = ["Subject", "Customer", "Number", "Status FU", "Reason"]
                for col,lebar in enumerate(lebar_kolom):
                    self.tableWidget.setColumnWidth(col,lebar)
                self.tableWidget.setHorizontalHeaderLabels(header)
                for row in range(len(crm)):
                    for col in range(3):
                        item = str(crm.iloc[row,col]).replace(".0","")
                        self.tableWidget.setItem(row,col,QTableWidgetItem(item))
                        data_crm.append(item)
                self.data_crm_ready.emit(data_crm)
                self.label.setText(f'Jumlah Data : {len(crm)}')
        except Exception as e :
            self.message('Error',str(e))
            self.lineEdit_crm.setText("")
    def import_wa(self):
        self.progressBar.setMinimum(0)
        self.progressBar.setMaximum(self.tableWidget.rowCount())
        path = QFileDialog.getOpenFileName(None,'Open File','','Excel File(*.xls *.xlsx)')
        try :
            if path :
                self.lineEdit_wa.setText(str(path[0]).replace('/','\\'))
                dt = pd.read_excel(self.lineEdit_wa.text())
                wa = dt[['Destination','Status']].dropna(how="all")
                #vlookup
                for row_crm in range(self.tableWidget.rowCount()):
                    phone_crm = self.tableWidget.item(row_crm,2).text()
                    match_found = False
                    for row_wa in range(len(wa)):
                        dest_wa = str(wa.iloc[row_wa]['Destination'])
                        QApplication.processEvents()
                        if phone_crm == dest_wa:
                            match_found = True
                            self.tableWidget.setItem(row_crm,3,QTableWidgetItem(str(wa.iloc[row_wa]['Status'])))
                            self.tableWidget.selectRow(row_crm)
                    if not match_found:
                        self.tableWidget.setItem(row_crm,3,QTableWidgetItem("N/A"))
                        self.tableWidget.selectRow(row_crm) 
                    self.progressBar.setValue(row_crm + 1)
                    self.label.setText("Mohon menunggu,sedang proses Validasi...")
                row = self.tableWidget.rowCount()
                tanpa_data = 0
                for r in reversed(range(row)):
                    if self.tableWidget.item(r,3).text() =="N/A" :
                        tanpa_data += 1
                        self.tableWidget.removeRow(r)
                    else:
                        dt_crm = self.tableWidget.item(r,0).text()
                        dt_status = self.tableWidget.item(r,3).text()
                        data_crm.append([dt_crm,dt_status])
                        
                self.message("Konfirmasi",f"{tanpa_data} Data tanpa status di hapus")
                self.label.setText(f'Jumlah Data : {self.tableWidget.rowCount()}')
                self.progressBar.setValue(0)
                self.tableWidget.selectRow(0)
        except Exception as e :
            self.message('Error',str(e))
            self.lineEdit_wa.setText("")
    def toggle(self):
        if self.is_running :
            self.is_running = False 
            self.btn_run.setText("Run")
        else:
            self.is_running = True
            self.btn_run.setText("Pause")
            icon4 = QtGui.QIcon()
            icon4.addPixmap(QtGui.QPixmap("icons/cil-media-pause.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
            self.btn_run.setIcon(icon4)
            self.process()
    def process(self):
        element.sales = self.combo_sales.currentText()
        element.transaksi = self.combo_type_transaksi.currentText()
        element.pekerjaan = self.combo_pekerjaan.currentText()
        element.opportunity = self.combo_oppotunity.currentText()
        element.followUp = self.combo_followUp.currentText()
        element.tgl = self.dateEdit.date()
        element.note = self.plainTextEdit.toPlainText()
        element.user = self.lineEdit_user.text()
        element.password = self.lineEdit_password.text()
        #Kirim data ke Qthread
        proses_data = ProsesData(data_crm)
        proses_data.update_status_sinyal.connect(self.update_status)
        proses_data.message_sinyal.connect(self.message)
        proses_data.start()
        while proses_data.is_running :
            QApplication.processEvents()
    def update_status(self,row,text):
        self.tableWidget.setItem(row,4,QTableWidgetItem(text))
        print(f"sinyal {text} baris {row + 1}")
        self.tableWidget.selectRow(row)
        self.progressBar.setValue(int((row + 1)/ self.tableWidget.rowCount())*100)
        print(self.tableWidget.rowCount())
    def delete_failed(self):
        r = self.tableWidget.rowCount()
        for i in range(r):
            if self.tableWidget.item(r,4).text() == "Done" :
                self.tableWidget.removeRow(i)
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Auto Input CRM"))
        self.label_7.setText(_translate("MainWindow", "<html><head/><body><p><span style=\" font-size:18pt; font-weight:600; color:#ffffff;\">AUTO INPUT CRM</span></p></body></html>"))
        self.label_8.setText(_translate("MainWindow", "<html><head/><body><p><span style=\" color:#ffffff;\">Connection</span></p></body></html>"))
        self.lineEdit_crm.setToolTip(_translate("MainWindow", "<html><head/><body><p>Laporan CRM</p></body></html>"))
        self.lineEdit_crm.setPlaceholderText(_translate("MainWindow", "Laporan CRM"))
        self.btn_crm.setText(_translate("MainWindow", "Browse"))
        self.lineEdit_wa.setToolTip(_translate("MainWindow", "<html><head/><body><p>Laporan WA</p></body></html>"))
        self.lineEdit_wa.setPlaceholderText(_translate("MainWindow", "laporan Broadcast"))
        self.btn_wa.setText(_translate("MainWindow", "Browse"))
        self.groupBox_source.setTitle(_translate("MainWindow", "Source"))
        self.label_2.setText(_translate("MainWindow", "Sales Source"))
        self.label_3.setText(_translate("MainWindow", "Tipe Transaksi"))
        self.label_4.setText(_translate("MainWindow", "Pekerjaan"))
        self.label_5.setText(_translate("MainWindow", "Opportunity"))
        self.label_6.setText(_translate("MainWindow", "Tanggal FU"))
        self.label_9.setText(_translate("MainWindow", "Type FollowUp"))
        self.dateEdit.setDisplayFormat(_translate("MainWindow", "dd MMMM yyyy"))
        self.plainTextEdit.setPlaceholderText(_translate("MainWindow", "Note (additional) :"))
        self.groupBox_user.setTitle(_translate("MainWindow", "User & Password"))
        self.lineEdit_user.setPlaceholderText(_translate("MainWindow", "User"))
        self.lineEdit_password.setPlaceholderText(_translate("MainWindow", "Password"))
        self.btn_stop.setText(_translate("MainWindow", "Stop"))
        self.btn_run.setText(_translate("MainWindow", "Run"))
        self.label.setText(_translate("MainWindow", "Progress :"))

        sales = ["CRM", "CHANEL","EVENT","SOCIAL MEDIA","WALK IN","CANVASING"]
        self.combo_sales.addItems(sales)
        source_transaksi = ["Reguler","PIC","PAC","Group Customer","Portal"]
        self.combo_type_transaksi.addItems(source_transaksi)
        source_pekerjaan = ["lain-lain", "Pelajar", "Ojek", "Wiraswasta/pedagang"]
        self.combo_pekerjaan.addItems(source_pekerjaan)
        source_opportunity = ["MEDIUM PROSPECT","HOT PROSPECT", "COLD PROSPECT"]
        self.combo_oppotunity.addItems(source_opportunity)
        source_followup = ["Follow Up 1", "Follow Up 2", "Follow Up 3"]
        self.combo_followUp.addItems(source_followup)
        self.checkBox.setText(_translate("MainWindow", "Hide Browser"))

import resources

class ProsesData(QThread):
    request_data_signal = pyqtSignal()
    proses_data_signal = pyqtSignal(list)
    update_status_sinyal = pyqtSignal(int,str)
    pause_sinyal = pyqtSignal()
    resume_sinyal = pyqtSignal()
    close_sinyal = pyqtSignal()
    message_sinyal = pyqtSignal(str,str)
    
    def __init__(self,data_crm):
        super(ProsesData,self).__init__()
        self.data_crm = data_crm
        self.is_running = True
        self.is_pause = False
    def run(self):
        try :
            opt = webdriver.ChromeOptions()
            opt.add_argument("--disable-gpu")
            opt.add_argument("--disable-infobar")
            opt.add_argument("--start-maximized")
            opt.add_argument("--incognito")
            service = chrome_service(ChromeDriverManager().install())
            service.creation_flags = CREATE_NO_WINDOW
            driver = webdriver.Chrome(service=service,options=opt)
        except SessionNotCreatedException :
            self.message_sinyal.emit("Info","Gagal memuat browser")
        except WebDriverException as e:
            self.message_sinyal.emit("Gagal memuat browser",str(e))
        except socket.gaierror as e :
            self.message_sinyal.emit("No Internet Connection","Sinyal Internet Terputus")
            
        wait = WebDriverWait(driver,30)
        driver.get("https://odm.daya-motor.com/web#page=0&limit=80&view_type=list&model=crm.lead&menu_id=1294&action=1408")
        user_login = wait.until(EC.presence_of_element_located((By.XPATH,"//input[@id='login']")))
        if user_login :
            user_login.send_keys(element.user)
            driver.find_element(By.XPATH,"//input[@id='password']").send_keys(element.password)
        else :
            self.message_sinyal.emit("Informasi","Gagal memuat halaman login")
            driver.quit()
        for index,i in enumerate(reversed(self.data_crm)):
            try :
                try:
                    # WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                    #     (By.XPATH, "//pre[contains(text(),'XmlHttpRequestError')]")))
                    # time.sleep(0.5)
                    # driver.find_element(By.XPATH, "//span[contains(text(),'Ok')]").click()
                    # logging.info("XmlHttpRequestError")
                    WebDriverWait(driver,15).until(EC.alert_is_present(), "Are you sure you want to leave this page ?")
                    alert = driver.switch_to.alert
                    alert.accept()
                except:
                    WebDriverWait(driver,100).until(EC.invisibility_of_element((By.CLASS_NAME,"blockUI blockOverlay")))
                    wait.until(EC.element_to_be_clickable((By.XPATH,"//div[contains(@class,'oe_searchview_clear')]")))
                    time.sleep(2)
                    driver.find_element(By.XPATH,"//div[contains(@class,'oe_searchview_clear')]").click()
                    driver.find_element(By.XPATH,"//div[@class='oe_searchview_input']").send_keys(i[0])
                    driver.find_element(By.XPATH,"//div[@class='oe_searchview_input']").send_keys(Keys.ENTER)
                    wait.until(EC.presence_of_element_located((By.XPATH,"//td[contains(text(),'"+ str(i[0]) +"')]")))
                    if driver.find_element(By.XPATH,"//tbody/tr[1]/td[8]").text not in (str(element.followUp),"Won","Lost") :
                        time.sleep(1)
                        driver.find_element(By.XPATH,"//td[contains(text(),'"+ str(i[0]) +"')]").click()
                        #Edit
                        try :
                            try:
                                wait.until(EC.presence_of_element_located((By.XPATH,"//button[contains(text(),'Edit')]")))
                                time.sleep(1)
                                driver.find_element(By.XPATH,"//button[contains(text(),'Edit')]").click()
                                time.sleep(1)
                            except :
                                ActionChains(driver).key_down(Keys.ALT).key_down(Keys.SHIFT).send_keys("e").key_up(Keys.ALT).key_up(Keys.SHIFT).perform()
                            time.sleep(1)
                            # print("on proses edit data konsumen..{}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            #Sales Location
                            sales_location = driver.find_element(By.XPATH,"//body[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[4]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[1]/table[1]/tbody[1]/tr[2]/td[2]/span[1]/div[1]/input[1]")
                            if len(sales_location.get_attribute("value")) == 0 :
                                sales_location.send_keys("RUANG TUNGGU")
                                time.sleep(0.5)
                                sales_location.send_keys(Keys.DOWN)
                                sales_location.send_keys(Keys.ENTER)
                                # print("on proses sales location..{}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            else : pass
                            #Sales activity
                            sales_source = driver.find_element(By.XPATH,"//body[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[4]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[1]/table[1]/tbody[1]/tr[3]/td[2]/span[1]/div[1]/input[1]")
                            if len(sales_source.get_attribute("value")) == 0 :
                                sales_source.send_keys(element.sales)
                                time.sleep(0.5)
                                sales_source.send_keys(Keys.DOWN)
                                sales_source.send_keys(Keys.ENTER)
                                # print("on proses sales activity..{}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            else :pass
                            time.sleep(1)
                            #tipe transaksi
                            if driver.find_elements(By.NAME,'tipe_transaksi')[0].is_displayed():
                                tipe_transaksi = Select(driver.find_elements(By.NAME,'tipe_transaksi')[0])
                                tipe_transaksi.select_by_visible_text(element.transaksi)
                                # print("on proses tipe transaksi {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            else:
                                logging.info('element tipe transaksi not found')
                            time.sleep(0.5)
                            wait.until(EC.presence_of_element_located((By.XPATH,
                                                                                    "//tr//tr[1]//td[2]//table[1]//tbody[1]//tr[8]//td[2]//span[1]//input[1]")))
                            KTP = driver.find_element(By.XPATH,
                                                                    "//tr//tr[1]//td[2]//table[1]//tbody[1]//tr[8]//td[2]//span[1]//input[1]")
                            time.sleep(1)
                            if len(KTP.get_attribute("value")) == 0:
                                KTP.send_keys("0")
                                # print("on proses KTP {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            else:
                                pass
                            try :
                                wait.until(EC.presence_of_element_located((By.XPATH,"//body[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[4]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/table[1]/tbody[1]/tr[13]/td[2]/span[1]/div[1]/span[2]")))
                                pekerjaan = driver.find_element(By.XPATH,"//body[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[4]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/table[1]/tbody[1]/tr[13]/td[2]/span[1]/div[1]/input[1]")
                                time.sleep(1)
                                pekerjaan.click()
                                time.sleep(1)
                                pekerjaan.send_keys(element.pekerjaan)
                                # print("on proses pekerjaan{}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            except Exception as x :
                                print("error pekerjaan")
                                print("Error in line {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                                logging.info("Error in line {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            #Save
                            time.sleep(1)
                            ActionChains(driver).key_down(Keys.ALT).key_down(Keys.SHIFT).send_keys("s").key_up(Keys.ALT).key_up(Keys.SHIFT).perform()
                            print("on peoses save data konsumen {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                        except Exception as e:
                            self.update_status_sinyal.emit(index,"Failed")
                            print(f"error baris : group edit =>{index}=>{i}")
                            print("Error in line {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            logging.info("Error in line {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            driver.find_element(By.LINK_TEXT,"Leads WOR").click()
                            continue
                    #Convert Opportunity
                        try :
                            convert_to_opportunity = driver.find_element(By.XPATH,"//span[contains(text(),'Convert to Opportunity')]")
                            convert_to_opportunity.click()
                            # print("on proses opportunity {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            try:
                                wait.until(EC.presence_of_element_located((By.XPATH,"// span[contains(text(), 'Create Opportunity')]")))
                                create_opportunity = driver.find_element(By.XPATH,"// span[contains(text(), 'Create Opportunity')]")
                                create_opportunity.click()
                                # print("on proses convert opportunity {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                                time.sleep(1)
                            except :
                                ActionChains(driver).key_down(Keys.ESCAPE).key_up(Keys.ESCAPE).perform()
                                # print("on proses convert opportunity {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                        except Exception as ex :
                            print("Error in line {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            self.update_status_sinyal.emit(index,"Failed")
                            print(f"error baris : group opportunity =>{index}=>{i}")
                            logging.info("Error in line {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            driver.find_element(By.LINK_TEXT,"Leads WOR").click()
                            continue
                    #Edit
                        try :
                            try :
                                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME,"blockUI blockOverlay")))
                                time.sleep(5)
                                button = driver.find_element(By.CSS_SELECTOR,"body > div.openerp.openerp_webclient_container > table > tbody > tr > td.oe_application > div > div:nth-child(2) > table > tbody > tr:nth-child(2) > td:nth-child(1) > div > div > span.oe_form_buttons_view > div > button")
                                button.click()
                                # print("on proses edit followUp {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            except :
                                time.sleep(1)
                                ActionChains(driver).key_down(Keys.ALT).key_down(Keys.SHIFT).send_keys("e").key_up(Keys.ALT).key_up(Keys.SHIFT).perform()
                                # print("on proses edit followUp {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))

                            wait.until(EC.presence_of_element_located((By.XPATH,"//select[@name='longshort']")))
                            opprtunity_status = Select(driver.find_element(By.XPATH,"//select[@name='longshort']"))
                            opprtunity_status.select_by_visible_text(element.opportunity)
                            # print("on proses opportunity {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                        except :
                            time.sleep(1)
                            driver.find_element(By.LINK_TEXT,"Leads WOR").click()
                            self.update_status_sinyal.emit(index,"Failed")
                            print(f"error baris : gorup edit opportunity =>{index}=>{i}")
                            logging.info("Error in line {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            # continue
                        #Button FollowUp
                        try :
                            angka = str(int(re.search(r'\d+', element.followUp)[0]) + 2)
                            if element.followUp == "Follow Up 1" :
                                wait.until(EC.presence_of_element_located((By.XPATH,"//header/button["+ angka +"]/span[1]")))
                                time.sleep(2)
                                followup = driver.find_element(By.XPATH,"//header/button["+ angka +"]/span[1]")
                                driver.execute_script('arguments[0].click();',followup)
                                # print("on proses button followUp {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            else:
                                wait.until(EC.presence_of_element_located((By.XPATH,"//header/button["+ angka +"]/span[1]")))
                                followup = driver.find_element(By.XPATH,"//header/button["+ angka +"]/span[1]")
                                driver.execute_script('arguments[0].click();',followup)
                                # print("on proses button followUp {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                        except Exception as err:
                            logging.info(err)
                            print("Error in line {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            logging.info("Error in line {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                            self.update_status_sinyal.emit(index,"Failed")
                            print(F"error baris : group btn followup =>{index}=>{i}")
                            # line_number += 1
                            driver.find_element(By.LINK_TEXT,"Leads WOR").click()
                            continue
                        
                        #Tanggal
                        wait.until(EC.visibility_of_all_elements_located((By.XPATH,"//input[@name='"+ str(element.followUp).lower().replace(" ","") +"']")))
                        ed = element.tgl
                        tgl = '{0}/{1}/{2}'.format(ed.month(), ed.day(), ed.year())
                        time.sleep(1)
                        driver.find_element(By.XPATH,"//input[@name='"+ str(element.followUp).lower().replace(" ","") +"']").send_keys(tgl)
                        # print("on proses tanggal followUp {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                        #Status_customer
                        time.sleep(1)
                        status_cust = Select(
                            driver.find_element(By.XPATH,"//select[@name='statuscustomer']"))
                        if str(i[1]).lower() == "sent" :
                            status_cust.select_by_index(1)
                            # print("on proses button status followUp {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                        else:
                            status_cust.select_by_index(2)
                            # print("on proses button status followUp {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                        try :
                            if element.note == "" :
                                pass
                            else :
                                wait.until(EC.presence_of_element_located((By.XPATH,"//tbody/tr[17]/td[2]/div[1]/textarea[1]")))
                                driver.find_element(By.XPATH,"//tbody/tr[16]/td[2]/div[1]/textarea[1]").send_keys(element.note)
                        except :
                            pass
                        #Save
                        try:
                            wait.until(EC.element_to_be_clickable(
                                (By.XPATH, "//div[@class='oe_view_manager_buttons']//div[1]//span[2]//button[1]")))
                            driver.find_element(By.XPATH,
                                "//div[@class='oe_view_manager_buttons']//div[1]//span[2]//button[1]").click()  # Save
                            # print("on proses save followUp {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                        except :
                            ActionChains(driver).key_down(Keys.ALT).key_down(Keys.SHIFT).send_keys("s").key_up(Keys.ALT).key_up(Keys.SHIFT).perform()
                        self.update_status_sinyal.emit(index,"Done")
                        driver.find_element(By.LINK_TEXT,"Leads WOR").click()
                        time.sleep(2)
                        continue
                    else:
                        self.update_status_sinyal.emit(index,"Done")
                        # print("on proses done {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                        continue
            except :
                print("Error in line {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                logging.info("Error in line {}{} ".format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[0]))
                self.update_status_sinyal.emit(index,"Failed")
                print(f"error baris : akhir code except =>{index}=>{i}")
                driver.find_element(By.LINK_TEXT,"Leads WOR").click()
                continue
                
        driver.quit()
        self.message_sinyal.emit("Info","Proses selesai")
        self.is_running = False
    def pause(self):
        self.is_pause = True
    def resume(self):
        self.is_pause = False
    def stop(self):
        self.is_running = False

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
