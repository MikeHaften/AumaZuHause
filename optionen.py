from PyQt5 import QtWidgets,uic
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout,QHBoxLayout, QSlider, QLabel, QFormLayout, QSpinBox,QFileDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QIcon
from PyQt5.QtWidgets import QMessageBox
import sys, os
from datetime import date

import datetime as dt
import datetime

import win32com.client as win32

import OptionenEX

import hashlib
import time
import uuid
import json
from datetime import datetime,timedelta
from cryptography.hazmat.primitives import serialization,hashes
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives.asymmetric import padding,rsa
import base64


MEINE_ADRESSE = 'mike.haftendorn@volkswagen.de'
PASSWORT = 'Ram1991MikAhnIsl'




def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
    
    


class Ui(QtWidgets.QDialog, OptionenEX.Ui_Dialog):
    def __init__(self,sqlmanager,parent=None):
        super(Ui,self).__init__(parent)
        
        self.setupUi(self)
        self.programmpfad = os.path.abspath(".")
        
        self.license_manager = LicenseManager()
        icon = QIcon(resource_path(self.programmpfad +"/icons/auma_farbe.png"))
        # Icon als Anwendungssymbol setzen
        self.setWindowIcon(icon)
        # Titelleisten-Icon temporär ändern
        self.setProperty("windowIcon", icon)
        
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint)  # Hinzufügen von Minimieren- und Maximieren-Buttons
        
        self.sqlmanager = sqlmanager
        
        self.backbutton = self.backToMainMenu
        self.backbutton.clicked.connect(self.backToMainmenu)
        self.Findbutton = self.findButton
        self.Findbutton.clicked.connect(self.finden)
        self.mailbutton = self.emailButton
        self.mailbutton.clicked.connect(self.emailtest)
        self.CloseConnectionbutton = self.closeConButton
        self.CloseConnectionbutton.clicked.connect(self.closeConnection)
        self.Connectionbutton = self.connectButton
        self.Connectionbutton.clicked.connect(self.connecttodatabase)
        self.But_SaveSysVar.clicked.connect(self.systemvariablenspeichern)
        self.matAddBut = self.pushButton_4
        self.matAddBut.clicked.connect(self.addMat)
        self.matUpdateBut = self.pushButton_3
        self.matUpdateBut.clicked.connect(self.matUpdate)
        self.matDelBut = self.pushButton_2
        self.matDelBut.clicked.connect(self.matDel)
        self.LizBut = self.pushButton_6
        self.LizBut.clicked.connect(self.generateLiz)
        self.CheckLizBut = self.pushButton_5
        self.CheckLizBut.clicked.connect(self.lizenzValid)
        self.saveFaBut = self.pushButton_7
        self.saveFaBut.clicked.connect(self.setFAParameter)
        self.aendern1But = self.pushButton_8
        self.aendern1But.clicked.connect(lambda:self.ordnerAendern(1))
        self.aendern2But = self.pushButton_9
        self.aendern2But.clicked.connect(lambda:self.ordnerAendern(2))
        self.aendern3But = self.pushButton_10
        self.aendern3But.clicked.connect(lambda:self.ordnerAendern(3))
        self.LizGenBut = self.pushButton_6
        
        self.farbwechsel = self.colorchange
        self.farbwechsel.clicked.connect(self.changecolor)
        
        self.loadDataBase = self.LoadButton
        self.loadDataBase.clicked.connect(self.loadData)
        
        self.takeData = self.pushButton
        self.takeData.clicked.connect(self.datenUebernehnmen)
        
        self.FontCombo1 = self.fontComboBox1
        self.FontCombo2 = self.fontComboBox2
        self.FontCombo3 = self.fontComboBox3
        self.FontCombo4 = self.fontComboBox4
        
        self.SpinU1 = self.spinBoxU1
        self.SpinU2 = self.spinBoxU2
        self.SpinL1 = self.spinBoxL1
        self.SpinL2 = self.spinBoxL2
        self.SpinL3 = self.spinBoxL3
        self.SpinB1 = self.spinBoxB1
        self.SpinB2 = self.spinBoxB2
        self.SpinB3 = self.spinBoxB3
        self.SpinT1 = self.spinBoxT1
        self.SpinT2 = self.spinBoxT2
        self.SpinT3 = self.spinBoxT3
        
        liste = self.sqlmanager.variableAusgeben(1)
        
        self.selector1 = RGBSelector()
        self.selector1.setRGBValue(liste[1])
        self.selector2 = RGBSelector()
        self.selector2.setRGBValue(liste[2])
        self.selector3 = RGBSelector()
        self.selector3.setRGBValue(liste[3])
        self.selector4 = RGBSelector()
        self.selector4.setRGBValue(liste[4])
        self.selector5 = RGBSelector()
        self.selector5.setRGBValue(liste[5])
        self.selector6 = RGBSelector()
        self.selector6.setRGBValue(liste[6])
        self.selector7 = RGBSelector()
        self.selector7.setRGBValue(liste[7])
        self.selector8 = RGBSelector()
        self.selector8.setRGBValue(liste[8])
        self.selector9 = RGBSelector()
        self.selector9.setRGBValue(liste[9])
        self.selector10 = QSpinBox()
        self.selector10.setValue(int(liste[10]))
        self.selector11 = QSpinBox()
        self.selector11.setValue(int(liste[11]))
        
        self.verticalLay2 = self.verticalLayout_2
        self.verticalLay3 = self.verticalLayout_3
        self.verticalLay4 = self.verticalLayout_4
        self.verticalLay5 = self.verticalLayout_5
        self.verticalLay2.addWidget(QLabel("Background"))
        self.verticalLay2.addWidget(self.selector1)
        self.verticalLay2.addWidget(QLabel("Background Frame"))
        self.verticalLay2.addWidget(self.selector2)
        self.verticalLay3.addWidget(QLabel("Background Widgets"))
        self.verticalLay3.addWidget(self.selector3)
        self.verticalLay3.addWidget(QLabel("Hover Widgets"))
        self.verticalLay3.addWidget(self.selector4)
        self.verticalLay3.addWidget(QLabel("Pressed Widgets"))
        self.verticalLay3.addWidget(self.selector5)
        self.verticalLay4.addWidget(QLabel("Border 1"))
        self.verticalLay4.addWidget(self.selector6)
        self.verticalLay4.addWidget(QLabel("Border1 Hover"))
        self.verticalLay4.addWidget(self.selector7)
        self.verticalLay4.addWidget(QLabel("Border 1 Pressed"))
        self.verticalLay4.addWidget(self.selector8)
        self.verticalLay5.addWidget(QLabel("Button Background"))
        self.verticalLay5.addWidget(self.selector9)
        self.verticalLay5.addWidget(QLabel("Border Radius 1"))
        self.verticalLay5.addWidget(self.selector10)
        self.verticalLay5.addWidget(QLabel("Border Radius 2"))
        self.verticalLay5.addWidget(self.selector11)
        
        self.tableWidget.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tableWidget.itemClicked.connect(self.matToLE)
        
        self.matID = self.lineEdit_3
        self.matArt = self.lineEdit
        self.matD = self.lineEdit_2
        self.matGew = self.lineEdit_4
        self.matPreis = self.lineEdit_5
        self.matVar1= self.lineEdit_6
        self.matVar2 = self.lineEdit_7
        self.matVar3 = self.lineEdit_8
        
        self.LiLizenzAktuell = self.lineEdit_11
        self.LiLizenzGeneriert = self.lineEdit_9
        
        #self.setStyleSheet(self.getStyle())
        
    def windowshow(self):
        status = ""
        try:
            status = "Stylesheet laden"
            #self.setStyleSheet(self.getStyle())
            status = "Fonts laden"
            self.fontdatenLaden()
            status = "Systemvariablen laden"
            self.getSystemVar()
            status = "Hilfsfunktionen anzeigen"
            self.mlUpdate()
            self.ordnerlabelsSetzen()
            self.show()  
        except:
            self.openErrorBox(("Hilfsfunktionen anzeigen fehlgeschlagen!!! Status: "+ status),sys.exc_info()[0],sys.exc_info()[1])

    def setMainmenuPointer(self,mainpointer):
        global MainmenuPointer
        MainmenuPointer = mainpointer
        
    def backToMainmenu(self):
        self.hide()
        MainmenuPointer.windowshow()
    
    def beendenButtonPressed(self):
        sys.exit(0)
        
    def closeConnection(self):
        print ("Close connection")
        self.sqlmanager.datenbanktrennen()
        
    def connecttodatabase(self):
        print("Verbindung zu Datenbank aufnehmen")
        self.sqlmanager.create_connection()
             
    def openErrorBox(self, error1,error2,error3):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("\n\n" + str(error1) + "\n\n\n" + str(error2) + "\n" + str(error3))
        msg.setWindowTitle("Es ist ein Fehler aufgetreten!!!")
        msg.exec_()
        
    ###########################################################
    ###########################################################
    ############## OPTIK TAB Funktionen #######################
    ###########################################################
    ###########################################################
    def changecolor(self):
        #print("Farbwechsel")
        self.sqlmanager.VariableVeraendern(self.selector1.getRGBvalue(),self.selector2.getRGBvalue(),self.selector3.getRGBvalue(),self.selector4.getRGBvalue(),self.selector5.getRGBvalue(),self.selector6.getRGBvalue(),self.selector7.getRGBvalue(),self.selector8.getRGBvalue(),self.selector9.getRGBvalue(),self.selector10.value(),self.selector11.value(),"","","","",1)
        
        self.setStyleSheet(self.getStyle())
    
    def getStyle(self):
        try:
            liste = self.sqlmanager.variableAusgeben(1)
            liste2 = self.sqlmanager.variableAusgeben(2)
            
            bg_dialog = liste[1]
            bg_frame = liste[2]
            bg_widgets = liste[3]
            bg_h_widgets = liste[4]
            bg_p_widgets = liste[5]
            border_1 = liste[6]
            border_h_1 = liste[7]
            border_p_1 = liste[8]
            buttons_2 = liste[9]
            border_radius1 = (str(liste[10])) + "px"
            border_radius2 = (str(liste[11])) + "px"

            fontU = liste2[1]
            fontUSize1 = liste2[2]
            fontUSize2 = liste2[3]
            fontL = liste2[4]
            fontLSize1 = (str(liste2[5])) + "px"
            fontLSize2 = liste2[6]
            fontLSize3 = liste2[7]
            fontB = liste2[8]
            fontBSize1 = (str(liste2[9])) + "px"
            fontBSize2 = liste2[10]
            fontBSize3 = liste2[11]
            fontT = liste2[12]
            fontTSize1 = (str(liste2[13])) + "px"
            fontTSize2 = liste2[14]
            fontTSize3 = liste2[15]

            
            main= '''
        QDialog {
    background: %s;
        }
        
        QFrame {
    background: %s;
    border-radius: 5px;
    border-style: solid;
    border-width: 2px;
    border-color: black;
        }''' % (bg_dialog,bg_frame)
            
            button= '''
        QPushButton {
    background: %s;
    border-radius:%s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
    font-family: %s;
    font-size: %s;   
        }
    
        QPushButton:hover {
    background:%s;
    border-style: outset;
    border-width: 2px;
    border-color: %s;
        }
    
        QPushButton:pressed {
    background:%s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
        }''' % (buttons_2, border_radius1, border_1, fontB, fontBSize1, buttons_2, border_h_1, buttons_2, border_p_1)
        
    
            lineedit= '''
        QLineEdit {
    background: %s;
    border-radius:%s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
    font-family: %s;
    font-size: %s;  
        }
    
        QLineEdit:hover {
    background: %s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
        }
    
        QLineEdit:pressed {
    background: %s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
        }'''% (bg_widgets, border_radius2, border_1, fontT, fontTSize1,bg_widgets,border_h_1,bg_widgets,border_p_1)
        
            combobox= '''
        QComboBox{
    background: %s;
    border-radius:%s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
    font-family: %s;
    font-size: %s;
    }
    
        QComboBox:hover{
    background: %s;
    border-style: outset;
    border-width: 2px;
    border-color: %s;
    }
    
        QComboBox:pressed{
    background: %s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
    }''' % (bg_widgets, border_radius2, border_1, fontT, fontTSize1,bg_widgets,border_h_1,bg_widgets,border_p_1)
    
            spinbox='''
        QSpinBox{
    background: %s;
    border-radius:%s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
    font-family: %s;
    font-size: %s;
    }
    
        QSpinBox:hover{
    background: %s;
    border-style: outset;
    border-width: 2px;
    border-color: %s;
    }
    
        QSpinBox:pressed{
    background: %s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
    }'''% (bg_widgets, border_radius2, border_1, fontT, fontTSize1,bg_widgets,border_h_1,bg_widgets,border_p_1)
    
            textedit='''
        QTextEdit{
    background: %s;
    border-radius:%s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
    font-family: %s;
    font-size: %s;
    }
    
        QTextEdit:hover{
    background: %s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
    }
    
        QTextEdit:pressed{
    background: %s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
    }'''% (bg_widgets, border_radius1, border_1, fontT, fontTSize1,bg_widgets,border_h_1,bg_widgets,border_p_1)
    
            radiobutton='''
        QRadioButton{
    background: %s;
    border-radius:%s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
    font-family: %s;
    font-size: %s;
    }
    
        QRadioButton:hover{
    background: %s;
    border-style: outset;
    border-width: 2px;
    border-color: %s;
    }
    
        QRadioButton:pressed{
    background: %s;
    border-style: solid;
    border-width: 2px;
    border-color: %s;
    }'''% (bg_widgets, border_radius2, border_1, fontT, fontTSize1,bg_widgets,border_h_1,bg_widgets,border_p_1)
    
            label='''
        QLabel{
    background: %s;
    border-width: 0px;
    border-radius:%s;
    font-family: %s;
    font-size: %s;
    }''' % (bg_widgets, border_radius2,fontL,fontLSize1)
    
            tablewidget='''
        QTableWidget{
    background: %s;
    border-radius:%s;
    border-width: 1px;
    border-style: solid;
    border-color: %s;
    }''' % (bg_widgets, border_radius1, border_1)
    
            test = main + button + lineedit + combobox + spinbox + textedit + radiobutton + label + tablewidget        
    
            return test
        except:
            print("stylesheet fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
    
    def datenUebernehnmen(self):
        print("Übernehmen")  
        try:
            font1 = self.FontCombo1.currentText()
            font2 = self.FontCombo2.currentText()
            font3 = self.FontCombo3.currentText()
            font4 = self.FontCombo4.currentText()
            
            u1 = self.SpinU1.value()
            u2 = self.SpinU2.value()
            l1 = self.SpinL1.value()
            l2 = self.SpinL2.value()
            l3 = self.SpinL3.value()
            b1 = self.SpinB1.value()
            b2 = self.SpinB2.value()
            b3 = self.SpinB3.value()
            t1 = self.SpinT1.value()
            t2 = self.SpinT2.value()
            t3 = self.SpinT3.value()
            
            self.sqlmanager.VariableVeraendern(font1,str(u1),str(u2),font2,str(l1),str(l2),str(l3),font3,str(b1),str(b2),str(b3),font4,str(t1),str(t2),str(t3),2)     
        except:
            print("datenUebernehnmen fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
    
    def fontdatenLaden(self):
        #print("Font daten laden")
        try:
            liste = self.sqlmanager.variableAusgeben(2)
            
            self.FontCombo1.setCurrentText(str(liste[1]))
            self.SpinU1.setValue(int(liste[2]))
            self.SpinU2.setValue(int(liste[3]))
            
            self.FontCombo2.setCurrentText(str(liste[4]))
            self.SpinL1.setValue(int(liste[5]))
            self.SpinL2.setValue(int(liste[6]))
            self.SpinL3.setValue(int(liste[7]))
            
            self.FontCombo3.setCurrentText(str(liste[8]))
            self.SpinB1.setValue(int(liste[9]))
            self.SpinB2.setValue(int(liste[10]))
            self.SpinB3.setValue(int(liste[11]))
            
            self.FontCombo4.setCurrentText(str(liste[12]))
            self.SpinT1.setValue(int(liste[13]))
            self.SpinT2.setValue(int(liste[14]))
            self.SpinT3.setValue(int(liste[15]))  
        except:
            print("fontdatenLaden fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            
    ###########################################################
    ###########################################################
    ############## SYSTEM VARIABLEN Funktionen ################
    ###########################################################
    ###########################################################
    
    def systemvariablenspeichern(self):
        
        try:
            if self.is_float(str(self.LE_Maschinenstunde.text())):
                maschinenstunde = float(str(self.LE_Maschinenstunde.text()))
            else:
                maschinenstunde = 0.0
                
            if self.is_float(str(self.LE_Arbeitsstunde.text())):
                arbeitsstunde = float(str(self.LE_Arbeitsstunde.text()))
            else:
                arbeitsstunde = 0.0
                
            if self.is_float(str(self.LE_Aufpreisfaktor.text())):
                aufpreisfaktor = float(str(self.LE_Aufpreisfaktor.text()))
            else:
                aufpreisfaktor = 0.0
            
            
            self.sqlmanager.VariableVeraendern(maschinenstunde,arbeitsstunde,aufpreisfaktor,"","","","","","","","","","","","",4)
        except:
            print("systemvariablenspeichern fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
    
    def getSystemVar(self):
        #print("System Var")
        try:
            self.LE_Maschinenstunde.setText(str(self.sqlmanager.variableAusgeben(4)[1]))
            self.LE_Arbeitsstunde.setText(str(self.sqlmanager.variableAusgeben(4)[2]))
            self.LE_Aufpreisfaktor.setText(str(self.sqlmanager.variableAusgeben(4)[3]))
        except:
            print("systemvariablen laden fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            self.openErrorBox(("systemvariablen laden fehlgeschlagen!!! Status: "+ status),sys.exc_info()[0],sys.exc_info()[1])
            
    def setFAParameter(self):
        try:
            druckerPrusa = "Prusa MK3S+"
            druckerHP = "HP JF 5200"
            afo = self.lineEdit_13.text()
            ap = self.lineEdit_14.text()
            APLprusa = self.lineEdit_15.text()
            APLhp = self.lineEdit_16.text()
            
            self.sqlmanager.VariableVeraendern(druckerPrusa,afo,ap,APLprusa,"","","","","","","","","","","",5)
            self.sqlmanager.VariableVeraendern(druckerHP,afo,ap,APLhp,"","","","","","","","","","","",6)
        except:
            print("setFAParameter fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            self.openErrorBox(("setFAParameter fehlgeschlagen!!! Status: "),sys.exc_info()[0],sys.exc_info()[1])
    ###########################################################
    ###########################################################
    ############## MATERIALPARAMETER Funktionen ###############
    ###########################################################
    ###########################################################
    
    def mlUpdate(self):
        # Materialliste updaten
        #print("materialliste laden")
        try:
            result = self.sqlmanager.getML(1)
            self.tableWidget.setRowCount(0)
            for row_number, row_data in enumerate(result):
                self.tableWidget.insertRow(row_number)
                for colum_number, data in enumerate(row_data):
                    if colum_number == 0:
                        self.tableWidget.setItem(row_number, 0, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 1:
                        self.tableWidget.setItem(row_number, 1, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 2:
                        self.tableWidget.setItem(row_number, 2, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 3:
                        self.tableWidget.setItem(row_number, 3, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 4:
                        self.tableWidget.setItem(row_number, 4, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 5:
                        self.tableWidget.setItem(row_number, 5, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 6:
                        self.tableWidget.setItem(row_number, 6, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 7:
                        self.tableWidget.setItem(row_number, 7, QtWidgets.QTableWidgetItem(str(data)))
                            
                            
        except:
            print("Materialliste laden fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            self.openErrorBox("Materialliste laden fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
            
    def resetLEmat(self):
        self.lineEdit.clear()
        self.lineEdit_2.clear()
        self.lineEdit_3.clear()
        self.lineEdit_4.clear()
        self.lineEdit_5.clear()
        self.lineEdit_6.clear()
        self.lineEdit_7.clear()
        self.lineEdit_8.clear()
        
    def matToLE(self):
        try:
            self.resetLEmat()
            selectedID = self.tableWidget.item(self.tableWidget.currentRow(),0).text()
            
            cursor = self.sqlmanager.matSuchenMitId(selectedID)
            selectedData = cursor.fetchone()
            
            self.matID.setText(str(selectedData[0]))
            self.matArt.setText(str(selectedData[1]))
            self.matD.setText(str(selectedData[2]))
            self.matGew.setText(str(selectedData[3]))
            self.matPreis.setText(str(selectedData[4]))
            self.matVar1.setText(str(selectedData[5]))
            self.matVar2.setText(str(selectedData[6]))
            self.matVar3.setText(str(selectedData[7]))

        except:
            print("matToLE fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            
    def addMat(self):
        self.sqlmanager.matAdd(self.matArt.text(),self.matD.text(),self.matGew.text(),self.matPreis.text(),self.matVar1.text(),self.matVar2.text(),self.matVar3.text())
        self.resetLEmat()
        self.mlUpdate()
        
    def matUpdate(self):
        selection = self.tableWidget.item(self.tableWidget.currentRow(),0).text()
        self.sqlmanager.matUpdate(self.matArt.text(),self.matD.text(),self.matGew.text(),self.matPreis.text(),self.matVar1.text(),self.matVar2.text(),self.matVar3.text(),selection)
        self.resetLEmat()
        self.mlUpdate()
        self.tableWidget.selectRow((int(selection))-1)
        self.matToLE()
        
    def matDel(self):
        selection = self.tableWidget.item(self.tableWidget.currentRow(),0).text()
        self.sqlmanager.matDel(selection)
        self.resetLEmat()
        self.mlUpdate()
    
    ###########################################################
    ###########################################################
    ############## HILFS Funktionen ###########################
    ###########################################################
    ###########################################################  

    def ordnerlabelsSetzen(self):
        print("Ordnerlabels setzen")  
        print(self.sqlmanager.variableAusgeben(3)[1])
        self.lineEdit_12.setText(self.sqlmanager.variableAusgeben(3)[1])
        self.lineEdit_17.setText(self.sqlmanager.variableAusgeben(3)[2])
        self.lineEdit_18.setText(self.sqlmanager.variableAusgeben(3)[3])

    def pick_new(self):
        dialog = QFileDialog()
        folder_path = dialog.getExistingDirectory(None, "Select Folder")
        
        return folder_path

    def ordnerAendern(self,variante):
        var1 = self.lineEdit_12.text()
        var2 = self.lineEdit_17.text()
        var3 = self.lineEdit_18.text()
        
        pfad = self.pick_new()
        
        if variante == 1:
            var1 = pfad
        elif variante == 2:
            var2 = pfad
        elif variante == 3:
            var3 = pfad
            
        self.sqlmanager.variableUpdaten(var1,var2,var3,"","","","","","","","","","","","",3)
        self.ordnerlabelsSetzen()
        print(pfad)  

    def loadData(self):
        # Datenordner nach Aufträgen durchleuchten und Aufträge Anlegen
        try:
            
            for i in range(1200,1350):
                for (root,dirs,files) in os.walk(self.sqlmanager.variableAusgeben(3)[1] +'\\'):
                    for x in dirs:
                        if x.isdigit():
                            y = int(x)
                            if y == i:
                                nameFull = (str(root)).split("\\")
                                nameDouble = nameFull[2].split(" ")
                                timestampFunc = dt.datetime.fromtimestamp(os.path.getmtime(str(root)+"\\"+ str(y)))
                                eingangsdatum= timestampFunc.strftime("%d-%m-%Y")
                                today = date.today()
                                wunschdatum = today.strftime("%d-%m-%Y")
                                
                                self.sqlmanager.auftragErstellenAusBackup(i,eingangsdatum,wunschdatum,nameDouble[0],nameDouble[1])                     
                print("Fertig" + str(i))
        except:
            print("Auftrag eintragen fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")

        print("!!!! ENDE !!!!")
            
    def if_integer(string):
        if string[0] == ('-', '+'):
            return string[1:].isdigit()
        else:
            return string.isdigit()   
            
    def datumWandel(self,datum):
        res1 = True
        res2 = True
        
        try:
            res1 = bool(datetime.datetime.strptime(datum,"%Y-%m-%d"))
        except ValueError:
            res1 = False
        try:
            res2 = bool(datetime.datetime.strptime(datum,"%d-%m-%Y"))
        except ValueError:
            res2 = False
        
        if (res1):
            DTwandel = datetime.datetime.strptime(datum,"%Y-%m-%d")
            to = DTwandel.strftime("%d-%m-%Y")
        else:
            DTwandel = datetime.datetime.strptime(datum,"%d-%m-%Y")
            to = DTwandel.strftime("%d-%m-%Y")
        
        return to
        
    def emailtest2(self):
        print("email2 start")
        try:
            outlook = win32.Dispatch('Outlook.Application')
            outlookNS = outlook.GetNameSpace('MAPI')
            
            mail = outlook.CreateItem(0)
            mail.To = '3d-druck.ks@volkswagen.de'
            mail.Subject = 'Message subject'
            mail.Body = 'Message body'
            mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
    
            # To attach a file to the email (optional):
            #attachment  = "Path to the attachment"
            #mail.Attachments.Add(attachment)
    
            mail.Send()
        except:
            print("emailtest2 fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            
    def is_float(self,value):
        if value is None:
            return False
        try:
            float(value)
            return True
        except:
            return False
            
    def emailtest(self):
        print("email testen")
        try:
            outlook = win32.Dispatch('Outlook.Application')
            outlookNS = outlook.GetNameSpace('MAPI')
            
            mail = outlook.CreateItem(0)
            mail.To = '3d-druck.ks@volkswagen.de'
            mail.Subject = 'Materialbestellung'
            mail.Body = 'Test'
            mail.HTMLBody = """<html>
            <head>
            <title>Title of the document</title>
            </head>

            <body>
            <h3>Der 3D-Druck braucht Material.</h3>
            <hr>
            <p>This is a paragraph.</p>
            <hr>
            <Table>
                <tr>
                    <th> Materialnummer </th>
                    <th> Durchmesser </th>
                    <th> Gewicht </th>
                    <th> Art </th>
                    <th> Farbe </th>
                    <th> Aktueller Bestand </th>
                    <th>                </th>
                    <th> Benötigte Menge </th>
                </tr>
                
                <tr>
                <td> M407134 </td>
                <td> 1,75mm </td>
                <td> 1Kg </td>
                <td> PLA </td>
                <td> Schneeweiß </td>
                <td> 1 </td>
                <td>               </td>
                <td> 1 </td>
                </tr>
                
            </table>
            
            </body>

            </html>"""
 
            mail.Send()
            
        except:
            print("emailtest fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            
    def finden(self):
        #print(self.sqlmanager.suchennachName("Pries","Marcel"))
        for i in range(len(self.sqlmanager.getAuftragsliste(2))):         
            if not self.sqlmanager.suchennachName(self.sqlmanager.getAuftragsliste(2)[i][3],self.sqlmanager.getAuftragsliste(2)[i][4]):
                self.sqlmanager.PersonHizufugen(self.sqlmanager.getAuftragsliste(2)[i][4],self.sqlmanager.getAuftragsliste(2)[i][3],self.sqlmanager.getAuftragsliste(2)[i][6],self.sqlmanager.getAuftragsliste(2)[i][5],self.sqlmanager.getAuftragsliste(2)[i][8],self.sqlmanager.getAuftragsliste(2)[i][9])
    
    def generateLiz(self):
        print("generateLiz1")
        expiration_date = datetime(2023, 12, 25)
        private_key = "private_key_placeholder"
        public_key = "public_key_placeholder"
        print("generateLiz2")
        license_info = {'expiration_date': expiration_date, 'license_key': None}
        print("generateLiz22")
        self.license_manager.generate_license_key(private_key, license_info)
        print("generateLiz3")
        if self.license_manager.verify_license_key(public_key, license_info):
            print("License is valid.")
        else:
            print("License is invalid.")
        print("generateLiz4")


        self.LiLizenzGeneriert.setText(str(public_key))
        
    def lizenzValid(self):
        print("prüfen")
        try:
            lizenz = self.LiLizenzAktuell.text()
            if self.license_manager.is_license_key_valid(lizenz):
                print("Der Lizenzschlüssel ist gültig.")
            else:
                print("Der Lizenzschlüssel ist abgelaufen.")             
        except:
            print("lizenzValid fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")      
    ###########################################################
    ###########################################################
    ############## RGB SELECTOR KLASSE ########################
    ###########################################################
    ###########################################################            
            
class RGBSelector(QWidget):
    def __init__(self):
        super().__init__()

        # Benutzeroberfläche erstellen
        self.red_slider = QSlider(Qt.Horizontal)
        self.green_slider = QSlider(Qt.Horizontal)
        self.blue_slider = QSlider(Qt.Horizontal)
        self.red_slider.setFixedSize(100,10)
        self.green_slider.setFixedSize(100,10)
        self.blue_slider.setFixedSize(100,10)
        self.red_slider.setMaximum(255)
        self.green_slider.setMaximum(255)
        self.blue_slider.setMaximum(255)
        
        self.label_red = QLabel()
        self.label_green = QLabel()
        self.label_blue = QLabel()
        self.preview_label = QLabel()
        self.preview_label.setFixedSize(50, 50)
        
        self.label_red.setText(str(0))
        self.label_green.setText(str(0))
        self.label_blue.setText(str(0))

        # Layout erstellen
        layout1 = QHBoxLayout()
        layout2 = QVBoxLayout()
        layout3 = QVBoxLayout()

        layout4 = QHBoxLayout()
        layout5 = QHBoxLayout()
        layout6 = QHBoxLayout()
        
        layout4.addWidget(self.label_red)
        layout4.addWidget(self.red_slider)
        layout5.addWidget(self.label_green)
        layout5.addWidget(self.green_slider)
        layout6.addWidget(self.label_blue)
        layout6.addWidget(self.blue_slider)
        
        layout3.addWidget(self.preview_label)
        
        layout2.addLayout(layout4)
        layout2.addLayout(layout5)
        layout2.addLayout(layout6)
        
        layout1.addLayout(layout2)
        layout1.addLayout(layout3)
        

        # Signal-Slot-Verbindungen
        self.red_slider.valueChanged.connect(self.update_preview)
        self.green_slider.valueChanged.connect(self.update_preview)
        self.blue_slider.valueChanged.connect(self.update_preview)

        # Layout auf das Widget setzen
        self.setLayout(layout1)
        
        self.update_preview()

    def update_preview(self):
        # Aktuelle Farbwerte aus den Schiebereglern auslesen
        self.red = self.red_slider.value()
        self.label_red.setText(str(self.red))
        self.green = self.green_slider.value()
        self.label_green.setText(str(self.green))
        self.blue = self.blue_slider.value()
        self.label_blue.setText(str(self.blue))

        # Farbe setzen und Vorschau aktualisieren
        self.color = QColor(self.red, self.green, self.blue)
        self.preview_label.setStyleSheet("background-color: {};".format(self.color.name()))
        
    def getRGBvalue(self):
        return self.color.name()
        
    def setRGBValue(self,hex_value="#008800"):
        try:        
            color1 = QColor(hex_value)
            
            red, green, blue, alpha = color1.getRgb()
                 
            self.red_slider.setValue(red)
            self.green_slider.setValue(green)
            self.blue_slider.setValue(blue)
            
            self.update_preview()
            
        except Exception as e:
            print("setRGBValue fehlgeschlagen!!!")
            print("Oops!", type(e).__name__, "occurred:", e)
            
class LicenseManager:
    def generate_license_key(self,private_key, license_info):
        print("generate_license_key1")
        # Ablaufdatum festlegen
        expiration_date_str = license_info.expiration_date.strftime("%Y-%m-%d")
        print("generate_license_key2")
        # Lizenzdaten erstellen
        license_data = {'expiration_date': expiration_date_str}
        license_data_str = json.dumps(license_data).encode('utf-8')
        print("generate_license_key3")
        # Lizenzschlüssel generieren (einfach SHA-256 Hash)
        license_key = hashlib.sha256(license_data_str).digest()
        license_info.license_key = license_key
    
    def verify_license_key(self,public_key, license_info):
        # Überprüfen des Ablaufdatums
        if license_info.expiration_date < datetime.now():
            print("License has expired.")
            return False
    
        # Überprüfen des Lizenzschlüssels (einfach SHA-256 Hash)
        expected_license_key = hashlib.sha256(json.dumps({'expiration_date': license_info.expiration_date.strftime("%Y-%m-%d")}).encode('utf-8')).digest()
    
        if license_info.license_key == expected_license_key:
            return True
        else:
            print("License verification failed.")
            return False