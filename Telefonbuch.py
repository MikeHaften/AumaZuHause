from PyQt5 import QtWidgets,uic
from PyQt5 import QtGui
from PyQt5.QtWidgets import QDialog
import sqlite3
import sys, os
import datetime
from datetime import date
import Telefon


class Ui(QtWidgets.QDialog, Telefon.Ui_Telefonbuch):
    def __init__(self,parent=None):
        super(Ui,self).__init__(parent)
        
        self.setupUi(self)
        
        self.backbutton = self.findChild(QtWidgets.QPushButton,'backToMainMenu')
        self.backbutton.clicked.connect(self.backToMainmenu)
        self.deletebutton = self.findChild(QtWidgets.QPushButton,'deletebut')
        self.deletebutton.clicked.connect(self.deleteFunction)
        self.addbutton = self.findChild(QtWidgets.QPushButton,'addbut')
        self.addbutton.clicked.connect(self.addFunction)
        self.savebutton = self.findChild(QtWidgets.QPushButton,'savebut')
        self.savebutton.clicked.connect(self.saveFunction)
        self.searchbutton = self.findChild(QtWidgets.QPushButton,'searchbut')
        self.searchbutton.clicked.connect(self.searchFunction)
        
    def deleteFunction(self):
        print("löschen")
        
    def addFunction(self):
        print("hinzufügen")
        
    def saveFunction(self):
        print("save")
        
    def searchFunction(self):
        print("search")
        
    def windowshow(self):
        self.show()

    def setMainmenuPointer(self,mainpointer):
        global MainmenuPointer
        MainmenuPointer = mainpointer
        
    def backToMainmenu(self):
        self.hide()
        MainmenuPointer.windowshow()
    
    def beendenButtonPressed(self):
        sys.exit(0)