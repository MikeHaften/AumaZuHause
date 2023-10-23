from PyQt5 import QtWidgets,uic
from PyQt5 import QtGui
from PyQt5.QtWidgets import QDialog
import sqlite3
import sys, os
import datetime
from datetime import date
import materialbestandUi



class Ui(QtWidgets.QDialog, materialbestandUi.Ui_materialbestand):
    def __init__(self,sqlpointer,parent=None):
        super(Ui,self).__init__(parent)
        
        self.setupUi(self)
        
        self.sqlpointer = sqlpointer
        
        self.backbutton = self.findChild(QtWidgets.QPushButton,'backToMainMenu')
        self.backbutton.clicked.connect(self.backToMainmenu)

        self.PreviewTable = self.findChild(QtWidgets.QTableWidget,'materialtable')
        self.PreviewTable.verticalHeader().setDefaultSectionSize(20)
        self.PreviewTable.setColumnWidth(0, 40)  # ID
        self.PreviewTable.setColumnWidth(1, 120)  # Materialnummer
        self.PreviewTable.setColumnWidth(2, 80)  # Durchmesser
        self.PreviewTable.setColumnWidth(3, 60)  # Gewicht
        self.PreviewTable.setColumnWidth(4, 90)  # Art
        self.PreviewTable.setColumnWidth(5, 110)  # Farbe
        self.PreviewTable.setColumnWidth(6, 120)  # Hersteller
        
    def MaterialDelete(self):
        print("Material Löschen")
        
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