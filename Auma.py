from PyQt5 import QtWidgets,uic
import sqlite3
import sys, os
from PyQt5.QtGui import QIcon

from PyQt5.QtWidgets import QMessageBox
import aumaUi


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
    
    
    
class Ui(QtWidgets.QMainWindow, aumaUi.Ui_MainWindow):
    def __init__(self,newTaskUi,taskmange,infoChef,stats,telefon,material,parent=None):
        super(Ui,self).__init__(parent)
        
        self.setupUi(self)
        
        icon = QIcon(resource_path(os.path.abspath(".") +"/icons/auma_farbe.png"))
        # Icon als Anwendungssymbol setzen
        self.setWindowIcon(icon)

        # Titelleisten-Icon temporär ändern
        self.setProperty("windowIcon", icon)
        
        self.newTask = newTaskUi
        self.taskmanager = taskmange
        self.informationChef = infoChef
        self.statistikDialog = stats
        self.telefonbuchDialog = telefon
        self.materialmanagement = material
        
        if not os.path.exists(resource_path(os.path.abspath(".") + "/auma.db")):
            global cursor , connection
            #print("Datenbank noch nicht erstellt")
            connection = sqlite3.connect(resource_path(os.path.abspath(".") + "/auma.db"))
            cursor = connection.cursor()
            
        else:
            #print("Datenbank schon vorhanden")
            connection = sqlite3.connect(resource_path(os.path.abspath(".") +"/auma.db"))
            cursor = connection.cursor()
        #Buttons

        self.beenden = self.findChild(QtWidgets.QPushButton,'beendenBut')
        self.beenden.clicked.connect(self.beendenButtonPressed)
        self.neuerAuftrag = self.findChild(QtWidgets.QPushButton,'NewTaskButton')
        self.neuerAuftrag.clicked.connect(self.NewTask)
        self.showAuftrag = self.findChild(QtWidgets.QPushButton,'AuftragAnz')
        self.showAuftrag.clicked.connect(self.showTask)
        self.Statistik = self.findChild(QtWidgets.QPushButton,'StatsButton')
        self.Statistik.clicked.connect(self.showStats)
        self.ChefInfo = self.findChild(QtWidgets.QPushButton,'infoChef')
        self.ChefInfo.clicked.connect(self.infoChefe)
        self.Telefonbuch = self.findChild(QtWidgets.QPushButton,'telefonbuch_but')
        self.Telefonbuch.clicked.connect(self.showPhonebook)
        self.Materialmanagment = self.findChild(QtWidgets.QPushButton,'materialmana')
        self.Materialmanagment.clicked.connect(self.showMaterialmanagement)
        self.Materialbestand = self.findChild(QtWidgets.QPushButton,'MaterialbestandBut')
       
    def beendenButtonPressed(self):
        sys.exit(0)

    def windowshow(self):
        self.ChefInfo.hide()
        self.Telefonbuch.hide()
        self.show()

    def NewTask(self):
        self.hide()
        self.newTask.windowshow()

    def showTask(self):
        self.hide()
        self.taskmanager.windowshow()

    def showStats(self):
        self.hide()
        self.statistikDialog.windowshow()

    def infoChefe(self):
        self.hide()
        self.informationChef.windowshow()

    def showPhonebook(self):
        self.hide()
        self.telefonbuchDialog.windowshow()

    def showMaterialmanagement(self):
        self.hide()
        self.materialmanagement.windowshow()
        
    def showMaterialBestand(self):
        self.hide()
        

    def closeEvent(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("ENDE!!!")
        msg.setWindowTitle("ENDE!!!")
        msg.exec_()
        