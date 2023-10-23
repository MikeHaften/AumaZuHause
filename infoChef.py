from PyQt5 import QtWidgets,uic
import sys,os
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import infoDiaUi


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class Ui(QtWidgets.QDialog, infoDiaUi.Ui_infoDia):
    def __init__(self,parent=None):
        super(Ui,self).__init__(parent)
        
        self.setupUi(self)
        self.programmpfad = os.path.abspath(".")
        
        icon = QIcon(resource_path(self.programmpfad +"/icons/auma_farbe.png"))
        # Icon als Anwendungssymbol setzen
        self.setWindowIcon(icon)
        # Titelleisten-Icon temporär ändern
        self.setProperty("windowIcon", icon)
        
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint)  # Hinzufügen von Minimieren- und Maximieren-Buttons
        
        self.backbutton = self.findChild(QtWidgets.QPushButton,'backToMainMenu')
        self.backbutton.clicked.connect(self.backToMainmenu)
        
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