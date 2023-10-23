from PyQt5 import QtWidgets,uic
import sys,os

from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import NewTask


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
    
    
    

class Ui(QtWidgets.QDialog, NewTask.Ui_NewTask):
    def __init__(self,sqlmanager,parent=None):
        super(Ui,self).__init__(parent)
        
        self.setupUi(self)
        self.programmpfad = os.path.abspath(".")
        
        icon = QIcon(resource_path(self.programmpfad +"/icons/auma_farbe.png"))
        # Icon als Anwendungssymbol setzen
        self.setWindowIcon(icon)
        # Titelleisten-Icon temporär ändern
        self.setProperty("windowIcon", icon)
        
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint)  # Hinzufügen von Minimieren- und Maximieren-Buttons
        
        self.sqlPointer = sqlmanager
        
        self.backbutton = self.backToMainMenu
        self.backbutton.clicked.connect(self.backToMainmenu)        
        self.saveTaskButton = self.save_task
        self.saveTaskButton.clicked.connect(lambda:self.saveTask(0))
        self.saveTaskNEWButton = self.saveandnew
        self.saveTaskNEWButton.clicked.connect(self.saveTaskAndNew)
        self.ChangePersDataButton = self.callChange
        self.ChangePersDataButton.clicked.connect(self.datenAktualisieren)
        
        self.LineName = self.l_Nachname
        self.LineName.textChanged.connect(self.filterName)
        self.LineVorname = self.l_Vorname
        self.LineAbteilung = self.l_Abteilung
        self.LineKostenstelle = self.l_Kostenstelle
        self.LineTelefon = self.l_Telefon
        self.LineEMail = self.l_Email
        self.LineBauteilname = self.l_Bauteil
        #self.LineEingangsdatum = self.lineEingangsdatum
        #self.LineWunschtermin = self.lineWunschtermin 
        self.TextBeschreibung = self.t_bauteil
        self.ComboVorlage = self.c_vorlage
        self.ComboMaterial = self.c_material
        self.ComboFB = self.c_fertbereich
        self.ComboFarbe = self.c_farbe
        #self.ComboBauteilgroße = self.c_bauteilgrose  
        self.SpinStuckzahl = self.s_stuck
        self.LabelAuftragsNR = self.ANummer  
        self.RadioJa = self.radioJa
        self.RadioNein = self.radioNein
        self.RadioNein.setChecked(True)     
        self.calendarBegin = self.calendarBegin
        self.calendarWish = self.calendarWish   
        self.PersTable = self.tableWidget
        self.PersTable.verticalHeader().setDefaultSectionSize(20)
        self.PersTable.setColumnWidth(0, 40)  # ID
        self.PersTable.setColumnWidth(1, 150)  # Nachname
        self.PersTable.setColumnWidth(2, 80)  # Vorname
        self.PersTable.horizontalHeader().setStretchLastSection(True)
        self.PersTable.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.PersTable.itemSelectionChanged.connect(self.loadDataFromTable)
        self.PersTable.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.PersTable.verticalHeader().setVisible(False)
        
    def windowshow(self):
        #self.getStyleFromHilfs()
        self.matLoad()
        self.reset_selection()
        self.setTaskNumber()
        self.refreshPersonList()
        self.RadioNein.setChecked(True)
        self.show()
        
    def setMainmenuPointer(self,mainpointer):
        global MainmenuPointer
        MainmenuPointer = mainpointer
        
    def backToMainmenu(self):
        self.hide()
        MainmenuPointer.windowshow()
    
    def beendenButtonPressed(self):
        sys.exit(0)        
        
    ###########################################################
    ###########################################################
    ############### AUFTRAG ERSTELLEN #########################
    ###########################################################
    ###########################################################   
          
    def saveTask(self,index):
        eilig = 0 
        
        if self.RadioJa.isChecked():
            eilig = 1
            print("EILIG AKTIVIERT")
        elif self.RadioNein.isChecked():
            eilig = 0
        
        if not (self.SpinStuckzahl.value() == 0):
            try:
                self.sqlPointer.auftragErstellen(self.calendarBegin.selectedDate().toString("dd-MM-yyyy"),
                                                 self.calendarWish.selectedDate().toString("dd-MM-yyyy"),self.LineName.text(),self.LineVorname.text(),
                                                 self.LineKostenstelle.text(),self.ComboFB.currentText(),self.LineAbteilung.text(),
                                                 self.LineTelefon.text(),self.LineEMail.text(),self.LineBauteilname.text(),
                                                 eilig,self.TextBeschreibung.toPlainText(),
                                                 self.ComboFarbe.currentText(),self.ComboVorlage.currentText(),self.ComboMaterial.currentText(),"unbekannt",0,0,self.SpinStuckzahl.value())
                                
            
                if self.sqlPointer.suchennachName(self.LineName.text(),self.LineVorname.text()):
                    self.sqlPointer.PersonHizufugen(self.LineVorname.text(),self.LineName.text(),self.LineKostenstelle.text(),self.LineAbteilung.text(),self.LineTelefon.text(),self.LineEMail.text())
                
            except:
                print("Neuen Auftrag eintragen fehlgeschlagen!!!")
                print("Oops!", sys.exc_info()[0], "occurred")
                print("Oops!", sys.exc_info()[1], "occurred")
                print("Oops!", sys.exc_info()[2], "occurred")
                
            if index == 0:
                self.openMessageBox(self.LabelAuftragsNR.text(),self.LineName.text(),self.LineVorname.text(),self.LineBauteilname.text())
                self.backToMainmenu()
                
            self.setTaskOrdner()
                            
        else:
            self.openMessageBox2()
        self.ClearAll() 
        
    def saveTaskAndNew(self):
        self.saveTask(1)
        self.openMessageBox(self.LabelAuftragsNR.text(),self.LineName.text(),self.LineVorname.text(),self.LineBauteilname.text())
        

    ###########################################################
    ###########################################################
    ############### INTERFACE BEREINIGEN ######################
    ###########################################################
    ###########################################################
    
    def ClearAll(self):
        try:
            print("säubern 1")
            self.l_Nachname.clear()
            self.LineVorname.clear()
            self.LineAbteilung.clear()
            self.LineKostenstelle.clear()
            self.LineTelefon.clear()
            self.LineEMail.clear()
            self.LineBauteilname.clear()
            self.TextBeschreibung.clear()       
            self.ComboVorlage.setCurrentIndex(0)
            self.ComboMaterial.setCurrentIndex(0)
            self.ComboFB.setCurrentIndex(0)
            self.ComboFarbe.setCurrentIndex(0)    
            self.SpinStuckzahl.setValue(0)   
            self.LabelAuftragsNR.clear()
            print("säubern 2")
        except:
            print("ClearAll fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            print("Oops!", sys.exc_info()[2], "occurred") 
            
    ###########################################################
    ###########################################################
    ############### PERSONENLISTE FUNKTIONEN ##################
    ###########################################################
    ###########################################################
    
    def refreshPersonList(self):
        try:
            result = self.sqlPointer.personenlisteRausgeben(2)
            self.PersTable.setRowCount(0)
            for row_number, row_data in enumerate(result):
                self.PersTable.insertRow(row_number)
                for colum_number, data in enumerate(row_data):
                    if colum_number == 0:
                        self.PersTable.setItem(row_number, 0, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 2:
                        self.PersTable.setItem(row_number, 1, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 1:
                        self.PersTable.setItem(row_number, 2, QtWidgets.QTableWidgetItem(str(data)))
         
        except:
            print("TaskTable Füllen fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
        
    def filterName(self):       
        try:
            rows = self.PersTable.rowCount()
            tempindex = (str(self.LineName.text())).lower()
            
            for i in range(rows):
                self.PersTable.hideRow(i)
                tasktableitem = (str(self.PersTable.item(i, 1).text())).lower()
                if (tasktableitem.startswith(tempindex)):
                    self.PersTable.showRow(i)
            
        except:
            print("AuftragslisteFiltern fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred") 
            
    def loadDataFromTable(self):
        #self.ClearAll()
        try:
            selectedID = self.PersTable.item(self.PersTable.currentRow(), 0).text()
            
            cursor = self.sqlPointer.perssuchenmitid(selectedID)
            selectedData = cursor.fetchone()
     
            self.LineName.setText(selectedData[2])
            self.LineVorname.setText(selectedData[1])
            self.LineAbteilung.setText(selectedData[4])
            self.LineKostenstelle.setText(selectedData[3])
            self.LineTelefon.setText(selectedData[5])
            self.LineEMail.setText(selectedData[6])
            
            self.setTaskNumber()
        except:
            print("datenAusTableLaden fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred") 
            
    def datenAktualisieren(self):     
        selected_items = self.PersTable.selectedItems()

        if len(selected_items) > 0:
            # Die erste ausgewählte Zelle enthält die Zeilennummer
            row = selected_items[0].row()

            # ID (erste Spalte) der ausgewählten Zeile abrufen
            item = self.PersTable.item(row, 0)  # 0 ist die Spaltennummer der ID
            id = item.text()
            #print('ID der ausgewählten Zeile:', id)
            
            self.sqlPointer.personUpdaten(id,self.LineVorname.text(),self.LineName.text(),self.LineKostenstelle.text(),self.LineAbteilung.text(),self.LineTelefon.text(),self.LineEMail.text())
            
        else:
            print('Keine Zeile ausgewählt')
        
    def reset_selection(self):
        # Auswahl in der QTableWidget zurücksetzen
        self.PersTable.clearSelection()
    
    ###########################################################
    ###########################################################
    ############### HILFS FUNKTIONEN ##########################
    ###########################################################
    ###########################################################
 
    def setTaskNumber(self):
        self.ident = self.sqlPointer.maxidfinden() +1
        self.LabelAuftragsNR.setText(str(self.ident))
        
    def setTaskOrdner(self):
        #print("Ordner erstellen mit Nummer: " + str(self.ident))
        self.sqlPointer.createFolder(self.LineName.text(),self.LineVorname.text(),int(self.ident))

    def openMessageBox(self,auftragsnummer,name,vorname,bauteilname):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Auftrag eingetragen!!!")
        msg.setWindowTitle("Auftrag eintragen erfolgreich!!!")
        msg.setDetailedText("Auftragsnummer: " + auftragsnummer + "\nName: " + name + "\nVorname: " + vorname + "\nBauteilname: " + bauteilname)
        msg.exec_()
        
    def openMessageBox2(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Stückzahl ist Null!!! Der Auftrag wird so nicht eingetragen!!!")
        msg.setWindowTitle("Stückzahl!!!")
        msg.exec_()
           
    def hilfspointer(self,pointer):
        self.hilfs = pointer
            
    def getStyleFromHilfs(self):
        self.setStyleSheet(self.hilfs.getStyle())
        
    def eiligFunction(self):
        print("Eilig")
        
    def matClear(self):
        self.c_material.clear()
    
    def matLoad(self):
        try:
            self.matClear()
            print("material laden")
            print(self.sqlPointer.getML(3))
            self.c_material.addItems(self.sqlPointer.getML(3))

        except:
            print("matLoad fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred") 
        
        