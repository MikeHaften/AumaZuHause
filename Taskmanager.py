from PyQt5 import QtWidgets
from PyQt5 import QtGui
import sys, os
import datetime
import shutil
import win32com.client
from PyQt5.QtGui import QPixmap, QIcon, QImage, QImageReader
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import Qt
import AuftragsboardUi
import win32com.client as win32
from PyQt5.Qt import QMimeData
from PyQt5.QtGui import QGuiApplication

from skimage.transform import resize
import matplotlib.pyplot as plt

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
    
    

class Ui(QtWidgets.QDialog, AuftragsboardUi.Ui_AuftragboardDialog):
    def __init__(self,sqlpointer,app,parent=None):
        super(Ui,self).__init__(parent)
        
        self.setupUi(self)
        self.programmpfad = os.path.abspath(".")
        print(self.programmpfad)
        
        icon = QIcon(resource_path(self.programmpfad +"/icons/auma_farbe.png"))
        # Icon als Anwendungssymbol setzen
        self.setWindowIcon(icon)
        # Titelleisten-Icon temporär ändern
        self.setProperty("windowIcon", icon)
              
        self.sqlpointer = sqlpointer
        self.app = app
        
        self.selectedID = 0
        self.lastSelectedID = 0
        self.butInvis = 1
        self.filterVariable = 8

        self.materialpreis = 30

        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint)  # Hinzufügen von Minimieren- und Maximieren-Buttons
        
        self.backbutton = self.backToMainMenu
        self.backbutton.clicked.connect(self.backToMainmenu)
        self.begleitbutton = self.begleit
        self.filterbutton = self.filterbut
        self.deletebutton1 = self.deleteButton
        self.deletebutton1.clicked.connect(self.auftragLoeschen)
        
        self.savebutton = self.savebut
        self.savebutton.clicked.connect(self.savefunction)
        self.openFolderbutton = self.openFolderBut
        self.openFolderbutton.clicked.connect(self.openFold)
        self.showFilterbutton = self.showFilter
        self.showFilterbutton.clicked.connect(self.filterShow)
        self.resetFilterbutton = self.resetFilter
        self.resetFilterbutton.clicked.connect(self.resetFilterFunc)
        self.Patentbutton = self.PatentButton
        self.Patentbutton.clicked.connect(lambda:self.patentOrderOefnen(1))
        self.copyTaskButton = self.pushButton_2
        self.copyTaskButton.clicked.connect(self.copyTask)
        self.FABelegButton = self.pushButton
        self.EMailButton = self.pushButton_3
        self.EMailButton.clicked.connect(self.email)
        self.LinkToClipButton = self.pushButton_4
        self.LinkToClipButton.clicked.connect(self.copyLinkClipboard)
        
        self.LiName = self.l_name
        self.LiVorname = self.l_vorname
        self.LiAbteilung = self.l_abteilung
        self.LiKostenstelle = self.l_kostenstelle
        self.LiTelefon = self.l_telefon
        self.LiEMail = self.l_email
        self.LiBauteil = self.l_bauteil
        self.LiStuckzahl = self.l_stuck
        self.LiDruckdauer = self.l_druckdauer
        self.LiGewicht = self.l_gewicht
        self.LiFiltername = self.l_filtername
        self.LiFiltername.textChanged.connect(self.nachNameFiltern)
        self.LiFilterbauteil = self.l_filterbauteil
        self.LWunschtermin = self.la_wish
        
        self.LiDatenaufbereitungZeit = self.l_data1
        self.LiNacharbeitZeit = self.l_nach1
        self.LiMaterial = self.l_mat
        self.LiMaterial.setEnabled(False)
        self.LiDatenaufbereitungKosten = self.l_data2
        self.LiDatenaufbereitungKosten.setEnabled(False)
        self.LiNacharbeitKosten = self.l_nach2
        self.LiNacharbeitKosten.setEnabled(False)
        self.LiDruckkosten = self.l_druckk
        self.LiDruckkosten.setEnabled(False)
        self.LiAngebotspreis = self.l_angpreis
        self.LiAngebotspreis.setEnabled(False)
        self.LiFolgeteilpreis = self.l_stpreis
        self.LiFolgeteilpreis.setEnabled(False)
        self.LiEinsparung = self.l_einsp
        self.LiEinsparung.setEnabled(False)
        self.LiEinsparungFolgeteil = self.l_einspfolge
        self.LiEinsparungFolgeteil.setEnabled(False)
        self.LiFertigdatum = self.lineFertigdatum
        self.LiFertigdatum.setEnabled(False)
        self.LIarbeitsstunde = self.lineEdit
        self.LIarbeitsstunde.setEnabled(False)
        self.LImaschinenstunde = self.lineEdit_2
        self.LImaschinenstunde.setEnabled(False)
        self.LIAufpreisfaktor = self.lineEdit_3
        self.LIAufpreisfaktor.setEnabled(False)
        
        self.LiEinData = self.lineEdit_11
        self.LiEinData.setEnabled(False)
        self.LiEinNach = self.lineEdit_12
        self.LiEinNach.setEnabled(False)
        self.LiEinDruck = self.lineEdit_13
        self.LiEinDruck.setEnabled(False)
        self.LiEinMat = self.lineEdit_14
        self.LiEinMat.setEnabled(False)
        self.LiEinHerst = self.lineEdit_15
        self.LiEinHerst.setEnabled(False)
        self.LiEinEinsp = self.lineEdit_16
        self.LiEinEinsp.setEnabled(False)
        
        self.LAuftragsnummer = self.Anr_label
        self.LAuftragsdauer = self.l_auftragd
        self.LEingangsdatum = self.Ein_label
        self.nameLabel = self.label_47
        self.nameBauteil = self.label_48
        self.nameFertigungsbereich = self.label_49
        self.nameEinsparung = self.label_51
        self.auftragszahl = self.label_20
        
        self.Lausgeliefert = self.lineEdit_4
        self.Lfertig = self.lineEdit_10
        self.Cinfill = self.comboBox
        self.tab = self.tabWidget
        self.L_Fa1 = self.lineEdit_5
        self.L_Fa2 = self.lineEdit_6
        self.L_Fa3 = self.lineEdit_7
        self.L_Fa4 = self.lineEdit_8
        self.L_Fa5 = self.lineEdit_9
        
        self.TextBeschreibung = self.l_baubesch
        
        self.ComboFarbe = self.c_farbe
        self.progressBar.setValue(int(0))
        
        self.ComboVorlage = self.c_vorlage
        self.ComboBauteilgrose = self.c_bauteilgrose
        self.ComboMaterial = self.c_material
        self.ComboFilterFB = self.c_filterFB
        self.LFertigungsb = self.c_Fertigungsbereich
        self.CStatus = self.c_Status
        self.FilterStatus = self.c_filterStatus
        self.FilterStatus.currentIndexChanged.connect(self.auftragslisteFiltern)
        self.FilterEinsparung = self.comboBox_8
        
        self.TaskTable.verticalHeader().setDefaultSectionSize(20)
        self.TaskTable.setColumnWidth(0, 45)  # ID
        self.TaskTable.setColumnWidth(1, 90)  # Datum
        self.TaskTable.setColumnWidth(2, 140)  # Name
        self.TaskTable.setColumnWidth(3, 160)  # Bauteil
        self.TaskTable.setColumnWidth(4, 40)  # Stückzahl
        self.TaskTable.setColumnWidth(5, 35)  # Status
        self.TaskTable.setColumnWidth(6, 100)  # Progress
        self.TaskTable.horizontalHeader().setStretchLastSection(True)
        self.TaskTable.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.TaskTable.itemClicked.connect(self.selectTask)
        #self.TaskTable.itemSelectionChanged.connect(self.selectTask)
        self.TaskTable.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.TaskTable.verticalHeader().setVisible(False)
        
        #Folgeteile Label und Lineedit ausblenden
        self.label_37.hide()
        self.label_38.hide()
        self.l_stpreis.hide()
        self.l_einspfolge.hide()
        
        self.label_56.hide()
        self.label_57.hide()
        
        self.vorschauBild = self.g_vorschaubild
        
        self.matLoad()
        
        #self.DeleteTaskbutton.hide()
        self.begleitbutton.hide()
              
        
    def windowshow(self):
        # StyleSheet von den Optionen laden // Den Filter für die Aufträge mit "9" initialisieren // Das TableWidget neu laden
        # Die Filterelemente die nicht gebraucht werden ausblenden // Filter ausführen // Fenster im Vollbildmodus öffenen
        try:
            status = "Datenbank nicht in Ordnung P1"
            self.arbeitsstunde = self.sqlpointer.variableAusgeben(4)[2]
            status = "Datenbank nicht in Ordnung P2"
            self.maschinenstunde = self.sqlpointer.variableAusgeben(4)[1]
            status = "Datenbank nicht in Ordnung P3"
            self.externerMultiplikator = self.sqlpointer.variableAusgeben(4)[3]
            status = "Arbeitsstunde Label konnte nicht gesetzt werden"
            self.LIarbeitsstunde.setText(str(self.arbeitsstunde))
            status = "Maschinenstunde Label konnte nicht gesetzt werden"
            self.LImaschinenstunde.setText(str(self.maschinenstunde))
            status = "Aufpreisfaktorstunde Label konnte nicht gesetzt werden"
            self.LIAufpreisfaktor.setText(str(self.externerMultiplikator))
            status = "Stylesheet setzen fehlgeschlagen"
            
            #self.getStyleFromHilfs()
            status = "Filterstatus auf nicht fertig setzen"
            self.FilterStatus.setCurrentIndex(9)
            self.tab.setCurrentIndex(0)
            status = "Auftragsboard aktualisieren fehlgeschlagen"
            self.tasktableRefresh()
            status = "Filter anzeigen"
            self.filterShow()
            status = "Auftragsliste Filtern"
            self.auftragslisteFiltern()
            status = "Fenster auf Bildschirm maximieren" 
            self.show()
            status = "Funktion geschafft"
            
        except:
            print("Auftragstable anzeigen fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            self.openErrorBox(("Auftragsbereich anzeigen fehlgeschlagen!!! Status: "+ status),sys.exc_info()[0],sys.exc_info()[1])
            
    def setMainmenuPointer(self,mainpointer):
        # Pointer zum Hauptmenu erstellen
        global MainmenuPointer
        MainmenuPointer = mainpointer
        
    def backToMainmenu(self):
        # Auftragsdialog verstecken und Hauptmenü anzeigen
        self.hide()
        MainmenuPointer.windowshow()
    
    def beendenButtonPressed(self):
        # Programm beenden
        sys.exit(0)
        
    ###########################################################
    ###########################################################
    ############## AUFTRAG SPEICHERN ##########################
    ###########################################################
    ###########################################################  
                   
    def savefunction(self):
        # Aktuelle Zeile erfassen // Auftrag in die SQL Tabelle speichern // Tablewidget mit den Aufträgen löschen / Auftrage neu in das TabelWidget laden 
        # Vorher ausgewählte Zeile in TableWidget selektieren // Auftragsliste Filtern
        row = self.TaskTable.currentRow()
        try:
            if (self.CStatus.currentIndex()==7):
                self.sqlpointer.fertigDatumSetzen(int(self.selectedID),datetime.date.today().strftime('%d-%m-%Y'))
                self.Lfertig.setText(self.LiStuckzahl.text())
                self.Lausgeliefert.setText(self.LiStuckzahl.text())
            
            self.sqlpointer.auftragSpeichern(int(self.selectedID),self.LEingangsdatum.text(),self.LWunschtermin.text(),self.LiName.text(),self.LiVorname.text(),
                                             self.LiAbteilung.text(),self.LiKostenstelle.text(),self.LFertigungsb.currentText(), self.LiTelefon.text(),self.LiEMail.text(),
                                             self.LiBauteil.text(),int(self.LiStuckzahl.text()),1,float(self.LiDruckdauer.text()),float(self.LiGewicht.text()),self.TextBeschreibung.toPlainText(),
                                             self.ComboFarbe.currentText(),self.ComboVorlage.currentText(),self.ComboBauteilgrose.currentText(),self.ComboMaterial.currentText(),float(self.LiDatenaufbereitungZeit.text()),
                                             float(self.LiNacharbeitZeit.text()),0,0,0,0,
                                             self.LiFertigdatum.text(),self.CStatus.currentIndex(),self.LImaschinenstunde.text(),self.LIarbeitsstunde.text(),self.LIAufpreisfaktor.text(),self.Lfertig.text(),self.Lausgeliefert.text(),self.Cinfill.currentText(),self.L_Fa1.text(),self.L_Fa2.text(),self.L_Fa3.text(),self.L_Fa4.text(),self.L_Fa5.text())

            
            if (self.CStatus.currentIndex()==7):
                self.sqlpointer.fertigDatumSetzen(int(self.selectedID),datetime.date.today().strftime('%d-%m-%Y'))
            
            prozent = ((int(self.Lfertig.text())/int(self.LiStuckzahl.text()))*100)
            self.progressBar.setValue(int(prozent))
            
            self.resetTasktable()
            self.tasktableRefresh()
            self.TaskTable.selectRow(row)
            self.auftragslisteFiltern()
            self.openMessageBox2("Auftrag gespeichert!!!")
        except:
            print("Auftragstable updaten fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            self.openErrorBox("Auftrag speichern fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
            
    ###########################################################
    ###########################################################
    ############## FOLGEAUFTRAG ERSTELLEN #####################
    ###########################################################
    ###########################################################         
    
    
    def copyTask(self):
        print("Auftrag kopieren")
        row = self.TaskTable.currentRow()
        try:
            print("in der func")
            heute = datetime.date.today().strftime('%d-%m-%Y')
            wunsch = (datetime.date.today() + datetime.timedelta(days=21)).strftime('%d-%m-%Y')
            print(wunsch)
            self.sqlpointer.auftragErstellen(heute,wunsch,self.LiName.text(),self.LiVorname.text(),self.LiKostenstelle.text(),
            self.LFertigungsb.currentText(),self.LiAbteilung.text(),self.LiTelefon.text(),self.LiEMail.text(),self.LiBauteil.text(),
            1,self.TextBeschreibung.toPlainText(),self.ComboFarbe.currentText(),self.ComboVorlage.currentText(),self.ComboMaterial.currentText(),
            self.ComboBauteilgrose.currentText(),self.LiDruckdauer.text(),self.LiGewicht.text(),self.LiStuckzahl.text(),self.LiDatenaufbereitungZeit.text(),
            self.LiNacharbeitZeit.text(),1)
            
            self.resetTasktable()
            self.tasktableRefresh()
            self.TaskTable.selectRow(row)
            self.auftragslisteFiltern()
            self.openMessageBox2("Auftrag kopiert!!!")
            
            maxId = self.sqlpointer.maxidfinden()
            ordnerPath = "N:/11 - Druckauftraege/"+ self.LiName.text() + " " + self.LiVorname.text() + "/" + str(maxId)
            
            OrdnerAlt = self.getFolder(self.LiName.text(),self.LiVorname.text(),int(self.selectedID))
            oldFolder = OrdnerAlt.replace("\\","/")
            destination = shutil.copytree(oldFolder, ordnerPath)  

            pfad = self.programmpfad + "/Auftragsbilder/Auftrag_"+ str(int(self.selectedID)) + ".JPG"   
            pfadneu = self.programmpfad + "/Auftragsbilder/Auftrag_"+ str(maxId) + ".JPG"
            
            if os.path.isfile(pfad):
                new_path = shutil.copy(pfad, pfadneu)
            

        except:
            print("copyTask fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            self.openErrorBox("copyTask fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])

    ###########################################################
    ###########################################################
    ############## AUFTRAG ANWÄHLEN UND LADEN #################
    ###########################################################
    ###########################################################  
            
    def selectTask(self):
        # Schiebt die aktuelle Variable,welche Zeile selektiert ist in eine neue Variable // Erfasst nochmal die aktuelle Selektierung
        # // Läd den aktuellen Auftrag // Überprüft ob in dem Auftragsordner eine PDF mit dem namen Patent.pdf existiert
        try:
            self.lastSelectedID = self.selectedID
            self.selectedID = self.TaskTable.item(self.TaskTable.currentRow(), 0).text()
            self.loadTask()
            self.patentOrderOefnen(2)
        except:
            self.openErrorBox("Auftrag auswählen fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
        
    def selectLastTask(self):
        # Auftrag im TableWidget über die ID selektieren
        self.TaskTable.selectRow(int(self.lastSelectedID))       
            
    def loadTask(self):
        # Setzt die Eingabemaske komplett zurück // Erfasst die aktuell selektierte ID // Läd den Task mitels der ID von der SQL Datenbank
        # // Füllt die Eingabemaske mit den aus der Datenbank geladenen Informationen
        try:
            status = ""
            # Alles Taskeingaben zurücksetzen
            status = "Alles zurücksetzen_Material laden_Pointer initialisieren"
            self.resetTaskpreview()
            self.matLoad()
            #print("Daten aus Tasktable laden")
            # ID des Auftrags extrahieren und in Variable speichern
            selectedID = self.TaskTable.item(self.TaskTable.currentRow(), 0).text()
            # In der SQL Tabelle nach der ID suchen und den Pointer auf das Objekt in die Variable Speichern
            cursor = self.sqlpointer.suchenmitid(selectedID)
            selectedData = cursor.fetchone()
         
            status = "Labels setzen"
            # Daten in die jeweiligen Felder eintragen
            self.LiName.setText(selectedData[3])
            self.LiVorname.setText(selectedData[4])
            self.LiAbteilung.setText(selectedData[5])
            self.LiKostenstelle.setText(selectedData[6])
            self.LiTelefon.setText(selectedData[8])
            self.LiEMail.setText(selectedData[9])
            self.LiBauteil.setText(selectedData[10])
            self.LiStuckzahl.setText(str(selectedData[11]))
            self.LiDruckdauer.setText(str(selectedData[13]))
            self.LiGewicht.setText(str(selectedData[14]))
            self.LiDatenaufbereitungZeit.setText(str(selectedData[20]))
            self.LiNacharbeitZeit.setText(str(selectedData[21]))
            self.TextBeschreibung.setText(selectedData[15])
            status = "Labels setzen2"
            self.ComboFarbe.setCurrentText(str(selectedData[16]))
            self.ComboVorlage.setCurrentText(str(selectedData[17]))
            self.ComboBauteilgrose.setCurrentText(str(selectedData[18]))
            self.LFertigungsb.setCurrentText(str(selectedData[7]))
            self.LAuftragsnummer.setText(str(selectedData[0]))
            self.LEingangsdatum.setText(str(selectedData[1]))
            self.LWunschtermin.setText(str(selectedData[2]))     
            self.LAuftragsdauer.setText(self.calAuftragsdauer(str(selectedData[1]),selectedData[0]))
            self.LiFertigdatum.setText(str(selectedData[26]))
            self.CStatus.setCurrentIndex(selectedData[27]) 
            status = "Labels3 setzen"
            if selectedData[31] == None:
                self.Lfertig.setText(str(0))
            else:
                self.Lfertig.setText(str(selectedData[31]))   
            if selectedData[32] == None:
                self.Lausgeliefert.setText(str(0))
            else:
                self.Lausgeliefert.setText(str(selectedData[32])) 
            if selectedData[33] == None: 
                self.Cinfill.setCurrentText(str(20))
            else:
                self.Cinfill.setCurrentText(str(selectedData[33]))           
            if selectedData[29] == None:
                self.LIarbeitsstunde.setText(str(self.sqlpointer.variableAusgeben(4)[2]))
            else:
                self.LIarbeitsstunde.setText(str(selectedData[29]))              
            if selectedData[28] == None:
                self.LImaschinenstunde.setText(str(self.sqlpointer.variableAusgeben(4)[1]))
            else:
                self.LImaschinenstunde.setText(str(selectedData[28]))              
            if selectedData[30] == None:
                self.LIAufpreisfaktor.setText(str(self.sqlpointer.variableAusgeben(4)[3]))
            else:
                self.LIAufpreisfaktor.setText(str(selectedData[30]))
                        
            status = "Bild und berechnungen"
            
            self.c_material.setCurrentText(str(selectedData[19]))
            self.LiDatenaufbereitungKosten.setText(str(self.berechnungDatenaufbereitung(1,(float(self.LiDatenaufbereitungZeit.text())),self.ComboVorlage.currentIndex())))
            self.LiEinData.setText(str(self.berechnungDatenaufbereitung(2,(float(self.LiDatenaufbereitungZeit.text())),self.ComboVorlage.currentIndex())))
            self.LiNacharbeitKosten.setText(str(self.berechnungNacharbeit(1,self.LiNacharbeitZeit.text())))
            self.LiEinNach.setText(str(self.berechnungNacharbeit(2,self.LiNacharbeitZeit.text())))
            status = "Bild und berechnungen2"
            self.LiDruckkosten.setText(str(self.berechnungDruckkosten(1,float(self.LiDruckdauer.text()),self.ComboMaterial.currentIndex())))
            self.LiEinDruck.setText(str(self.berechnungDruckkosten(2,float(self.LiDruckdauer.text()),self.ComboMaterial.currentIndex())))
            self.LiMaterial.setText(str(self.berechnungMaterialkosten(1,int(self.LiGewicht.text()),self.ComboMaterial.currentIndex())))
            self.LiEinMat.setText(str(self.berechnungMaterialkosten(2,int(self.LiGewicht.text()),self.ComboMaterial.currentIndex())))
            self.LiAngebotspreis.setText(str(self.berechnungPreis(1)))
            self.LiEinHerst.setText(str(self.berechnungPreis(2)))
            status = "Bild und berechnungen3"
            self.LiEinsparung.setText(str(self.berechnungEinsparung(1)))
            self.LiEinEinsp.setText(str(self.berechnungEinsparung(2)))
            self.LiEinsparungFolgeteil.setText(str(self.berechnungEinsparungFolgeteile()))
            self.LiFolgeteilpreis.setText(str(self.berechnungFolgeteileStuck()))
            
            self.L_Fa1.setText(str(selectedData[34]))
            self.L_Fa2.setText(str(selectedData[35]))
            self.L_Fa3.setText(str(selectedData[36]))
            self.L_Fa4.setText(str(selectedData[37]))
            self.L_Fa5.setText(str(selectedData[38]))
            
            status = "Progressbar"
            prozent = ((int(self.Lfertig.text())/int(self.LiStuckzahl.text()))*100)
            self.progressBar.setValue(int(prozent))

            self.bildAktualisieren(selectedData[0])  
        except:
            self.openErrorBox("Auftrag Laden fehlgeschlagen!!! " + status,sys.exc_info()[0],sys.exc_info()[1])
            
    ###########################################################
    ###########################################################
    ############## INTERFACE BEREINIGEN #######################
    ############ AUFTRAGSLISTE NEU LADEN ######################
    ###########################################################  
    
    def resetTaskpreview(self):
        # Setzt die Eingabemaske komplett zurück
        self.LiName.clear()
        self.LiVorname.clear()
        self.LiAbteilung.clear() 
        self.LiKostenstelle.clear() 
        self.LiTelefon.clear() 
        self.LiEMail.clear() 
        self.LiBauteil.clear()
        self.LiStuckzahl.clear()
        self.LiDruckdauer.clear()
        self.LiGewicht.clear()
        self.LiFilterbauteil.clear() 
        self.LiDatenaufbereitungZeit.clear() 
        self.LiNacharbeitZeit.clear() 
        self.LiMaterial.clear() 
        self.LiDatenaufbereitungKosten.clear() 
        self.LiNacharbeitKosten.clear() 
        self.LiDruckkosten.clear() 
        self.LiAngebotspreis.clear() 
        self.LiFolgeteilpreis.clear() 
        self.LiEinsparung.clear() 
        self.LiEinsparungFolgeteil.clear() 
        self.TextBeschreibung.clear()
        self.Lausgeliefert.clear()
        self.Lfertig.clear()
        self.L_Fa1.clear()
        self.L_Fa2.clear()
        self.L_Fa3.clear()
        self.L_Fa4.clear()
        self.L_Fa5.clear()
        self.c_farbe.setCurrentIndex(0)
        self.c_vorlage.setCurrentIndex(0)
        self.c_bauteilgrose.setCurrentIndex(0)
        self.c_material.setCurrentIndex(0)
        self.Cinfill.setCurrentIndex(0)
        self.c_Status.setCurrentIndex(0)
        self.c_Fertigungsbereich.setCurrentIndex(0)
        self.LiFertigdatum.clear()
        self.LIarbeitsstunde.clear()
        self.LImaschinenstunde.clear()
        self.LIAufpreisfaktor.clear()
        self.LiEinData.clear()
        self.LiEinNach.clear()
        self.LiEinDruck.clear()
        self.LiEinMat.clear()
        self.LiEinHerst.clear()
        self.LiEinEinsp.clear()
        
        self.matLoad()
        
    def resetTasktable(self):
        # Tasktable leeren
        self.TaskTable.clearSelection()
        if self.TaskTable.rowCount() == 1:
            self.TaskTable.removeRow(1)
        elif self.TaskTable.rowCount() > 1:
            for i in reversed(range(self.TaskTable.rowCount())):
                self.TaskTable.removeRow(i)
        
    def tasktableRefresh(self):
        # Läd die Auftragsliste aus der SQL Tabelle und fügt sie in das TableWidget
        try:
            result = self.sqlpointer.getAuftragsliste(1)
            self.TaskTable.setRowCount(0)
            for row_number, row_data in enumerate(result):
                self.TaskTable.insertRow(row_number)
                for colum_number, data in enumerate(row_data):
                    if colum_number == 0:
                        self.TaskTable.setItem(row_number, 0, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 1:
                        self.TaskTable.setItem(row_number, 1, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 3:
                        self.TaskTable.setItem(row_number, 2, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 10:
                        self.TaskTable.setItem(row_number, 3, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 11:
                        stueckzahl = str(data)
                        self.TaskTable.setItem(row_number, 4, QtWidgets.QTableWidgetItem(str(data)))    
                    if colum_number == 27:
                        status = str(data)
                        self.TaskTable.setItem(row_number, 5, QtWidgets.QTableWidgetItem(str(data)))
                        if data == 0:
                            self.TaskTable.item(row_number, 5).setBackground(QtGui.QColor(255, 0, 0))
                            self.TaskTable.item(row_number, 5).setForeground(QtGui.QColor(255, 0, 0))
                        if data == 1:
                            self.TaskTable.item(row_number, 5).setBackground(QtGui.QColor(255,165,0))
                            self.TaskTable.item(row_number, 5).setForeground(QtGui.QColor(255,165,0))
                        if data == 2:
                            self.TaskTable.item(row_number, 5).setBackground(QtGui.QColor(255,165,0))
                            self.TaskTable.item(row_number, 5).setForeground(QtGui.QColor(255,165,0))
                        if data == 3:
                            self.TaskTable.item(row_number, 5).setBackground(QtGui.QColor(255, 255, 0))
                            self.TaskTable.item(row_number, 5).setForeground(QtGui.QColor(255, 255, 0))
                        if data == 4:
                            self.TaskTable.item(row_number, 5).setBackground(QtGui.QColor(255, 255, 0))  
                            self.TaskTable.item(row_number, 5).setForeground(QtGui.QColor(255, 255, 0))
                        if data == 5:
                            self.TaskTable.item(row_number, 5).setBackground(QtGui.QColor(255, 255, 0))  
                            self.TaskTable.item(row_number, 5).setForeground(QtGui.QColor(255, 255, 0))
                        if data == 6:
                            self.TaskTable.item(row_number, 5).setBackground(QtGui.QColor(124,252,0))  
                            self.TaskTable.item(row_number, 5).setForeground(QtGui.QColor(124,252,0))
                        if data == 7:
                            self.TaskTable.item(row_number, 5).setBackground(QtGui.QColor(124,252,0))
                            self.TaskTable.item(row_number, 5).setForeground(QtGui.QColor(124,252,0))
                    if colum_number == 31:
                        fertig = str(data)
                 
                if int(status) != 7:
                    if stueckzahl == "None":
                        stueckzahl = 0
                    if fertig == "None":
                        fertig = 0
                    if int(stueckzahl) == 0:
                        prozent = 0
                    else:
                        prozent = (int(fertig)/int(stueckzahl))*100          
                    pro = QtWidgets.QProgressBar()
                    pro.setAlignment(Qt.AlignCenter)
                    pro.setStyleSheet('''QProgressBar{
                                                      border: 1px solid;
                                                      border-color:rgb(255, 255, 255);
                                                      font-weight: bold;
                                                      }
                                                      QProgressBar::chunk {
                                                      background-color:  rgb(0, 200, 0);
                                                      }''')
                    pro.setValue(int(prozent))
                    self.TaskTable.setCellWidget(row_number, 6, pro)
                
 
                                                  
        except:
            self.openErrorBox("Auftragsliste Aktualisieren fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
            
    def matLoad(self):
        try:
            self.c_material.clear()
            self.c_material.addItems(self.sqlpointer.getML(3))
        except:
            print("matLoad fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred") 
            
    ###########################################################
    ###########################################################
    ############## AUFTRÄGE FILTERN ###########################
    ###########################################################
    ###########################################################  
    
    def auftragslisteFiltern(self):
        # Initialisiert sie Zahl für die momentan angezeigten Aufträge // Läd die Menge der im Auftragstable angezeigten Aufträge
        # // Erfasst den aktuellen Filtrstatus // Filtert die Aufträge anhand der Kriterien // Setzt die Auftragszahl in das Label
        anzahl = 0
        self.auftragszahl.clear()
        
        aktuelles_datum = datetime.datetime.now()
        jahrNow = aktuelles_datum.year
        
        
        try:
            rows = self.TaskTable.rowCount()
            tempindex = self.FilterStatus.currentIndex()
            if tempindex == 11:
                liste = self.sqlpointer.sucheFehlendeDaten()
            if tempindex == 13:
                liste2 = self.sqlpointer.fehlendeFAnummer()
            
            for i in range(rows):
                #Alle Aufträge verstecken
                self.TaskTable.hideRow(i)
                if (int(self.TaskTable.item(i, 5).text()) == tempindex):
                    # Wenn Filterstatus gleich Auftragsstatus dann Auftrag anzeigen 
                    self.TaskTable.showRow(i)
                    anzahl += 1
                elif tempindex == 8:
                    # Wenn Filterstatus gleich "alles anzeigen" dann zeige alle Aufträge an
                    self.TaskTable.showRow(i)
                    anzahl += 1 
                elif tempindex == 9:
                    # Wenn Filterstatus gleich "nicht ausgeliefert anzeigen" dann zeige alle Aufträge an die noch nicht den Status "Ausgeliefert" haben
                    if (int(self.TaskTable.item(i, 5).text()) <= 6):
                        self.TaskTable.showRow(i)
                        anzahl += 1
                elif tempindex == 10:
                    # Wenn Filterstatus gleich "Dieses jahr anzeigen" dann zeige alle Aufträge an die dieses Jahr eingetragen wurden
                    datum_objekt = datetime.datetime.strptime(self.TaskTable.item(i,1).text(), "%d-%m-%Y")
                    jahr = datum_objekt.year
                    if (jahrNow==jahr):
                        self.TaskTable.showRow(i)
                        anzahl += 1
                elif tempindex == 11:
                    # Wenn Filterstatus gleich "Fehlende Daten" dann zeige alle Aufträge an die noch keine Daten bei Gewicht oder Druckdauer haben
                    for k in liste:
                        #print(type(i))
                        if(self.TaskTable.item(k,0).text()) == str(k):   
                            self.TaskTable.showRow(k)
                elif tempindex == 12:
                    # Wenn Filterstatus gleich "Fehlendes Bild" dann zeige alle Aufträge an die noch kein Bild haben
                    if self.auftragsbildSuchen(self.TaskTable.item(i,0).text()) == False:
                        self.TaskTable.showRow(i)
                elif tempindex == 13:
                    # Wenn Filterstatus gleich "Fehlende FA Beleg Daten" dann zeige alle Aufträge an die noch keine Daten haben
                    for j in liste2:
                        if(self.TaskTable.item(i,0).text()) == str(j):   
                            self.TaskTable.showRow(i)
                         
            self.auftragszahl.setText(str(anzahl))
        except: 
            self.openErrorBox("Auftragsliste filtern fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
            
    def nachNameFiltern(self): 
        # Initialisiert sie Zahl für die momentan angezeigten Aufträge // Läd die Menge der im Auftragstable angezeigten Aufträge
        # // Setzt den Filterwert // Filtert anhand der angegebenen Filterkriterien
        anzahl = 0
        self.auftragszahl.clear()
        try:
            rows = self.TaskTable.rowCount()
            tempindex = (str(self.LiFiltername.text())).lower()
            
            for i in range(rows):
                self.TaskTable.hideRow(i)
                tasktableitem = (str(self.TaskTable.item(i, 2).text())).lower()
                if (tasktableitem.startswith(tempindex)):
                    self.TaskTable.showRow(i)
                    anzahl += 1
            
        except:
            self.openErrorBox("Aufträge nach Name filtern fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
        
        self.auftragszahl.setText(str(anzahl)) 
        
    def filterShow(self):
        # Filter Buttons und Eingaben anzeigen oder verbergen
        if (self.butInvis == 1):
            self.showFilterbutton.setText("Weitere Filter anzeigen")
            self.filterbutton.hide()
            self.nameLabel.hide()
            self.nameBauteil.hide()
            self.nameFertigungsbereich.hide()
            self.nameEinsparung.hide()
            self.LiFiltername.hide()
            self.LiFilterbauteil.hide()
            self.ComboFilterFB.hide()
            self.FilterEinsparung.hide()
            self.butInvis = 0
        else:
            self.showFilterbutton.setText("Weniger Filter anzeigen")
            self.filterbutton.show()
            self.nameLabel.show()
            self.nameBauteil.show()
            self.nameFertigungsbereich.show()
            self.nameEinsparung.show()
            self.LiFiltername.show()
            self.LiFilterbauteil.show()
            self.ComboFilterFB.show()
            self.FilterEinsparung.show()
            self.butInvis = 1
            
    def resetFilterFunc(self):
        # Filter auf Variablenwert = 8 zurücksetzen
        try:
            self.filterVariable = 8
            self.FilterStatus.setCurrentIndex(self.filterVariable)
            self.auftragslisteFiltern()
        except:
            self.openErrorBox("Filter zurücksetzen fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
            
    ###########################################################
    ###########################################################
    ############## BERECHUNGSFUNKTIONEN #######################
    ###########################################################
    ###########################################################

    def calAuftragsdauer(self,eingangsdatum,ident):
        # Auftragsdauer berechnen mit dem Eingangsdatum und dem Aktuellen Datum
        try:
            datum = (self.sqlpointer.suchenmitid(ident)).fetchone()[26]
            today = datetime.datetime.today()
            ein_date_object = datetime.datetime.strptime(eingangsdatum, '%d-%m-%Y').date()
                
            if (datum == ""):
                diff = today.date() - ein_date_object
            else:
                end_date_object = datetime.datetime.strptime(datum, '%d-%m-%Y').date()
                diff = end_date_object - ein_date_object
                
                if int(str(diff.days)) < 0:
                    self.sqlpointer.fertigDatumSetzen(ident,"")
                    
            diffString = str(diff.days)
            #print(diffString)
            
            diffString = diffString + " Tage"
        
            return diffString
        except:

            self.openErrorBox("Auftragsdauer berechnen fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])    
            
            
    def berechnungDatenaufbereitung(self,variante, eingabeDatenaufbereitung, vorlage):
        try:   
            if variante == 1:
                datenaufbereitung = round((eingabeDatenaufbereitung * self.getVorlageFaktor(vorlage) * float(self.LIarbeitsstunde.text())),2)
            elif variante == 2:
                datenaufbereitung = round((eingabeDatenaufbereitung * self.getVorlageFaktor(vorlage) * float(self.LIarbeitsstunde.text()))/int(self.LiStuckzahl.text()),2)
            return datenaufbereitung
        except:
            self.openErrorBox("Berechnung Datenaufbereitung fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
                      
    def berechnungNacharbeit(self,variante, eingabeNachbereitung):
        try:
            if variante == 1:
                nacharbeitskosten = round(((float(eingabeNachbereitung) * float(self.LIarbeitsstunde.text())) * int(self.LiStuckzahl.text())),2)
            elif variante == 2:
                nacharbeitskosten = round(((float(eingabeNachbereitung) * float(self.LIarbeitsstunde.text()))),2)
            return nacharbeitskosten
        except:
            self.openErrorBox("Berechnung Nacharbeit fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
            
    def berechnungDruckkosten(self,variante, eingabeDruckdauer, material): 
        if variante == 1:
            druckkosten = round((eingabeDruckdauer * float(self.LImaschinenstunde.text()) * int(self.LiStuckzahl.text())),2)
        elif variante == 2:
            druckkosten = round((eingabeDruckdauer * float(self.LImaschinenstunde.text()) ),2)
        return druckkosten
        
    def berechnungMaterialkosten(self,variante,eingabeGewicht, material):
        try:
            if variante == 1:
                matPreis, matGramm = self.sqlpointer.getMatData(self.ComboMaterial.currentText())
                materialkosten2 = round((eingabeGewicht * (float(matPreis)/float(matGramm)) * int(self.LiStuckzahl.text())),2)
            elif variante == 2:
                matPreis, matGramm = self.sqlpointer.getMatData(self.ComboMaterial.currentText())
                materialkosten2 = round((eingabeGewicht * (float(matPreis)/float(matGramm))),2)
        except:
            self.openErrorBox("berechnungMaterialkosten fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
        
        return materialkosten2
        
    def berechnungPreis(self,variante):
        folgeteil = 1
        
        if variante == 1:
            preis = round(((float(self.LiDatenaufbereitungKosten.text()) * folgeteil) + float(self.LiDruckkosten.text()) + float(self.LiNacharbeitKosten.text()) + float(self.LiMaterial.text())),2)
        elif variante == 2:
            preis = round(((float(self.LiDatenaufbereitungKosten.text()) * folgeteil) + float(self.LiDruckkosten.text()) + float(self.LiNacharbeitKosten.text()) + float(self.LiMaterial.text()))/int(self.LiStuckzahl.text()),2)
        return preis
    
    def berechnungEinsparung(self,variante): 
        preis = float(self.LiAngebotspreis.text())
        externerPreis = preis * float(self.LIAufpreisfaktor.text())
        einsparung = round((externerPreis - preis),2)
        if variante == 1:
            einsparung = einsparung
        elif variante == 2:
            einsparung = round(einsparung / int(self.LiStuckzahl.text()),2)

        return einsparung
        
    def berechnungEinsparungFolgeteile(self):   
        preisEinzel = round((float(self.LiDruckkosten.text()) + float(self.LiNacharbeitKosten.text()) + float(self.LiMaterial.text())),2)  
        einsparungFolge = round((preisEinzel * float(self.sqlpointer.variableAusgeben(4)[3])) * int(self.LiStuckzahl.text()),2)

        return einsparungFolge
        
    def berechnungFolgeteileStuck(self):  
        preisEinzel = round((float(self.LiDruckkosten.text()) + float(self.LiNacharbeitKosten.text()) + float(self.LiMaterial.text())),2)  
        stuckpreisFolge = round((preisEinzel / int(self.LiStuckzahl.text())),2)

        return stuckpreisFolge

    ###########################################################
    ###########################################################
    ############### HILFSFUNKTIONEN ###########################
    ###########################################################
    ###########################################################     
    
    def email(self):
        try:
            selectedID = self.TaskTable.item(self.TaskTable.currentRow(), 0).text()
            cursor = self.sqlpointer.suchenmitid(selectedID)
            selectedData = cursor.fetchone()
            
            email = selectedData[9]
            bauteilname = selectedData[10]
            name = selectedData[3]
            
            htmltest_anfang = """
            <!DOCTYPE html>
            <html>  
                <body>
                    Guten Tag Herr {name},
                    <div>
                    <p style="font-family:Garamond;">Ihre 3D-Druck Teile "{bauteilname}" sind fertig und können abgeholt werden.
                    <br></br>
                    <br></br>
                    Wir sitzen in der <b>Halle 2 HG F4</b></p>
                    </div>
                            
                    <p style="font-family:Garamond;font-size:20">Mit freundlichen Gr&uuml;&szlig;en
                    <br></br>
                    <p1 style="font-family:Garamond;color:red"> <b>Mike Haftendorn</b> </p1>
                    <br></br>
                    <div style="font-family:Garamond">
                    <b>3D-Druck</b> Modell und Schablonenbau (CPK-U2/W2B) Halle 2 HG F4
                    <br></br>
                    Volkswagen Aktiengesellschaft  Brieffach 014/ 42950 D-34219 Baunatal
                    <br></br>
                    Tel: 0561 490 102449
                    <br></br>
                    <b>3d-druck.ks@volkswagen.de</b></p>
                    </div>
                    <br></br>
                    <p style="font-family:Garamond;font-size:12">Dies ist eine automatisch erstellte E-Mail aus der Auftragsverwaltungssoftware des 3D-Drucks</p1>
                    <br></br>
                    <p  style="font-family:Garamond;font-size:12">Volkswagen Aktiengesellschaft Sitz: Wolfsburg Registergericht: Amtsgericht Braunschweig  HRB Nr.: 100484 Vorsitzender des Aufsichtsrats: Hans Dieter P&ouml;tsch Vorstand: Oliver Blume (Vorsitzender), Arno Antlitz, Ralf Brandst&auml;tter, Gernot D&ouml;llner, Manfred D&ouml;ss, Gunnar Kilian, Thomas Sch&auml;fer, Thomas Schmall-von Westerholt, Hauke Stars Wichtiger Hinweis: Die vorgenannten Angaben werden jeder E-Mail automatisch hinzugef&uuml;gt und lassen keine R&uuml;ckschl&uuml;sse auf den Rechtscharakter der E-Mail zu.   Informationen zum Umgang mit Ihren personenbezogenen Daten finden Sie unter https://www.volkswagen.de/de/mehr/rechtliches/datenschutzerklaerung-allgemeine-kommunikation.html</p>
                </body>
            </html>
    
                    """.format(bauteilname=bauteilname,name=name)
    
            
            htmltestgesamt = htmltest_anfang 
        
        
            outlook = win32.Dispatch('Outlook.Application')
            outlookNS = outlook.GetNameSpace('MAPI')
            
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = '3D Druck Teile sind fertig'
            mail.Body = 'Test'
            mail.HTMLBody = htmltestgesamt   
    
            mail.Display(True)
            
        except:
            print("emailtest fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")    
        
    def bildAktualisieren(self,auftragsnummer):
        # Läd das Auftragsbild mittels der Auftragsnummer
        try:
            pfad = self.sqlpointer.variableAusgeben(3)[3] + "/Auftrag_"+ str(auftragsnummer) + ".JPG"
            if (os.path.isfile(pfad)):
                print("Auftrag")
                bildstring = self.sqlpointer.variableAusgeben(3)[3] + "/Auftrag_" + str(auftragsnummer) + ".JPG"
                self.vorschauBild.setPixmap(QPixmap(bildstring))
                print("Auftrag3")
            else:
                print("Kein Bild")
                bildstring = self.sqlpointer.variableAusgeben(3)[3] + "/KEIN_BILD.JPG"
                self.vorschauBild.setPixmap(QPixmap(bildstring))
        except:
            self.openErrorBox("Bild aktualisieren fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])

        image_reader = QImageReader('e:/test.jpg')  # Passe den Pfad zu deinem Bild an
        if image_reader.size().isValid():
            image = image_reader.read()
            pixmap = QPixmap(image)
            self.label_69.setPixmap(pixmap)
        else:
            print(f"Fehler beim Lesen des Bildes: {image_reader.errorString()}")    
            
    def load_operation_that_might_fail(self):
        print("hier")
        with open('e:/test.jpg', 'rb') as f:
            return f.read()

    def auftragsbildSuchen(self,auftragsnummer):
        #print(auftragsnummer)
        try:
            pfad = self.sqlpointer.variableAusgeben(3)[3] + "/Auftrag_"+ str(auftragsnummer) + ".JPG"
            
            if (os.path.isfile(pfad)):
                ausgabe = True
            else:
                ausgabe = False
            return ausgabe

        except:
            self.openErrorBox("Bild aktualisieren fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
                     
    def openFold(self):
        # Auftragsordner mit den Eingaben für Name und Vorname sowie ID von SQL laden und öffnen
        try:
            destination = self.sqlpointer.getFolderName(self.LiName.text(),self.LiVorname.text(),int(self.selectedID))
            tempdestination = destination.replace("/","\\")
            fullstring = tempdestination
            
            path = fullstring
            path = os.path.realpath(path)
            os.startfile(path) 
        except:
            self.openErrorBox("Ordner öffnen fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
            
    def copyLinkClipboard(self):
        try:
            destination = self.sqlpointer.getFolderName(self.LiName.text(),self.LiVorname.text(),int(self.selectedID))
            tempdestination = destination.replace("/","\\")
            fullstring = os.path.realpath(tempdestination)
            
            data = QMimeData()
            data.setText(fullstring)
            self.app.clipboard().setMimeData(data)
        except:
            self.openErrorBox("copyLinkClipboard fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
        
            
    def getFolder(self,name,vorname,ident):
        try:
            destination = self.sqlpointer.getFolderName(self.LiName.text(),self.LiVorname.text(),int(self.selectedID))
            tempdestination = destination.replace("/","\\")
            fullstring = tempdestination
            
            path = fullstring
            path = os.path.realpath(path)
            #os.startfile(path) 
            
            return path
        except:
            self.openErrorBox("Ordner öffnen fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
              
    def auftragLoeschen2(self,name,vorname,auftragsnummer):
        #Funktioniert noch nicht // Einfach über die Datenbank Löschen
        destination = self.sqlpointer.getFolderName(name,vorname,auftragsnummer)
        tempdestination = destination.replace("/","\\")
        fullstring = os.path.abspath(".") + "\\" + tempdestination
        
        try:
            shutil.rmtree(fullstring)
        except:
            self.openErrorBox("Auftrag löschen fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
            
        self.sqlpointer.deletefromAuftragsliste(auftragsnummer)
       
    def openMessageBox(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Auftrag auswählen!!!")
        msg.setWindowTitle("Keinen Auftrag Ausgewählt oder Ordner nicht gefunden!!!")
        msg.exec_()
        
    def openMessageBox2(self, message):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText(message)
        msg.setWindowTitle(message)
        msg.exec_()
        
    def openErrorBox(self, error1,error2,error3):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("\n\n" + str(error1) + "\n\n\n" + str(error2) + "\n" + str(error3))
        msg.setWindowTitle("Es ist ein Fehler aufgetreten!!!")
        msg.exec_()
        
    def auftragLoeschen(self):
        # Auftrag löschen // Auftragstablewidget leeren // Auftragstaqblewidget neu füllen // Filter anwenden
        self.sqlpointer.deletefromAuftragsliste(self.selectedID)
        self.resetTasktable()
        self.tasktableRefresh()
        self.auftragslisteFiltern()
        
    def getMaterialFaktor(self,material):
        switcher={
                  0:1.0, #PLA
                  1:1.1, #PLAPlus
                  2:1.2, #PETG
                  3:1.4, #GREENTEC
                  4:1.5, #NYLONGF
                  5:1.5, #NYLONCF
                  6:1.8  #FILAFLEX
        }
        return switcher.get(material,"Eingabe nicht korrekt!!")       
    
    def getVorlageFaktor(self,vorlage):
        switcher={
                  0:1.4, #Skizze
                  1:1.3, #Zeichnung
                  2:1.5, #Musterteil
                  3:1.0, #CAD Daten
                  4:1.1  #KVS
        }
        return switcher.get(vorlage,"Eingabe nicht korrekt!!")        
      
    def hilfspointer(self,pointer):
        self.hilfs = pointer
            
    def getStyleFromHilfs(self):
        self.setStyleSheet(self.hilfs.getStyle())
        
    def patentOrderOefnen(self,variante):
        # Öffnet den Ordner in dem das Dokument Patent.pdf liegt // Bei Variante 2 wird nur geschaut ob eine Patent.pdf Datei vorhanden ist und visuell über den Button angezeigt
        file_Online = False
        try:     
            destination = self.sqlpointer.getFolderName(self.LiName.text(),self.LiVorname.text(),int(self.selectedID))
            tempdestination = destination.replace("/","\\")
            fullstring = tempdestination + "/Patent.pdf"
            
            # Pfad zur Datei, die Sie prüfen möchten
            path = fullstring
            
            # Prüfen, ob die Datei existiert
            if os.path.isfile(path):
                #print("Die Datei existiert.")
                file_Online = True
            else:
                file_Online = False
                #print("Die Datei existiert nicht.")
                
            if (variante==1):
                if(file_Online==True):
                     # Adobe Acrobat-Anwendung erstellen  
                    acrobat = win32com.client.Dispatch("AcroExch.App")
                        
                    av_doc = win32com.client.Dispatch("AcroExch.AVDoc")
                    av_doc.Open(path, "")
                    pd_doc = av_doc.GetPDDoc()
                    
                    # PDF-Dokument anzeigen
                    av_doc.BringToFront()
                elif (file_Online==False):
                    self.openMessageBox2("Keine Patentdatei vorhanden!!!")
            if (variante==2):
                if(file_Online == False):
                    self.Patentbutton.setStyleSheet("QPushButton { background-color: #FF4500; }")
                elif (file_Online == True):
                    self.Patentbutton.setStyleSheet("QPushButton { background-color: #00EE00; }")
            
        except:
            self.openErrorBox("Patentordner öffnen fehlgeschlagen!!!",sys.exc_info()[0],sys.exc_info()[1])
            
            