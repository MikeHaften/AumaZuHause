from PyQt5 import QtWidgets,uic
import sys,os
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon

import win32com.client as win32
import materialbestandUi


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)



class Ui(QtWidgets.QDialog, materialbestandUi.Ui_materialbestand):
    def __init__(self,sqlpointer,parent=None):
        super(Ui,self).__init__(parent)
        
        self.setupUi(self)
        
        self.programmpfad = os.path.abspath(".")
        
        icon = QIcon(resource_path(self.programmpfad +"/icons/auma_farbe.png"))
        # Icon als Anwendungssymbol setzen
        self.setWindowIcon(icon)
        # Titelleisten-Icon temporär ändern
        self.setProperty("windowIcon", icon)
        
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint)  # Hinzufügen von Minimieren- und Maximieren-Buttons
        
        self.sqlpointer = sqlpointer
        
        self.backbutton = self.backToMainMenu
        self.backbutton.clicked.connect(self.backToMainmenu)
        self.hinzubutton = self.AddButton
        self.hinzubutton.clicked.connect(self.MaterialHinzu)
        self.Updatebutton = self.UpdateButton
        self.Updatebutton.clicked.connect(self.materialUpdaten)
        self.deletebutton = self.delButton
        self.deletebutton.clicked.connect(self.MaterialDelete)
        self.Minusbutton = self.MinusButton
        self.Minusbutton.clicked.connect(lambda:self.plusButtonFunktion(2))
        self.Plusbutton = self.PlusButton
        self.Plusbutton.clicked.connect(lambda:self.plusButtonFunktion(1))
        
        self.delete2button = self.delete2But
        self.delete2button.clicked.connect(self.MaterialDelete2)
        self.BestellHinzuButton = self.bestellHinzuBut
        self.BestellHinzuButton.clicked.connect(self.bestellungHinzufugen)
        self.BestellenButton = self.BestellenBut
        self.BestellenButton.clicked.connect(self.bestellenAusloesen)
        
        self.LiMatNummer = self.MatNrEdit
        self.LIGewicht = self.GewEdit
        self.LiFarbe = self.FarbeEdit
        self.LiArt = self.ArtEdit
        
        self.LIBestand = self.li_bestand
        #self.LiEntnommen = self.li_entnommen
        self.LiBestellt = self.li_bestellt
        
        self.bestellmenge = self.li_counter
        
        self.CHersteller = self.HerstBox
        self.Cdurchmesser = self.DurchBox
        
        self.CFilterBox = self.FilterBox
        self.CFilterBox.currentIndexChanged.connect(self.materiallisteFiltern)
        
        self.MaterialTable = self.materialtable
        self.MaterialTable.verticalHeader().setDefaultSectionSize(20)
        self.MaterialTable.setColumnWidth(0, 40)  # ID
        self.MaterialTable.setColumnWidth(1, 120)  # Materialnummer
        self.MaterialTable.setColumnWidth(2, 100)  # Durchmesser
        self.MaterialTable.setColumnWidth(3, 100)  # Gewicht
        self.MaterialTable.setColumnWidth(4, 100)  # Art
        self.MaterialTable.setColumnWidth(5, 120)  # Farbe
        self.MaterialTable.setColumnWidth(6, 150)  # Hersteller
        self.MaterialTable.setColumnWidth(7, 120)  # Bestand
        self.MaterialTable.setColumnWidth(8, 120)  # Bestellt
        
        self.MaterialTable.horizontalHeader().setStretchLastSection(True)
        self.MaterialTable.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.MaterialTable.itemClicked.connect(self.selectMaterial)
        self.MaterialTable.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.MaterialTable.verticalHeader().setVisible(False)
        
        
        self.MaterialTable2 = self.materialtable2
        self.MaterialTable2.verticalHeader().setDefaultSectionSize(20)
        self.MaterialTable2.setColumnWidth(0, 40)  # ID
        self.MaterialTable2.setColumnWidth(1, 120)  # Materialnummer
        self.MaterialTable2.setColumnWidth(2, 100)  # Durchmesser
        self.MaterialTable2.setColumnWidth(3, 100)  # Gewicht
        self.MaterialTable2.setColumnWidth(4, 100)  # Art
        self.MaterialTable2.setColumnWidth(5, 120)  # Farbe
        self.MaterialTable2.setColumnWidth(6, 150)  # Hersteller
        self.MaterialTable2.setColumnWidth(7, 90)  # Menge
        
        self.MaterialTable2.horizontalHeader().setStretchLastSection(True)
        self.MaterialTable2.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.MaterialTable2.itemClicked.connect(self.selectMaterial2)
        self.MaterialTable2.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.MaterialTable2.verticalHeader().setVisible(False)
            
    def materialtableRefresh(self):
        # Obere Liste aus SQL Tabelle Laden und in das ListenWidget einfügen
        try:
            result = self.sqlpointer.getMaterialliste()
            self.MaterialTable.setRowCount(0)
            for row_number, row_data in enumerate(result):
                self.MaterialTable.insertRow(row_number)
                for colum_number, data in enumerate(row_data):
                    if colum_number == 0:
                        self.MaterialTable.setItem(row_number, 0, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 1:
                        self.MaterialTable.setItem(row_number, 1, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 2:
                        self.MaterialTable.setItem(row_number, 2, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 3:
                        self.MaterialTable.setItem(row_number, 3, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 4:
                        self.MaterialTable.setItem(row_number, 4, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 5:
                        self.MaterialTable.setItem(row_number, 5, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 6:
                        self.MaterialTable.setItem(row_number, 6, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 7:
                        self.MaterialTable.setItem(row_number, 7, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 8:
                        self.MaterialTable.setItem(row_number, 8, QtWidgets.QTableWidgetItem(str(data)))           
                            
        except:
            print("TaskTable Füllen fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
                             
    def materialtable2Refresh(self):
        # Untere Liste aus SQL Tabelle Laden und in das ListenWidget einfügen
        try:
            result = self.sqlpointer.getMaterialliste2()
            self.MaterialTable2.setRowCount(0)
            for row_number, row_data in enumerate(result):
                self.MaterialTable2.insertRow(row_number)
                for colum_number, data in enumerate(row_data):
                    if colum_number == 0:
                        self.MaterialTable2.setItem(row_number, 0, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 1:
                        self.MaterialTable2.setItem(row_number, 1, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 2:
                        self.MaterialTable2.setItem(row_number, 2, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 3:
                        self.MaterialTable2.setItem(row_number, 3, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 4:
                        self.MaterialTable2.setItem(row_number, 4, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 5:
                        self.MaterialTable2.setItem(row_number, 5, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 6:
                        self.MaterialTable2.setItem(row_number, 6, QtWidgets.QTableWidgetItem(str(data)))
                    if colum_number == 7:
                        self.MaterialTable2.setItem(row_number, 7, QtWidgets.QTableWidgetItem(str(data)))

        except:
            print("TaskTable Füllen fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
        
    def resetMaterialtable(self):
        # Obere Liste einmal Leeren
        self.MaterialTable.clearSelection()
        if self.MaterialTable.rowCount() == 1:
            self.MaterialTable.removeRow(1)
        elif self.MaterialTable.rowCount() > 1:
            for i in reversed(range(self.MaterialTable.rowCount())):
                self.MaterialTable.removeRow(i)   
                
    def resetMaterialtable2(self):
        # Untere Liste einmal Leeren
        self.MaterialTable2.clearSelection()
        if self.MaterialTable2.rowCount() == 1:
            self.MaterialTable2.removeRow(1)
        elif self.MaterialTable2.rowCount() > 1:
            for i in reversed(range(self.MaterialTable2.rowCount())):
                self.MaterialTable2.removeRow(i)
                
    def MaterialHinzu(self):
        # Material der SQL Tabelle hinzufügen // Materialtabelle einmal Leeren // Materialtabelle neu Laden // Eingabemaske einmal zurücksetzen
        self.sqlpointer.MaterialHizufugen(self.LiMatNummer.text(),self.Cdurchmesser.currentText(),self.LIGewicht.text(),self.LiArt.text(),self.LiFarbe.text(),self.CHersteller.currentText())    
        self.refreshListe1()
        
    def selectMaterial(self):
        # ID der ausgewählten Zeile in der oberen Liste erfassen // Funktion zum laden der Daten in die Eingabemaske ausführen
        try:
            self.selectedID = self.MaterialTable.item(self.MaterialTable.currentRow(), 0).text()
            self.loadMaterial()
        except:
            print("selectMat fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            
    def selectMaterial2(self):
        # ID der augewählten Zeile in der unteren Liste erfassen
        try:
            self.selectedID2 = self.MaterialTable2.item(self.MaterialTable2.currentRow(), 0).text() 
        except:
            print("selectMat2 fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            
    def loadMaterial(self):
        # Mit der ID die Daten aus der SQL Tabelle Laden und in die Eingabemaske übertragen
        self.resetMaterialUI() 
        selectedID = self.MaterialTable.item(self.MaterialTable.currentRow(), 0).text()
        cursor = self.sqlpointer.suchenmitMaterialid(selectedID)
        
        selectedData = cursor.fetchone()
        
        self.LiMatNummer.setText(str(selectedData[1]))
        self.Cdurchmesser.setCurrentText(selectedData[2])
        self.LIGewicht.setText(selectedData[3])
        self.LiArt.setText(selectedData[4])
        self.LiFarbe.setText(selectedData[5])
        self.CHersteller.setCurrentText(selectedData[6])
        self.LIBestand.setText(str(selectedData[7]))
        self.LiBestellt.setText(str(selectedData[8]))
      
    def resetMaterialUI(self):
        # Eingabemaske einmal Leeren
        self.LiMatNummer.clear()
        self.LIGewicht.clear()
        self.LiFarbe.clear()
        self.LiArt.clear()
        self.CHersteller.setCurrentIndex(0)
        self.Cdurchmesser.setCurrentIndex(0)
        self.LIBestand.clear()
        self.LiBestellt.clear()
        
    def materialUpdaten(self):
        # Eigenschaften eines Materials in der SQL Tabelle durch die Eingabemaske aktualisieren // Liste einmal Leeren // Liste neu laden
        self.sqlpointer.MaterialUpdate(self.LiMatNummer.text(),self.Cdurchmesser.currentText(),self.LIGewicht.text(),self.LiArt.text(),self.LiFarbe.text(),self.CHersteller.currentText(),self.selectedID,int(self.LIBestand.text()),int(self.LiBestellt.text()))
        self.refreshListe1()
            
    def MaterialDelete(self):
        # Material anhand der ID aus der ersten Liste löschen // Einmal die Eingabemaske leeren // Einmal die Materialliste leeren // Und Materialliste neu laden
        self.sqlpointer.deletefromMaterialliste(self.selectedID)
        self.refreshListe1()
        
    def MaterialDelete2(self):
        # Material anhand der ID aus der zweiten Liste löschen // Einmal die Eingabemaske leeren // Einmal die Materialliste leeren // Und Materialliste neu laden
        self.sqlpointer.deleteFromBestellliste(self.selectedID2)
        self.refreshListe2()
        
    def plusButtonFunktion(self,variante):
        # ID der ausgewählten Zeile in der ersten Liste erfassen // Mit der ID das augewählte Material aus der SQL Tabelle laden 
        # // Den Bestand um 1 erhöhen und die bestellung um 1 verringern bei Variante 1 // Bei Variante 2 den Bestand um 1 veringern
        # // Daten in die SQL Tabelle Updaten // Liste1 neu Laden // Vorherige Zeile wieder Selektieren
        try:
            selectedID = self.MaterialTable.item(self.MaterialTable.currentRow(), 0).text()
            cursor = self.sqlpointer.suchenmitMaterialid(selectedID)
            selectedData = cursor.fetchone()
            
            if (variante==1):
                bestand = int(selectedData[7]) + 1
                if int(selectedData[8]) >= 1: 
                    bestellt = int(selectedData[8]) -1
                else:
                    bestellt = int(selectedData[8])
            elif (variante==2):
                if int(selectedData[7]) >= 1:
                    bestand = int(selectedData[7]) - 1
                    bestellt = int(selectedData[8])
            
            self.sqlpointer.MaterialUpdate(str(selectedData[1]),str(selectedData[2]),str(selectedData[3]),str(selectedData[4]),str(selectedData[5]),str(selectedData[6]),str(selectedData[0]),bestand,bestellt)
            self.refreshListe1()
            
            self.selectRowByID(selectedID)
        except:
            print("plus fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")        
        
    def bestellungHinzufugen(self):
        # ID der ausgewählten Zeile in der ersten Liste erfassen // Mit der ID das augewählte Material aus der SQL Tabelle laden
        # // In der unteren Liste schauen ob die Eingabe schon existiert // Wenn ja, dann wird die Bestellung unten aktualisiert 
        # // Wenn Nein, dann wird ein neues Materialobjekt hinzuigefügt // Danach wird die zweite Liste neu geladen // Die bestellmenge wird gelöscht
        try:
            selectedID = self.MaterialTable.item(self.MaterialTable.currentRow(), 0).text()
            cursor = self.sqlpointer.suchenmitMaterialid(selectedID)
            selectedData = cursor.fetchone() 

            if (self.sqlpointer.suchenmitmaterialnummer(str(selectedData[1]),1)):        
                bestellmenge = int(self.bestellmenge.text()) + self.sqlpointer.suchenmitmaterialnummer(str(selectedData[1]),2)
                
                self.sqlpointer.MaterialbestellungUpdate(str(selectedData[1]),str(selectedData[2]),str(selectedData[3]),str(selectedData[4]),
                                                         str(selectedData[5]),str(selectedData[6]),bestellmenge)
                
            else:
                self.sqlpointer.MaterialbestellungHizufugen(str(selectedData[1]),str(selectedData[2]),str(selectedData[3]),str(selectedData[4]),str(selectedData[5]),str(selectedData[6]),int(self.bestellmenge.text()))
                      
            self.refreshListe2()
            self.bestellmenge.clear()
            
        except:
            print("bestellungHinzufugen fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")            
            
    def materiallisteFiltern(self):
        # Erfassen der Zeilenmenge und des aktuellen Text von der Combobox zum Filtern // Zeigt entweder alle Zeilen an
        # // Oder Zeigt anhand der auswahl der Combobox die jeweiligen Zeilen an
        try:
            rows = self.MaterialTable.rowCount()
            tempindex = self.CFilterBox.currentText()
            
            if tempindex == "alle":
                for j in range(rows):
                    self.MaterialTable.showRow(j)
            else:
                for i in range(rows):
                    self.MaterialTable.hideRow(i)
                    tasktableitem = str(self.MaterialTable.item(i, 4).text())
                    if tasktableitem == tempindex:
                        self.MaterialTable.showRow(i)
                    
        except:
            print("materiallisteFiltern fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred") 
            
    def bestellenAusloesen(self):
        # Läd die Bestellliste aus der SQL Tabelle // Führt die E-Mail Funktion aus um eine E-Mail aus der Liste zu generieren
        # // Aktualisiert die Menge in der SQL Tabelle welche bestellt sind // Löscht die untere Bestellliste
        # // Aktualisiert die Zweite Liste // Aktualisiert die erste Liste
        bestellliste = self.sqlpointer.getMaterialliste2()
        self.email(bestellliste)       
        
        for i in bestellliste:
            mengegesamt = int(i[7]) + int(self.sqlpointer.getMaterialVariable(i[1],8))
            self.sqlpointer.BestelltUpdate(str(i[1]),mengegesamt)
            
            self.sqlpointer.deleteFromBestellliste(i[0])
            
            self.resetMaterialtable2()
            self.materialtable2Refresh()
            
            self.resetMaterialtable()
            self.materialtableRefresh()
        
    def email(self,bestellliste):
        # Mit der Bestellliste einen HTML String erzeugen der dann mit Outlook versendet wird
        liste = bestellliste
        htmltest_mitte = ""
        
        for i in liste:
            htmltest_mitte2 = "<tr><td>" + str(i[1]) + "</td><td>" + str(i[2]) + "</td><td>" + str(i[3]) + "</td><td>" + str(i[4]) + "</td><td>" + str(i[5]) + "</td><td>" + str(self.sqlpointer.getMaterialVariable(i[1],7)) + "</td><td>               </td><td>" + str(i[7]) + "</td></tr>"
            htmltest_mitte = htmltest_mitte + htmltest_mitte2
        
        htmltest_anfang = """<html>
            <head>
            <title>Title of the document</title>
            </head>

            <body>
            <h3>Der 3D-Druck braucht Material.</h3>
            <hr>
            
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
                """

        htmltest_ende = """          
            </table>
            
            </body>

            </html>"""
        
        htmltestgesamt = htmltest_anfang + htmltest_mitte + htmltest_ende
        
        try:
            outlook = win32.Dispatch('Outlook.Application')
            outlookNS = outlook.GetNameSpace('MAPI')
            
            mail = outlook.CreateItem(0)
            mail.To = '3d-druck.ks@volkswagen.de;VWAG R: KS, MAWIS-WZB-PRESS'
            mail.Subject = 'Materialbestellung'
            mail.Body = 'Test'
            mail.HTMLBody = htmltestgesamt   
    
            mail.Send()
            
        except:
            print("emailtest fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")    
        
        
    def windowshow(self):
        # Läd das Stylesheet für den Dialog // Aktualisiert die Liste 1 // Aktualisiert die Liste 2 
        # // Zeigt den Dialog an
        #self.getStyleFromHilfs()
        self.refreshListe1()
        self.refreshListe2()
        self.show()

    def setMainmenuPointer(self,mainpointer):
        # Setzt einen Pointer um auf Funktionen vom Hauptmenü zugreifen zu können
        global MainmenuPointer
        MainmenuPointer = mainpointer
        
    def backToMainmenu(self):
        # Versteckt den Dialog // Öffnet den Dialog vom Hauptmenü
        self.hide()
        MainmenuPointer.windowshow()
    
    def beendenButtonPressed(self):
        # Beendet das Programm
        sys.exit(0)
        
    def hilfspointer(self,pointer):
        # Setzt den Pointer um auf Funktionen aus den Opionen zugreifen zu können
        self.hilfs = pointer
        
    def getStyleFromHilfs(self):
        # Läd aus den Optionen das Stylesheet für den Dialog
        self.setStyleSheet(self.hilfs.getStyle())
        
    def selectRowByID(self,desired_id):
        # Selektiert die Zeile in der Oberen Liste anhand der ID
        for row in range(self.MaterialTable.rowCount()):
            item = self.MaterialTable.item(row, 0)  # Assumes ID column is at index 0
            if item is not None and item.text() == desired_id:
                self.MaterialTable.selectRow(row)
                break
            
    def refreshListe1(self):
        # Liste 1 löschen // Liste 1 neu laden // Eingabemaske löschen
        self.resetMaterialtable()
        self.materialtableRefresh()
        self.resetMaterialUI()
        
    def refreshListe2(self):
        # Liste 2 löschen // Liste 2 neu laden
        self.resetMaterialtable2()
        self.materialtable2Refresh()
