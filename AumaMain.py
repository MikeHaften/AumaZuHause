import Auma
import nTask
import Taskmanager
import infoChef
import optionen
import Telefonbuch
import Materialmanagement
import sqlManager

from PyQt5 import QtWidgets
import sys



app = QtWidgets.QApplication(sys.argv)


sqlManagerObj = sqlManager.sqlmanager()

newTaskObj = nTask.Ui(sqlManagerObj)
TaskmanagerObj = Taskmanager.Ui(sqlManagerObj,app)
InfoChefObj = infoChef.Ui()
OptionenObj = optionen.Ui(sqlManagerObj)
TelefonbuchObj = Telefonbuch.Ui()
MaterialmObj = Materialmanagement.Ui(sqlManagerObj)

window = Auma.Ui(newTaskObj,TaskmanagerObj,InfoChefObj,OptionenObj,TelefonbuchObj,MaterialmObj)

newTaskObj.setMainmenuPointer(window)
TaskmanagerObj.setMainmenuPointer(window)
InfoChefObj.setMainmenuPointer(window)
OptionenObj.setMainmenuPointer(window)
TelefonbuchObj.setMainmenuPointer(window)
MaterialmObj.setMainmenuPointer(window)

MaterialmObj.hilfspointer(OptionenObj)
newTaskObj.hilfspointer(OptionenObj)
TaskmanagerObj.hilfspointer(OptionenObj)

window.windowshow()


app.exec_()

"""
----------To-Do Liste ----------

--- Bilderskalierung

--- Pfad der Bilder und Aufträge über das Hilfsmenü ändern

--- Button um FA Daten an AV zu schicken

--- Bei gleichzeitiger bedienung von zwei oder mehreren pc's prüfen ob es möglich ist dem anderen eine aktualisierungsnachricht zu geben

--- Anzeige so verbessern das gleich ersichtlich wenn das Bild oder die Druckdauer/Materialmenge fehlt

--- Weiteren Auftrag eingeben Programmieren

--- Fehlermeldungen als Messagebox im Programm ausgeben

--- Statistiken ergänzen


"""