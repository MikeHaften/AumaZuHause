from __future__ import print_function
from __future__ import unicode_literals

import argparse
import sqlite3
import shutil
import time
import datetime
import sys,os
from calendar import monthrange

from sqlite3 import Error

DESCRIPTION = """
              Create a timestamped SQLite database backup, and
              clean backups older than a defined number of days
              """

NO_OF_DAYS = 7



def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

    

class sqlmanager:
    def __init__(self):     
        self.conn = None
        self.cursor = None
        self.dbname = resource_path(os.path.abspath(".") + "/auma.db")
        self.backupdir = resource_path(os.path.abspath(".") + "\Backup")
        
        # Backup erstellen
        self.clean_data(self.backupdir)
        self.sqlite3_backup(self.dbname,self.backupdir)

        # print("Sqlmanager erstellt")
        self.create_connection()
        self.tableerstellen()

    # -------------------------------------------------------------------------------
    # ----------- VERBINDUNG HERSTELLEN ---------------------------------------------
    # -------------------------------------------------------------------------------

    def create_connection(self):
        """ create a database connection to a SQLite database """
        self.conn = None
        self.cursor = None
        try:
            self.conn = sqlite3.connect(self.dbname)
            self.cursor = self.conn.cursor()
            #print("Verbindung zu Datenbank hergestellt!!!")
        except Error as e:
            print(e)
            
    def create_Connection2(self):
        """ create a database connection to a SQLite database """
        conn = None
        cursor = None
        try:
            conn = sqlite3.connect(self.dbname)
            cursor = conn.cursor()
            #print("Verbindung zu Datenbank hergestellt!!!")
        except Error as e:
            print(e)
        
        return conn,cursor
        
    def datenbanktrennen(self):
        self.conn.close()
        #print("Datenbank wurde getrennt")

    # -------------------------------------------------------------------------------
    # ----------- TABLES ERSTELLEN WENN NOCH NICHT PASSIERT -------------------------
    # -------------------------------------------------------------------------------

    def tableerstellen(self):
        conn, cursor = self.create_Connection2()
        # print("In Table erstellen")
        cursor.execute("""CREATE TABLE IF NOT EXISTS auftragsliste (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        EINGANGSDATUM TEXT,
        WUNSCHDATUM TEXT,
        NAME TEXT,
        VORNAME TEXT,
        ABTEILUNG TEXT,
        KOSTENSTELLE TEXT, 
        FERTIGUNGSBEREICH TEXT,
        TELEFONNUMMER TEXT,
        EMAIL TEXT,
        BAUTEILNAME TEXT,
        STUCKZAHL INTEGER,
        EILIG INTEGER,
        DRUCKDAUER REAL,
        GEWICHT INTEGER,
        BESCHREIBUNG TEXT,
        FARBE TEXT,
        VORLAGE TEXT,
        BAUTEILGROSE TEXT,
        MATERIAL TEXT,
        DATENAUFBEREITUNG REAL,
        NACHBEARBEITUNG REAL,
        FOLGETEIL INTEGER,
        PREIS REAL,
        PREISEXTERN REAL,
        EINSPARUNG REAL,
        FERTIGDATUM TEXT,
        STATUS INTEGER,
        MASCHINENSTUNDE REAL,
        ARBEITSSTUNDE REAL,
        AUFPREISFAKTOR REAL,
        FERTIG INTEGER,
        AUSGELIEFERT INTEGER,
        INFILL INTEGER,
        FA1 TEXT,
        FA2 TEXT,
        FA3 TEXT,
        FA4 TEXT,
        FA5 TEXT
        )""")
        # print("Auftragsliste Table created successfully")

        cursor.execute("""CREATE TABLE IF NOT EXISTS personenliste (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        VORNAME TEXT,
        NACHNAME TEXT,
        KOSTENSTELLE TEXT,
        ABTEILUNG TEXT,
        TELEFON TEXT,
        EMAIL TEXT
        )""")
        
        
        cursor.execute("""CREATE TABLE IF NOT EXISTS materialbestandliste (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        MATERIALNUMMER TEXT,
        DURCHMESSER TEXT,
        GEWICHT TEXT,
        ART TEXT,
        FARBE TEXT,
        HERSTELLER TEXT,
        BESTAND INTEGER,
        BESTELLT INTEGER
        )""")
        
        cursor.execute("""CREATE TABLE IF NOT EXISTS materialbestellliste (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        MATERIALNUMMER TEXT,
        DURCHMESSER TEXT,
        GEWICHT TEXT,
        ART TEXT,
        FARBE TEXT,
        HERSTELLER TEXT,
        MENGE TEXT
        )""")
        
        cursor.execute("""CREATE TABLE IF NOT EXISTS materialliste (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ART TEXT,
        DURCHMESSER TEXT,
        GEWICHT TEXT,
        PREIS REAL,
        VAR1 TEXT,
        VAR2 TEXT,
        VAR3 TEXT
        )""")
        
        
        cursor.execute("""CREATE TABLE IF NOT EXISTS programmvariablen (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        Variable1 TEXT,
        Variable2 TEXT,
        Variable3 TEXT,
        Variable4 TEXT,
        Variable5 TEXT,
        Variable6 TEXT,
        Variable7 TEXT,
        Variable8 TEXT,
        Variable9 TEXT,
        Variable10 TEXT,
        Variable11 TEXT,
        Variable12 TEXT,
        Variable13 TEXT,
        Variable14 TEXT,
        Variable15 TEXT
        )""")
        
        self.datenbanktrennen()
    # -------------------------------------------------------------------------------
    # ----------- AUFTRAGSLISTE FUNKTIONEN ------------------------------------------
    # -------------------------------------------------------------------------------
  
    def auftragErstellen(self,eingangsdatum,
                         wunschdatum,name,vorname,
                         kostenstelle,fertigungsbereich,abteilung,
                         telefon,email,bauteil,
                         eilig,beschreibung,
                         farbe,vorlage,material,bauteilgrose = "unbekannt",
                         druckdauer = 0,gewicht = 0,stuckzahl = 0,
                         datenaufbereitung = 0,nacharbeit = 0,
                         folgeauftrag = 0,preis = 0,preisextern = 0,
                         einsparung = 0,auslieferungsdatum = "", status= 0,
                         maschinenstunde = 0,arbeitsstunde=0,aufpreisfaktor=0,fertig=0,ausgeliefert=0,infill=20,fa1="",fa2="",fa3="",fa4="",fa5=""):
        
        maschinenstunde = self.variableAusgeben(4)[1]
        arbeitsstunde = self.variableAusgeben(4)[2]
        aufpreisfaktor = self.variableAusgeben(4)[3]
        
        try:
            conn, cursor = self.create_Connection2()
            cursor.execute("""INSERT INTO 'auftragsliste'
            (eingangsdatum,
            wunschdatum,
            name,
            vorname,
            abteilung,
            kostenstelle,
            fertigungsbereich,
            telefonnummer,
            email,
            bauteilname,
            stuckzahl,
            eilig,
            druckdauer,
            gewicht,
            beschreibung,
            farbe,
            vorlage,
            bauteilgrose,
            material,
            datenaufbereitung,
            nachbearbeitung,
            folgeteil,
            preis,
            preisextern,
            einsparung,
            fertigdatum,
            status,
            maschinenstunde,
            arbeitsstunde,
            aufpreisfaktor,
            fertig,
            ausgeliefert,
            infill,
            fa1,
            fa2,
            fa3,
            fa4,
            fa5
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
                eingangsdatum,
                wunschdatum, name, vorname, abteilung,
                kostenstelle, fertigungsbereich,
                telefon, email,bauteil,stuckzahl,
                eilig, druckdauer, gewicht, beschreibung,
                farbe, vorlage, bauteilgrose, material,
                datenaufbereitung, nacharbeit,
                folgeauftrag, preis, preisextern,
                einsparung, auslieferungsdatum,status,maschinenstunde,arbeitsstunde,aufpreisfaktor,
                fertig,ausgeliefert,infill,fa1,fa2,fa3,fa4,fa5))
        except Error as e:
            print(e)

        conn.commit() 
        self.datenbanktrennen()
        
    def auftragErstellenAusBackup(self,identifikation,eingangsdatum="22-02-2023",
                         wunschdatum="22-02-2023",name="testermann",vorname="max",
                         kostenstelle="4381",fertigungsbereich="FB6",abteilung="COK",
                         telefon= "0123455667",email="max.testermann@volkswagen.de",bauteil="Testteil",
                         eilig= 0,beschreibung= "Test",
                         farbe= "Rot",vorlage = "Skizze",material= "PLA" ,bauteilgrose = "Unbekannt",
                         druckdauer = 0,gewicht = 0,stuckzahl = 0,
                         datenaufbereitung = 0,nacharbeit = 0,
                         folgeauftrag = 0,preis = 0,preisextern = 0,
                         einsparung = 0,auslieferungsdatum = "Datum", status= 0):
        # IST NICHT MEHR AKTUELL
        try:
            conn, cursor = self.create_Connection2()
            cursor.execute("""INSERT INTO 'auftragsliste'
            (id,
            eingangsdatum,
            wunschdatum,
            name,
            vorname,
            abteilung,
            kostenstelle,
            fertigungsbereich,
            telefonnummer,
            email,
            bauteilname,
            stuckzahl,
            eilig,
            druckdauer,
            gewicht,
            beschreibung,
            farbe,
            vorlage,
            bauteilgrose,
            material,
            datenaufbereitung,
            nachbearbeitung,
            folgeteil,
            preis,
            preisextern,
            einsparung,
            fertigdatum,
            status
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
                identifikation,eingangsdatum,
                wunschdatum, name, vorname, abteilung,
                kostenstelle, fertigungsbereich,
                telefon, email,bauteil,stuckzahl,
                eilig, druckdauer, gewicht, beschreibung,
                farbe, vorlage, bauteilgrose, material,
                datenaufbereitung, nacharbeit,
                folgeauftrag, preis, preisextern,
                einsparung, auslieferungsdatum,status))
        except Error as e:
            print(e)

        conn.commit()   
        self.datenbanktrennen()
        
    def getAuftragsliste(self,variante):
        conn, cursor = self.create_Connection2()
        result = cursor.execute("SELECT * FROM auftragsliste ORDER BY id ASC")
        conn.commit()
        self.datenbanktrennen()
        if variante==1:
            ausgabe = result
        elif variante==2:
            ausgabe = result.fetchall()
        
        return ausgabe
        
    def suchenmitid(self, identifikation):
        conn, cursor = self.create_Connection2()
        sqlstring = "SELECT * FROM auftragsliste WHERE id= (?)"
        result = cursor.execute(sqlstring, (identifikation,))  
        
        conn.commit()
        self.datenbanktrennen()
        return result   
        
    def maxidfinden(self):
        try:
            conn, cursor = self.create_Connection2()
            sqlstring = "SELECT max(ID) FROM 'auftragsliste'"
            result = cursor.execute(sqlstring,)  
            conn.commit()
            ausgabe = result.fetchone()[0]
            self.datenbanktrennen()
        except:
            print("Max ID finden fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            
        return ausgabe
        
    def sucheFehlendeDaten(self):
        conn, cursor = self.create_Connection2()
        sqlstring = "SELECT * FROM auftragsliste WHERE druckdauer=0.0 AND gewicht=0 ORDER BY id ASC"
        
        result = cursor.execute(sqlstring)
        conn.commit()
        self.datenbanktrennen()
        ausgabe = result.fetchall()
        liste = []
        for i in ausgabe:
            #print(i[0])
            liste.append(i[0])
        
        return liste
        
    def fehlendeFAnummer(self):
        conn, cursor = self.create_Connection2()
        sqlstring = "SELECT * FROM auftragsliste WHERE fa1 IS NULL OR fa1='' ORDER BY id ASC"
        
        result = cursor.execute(sqlstring)
        conn.commit()
        self.datenbanktrennen()
        ausgabe = result.fetchall()
        liste = []
        for i in ausgabe:
            #print(i[0])
            liste.append(i[0])
        
        return liste
        
            
    def auftragSpeichern(self,auftragsnummer,eingangsdatum,wunschdatum,name,vorname,
                         abteilung,kostenstelle,fertigungsbereich,telefonnummer,email,
                         bauteilname,stuckzahl,eilig,druckdauer,gewicht,beschreibung,
                         farbe,vorlage,bauteilgrose,material,datenaufbereitung,
                         nachbearbeitung,folgeteil,preis,preisextern,einsparung,
                         fertigdatum,status,maschinenstunde,arbeitsstunde,aufpreisfaktor,
                         fertig,ausgeliefert,infill,fa1,fa2,fa3,fa4,fa5):
        try:
            conn, cursor = self.create_Connection2()
            cursor.execute("""UPDATE auftragsliste SET eingangsdatum=?,
            wunschdatum=?,
            name=?,
            vorname=?,
            abteilung=?,
            kostenstelle=?,
            fertigungsbereich=?,
            telefonnummer=?,
            email=?,
            bauteilname=?,
            stuckzahl=?,
            eilig=?,
            druckdauer=?,
            gewicht=?,
            beschreibung=?,
            farbe=?,
            vorlage=?,
            bauteilgrose=?,
            material=?,
            datenaufbereitung=?,
            nachbearbeitung=?,
            folgeteil=?,
            preis=?,
            preisextern=?,
            einsparung=?,
            fertigdatum=?,
            status=?,
            maschinenstunde=?,
            arbeitsstunde=?,
            aufpreisfaktor=?,
            fertig=?,
            ausgeliefert=?,
            infill=?,
            fa1=?,
            fa2=?,
            fa3=?,
            fa4=?,
            fa5=? WHERE id = ?""",
                                (eingangsdatum, wunschdatum, name, vorname, abteilung, kostenstelle, fertigungsbereich, telefonnummer,
                                 email, bauteilname, stuckzahl, eilig, druckdauer, gewicht,
                                 beschreibung, farbe, vorlage, bauteilgrose, material, datenaufbereitung, nachbearbeitung,
                                 folgeteil, preis, preisextern, einsparung, fertigdatum, status,maschinenstunde,arbeitsstunde,aufpreisfaktor,fertig,ausgeliefert,infill,fa1,fa2,fa3,fa4,fa5,
                                 auftragsnummer))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
        
    def faDatenUpdate(self,ident,fa1,fa2,fa3,fa4,fa5):
        print("fa daten updaten")
        try:
            conn, cursor = self.create_Connection2()
            cursor.execute("""UPDATE auftragsliste SET fa1=?,
            fa2=?,
            fa3=?,
            fa4=?,
            fa5=? WHERE id = ?""",
                                (fa1,fa2,fa3,fa4,fa5,
                                 ident))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
        
    def fertigDatumSetzen(self,ident,datum):
        query = "UPDATE auftragsliste SET fertigdatum = ? WHERE id = ?"
        try:
            conn, cursor = self.create_Connection2()
            # SQL-Abfrage ausführen und Parameter binden
            cursor.execute(query, (datum, ident))
            conn.commit()
            self.datenbanktrennen()
            if cursor.rowcount > 0:
                #print("Eigenschaft erfolgreich aktualisiert.")
                pass
            else:
                print("Kein Eintrag mit der angegebenen ID gefunden.")
        except sqlite3.Error as e:
            print("Fehler beim Aktualisieren der Eigenschaft:", e) 
                  
    def deletefromAuftragsliste(self, identifikation):
        try:
            conn, cursor = self.create_Connection2()
            cursor.execute("DELETE FROM auftragsliste WHERE id=?", (identifikation,))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
        
    ###########################################################
    ############ Personenliste Aufgaben #######################
    ###########################################################
    
    def PersonHizufugen(self,vorname,name,kostenstelle,abteilung,telefon,email):  
        try:
            conn, cursor = self.create_Connection2()
            cursor.execute("""INSERT INTO 'personenliste'
            (vorname,
            nachname,
            kostenstelle,
            abteilung,
            telefon,
            email
            ) VALUES (?,?,?,?,?,?)""", (
                vorname,name,
                kostenstelle, abteilung, telefon, email))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
        
    def suchennachName(self, name, vorname):
        try:
            conn, cursor = self.create_Connection2()
            sqlstring = "SELECT * FROM personenliste WHERE nachname= (?) AND vorname= (?)"
            result = cursor.execute(sqlstring, (name, vorname,))           
            conn.commit()            
            ausgabe = result.fetchone()
            self.datenbanktrennen()
            
            if ausgabe is not None:
                ausgabe2 = False
            else:
                ausgabe2 = True
            
        except:
            print("suchennachName fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            
        return ausgabe2   
    
    def personStartsWith(self,name):
        conn, cursor = self.create_Connection2()
        try:
            sqlstring = "SELECT * FROM personenliste WHERE nachname LIKE '" + name + "%'"
            result = cursor.execute(sqlstring)
            conn.commit()
            ausgabe = result.fetchall()
            self.datenbanktrennen()
        except:
            print("suchennachName fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            
        return ausgabe
        
    def personenlisteRausgeben(self,variante):
        result = None
        conn, cursor = self.create_Connection2()
        try:
            result = cursor.execute("SELECT * FROM personenliste ORDER BY nachname ASC")
            conn.commit()
            if variante==1:
                ausgabe = result.fetchall()
            elif variante==2:
                ausgabe = result
            
            self.datenbanktrennen()
        except Error as e:
            print(e)

        return ausgabe
        
    def perssuchenmitid(self, identifikation):
        conn, cursor = self.create_Connection2()
        sqlstring = "SELECT * FROM personenliste WHERE id= (?)"
        result = cursor.execute(sqlstring, (identifikation,))  
        
        conn.commit()
        self.datenbanktrennen()
        return result
        
    def personUpdaten(self,identifikation,vorname,name,kostenstelle,abteilung,
                         telefon,email):
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("""UPDATE personenliste SET vorname=?,
            nachname=?,
            kostenstelle=?,
            abteilung=?,
            telefon=?,
            email=? WHERE id = ?""",
                                (vorname, name, kostenstelle, abteilung, telefon, email, identifikation))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
            
    ###########################################################
    ############ Materialliste Aufgaben #######################
    ###########################################################
    
    def MaterialHizufugen(self,materialnummer,durchmesser,gewicht,art,farbe,hersteller,bestand=0,bestellt=0):
        
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("""INSERT INTO 'materialbestandliste'
            (materialnummer,
            durchmesser,
            gewicht,
            art,
            farbe,
            hersteller,
            bestand,
            bestellt
            ) VALUES (?,?,?,?,?,?,?,?)""", (
                materialnummer,durchmesser,
                gewicht, art, farbe, hersteller,bestand,bestellt))
        except Error as e:
            print(e)

        conn.commit() 
        self.datenbanktrennen()
        
              
    def getMaterialliste(self):
        conn, cursor = self.create_Connection2()
        result = cursor.execute("SELECT * FROM materialbestandliste ORDER BY id ASC")
        conn.commit()
        self.datenbanktrennen()
        return result
        
    def suchenmitMaterialid(self, identifikation):
        conn, cursor = self.create_Connection2()
        sqlstring = "SELECT * FROM materialbestandliste WHERE id= (?)"
        result = cursor.execute(sqlstring, (identifikation,))  
        
        conn.commit()
        self.datenbanktrennen()
        return result
        
    def getMaterialVariable(self,materialnummer,variante):
        conn, cursor = self.create_Connection2()
        try:
            result = cursor.execute("SELECT * FROM materialbestandliste WHERE materialnummer=?", (materialnummer,))
            if variante == 7:
                ausgabe = result.fetchone()[7]
            if variante == 8:
                ausgabe = result.fetchone()[8] 
            conn.commit()
            self.datenbanktrennen()
        except Error as e:
            print(e)
            
        return ausgabe 
        
    def MaterialUpdate(self,materialnummer,durchmesser,gewicht,art,farbe,
                         hersteller,identifikation,bestand=0,bestellt=0):
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("""UPDATE materialbestandliste SET materialnummer=?,
            durchmesser=?,
            gewicht=?,
            art=?,
            farbe=?,
            hersteller=?,
            bestand=?,
            bestellt=? WHERE id = ?""",
                                (materialnummer, durchmesser, gewicht, art, farbe, hersteller,bestand,bestellt,
                                 identifikation))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
        
    def BestelltUpdate(self,materialnummer,bestellt):
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("""UPDATE materialbestandliste SET 
            bestellt=? WHERE materialnummer = ?""",
                                (bestellt,
                                 materialnummer))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()          
            
    def deletefromMaterialliste(self, identifikation):
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("DELETE FROM materialbestandliste WHERE id=?", (identifikation,))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
        
    ###########################################################
    ############ Materialbestellliste Aufgaben ################
    ###########################################################
    
    def MaterialbestellungHizufugen(self,materialnummer,durchmesser,gewicht,art,farbe,hersteller,menge):  
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("""INSERT INTO 'materialbestellliste'
            (materialnummer,
            durchmesser,
            gewicht,
            art,
            farbe,
            hersteller,
            menge
            ) VALUES (?,?,?,?,?,?,?)""", (
                materialnummer,durchmesser,
                gewicht, art, farbe, hersteller,menge))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
        
    def getMaterialliste2(self):
        conn, cursor = self.create_Connection2()
        result = cursor.execute("SELECT * FROM materialbestellliste ORDER BY id ASC")
        conn.commit()
        ausgabe = result.fetchall()
        self.datenbanktrennen()
        
        return ausgabe   

    def suchenmitmaterialnummer(self, material, variante):
        # Variante 1 = Suchen mit Materialnummer und True oder False ausgabe bei Fund
        # Variante 2 = Suchen mit Materialnummer und Ausgabe der Menge
        # Variante 3 = Suchen mit Materialnummer und Ausgabe der ID
        conn, cursor = self.create_Connection2()
        if (variante == 1):
            try:
                result = cursor.execute("SELECT * FROM materialbestellliste WHERE materialnummer=?", (material,))
                ausgabe = result.fetchone()[1]
                conn.commit()
                self.datenbanktrennen()
                
                if ausgabe is not None:
                    ausgabe2 = True
                else:
                    ausgabe2 = False
                
            except:
                ausgabe2 = False
    
            return ausgabe2 
            
        if (variante == 2):
            try:
                result = cursor.execute("SELECT * FROM materialbestellliste WHERE materialnummer=?", (material,))
                ausgabe = result.fetchone()[8]
                conn.commit()
                self.datenbanktrennen()
            except:
                ausgabe = 0
    
            return ausgabe 
            
        if (variante == 3):
            try:
                result = cursor.execute("SELECT * FROM materialbestellliste WHERE materialnummer=?", (material,))
                ausgabe = result.fetchone()[0]
                conn.commit()
                self.datenbanktrennen()
            except:
                ausgabe = 0
    
            return ausgabe    
            
    def deleteFromBestellliste(self,identifikation):
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("DELETE FROM materialbestellliste WHERE id=?", (identifikation,))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()            
 
    def MaterialbestellungUpdate(self,materialnummer,durchmesser,gewicht,art,farbe,
                         hersteller,identifikation,menge):
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("""UPDATE materialbestellliste SET materialnummer=?,
            durchmesser=?,
            gewicht=?,
            art=?,
            farbe=?,
            hersteller=?,
            menge=? WHERE id = ?""",
                                (materialnummer, durchmesser, gewicht, art, farbe, hersteller,menge,
                                 identifikation))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()           
            
    ###########################################################
    ###########################################################
    ############## Programmvariablen verwalten ################
    ###########################################################
    ###########################################################
    
    def VariableVeraendern(self,var1,var2,var3,var4,var5,var6,var7,var8,var9,var10,var11,var12,var13,var14,var15,ident):
        #print("Verändern")
        conn, cursor = self.create_Connection2()
        try:
            query = "INSERT OR REPLACE INTO programmvariablen (id, Variable1, Variable2, Variable3, Variable4, Variable5, Variable6, Variable7, Variable8, Variable9, Variable10, Variable11, Variable12, Variable13, Variable14, Variable15) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            cursor.execute(query, (ident, var1, var2, var3, var4, var5, var6, var7, var8, var9, var10, var11, var12, var13, var14, var15))        
            conn.commit()
            self.datenbanktrennen()
        except Error as e:
            print(e)

    def VariableHizufugen(self,var1,var2,var3,var4,var5,var6,var7,var8,var9,var10,var11,var12,var13,var14,var15):
        
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("""INSERT INTO 'programmvariablen'
            (variable1,
            variable2,
            variable3,
            variable4,
            variable5,
            variable6,
            variable7,
            variable8,
            variable9,
            variable10,
            variable11,
            variable12,
            variable13,
            variable14,
            variable15
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
                var1,var2,var3, var4, var5, var6,var7,var8,var9,var10, var11, var12, var13,var14,var15))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
        
        
    def variableAusgeben(self, identifikation):
        conn, cursor = self.create_Connection2()
        sqlstring = "SELECT * FROM programmvariablen WHERE id= (?)"
        result = cursor.execute(sqlstring, (identifikation,))
        ausgabe = result.fetchone()
        
        conn.commit()
        self.datenbanktrennen()
        return ausgabe        
        
        
    def variableUpdaten(self,var1,var2,var3,var4,var5,
                         var6,var7,var8,var9,var10,
                         var11,var12,var13,var14,var15,identifikation):
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("""UPDATE programmvariablen SET variable1=?,
            variable2=?,
            variable3=?,
            variable4=?,
            variable5=?,
            variable6=?,
            variable7=?,
            variable8=?,
            variable9=?,
            variable10=?,
            variable11=?,
            variable12=?,
            variable13=?,
            variable14=?,
            variable15=? WHERE id = ?""",
                                (var1, var2, var3, var4, var5, var6, var7, var8,
                                 var9, var10, var11, var12, var13, var14,
                                 var15,identifikation))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
        
    def variableloeschen(self,identifikation):
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("DELETE FROM programmvariablen WHERE id=?", (identifikation,))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
            
    ###########################################################
    ############ MATERIALLISTE BERECHNUNGSPARAMETER ###########
    ###########################################################
    
    def getML(self,variante):
        conn, cursor = self.create_Connection2()
        result = cursor.execute("SELECT * FROM materialliste ORDER BY id ASC")
        conn.commit()
        if variante==1:
            ausgabe = result
        elif variante==2:
            ausgabe = result.fetchall()
        elif variante==3:
            liste = []
            tempListe = result.fetchall()
            for i in tempListe:
                liste.append(i[1])
            ausgabe = liste
        
        self.datenbanktrennen()
        return ausgabe
    
    def matSuchenMitId(self, identifikation):
        conn, cursor = self.create_Connection2()
        sqlstring = "SELECT * FROM materialliste WHERE id= (?)"
        result = cursor.execute(sqlstring, (identifikation,))  
        
        conn.commit()
        self.datenbanktrennen()
        return result
        
    def matAdd(self,art,d,gew,preis,var1,var2,var3):
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("""INSERT INTO 'materialliste'
            (art,
            durchmesser,
            gewicht,
            preis,
            var1,
            var2,
            var3
            ) VALUES (?,?,?,?,?,?,?)""", (
                art,d,
                gew, preis, var1, var2, var3))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
        
    def matUpdate(self,art,d,gew,preis,var1,
                         var2,var3,identifikation):
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("""UPDATE materialliste SET art=?,
            durchmesser=?,
            gewicht=?,
            preis=?,
            var1=?,
            var2=?,
            var3=? WHERE id = ?""",
                                (art, d, gew, preis, var1, var2, var3,
                                 identifikation))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
        
    def getMatData(self, mat):
        conn, cursor = self.create_Connection2()
        sqlstring = "SELECT * FROM materialliste WHERE art= (?)"
        result = cursor.execute(sqlstring, (mat,))
        ausgabe1 = result.fetchone()
        ausgabePreis = ausgabe1[4]
        ausgabeGewicht = ausgabe1[3]
        
        conn.commit()
        self.datenbanktrennen()
        return ausgabePreis, ausgabeGewicht
        
    def matDel(self,identifikation):
        conn, cursor = self.create_Connection2()
        try:
            cursor.execute("DELETE FROM materialliste WHERE id=?", (identifikation,))
        except Error as e:
            print(e)

        conn.commit()
        self.datenbanktrennen()
        
    ###########################################################
    ###########################################################
    ############## Backup erstellen ###########################
    ###########################################################
    ###########################################################

    def sqlite3_backup(self, dbfile, backupdir):
        """Create timestamped database copy"""

        if not os.path.isdir(backupdir):
            raise Exception("Backup directory does not exist: {}".format(backupdir))

        backup_file = os.path.join(backupdir, os.path.basename(dbfile) +
                                   time.strftime("-%Y-%m-%d-%H-%M"))

        connection = sqlite3.connect(dbfile)
        cursor = connection.cursor()

        # Lock database before making a backup
        cursor.execute('begin immediate')
        # Make new backup file
        shutil.copyfile(dbfile, backup_file)
        print("\nCreating {}...".format(backup_file))
        # Unlock database
        connection.rollback()

    def clean_data(self, backup_dir):
        """Delete files older than NO_OF_DAYS days"""

        print("\n------------------------------")
        print("Cleaning up old backups")

        for filename in os.listdir(backup_dir):
            backup_file = os.path.join(backup_dir, filename)
            if os.stat(backup_file).st_ctime < (time.time() - NO_OF_DAYS * 86400):
                if os.path.isfile(backup_file):
                    os.remove(backup_file)
                    print("Deleting {}...".format(backup_file))

    def get_arguments(self):
        """Parse the commandline arguments from the user"""

        parser = argparse.ArgumentParser(description=DESCRIPTION)
        parser.add_argument('db_file',
                            help='the database file that needs backed up')
        parser.add_argument('backup_dir',
                            help='the directory where the backup'
                                 'file should be saved')
        return parser.parse_args()
        
    def getmengevonstatusdatum(self, status, indikator):
        # print("getmengevonstatusdatum")

        ausgabe = 0

        if indikator == 1:
            ausgabe = "monat"
            anfang = datetime.datetime.today().strftime("%Y") + "-" + datetime.datetime.today().strftime("%m") + "-01"
            ende = datetime.datetime.today().strftime("%Y") + "-" + datetime.datetime.today().strftime(
                "%m") + "-" + str(monthrange(int(datetime.datetime.today().strftime("%Y")),
                                             int(datetime.datetime.today().strftime("%m")))[1])

            ausgabe = self.getmengevonstatusdatum2(anfang, ende, status)

        elif indikator == 2:
            ausgabe = "jahr"
            anfang = datetime.datetime.today().strftime("%Y") + "-01-01"
            ende = datetime.datetime.today().strftime("%Y") + "-12-31"

            ausgabe = self.getmengevonstatusdatum2(anfang, ende, status)

        elif indikator == 3:
            ausgabe = "letztes jahr"
            lastjahrint = int(datetime.datetime.today().strftime("%Y")) - 1
            anfang = str(lastjahrint) + "-01-01"
            ende = str(lastjahrint) + "-12-31"

            ausgabe = self.getmengevonstatusdatum2(anfang, ende, status)

        return ausgabe
        
    def getstueckzahl(self, indikator):
        tempJob = self.suchenmitid(indikator)
        ausgabe = tempJob.fetchone()[2]
        return ausgabe
        
    # -------------------------------------------------------------------------------
    # -----------  Datenordner für Auftrag erstellen   ------------------------------
    # -------------------------------------------------------------------------------
    
    
    def createFolder(self, name, vorname, auftragsnummer):
        try:
            
            ordnerPath = self.variableAusgeben(3)[1] + "/"+ str(name) + " " + str (vorname) + "/" + str(auftragsnummer)
        
            if not os.path.exists(ordnerPath):
                os.makedirs(ordnerPath)
        except:
            print("Ordner erstellen fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
 
            
    def getFolderName(self, name, vorname, auftragsnummer):
        try:     
            ordnerPath = self.variableAusgeben(3)[1] + "/"+ str(name) + " " + str (vorname) + "/" + str(auftragsnummer)

            if not os.path.exists(ordnerPath):
                os.makedirs(ordnerPath)
                print("Odner nicht vorhanden")
        except:
            print("Ordnername zurückgeben fehlgeschlagen!!!")
            print("Oops!", sys.exc_info()[0], "occurred")
            print("Oops!", sys.exc_info()[1], "occurred")
            
        return ordnerPath
        
    def datenschieben(self, von_ordner,nachname,vorname,auftragsnummer):

        destination = self.getFolderName(nachname,vorname,auftragsnummer)
        
        basepath = os.path.abspath(".")
        tempdestination = destination.replace("/","\\")
        
        liste = von_ordner.getListe()

        for i in liste:
            try:    
                source = i
                #print("Von " + source + " nach " + (basepath + "\\" + tempdestination + "\\" + os.path.basename(source)) + " kopieren")
                shutil.copyfile(source,(basepath + "\\" + tempdestination + "\\" + os.path.basename(source)))
                #print(i)
                von_ordner.clearTable()
                #print("File copied successfully.")
                
                # If source and destination are same
            except shutil.SameFileError:
                print("Source and destination represents the same file.")
 
                # If destination is a directory.
            except IsADirectoryError:
                print("Destination is a directory.")
 
                # If there is any permission issue
            except PermissionError:
                print("Permission denied.")
                print("Oops!", sys.exc_info()[0], "occurred")
                print("Oops!", sys.exc_info()[1], "occurred")
 
                # For other errors
            except:
                print("Error occurred while copying file.")
                raise
                
            

    