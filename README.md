# School-Helper
Der “Musterlösung-Helfer” erstellt mit wenigen Klicks zwei Versionen eines Dokuments:
- Exemplar für Schülerinnen und Schüler
- Exemplar für Lehrkraft (Musterlösung)
Mit Hilfe von Makros für Microsoft-Word können unterschiedliche Ansichten ausgewählt werden und schnelle PDF-Exporte durchgeführt werden.

Die Erweiterung “Notenschlüssel-Tabelle” erstellt automatisch nach Eingabe der Gesamtpunktzahl eine Tabelle für den Notenschlüssel:
- Lineare Notenschlüssel (erste 4 gibt es bei 50%)
- 50%-Notenschlüssel (letzte 4 gibt es bei 50%)

# Voraussetzungen
- **Software**
    - Microsoft Word (getestet ab Version 2010)
- **Dokumentenaufbau**
    - Musterlösung muss in rot (RGB: 255,0,0) eingetragen werden
    - Roter Text wird in **Fließtext, Kopf-/Fußzeile, Textboxen und Formen** ersetzt. Text in **Tabellen** wird nur in wenigen Außnahmefällen *nicht* beachtet.


# Installation
- **Quellcode herunterladen**    
- **Makro hinzufügen**
    - Empfohlen: Global (für alle Dokumente): Entwicklertools ⇒ Visual Basic ⇒ Normal ⇒ Module ⇒ Rechts Klick ⇒ Datei importieren ⇒ SolutionHelper.bas auswählen
    - Lokal (für aktuelles Dokument): Entwicklertools ⇒ Visual Basic ⇒ Project (*\<Dokumentname\>*) ⇒ Module ⇒ Rechts Klick auf Module ⇒ Datei importieren ⇒ SolutionHelper.bas auswählen  
      ⚠️ *Word-Dokument muss anschließend als \*.docm (Dokument mit Makros) gespeichert werden.* ⚠️        
      
- **Makro zu Menüband hinzufügen**
    Wordoptionen ⇒ Menüband anpassen ⇒ Befehle auswählen: Makros ⇒ gewünschte Registerkarte und Gruppe auswählen.
    (Ggf. Beschriftung und Icon anpassen) z. B.  
    ![MSWord_SolutionHelper](https://github.com/mexterng/MSWord_Solution-Helper/assets/16732689/03f501ba-6120-41e6-a107-2549e2d8157e)

    

# Verwendung
- **ExportAll()**
    1. Öffnet “Speichern unter …”-Dialog für Version für Lehrkraft
    2. Öffnet “Speichern unter …”-Dialog für Version für Schülerinnen und Schüler
    
  ---
  
- **ExportLK()**
    Öffnet “Speichern unter …”-Dialog für Version für Lehrkraft
- **ExportSuS()**
    Öffnet “Speichern unter …”-Dialog für Version für Schülerinnen und Schüler
  
  ---
  
- **ChangeRedToWhite()**
    Ersetzt alle *roten* Zeichen in *weiße* Zeichen (Ansicht für Schülerinnen und Schüler)
- **ChangeWhiteToRed()**
    Ersetzt alle *weißen* Zeichen in *rote* Zeichen (Ansicht für Lehrkraft)

    ---
  
- **GenerateGradeDistributionLinear()**
    Erstellt Lineare Notenschlüssel (erste 4 gibt es bei 50%) 
- **GenerateGradeDistributionFiftyPercent()**
    50%-Notenschlüssel (letzte 4 gibt es bei 50%)
  

