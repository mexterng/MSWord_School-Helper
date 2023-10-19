# Solution-Helper
Der “Musterlösung-Helfer” erstellt mit wenigen Klicks zwei Versionen eines Dokuments:
- Exemplar für Schülerinnen und Schüler
- Exemplar für Lehrkraft (Musterlösung)

Mit Hilfe von Makros für Microsoft-Word können unterschiedliche Ansichten ausgewählt werden und schnelle PDF-Exporte durchgeführt werden.


# Voraussetzungen
- **Software**
    - Microsoft Word (getestet ab Version 2010)
- **Dokumentenaufbau**
    - Musterlösung muss in rot (RGB: 255,0,0) eingetragen werden
    - Roter Text wird nur in **Fließtext, Kopf-/Fußzeile, Textboxen und Formen** ersetzt


# Installation
- **Quellcode herunterladen**    
- **Makro hinzufügen**
    - Empfohlen: Global (für alle Dokumente): Entwicklertools ⇒ Visual Basic ⇒ Normal ⇒ Module ⇒ Rechts Klick ⇒ Datei importieren ⇒ DocumentHelper.bas auswählen
    - Lokal (für aktuelles Dokument): Entwicklertools ⇒ Visual Basic ⇒ Project (*\<Dokumentname\>*) ⇒ Module ⇒ Rechts Klick auf Module ⇒ Datei importieren ⇒ DocumentHelper.bas auswählen
       ⚠️ *Word-Dokument muss anschließend als \*.docm (Dokument mit Makros) gespeichert werden.* ⚠️        
      
- **Makro zu Menüband hinzufügen**
    Wordoptionen ⇒ Menüband anpassen ⇒ Befehle auswählen: Makros ⇒ gewünschte Registerkarte und Gruppe auswählen.
    (Ggf. Beschriftung und Icon anpassen)
    

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

