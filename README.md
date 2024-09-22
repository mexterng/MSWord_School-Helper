# School-Helper
*english version below*

Der “Musterlösung-Helfer” erstellt mit wenigen Klicks zwei Versionen eines Dokuments:
- Exemplar für Schülerinnen und Schüler
- Exemplar für Lehrkraft (Musterlösung)
Mit Hilfe von Makros für Microsoft-Word können unterschiedliche Ansichten ausgewählt werden und schnelle PDF-Exporte durchgeführt werden.

Die Erweiterung “Notenschlüssel-Tabelle” erstellt automatisch nach Eingabe der Gesamtpunktzahl eine Tabelle für den Notenschlüssel:
- Lineare Notenschlüssel (erste 4 gibt es bei 50%)
- 50%-Notenschlüssel (letzte 4 gibt es bei 50%)

# Voraussetzungen
*english version below*

- **Software**
    - Microsoft Word (getestet ab Version 2010)
- **Dokumentenaufbau**
    - Musterlösung muss in rot (RGB: 255,0,0) eingetragen werden
    - Roter Text wird in **Fließtext, Kopf-/Fußzeile, Textboxen und Formen** ersetzt. Text in **Tabellen** wird nur in wenigen Außnahmefällen *nicht* beachtet.


# Installation
*english version below*

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
*english version below*

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
  

---
---
# School Helper

The “Solution Helper” creates two versions of a document with just a few clicks:
- Version for Students
- Version for Teachers (Sample Solution)

With the help of macros for Microsoft Word, different views can be selected and quick PDF exports can be performed.

The extension “GradeDistribution” automatically creates a table for the grade key after entering the total score:
- Linear grade key (first 4 are available at 50%)
- 50% grade key (last 4 are available at 50%)

## Requirements
- **Software**
    - Microsoft Word (tested from version 2010)
- **Document Structure**
    - Sample solution must be entered in red (RGB: 255,0,0)
    - Red text will be replaced in **body text, headers/footers, text boxes, and shapes**. Text in **tables** will only be ignored in a few exceptional cases.

## Installation
- **Download Source Code**    
- **Add Macro**
    - Recommended: Globally (for all documents): Developer Tools ⇒ Visual Basic ⇒ Normal ⇒ Module ⇒ Right-click ⇒ Import File ⇒ Select SolutionHelper.bas
    - Locally (for the current document): Developer Tools ⇒ Visual Basic ⇒ Project (*\<DocumentName\>*) ⇒ Module ⇒ Right-click on Module ⇒ Import File ⇒ Select SolutionHelper.bas  
      ⚠️ *Word document must then be saved as \*.docm (Document with macros).* ⚠️        
      
- **Add Macro to Ribbon**
    Word Options ⇒ Customize Ribbon ⇒ Choose Commands: Macros ⇒ Select desired tab and group.
    (Optionally adjust label and icon) e.g.  
    ![MSWord_SolutionHelper](https://github.com/mexterng/MSWord_Solution-Helper/assets/16732689/03f501ba-6120-41e6-a107-2549e2d8157e)

## Usage
- **ExportAll()**
    1. Opens “Save As…” dialog for the teacher's version
    2. Opens “Save As…” dialog for the students' version
    
  ---
  
- **ExportLK()**
    Opens “Save As…” dialog for the teacher's version
- **ExportSuS()**
    Opens “Save As…” dialog for the students' version
  
  ---
  
- **ChangeRedToWhite()**
    Replaces all *red* characters with *white* characters (view for students)
- **ChangeWhiteToRed()**
    Replaces all *white* characters with *red* characters (view for teachers)

    ---
  
- **GenerateGradeDistributionLinear()**
    Creates linear grade keys (first 4 are available at 50%) 
- **GenerateGradeDistributionFiftyPercent()**
    50% grade key (last 4 are available at 50%)
