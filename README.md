# shiftplaner
Shift plan control tool

#### 1 Anaconda installieren
- Passenden Installer runterladen https://www.anaconda.com/products/individual
- Installationsverzeichnis auswählen und Pfad zur installierten python.exe aufsschreiben (z.B. **C:\\Anaconda3**)
- Anleitung zur Installation auf Linux https://docs.continuum.io/anaconda/install/linux/
- Anleitung zur Installation auf Windows https://docs.continuum.io/anaconda/install/windows/

---
#### 2 Tesseract installieren
##### Linux
> sudo apt-get install tesseract-ocr -y

##### Windows
- Installer runterladen	https://github.com/UB-Mannheim/tesseract/wiki
- Installationsverzeichnis auswählen und Pfad zur installierten *tesseract.exe* aufsschreiben (z.B. **C:\\Tesseract-OCR**)

WARNUNG: Tesseract sollte entweder in das Verzeichnis, das bei der Installation vorgeschlagen wird, oder in ein neues Verzeichnis installiert werden. Das Deinstallationsprogramm entfernt das gesamte Installationsverzeichnis. Wenn Tesseract in ein bestehendes Verzeichnis installiert wurde, wird dieses Verzeichnis mit all seinen Unterverzeichnissen und Dateien entfernt.

---
#### 3 conda und tesseract zu Umgebungsvariablen hinzufügen
##### Windows
- Windows-Taste drücken, "Umgebung" eingeben 
- "Systemumgebungsvariablen bearbeiten" anklicken (oder Enter drücken)
- Im neuen Fenster "Systemeigenschaften" auf den Reiter "Erweitert" gehen
- Button "Umgebungsvariablen" klicken
- Im neuen Fenster "Umgebungsvariablen" im oberen Abschnitt "Benutzervariablen für ..." nach der Variablen "Path" (oder "PATH") suchen
- bei vorhandener "Path" Variable auf "Bearbeiten..." Button klicken
	1. Doppelklick auf leere Zeile -> Pfad aus Schritt 1 einfügen und mit Enter bestätigen (z.b. **C:\\Anaconda3**)
	3. Doppelklick auf leere Zeile -> Pfad aus 1 + "\\condabin" einfügen und mit Enter bestätigen (z.b. **C:\\Anaconda3\\condabin**)
	4. Doppelklick auf leere Zeile -> Pfad aus Schritt 2 einfügen und mit Enter bestätigen (z.b. **C:\\Tesseract-OCR**)
-  wenn keine "Path" Variable vorhanden ist, auf "Neu..." Button klicken:
	- Name der Variablen: "Path"
	- Wert der Variablen: Pfade aus Schritt 1 und 2, getrennt durch ein Semikolon (z.B. **"C:\\Anaconda3;C:\\Anaconda3\\condabin;C:\\Tesseract-OCR"**)

---
#### 4.0 Einrichtung einer neuen Python Umgebung (optional)
> conda create -n Takeaway python=3.8.5

BEACHTEN: Um in der Takeaway Python Umgebung zu arbeiten, muss in die Umgebung **vor** der Nutzung aktiviert werden:
- Terminal/Console/Command Prompt öffnen
	###### Linux
	- gleichzeitiges drücken von "STRG" + "ALT" + "T"
	###### Windows
	- Windows-Taste drücken, "cmd" eingeben, mit Enter bestätigen

- Im Terminal/Console/Command Prompt eingeben und mit Enter bestätigen
> conda activate Takeaway

---
#### 4 Installation der python packages
> conda install -c conda-forge pandas pillow xlrd xlsxwriter opencv fuzzywuzzy pytesseract -y
---
#### 5 report.config erstellen im Schichtplan Ordner
- Neue Textatei erstellen
- In Texteditor öffnen
- 
