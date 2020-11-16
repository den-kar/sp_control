### 1 Installation

#### 1.1 Anaconda
- Installer [runterladen](https://www.anaconda.com/products/individual) -> bei *Anaconda Installers*
- Installationsanleitung für [Linux](https://docs.continuum.io/anaconda/install/linux/) oder [Windows](https://docs.continuum.io/anaconda/install/windows/) befolgen

#### 1.2 Tesseract
##### Linux
- Terminal öffnen mit `STRG + ALT + T`
- Tesseract installieren mit 
> sudo apt-get install tesseract-ocr -y

##### Windows
[Schritt-für-Schritt Anleitung](https://medium.com/quantrium-tech/installing-and-using-tesseract-4-on-windows-10-4f7930313f82). Dabei folgendes beachten:
- Das setzen der Umgebungsvariablen ist nicht nötig und kann übersprungen werden.
- [Alternative Installer Quelle](https://github.com/UB-Mannheim/tesseract/wiki) -> bei *The latest installers can be downloaded here*
- Pfad zur installierten *tesseract.exe* notieren (z.B. **C:\\Tesseract-OCR\\tesseract.exe**)

WARNUNG: Tesseract sollte entweder in das Verzeichnis, das bei der Installation vorgeschlagen wird, oder in ein neues Verzeichnis installiert werden. Das Deinstallationsprogramm entfernt das gesamte Installationsverzeichnis. Wenn Tesseract in ein bestehendes Verzeichnis installiert wurde, wird dieses Verzeichnis mit all seinen Unterverzeichnissen und Dateien entfernt.


---
### 2 Python einrichten

#### 2.1 - Einrichtung einer eigenen Takeaway Python Umgebung
Terminal öffnen
- Linux: `STRG + ALT + T` drücken
- Windows:
  1. `Windows-Taste` drücken
  1. `Anaconda Prompt` eingeben
  1. mit `Enter` bestätigen

Zur Erstellung der neuen Python Umgebung **Takeaway** folgende Zeile ins Terminal einfügen und mit `Enter` bestätigen 
 > conda create -n Takeaway python=3.8.5 -y

#### 2.2 Python Umgebung in Terminal aktivieren
Um in der gewünschten conda Python Umgebung zu arbeiten, muss in die Umgebung in **jedem neu geöffneten Terminal** aktiviert werden mit
> conda activate Takeaway

Wenn im Terminal die aktive Zeile mit `(Takeaway)` anfängt, ist die Umgebung aktiviert

#### 2.2 site-packages installieren
In Terminal mit aktivierter Python Umgeben eingeben
> conda install -c conda-forge pandas pillow xlrd xlsxwriter opencv fuzzywuzzy pytesseract -y

---
### 3 Schichtplan Arbeitsordner einrichten
Der Arbeitsordner hat folgende Struktur 

```v
sp_control-master
├── Rider_Ersterfassung (wird autom. erstellt)
│   └── Rider_Ersterfassung_<STADTNAME>.xlsx (je Stadt eine xlsx)
├── Schichtplan_bearbeitet (wird autom. erstellt)
│   └── KW<KALENDERWOCHE>_<STADTNAME>_<YYYY_mm_dd_HH_MM_SS>.xlsx (je Report eine xlsx)
├── Schichtplan_Daten
│   └── KW<KALENDERWOCHE> (ein Ordner je Schichtplan Datenpaket)
│       ├── .xlsx files (Schichtplan, Verfügbarkeiten, etc.)
│       └── .zip files (Verfügbarkeiten Screenshots)
├── config_report.json
├── Rider_Ersterfassung.xlsx (optional)
└── sp_control.py
```

#### Weg A - Arbeitsordner aus git-respository erstellen

##### 3.1A **git-repository runterladen**
  - auf [sp_control Projektseite](https://github.com/den-kar/sp_control) gehen
  - grünes Feld `Code` anklicken
  - im neu geöffneten Frame `Download ZIP` anklicken
  - beinhaltetes Ordner `sp_control` in das gewünschte lokale Verzeichnis entpacken

##### 3.2A **config_report_muster.json bearbeiten**
  - Datei an den eigenen Standort anpassen
  - Datei umbenennen in **config_report.json**
  - Beschreibung unter **3.2B**

##### 3.3A Rider_Ersterfassung_Muster.xlsx bearbeiten (optional)
  - Beschreibung unter **3.3B**

#### Weg B - Dateien manuell erstellen

##### 3.1B sp_control.py erstellen
1. [sp_control.py](https://github.com/den-kar/sp_control/blob/master/sp_control.py) auf github öffnen
1. auf `Raw` klicken (eine Zeile über dem Code, rechte Seite)
1. kompletten Text markieren `STRG + A` und in Zwischenablage kopieren `STRG + C`
1. Neue Textatei erstellen
1. In Texteditor öffnen
1. Zwischenablage-Inhalt einfügen `STRG + V`
1. unter *sp_control.py* speichern in angezeigter Stelle im Verzeichnisbaum

##### 3.2B config_report.json erstellen
1. Neue Textatei erstellen
1. In Texteditor öffnen
1. Datei-Inhalt einfügen
```yaml
{
    "tesseract": {
        "cmd_path": "C:\\Tesseract-OCR\\tesseract.exe"
    }
    "cities": ["Frankfurt", "Offenbach"],
    "aliases": {
        "Frankfurt": ["frankfurt", "ffm", "frankfurt am main"],
        "Offenbach": ["offenbach", "of", "offenbach am main"],
        "avail": ["Verfügbarket", "Verfügbarkeiten"],
        "month": ["Monatsstunden", "Stunden"],
        "shift": ["Schichtplan", "Schichtplanung"]
    }
}
```
4. unter *config_report.json* speichern in angezeigter Stelle im Verzeichnisbaum
- alle eingegeben Werte mit *doppelten Anführungszeichen* umschließen
- nur unter Windows *muss* der **cmd_path** eingetragen werden
  - hierfür den in 1.2 notierten Installationspfad der `tesseract.exe` verwenden
  - Ordner-Trennzeichen (Backslash) doppelt eingeben `\\`
- Die Werte für **cities** werden im shiftplaner-Tool als default-Werte genutzt, wenn sonst keine Städte angegeben werden
- Jede Stadt in **cities** *benötigt* eine eigene Zeile in **aliases** nach vorgegebenem Muster 

##### 3.3B Rider_Ersterfassung.xlsx (optional)
Die Datei ist optional und wird nur beim aller ersten Run genutzt. Die Daten werden erfasst, um alle neuen Daten aus dem Schichtplan Datenpaket erweitert und anschließend unter `Rider_Erfassung/Rider_Ersterfassung_\<STADTNAME>.xlsx` gespeichert. 

- xlsx Datei mit bekannten Ridern
- ein Reiter je Stadt

- Excel Kopfzeile mit Beispiel Eintrag

rider name | contract type | min | city | first_entry | last_entry
-|-|-|-|-|-
Jane Doe | Minijob | 5 | Offenbach | 2019-06-01 | 2020-11-16

  - first_entry und last_entry Datumsformat: **YYYY-MM-DD**
  - Verwendete wöchtenliche Mindeststunden

Vertragsart | Stunden
-|-
Foodora_Working Student | 12
Midijob                 | 12
Minijob                 |  5
Minijobber              |  5
Mini-Jobber             |  5
TE Midijob              | 12
TE Minijob              |  5
TE Teilzeit             | 30
TE Werkstudent          | 12
TE WS                   | 12
Teilzeit                | 30
Vollzeit                | 30
Werk Student            | 12

---
### Anwendung

**sp_control** wird von einem Terminal mit eingerichteter und aktivierter Python-Umgebung aus genutzt und bietet durch Eingabe von Parametern zuschaltbare Funktionen.
```d
report.py [-h] --kalenderwoche KALENDERWOCHE [--last_kw LAST_KW] [--cities [CITIES [CITIES ...]]] [--getavails] [--mergeperday] [--unzip_only]
```

#### Parameter
```v
Pflichtparameter:
--kalenderwoche, -k  erste zu bearbeitende Kalenderwoche als Zahl

Optionale Parameter:
--last_kw, -lm       letzte zu bearbeitende Kalenderwoche als Zahl
--cities, -c         zu bearbeitende Städte; einzeln aufführen, trennen mit Leerzeichen
--getavails, -a      Auslesen von Verfügbarkeiten aus Screenshots aktivieren
--mergeperdaz, -m    Erzeugen zusammengefasster Screenshots je Stadt und Tag
--unzip_only, -z     zip Dateien mit Screenshots entpacken und shiftplaner beenden
```

#### Terminal mit Python Umgebung öffnen und zum Arbeitsorder navigieren
**Linux**
  1. `STRG + ALT + T` drücken
  1. `conda activate Takeaway` eingeben
  1. zum Datei Verzeichnis navigieren mit `cd ~/pfad/zu/sp_control-master`

**Windows**
  1. `Windows-Taste` drücken
  1. `Anaconda Prompt` eingeben
  1. mit `Enter` bestätigen
  1. zum Datei Verzeichnis navigieren mit `cd <LAUFWERK>:\pfad\zu\sp_control-master`

#### Beispielanwendungen

**Kompletter Report `KW<KALENDERWOCHE>_<STADTNAME><YYYY_mm_dd_HH_MM_SS>.xlsx`**
  - Kalenderwoche 47, default Städte, Verfügbarkeiten auslesen
> sp_control.py -kw 47 -a

**Erstellung vollständiger `Rider_Ersterfassung_<STADTNAME>.xlsx` Datei**
  - ab KW 1, bis KW 47, nur Frankfurt, ohne auslesen der Verfügbarkeiten
> sp_control.py -kw 1 -l 47 -c Frankfurt
