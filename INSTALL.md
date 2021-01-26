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

#### 2.3 site-packages installieren
In Terminal mit aktivierter Python Umgeben eingeben
> conda install -c conda-forge pandas pillow xlrd xlsxwriter openpyxl opencv fuzzywuzzy pytesseract matplotlib -y

Anschließend im selben Terminal eingeben
> pip install msoffcrypto-tool

---
### 3 Schichtplan Arbeitsordner einrichten

**Der Arbeitsordner hat folgende Struktur, teile werden automatisch erstellt**
```md
sp_control-master
├── Rider_Ersterfassung       (wird autom. erstellt)
│   └── Rider_Ersterfassung_<STADTNAME>.xlsx           (je Stadt eine xlsx)
├── Schichtplan_bearbeitet    (wird autom. erstellt)
│   └── KW<KALENDERWOCHE>_<STADTNAME>_<DATUMZEIT>.xlsx (je Stadt und KW eine xlsx)
├── Schichtplan_Daten         (manuell erstellen)
│   └── <JAHR>                (Ordner, Name ist 4-stellige Jahreszahl)
│       └── KW<KALENDERWOCHE> (ein Ordner je Schichtplan Datenpaket)
│           ├── .xlsx files   (Schichtplan, Verfügbarkeiten, Monatsstunden)
│           └── .zip files    (Verfügbarkeiten Screenshots)
├── config_report.json        (unter Linux optional, unter Windows "cmd_path" Parameter ausfüllen)
├── Rider_Ersterfassung.xlsx  (optional)
└── sp_control.py
```

#### Weg A - Dateien aus git-respository herunterladen

##### 3.1A git-repository runterladen
  1. auf [sp_control Projektseite](https://github.com/den-kar/sp_control) gehen
  1. grünes Feld `Code` anklicken
  1. im neu geöffneten Frame `Download ZIP` anklicken
  1. gezippten Ordner `sp_control-master` in das gewünschte lokale Verzeichnis entpacken

##### 3.2A **config_report_muster.json bearbeiten**
  - Datei Inhalt an den eigenen Standort anpassen -> **aliases**, **cities**, ggf. **cmd_path**
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
```json
{
  "cmd_path": "C:\\Tesseract-ocr\\tesseract.exe",
  "cities": ["Frankfurt", "Offenbach"],
  "password": "",
  "aliases": {
    "Darmstadt": ["darmstadt", "da"],
    "Frankfurt": ["frankfurt", "ffm", "frankfurt am main"],
    "Fürth": ["fürth", "fuerth", "fue"],
    "Nürnberg": ["nürnberg", "nuernberg", "nuremberg", "nbg", "nue"],
    "Offenbach": ["offenbach", "of", "offenbach am main"],
    "Stuttgart": ["stuttgart", "stg"],
    "Bielefeld": ["bielefeld", "osnabrБck", "bi"],
    "Münster": ["münster", "mБnster", "ms"],
    "Osnabrück": ["osnabrück", "os"]
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
Die Datei ist optional und wird nur beim aller ersten Run genutzt. Die Daten werden erfasst, um alle neuen Daten aus dem Schichtplan Datenpaket erweitert und anschließend unter `Rider_Erfassung/Rider_Ersterfassung_<STADTNAME>.xlsx` gespeichert. 

- xlsx Datei mit bekannten Ridern
- ein Reiter je Stadt
- first_entry und last_entry Datumsformat: **YYYY-MM-DD**

*Excel Kopfzeile mit Beispiel Eintrag*
rider name | contract type | min | city | first_entry | last_entry | current contract first entry | prev contracts | similar names
-|-|-|-|-|-|-|-|-
Jane Doe | Minijob | 5 | Offenbach | 2019-06-01 | 2020-11-16 | 2019-06-01
John Wayne | ... | ... | ... | ... | ... | ...

*Verwendete wöchtenliche Mindeststunden*
Vertragsart | Stunden
-|-
Foodora_Working Student | 12
Midijob | 12
Minijob | 5
Minijobber | 5
Mini-Jobber | 5
TE Midijob | 12
TE Minijob | 5
TE Teilzeit | 30
TE Werkstudent | 12
TE WS | 12
Teilzeit | 30
Vollzeit | 30
Werk Student | 12

---

### 4 Anaconda Prompot einrichten (nur Windows, optional)
1. `Windows`-Taste drücken
1. `Anaconda Prompt` eingeben
1. Rechts-Klick auf das `Anaconda Prompt` Feld
1. im Kontextfenster auf `open file location` klicken
1. im neuen Fenster Rechts-Klick auf `Anaconda Prompt Verknüpfung`
1. im Kontextfenster auf `Eigenschaften` klicken
1. im neuen Fenster auf den Reiter `Verknüpfung` gehen
1. unter `Ziel:` hinter `\Anaconda3`, ganz ans Ende `\envs\Takeaway` hinzufügen
1. unter `Ausführen:` den Pfad einfügen, in dem `sp_control.py` abliegt (z.B. `C:\Takeaway\sp_control-master`) 
1. auf das `OK` Feld klicken


---
### Anwendung

**sp_control** wird von einem Terminal mit eingerichteter und aktivierter Python-Umgebung aus genutzt und bietet durch Eingabe von Parametern zuschaltbare Funktionen.

```
usage: sp_control.py [-h] [-y START_YEAR] [-z LAST_YEAR] [-k START_KW] [-l LAST_KW] [-c [CITIES [CITIES ...]]] [-a] [-t] [-m] [-e] [-v]
```

#### Parameter
```
  -h, --help            show this help message and exit
  -y START_YEAR, --start_year START_YEAR
                        Jahr der zu prüfenden Daten, default: heutiges Jahr
  -z LAST_YEAR, --last_year LAST_YEAR
                        Letztes Jahr der zu prüfenden Daten, default: heutiges Jahr
  -k START_KW, --start_kw START_KW
                        Kalenderwoche der zu prüfenden Daten, default: 1
  -l LAST_KW, --last_kw LAST_KW
                        Letzte zu bearbeitende Kalenderwoche als Zahl
  -c [CITIES [CITIES ...]], --cities [CITIES [CITIES ...]]
                        Zu prüfende Stadt oder Städte, Stadtnamen trennen mit einem Leerzeichen, default: [Frankfurt Offenbach]
  -a, --get_avail       Aktiviert das Auslesen mitgeschickter Screenshots
  -t, --tidy_only       Räumt alle Verfügbarkeiten Screenshot Dateien auf
  -m, --mergeperday     Erstellt je Stadt und Tag eine zusammegesetzte Verfügbarkeiten-Screenshot-Datei
  -e, --ersterfassung   Erstellt nur die Rider_Ersterfassung Datei, ohne SP-Report
  -v, --visualize       Daten Visualisierung, erstellt Plots der vergebenen Schichten
```

#### Terminal mit Python Umgebung öffnen und zum Arbeitsorder navigieren
**Linux**
  1. `STRG + ALT + T` drücken
  1. `conda activate Takeaway` eingeben
  1. zum Datei Verzeichnis navigieren mit `cd ~/pfad/zu/sp_control-master`
  1. Tool starten mit `python sp_control.py`, gewünschte Parametern hinzufügen


**Windows**
  1. _Anaconda Prompt_ öffnen
  - `Windows-Taste` drücken
  - `Anaconda Prompt` eingeben
  - mit `Enter` bestätigen

  2. Python Umgebung aktivieren (überspringen, wenn **Schritt 4** durchgeführt wurde)
  - in _Anaconda Prompt_ `conda activate Takeaway` eingeben
  - mit `Enter` bestätigen

  3. zum Datei Verzeichnis navigieren (überspringen, wenn **Schritt 4** durchgeführt wurde)
  - wenn man den Ordner im Explorer öffnet und in die Adresszeile klickt, kann man den benötigten Pfad einfach kopieren
  - in _Anaconda Prompt_ eingeben `cd <Laufwerk>:\pfad\zu\sp_control-master`
  - mit `Enter` bestätigen

  4. Tool starten
   - `python sp_control.py` mit gewünschten Parametern eingeben
   - mit `Enter` bestätigen

---

### Beispielanwendungen

#### Standard Wochenreport, aktuelle Woche, mit Daten Visualisierung, default Städte
> python sp_control.py -a -v
  - liest Daten der folgenden KW vom Zeitpunkt der Ausführung
  - default Städte
    - default Städte können in der `config_report.json` auf die eigene Region angepasst werden, der Standard Wert ist ["Frankfurt", "Offenbach"]
  - Verfügbarkeiten auslesen
  - Schichten visualisieren
  - Report Output Pfad `Schichtplan_bearbeitet/`
  - Beispiel Report Dateinamen 
     - `KW4_Frankfurt_2021_01_21_17_56_51.xlsx`
     - `KW4_Offenbach_2021_01_21_17_56_51.xlsx`
  - Beispiel Plots Output Pfad `Schichtplan_Daten/2021/KW4/Analyse/`
  - Beispiel Plot Output Dateinamen
    - `Frankfurt_[2021-01-25 - 2021-01-31].png` (je Tag eine Datei)
    - `Offenbach_[2021-01-25 - 2021-01-31].png` (je Tag eine Datei)

#### Erstellung vollständiger `Rider_Ersterfassung_<STADTNAME>.xlsx` Datei
> python sp_control.py -y 2019 -k 29 -z 2021 -l 4 -e -c Frankfurt
  - ab KW 29 / 2019
  - bis KW 4 / 2021
  - Stadt Frankfurt
  - speichert nur Ersterfassung Datei 
    - keine Schichtplan-Reports
    - ohne auslesen der Verfügbarkeiten
  - Ersterfassung Datei Pfad `Rider_Ersterfassung/`
  - Ersterfassung Dateiname
    - `Rider_Ersterkennung_Frankfurt.xlsx`
  - weitere Parameter nach Bedarf zuschaltbar (z.B. -v für Daten Visualisierung)

