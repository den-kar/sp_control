## sp_control
Tool zur Erstellung eines Schichtplan-Kontroll-Reports. **sp_control** wird von einem Terminal mit eingerichteter und aktivierter Python-Umgebung aus genutzt und bietet durch Eingabe von Parametern zuschaltbare Funktionen.

---
### Installation

Vollständige Anleitung [INSTALL.md](https://github.com/den-kar/sp_control/blob/master/INSTALL.md)

---
### Schichtplan Arbeitsordner Struktur
```md
sp_control-master
├── Rider_Ersterfassung       (wird autom. erstellt)
│   └── Rider_Ersterfassung_<STADTNAME>.xlsx           (je Stadt eine xlsx)
├── Schichtplan_bearbeitet    (wird autom. erstellt)
│   └── KW<KALENDERWOCHE>_<STADTNAME>_<DATUMZEIT>.xlsx (je Stadt und KW eine xlsx)
├── Schichtplan_Daten         (manuell erstellen, Unterordner ebenfalls)
│   └── <JAHR>                (manuell erstellen, Ordner, Name ist 4-stellige Jahreszahl)
│       └── KW<KALENDERWOCHE> (manuell erstellen, ein Ordner je Schichtplan Datenpaket)
│           ├── .xlsx files   (Schichtplan, Verfügbarkeiten, Monatsstunden)
│           ├── Analyse       (wird erstellt bei Daten Visualisierung)
│           ├── logs          (wird erstellt bei Reporterstellung, enthält log Daten)
│           └── Screenhots    (wird erstellt bei Reporterstellung, Verfügbarkeiten Screenshots als zip, einzelne jpgs oder pngs)
├── config_report.json        (eigene Städte unter "cities" eintragen, unter Windows zusätzlich "cmd_path" Parameter ausfüllen)
├── Rider_Ersterfassung.xlsx  (optional)
└── sp_control.py
```

---
### Anwendung

**sp_control** wird von einem Terminal mit eingerichteter und aktivierter Python-Umgebung aus genutzt und bietet durch Eingabe von Parametern zuschaltbare Funktionen.

```
usage: python sp_control.py [-h] [-y START_YEAR] [-z LAST_YEAR] [-k START_KW] [-l LAST_KW] [-c [CITIES [CITIES ...]]] [-a] [-t] [-m] [-e] [-v]
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

#### Terminal mit Python Umgebung öffnen
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

  2. Python Umgebung aktivieren (überspringen, wenn _Anaconda Prompt_ eingerichtet ist)
  - in _Anaconda Prompt_ `conda activate Takeaway` eingeben
  - mit `Enter` bestätigen

  3. zum Datei Verzeichnis navigieren (überspringen, wenn _Anaconda Prompt_ eingerichtet ist)
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

