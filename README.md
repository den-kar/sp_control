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
│   └── <JAHR>                (Ordner, Name ist 4-stellige Jahreszahl)
│       └── KW<KALENDERWOCHE> (ein Ordner je Schichtplan Datenpaket)
│           ├── .xlsx files   (Schichtplan, Verfügbarkeiten, Monatsstunden)
│           └── screenhots    (Verfügbarkeiten Screenshots als zip, einzelne jpgs oder pngs)
├── config_report.json        (eigene Städte unter "cities" eintragen, unter Windows zusätzlich "cmd_path" Parameter ausfüllen)
├── Rider_Ersterfassung.xlsx  (optional)
└── sp_control.py
```

---
### Anwendung

**sp_control** wird von einem Terminal mit eingerichteter und aktivierter Python-Umgebung aus genutzt und bietet durch Eingabe von Parametern zuschaltbare Funktionen.

```md
usage: sp_control.py [-h] [-y YEAR] [-kw KALENDERWOCHE] [-lkw LAST_KW] [-c [CITIES [CITIES ...]]] [-a] [-to] [-m] [-eeo]
```

#### Parameter
```md
optional arguments:
  -h, --help            show this help message and exit
  -y YEAR, --year YEAR  Jahr der zu prüfenden Daten, default: heutiges Jahr
  -kw KALENDERWOCHE, --kalenderwoche KALENDERWOCHE
                        Kalenderwoche der zu prüfenden Daten, default: 1
  -lkw LAST_KW, --last_kw LAST_KW
                        Letzte zu bearbeitende Kalenderwoche als Zahl
  -c [CITIES [CITIES ...]], --cities [CITIES [CITIES ...]]
                        Zu prüfende Stadt oder Städte, Stadtnamen trennen mit einem Leerzeichen, default: [Frankfurt Offenbach]
  -a, --get_avail       Aktiviert das Auslesen mitgeschickter Screenshots
  -to, --tidy_only      Räumt alle Verfügbarkeiten Screenshot Dateien auf
  -m, --mergeperday     Erstellt je Stadt und Tag eine zusammegesetzte Verfügbarkeiten-Screenshot-Datei
  -eeo, --ersterkennung_only
                        Erstellt nur die Rider_Ersterkennung Datei, ohne SP-Report
  -v, --visualize_shifts
                        Daten Visualisierung, erstellt Plots der vergebenen Schichten
```

#### Terminal mit Python Umgebung öffnen
**Linux**
  1. `STRG + ALT + T` drücken
  1. `conda activate Takeaway` eingeben
  1. zum Datei Verzeichnis navigieren mit `cd ~/pfad/zu/sp_control-master`

**Windows**
  1. _Anaconda Prompt_ öffnen
  - `Windows-Taste` drücken
  - `Anaconda Prompt` eingeben
  - mit `Enter` bestätigen

  2. Python Umgebung aktivieren
  - in _Anaconda Prompt_ `conda activate Takeaway` eingeben
  - mit `Enter` bestätigen

  3. zum Datei Verzeichnis navigieren
  - wenn man den Ordner im Explorer öffnet und in die Adresszeile klickt, kann man den benötigten Pfad einfach kopieren
  - in _Anaconda Prompt_ eingeben `cd <Laufwerk>:\pfad\zu\sp_control-master`
  - mit `Enter` bestätigen

---
### Beispielanwedungen

**Standard Wochenreport mit Daten Visualisierung**
  - Jahr 2020, Kalenderwoche 47, Stadt Frankfurt, Verfügbarkeiten auslesen, Schichten visualisieren
  - Report Output Pfad `Schichtplan_bearbeitet/`
    - Dateiname `KW47_Frankfurt_<ERSTELLUNGSDATUM>.xlsx`
  - Plots Output Pfad `Schichtplan_Daten/2020/KW47/Analyse`
    - Dateiname `Frankfurt_KW47_[1_Montag - 7_Sonntag].png`
> sp_control.py -y 2020 -kw 47 -a -v -c Frankfurt

**Erstellung vollständiger `Rider_Ersterfassung_<STADTNAME>.xlsx` Datei**
  - ab KW 1, bis KW 53, default Städte, speichert nur Ersterfassung Datei -> keine Schichtplan-Reports, ohne auslesen der Verfügbarkeiten
  - Default Städte können in der `config_report.json` auf die eigene Region angepasst werden, der Standard Wert ist ["Frankfurt", "Offenbach"]
  - Ersterfassung Datei Pfad `Rider_Ersterfassung/`
    - Dateiname
      - `Rider_Ersterkennung_Frankfurt.xlsx`
      - `Rider_Ersterkennung_Offenbach.xlsx`
  - weitere Parameter nach Bedarf zuschaltbar (z.B. -v für Daten Visualisierung)
> sp_control.py -kw 1 -lkw 53 -eeo

