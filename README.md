## sp_control
Tool zur Erstellung eines Schichtplan-Kontroll-Reports. **sp_control** wird von einem Terminal mit eingerichteter und aktivierter Python-Umgebung aus genutzt und bietet durch Eingabe von Parametern zuschaltbare Funktionen.

---
### Installation

Vollständige Anleitung [INSTALL.md](https://github.com/den-kar/sp_control/blob/master/INSTALL.md)

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
  1. `Windows-Taste` drücken
  1. `Anaconda Prompt` eingeben
  1. mit `Enter` bestätigen
  1. zum Datei Verzeichnis navigieren mit `cd <LAUFWERK>:\pfad\zu\sp_control-master`

#### Beispielanwendungen

**Standard Wochenreport**
  - Kalenderwoche 47, default Städte, Verfügbarkeiten auslesen
> sp_control.py -kw 47 -a

**Erstellung vollständiger `Rider_Ersterfassung_<STADTNAME>.xlsx` Datei**
  - ab KW 1, bis KW 47, nur Frankfurt, ohne auslesen der Verfügbarkeiten
> sp_control.py -kw 1 -l 47 -c Frankfurt

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
