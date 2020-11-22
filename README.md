## sp_control
Tool zur Erstellung eines Schichtplan-Kontroll-Reports. **sp_control** wird von einem Terminal mit eingerichteter und aktivierter Python-Umgebung aus genutzt und bietet durch Eingabe von Parametern zuschaltbare Funktionen.

---
### Installation

Vollständige Anleitung [INSTALL.md](https://github.com/den-kar/sp_control/blob/master/INSTALL.md)

---
### Anwendung

```d
python sp_control.py [-h] --kalenderwoche KALENDERWOCHE [--last_kw LAST_KW] [--cities [CITIES [CITIES ...]]] [--getavails] [--mergeperday] [--unzip_only]
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
```v
sp_control-master
├── Rider_Ersterfassung       (wird autom. erstellt)
│   └── Rider_Ersterfassung_<STADTNAME>.xlsx           (je Stadt eine xlsx)
├── Schichtplan_bearbeitet    (wird autom. erstellt)
│   └── KW<KALENDERWOCHE>_<STADTNAME>_<DATUMZEIT>.xlsx (je Stadt und KW eine xlsx)
├── Schichtplan_Daten         (manuell erstellen)
│   └── KW<KALENDERWOCHE>     (ein Ordner je Schichtplan Datenpaket)
│       ├── .xlsx files       (Schichtplan, Verfügbarkeiten, Monatsstunden)
│       └── .zip files        (Verfügbarkeiten Screenshots)
├── config_report.json
├── Rider_Ersterfassung.xlsx  (optional)
└── sp_control.py
```
