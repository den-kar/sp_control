# sp_control
Tool zur Erstellung eines Schichtplan-Kontroll-Reports. **sp_control** wird von einem Terminal mit eingerichteter und aktivierter Python-Umgebung aus genutzt und bietet durch Eingabe von Parametern zuschaltbare Funktionen.

---
### Anwendung

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

#### Terminal mit Python Umgebung öffnen
**Linux**
  1. `STRG + ALT + T` drücken
  1. `conda activate Takeaway` eingeben

**Windows**
  1. `Windows-Taste` drücken
  1. `Anaconda Prompt` eingeben
  1. mit `Enter` bestätigen

#### Beispielanwendungen

**Standard Wochenreport**
  - Kalenderwoche 47, default Städte, Verfügbarkeiten auslesen
> report.py -kw 47 -a

**Erstellung vollständiger `Rider_Ersterfassung_<STADTNAME>.xlsx` Datei**
  - ab KW 1, bis KW 47, nur Frankfurt, ohne auslesen der Verfügbarkeiten
> report.py -kw 1 -l 47 -c Frankfurt

---
### Schichtplan Arbeitsordner Struktur
```v
Schichtplan
├── Schichtplan_bearbeitet (wird autom. erstellt)
│   └── KW<Kalenderwoche>_<Stadtname>.xlsx (je Report eine xlsx)
├── Rider_Ersterkennung (wird autom. erstellt)
│   └── Rider_Ersterkennung_<Stadtname>.xlsx (je Stadt eine xlsx)
├── KW<Kalenderwoche> (ein Ordner je Schichtplan Datenpaket)
│   ├── .xlsx files (Schichtplan, Verfügbarkeiten, etc.)
│   └── .zip files (Verfügbarkeiten Screenshots)
├── config_report.json
├── Rider_Ersterfassung.xlsx (optional)
└── report.py
```

---
### Installation

Vollständige Anleitung [INSTALL.md](.../INSTALL.md)
