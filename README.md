# EBA ITS DPM Viewer

Eine Streamlit-Webanwendung zur originalgetreuen Darstellung von EBA-ITS-Excel-Dateien nach dem Data Point Model (DPM). Leere Eingabezellen werden automatisch mit DPM-Koordinaten beschriftet. Zeilen- und Spaltenbeschriftungen sowie Eingabezellen können manuell mit Freitextannotationen versehen werden, die als interaktive Tooltips beim Mouse-over erscheinen. Alle Koordinaten und Annotationen werden in einer CSV-Datei verwaltet.

---

## Projektstruktur

```
eba_viewer/
├── app.py                   # Streamlit-Hauptanwendung
├── excel_parser.py          # Excel-Verarbeitung und Koordinaten-Erkennung
├── renderer.py              # HTML-Rendering der Tabellen
├── export_coordinates.py    # CSV-Export aller DPM-Koordinaten und Labels
├── eba_styles.css           # Alle CSS-Regeln (auskommentiert, zentral anpassbar)
├── requirements.txt         # Python-Abhängigkeiten
├── README.md
└── data/
    ├── eba_template.xlsx    # Excel-Datei hier ablegen
    └── coordinates.csv      # wird automatisch generiert / manuell gepflegt
```

---

## Schnellstart

```bash
# 1. Abhängigkeiten installieren
pip install -r requirements.txt

# 2. Excel-Datei ablegen
mkdir -p data
cp /pfad/zur/datei.xlsx data/eba_template.xlsx

# 3. App starten
streamlit run app.py
```

Beim ersten Start wird `data/coordinates.csv` automatisch erzeugt.

---

## Konfiguration

**Option A – Umgebungsvariable (empfohlen für Produktion)**
```bash
export EBA_EXCEL_PATH="/vollständiger/pfad/zur/datei.xlsx"
streamlit run app.py
```

**Option B – Datei im `data/`-Ordner**
Die App sucht standardmäßig nach `data/eba_template.xlsx` relativ zum `app.py`-Verzeichnis.

**Option C – Direktanpassung in `app.py`**
```python
EXCEL_PATH = "/vollständiger/pfad/zur/datei.xlsx"
```

Der Name des Index-Blatts ist ebenfalls konfigurierbar (Zeile 25 in `app.py`):
```python
INDEX_SHEET = "Index"
```

Falls die Datei beim Start nicht gefunden wird, erscheint eine Fehlerseite mit einem temporären Upload-Formular.

---

## Features

### Navigation
- **Sidebar** mit allen sichtbaren Tabellenblättern — ausgeblendete Blätter (`hidden`, `veryHidden`) werden automatisch gefiltert
- **Index-Seite** als Startseite mit Tabellenansicht des Index-Blatts und Direktlink-Karten zu allen Blättern
- **Rücknavigation** von jedem Blatt zum Index über einen Button oben rechts
- **Tab-Leiste** für schnellen Wechsel zwischen Blättern

### Darstellung
- Originalgetreue Wiedergabe aller Zellformate: Hintergrundfarben, Schriftfarbe, Fett/Kursiv, Textausrichtung, Rahmen, Merge-Bereiche (rowspan/colspan)
- Korrekte Auflösung aller Excel-Farbformate: ARGB-RGB, Theme-Farben mit Tint/Shade, Indexed Colors (Legacy)
- Weiße, schwarze und near-white Füllungen (alle RGB-Kanäle ≥ 248) werden als „keine Füllung" behandelt
- Vollständig leere Zeilen und Spalten werden automatisch entfernt

### Einheitliche Typografie
- Eine einzige Schriftgröße (`10pt`) für alle Tabelleninhalte — Excel-Schriftgrößen werden nicht übernommen
- Ein einziger Font-Stack: `'Segoe UI', 'Inter', 'Calibri', system-ui, sans-serif`

### DPM-Koordinaten

Leere, ungefärbte Eingabezellen erhalten automatisch eine Koordinatenbezeichnung nach dem DPM-Schema.

**Format:** `<Blattname>_<Zeilencode>_<Spaltencode>`  
Beispiel: `C 01.00_0010_0020`

**Erkennungslogik:**
- *Spaltencodes:* Die ersten 15 Zeilen jedes Blatts werden durchsucht. Die erste Zeile mit mindestens einem vierstelligen numerischen Code (`0010`, `0020` …) gilt als Spalten-Header.
- *Zeilencodes:* Die ersten 5 Spalten werden verglichen. Die Spalte mit den meisten vierstelligen Codes wird als Zeilen-Header-Spalte verwendet.
- Eine Zelle erhält eine Koordinate **nur wenn** sie im Schnittbereich einer codierten Zeile und Spalte liegt **und** weder Hintergrundfüllung noch Textinhalt hat.

Die Koordinate erscheint als kleine blaue Beschriftung oben in der Zelle und im Tooltip.

### Annotationen & Tooltips

Jede Eingabezelle sowie jede Zeilen- und Spaltenbeschriftung kann mit einem Freitext-Kommentar versehen werden. Die Annotation erscheint beim Mouse-over als Sprechblase.

**Visueller Hinweis:** Zellen mit Annotation tragen einen kleinen orangenen Punkt (●) in der oberen rechten Ecke.

**Drei Arten von annotierbaren Einträgen:**

| Typ | Key-Format | Beispiel |
|---|---|---|
| Eingabezelle | `<Blatt>_<Zeilencode>_<Spaltencode>` | `C 01.00_0010_0020` |
| Spaltenbeschriftung | `<Blatt>_col_<Spaltencode>` | `C 01.00_col_0020` |
| Zeilenbeschriftung | `<Blatt>_row_<Zeilencode>` | `C 01.00_row_0010` |

**Annotation hinzufügen:**
1. `data/coordinates.csv` in Excel oder einem Texteditor öffnen
2. In der Spalte `annotation` den gewünschten Text eintragen
3. CSV speichern → App neu starten

Annotationen werden bei einem Re-Export der CSV **automatisch erhalten** – manuelle Einträge gehen nicht verloren, wenn die XLSX neu geladen wird.

### CSV-Export der Koordinaten

Beim Start der App wird `data/coordinates.csv` automatisch erzeugt bzw. aktualisiert.

**Speicherort:** `data/coordinates.csv` (relativ zum `app.py`-Verzeichnis)

**Spalten:**

| Spalte | Beschreibung | Beispiel |
|---|---|---|
| `key` | Eindeutiger Schlüssel (Koordinate oder Label-Key) | `C 01.00_0010_0020` |
| `type` | Eintragstyp: `cell`, `col_label`, `row_label` | `cell` |
| `sheet` | Blattname | `C 01.00` |
| `row_code` | Vierstelliger Zeilencode (leer bei `col_label`) | `0010` |
| `col_code` | Vierstelliger Spaltencode (leer bei `row_label`) | `0020` |
| `annotation` | Freitext-Kommentar (manuell befüllbar) | `Buchwert gem. IAS 39` |

**Standalone-Nutzung** (ohne App):
```bash
python export_coordinates.py
python export_coordinates.py /pfad/zur/datei.xlsx /pfad/zur/ausgabe.csv
```

**Als Modul:**
```python
from export_coordinates import export_coordinates
n = export_coordinates("data/meine_datei.xlsx", "data/coordinates.csv")
print(f"{n} Einträge exportiert")
```

---

## Styling anpassen (`eba_styles.css`)

Alle visuellen Parameter der Tooltips und annotierbaren Zellen sind in `eba_styles.css` zentral zusammengefasst und kommentiert. Die Datei wird beim Rendern jedes Blatts eingelesen – Änderungen wirken sich nach einem App-Neustart sofort aus.

| CSS-Klasse | Steuert |
|---|---|
| `.eba-coord-cell` | Eingabezellen mit DPM-Koordinate (position, overflow) |
| `.eba-label-cell` | Zeilen-/Spaltenbeschriftungen mit 4-stelligem Code |
| `.eba-badge` | Orangener Hinweis-Punkt bei vorhandener Annotation |
| `.eba-tooltip` | Sprechblase: Position, Größe, Farbe, Schatten, Schrift |
| `.eba-tooltip-coord` | Technischer Key im Tooltip (oben, blau) |
| `.eba-tooltip-divider` | Trennlinie zwischen Key und Annotationstext |
| `.eba-tooltip-text` | Annotationstext im Tooltip (unten, fast-weiß) |

---

## Modulbeschreibung

### `app.py`
Streamlit-Hauptanwendung. Enthält Seitenkonfiguration, globales CSS, Session-State-Routing, Sidebar, Index-Seite und Blatt-Ansicht. Lädt Workbook-Daten über `@st.cache_resource` (einmaliges Parsen pro Anwendungsstart), triggert den CSV-Export und wendet anschließend Annotationen aus der CSV auf die geparsten Zellen an (`_apply_annotations`).

### `excel_parser.py`
Liest Excel-Dateien mit `openpyxl` und gibt strukturierte `SheetData`-Objekte zurück.

| Funktion / Klasse | Aufgabe |
|---|---|
| `parse_workbook(path)` | Lädt die Arbeitsmappe, filtert ausgeblendete Blätter, parst alle sichtbaren Blätter, ruft `_build_coordinates` auf |
| `_parse_sheet(ws, theme_colors)` | Liest Zellen, Merge-Bereiche, Spaltenbreiten, Zeilenhöhen; entfernt leere Zeilen/Spalten |
| `_extract_style(cell, theme_colors)` | Extrahiert Füllfarbe, Schriftformat, Ausrichtung, Rahmen, Zahlenformat |
| `_resolve_color(...)` | Konvertiert ARGB-, Theme- und Indexed-Farben in 6-stellige Hex-Strings |
| `_is_near_white(hex6)` | `True` wenn alle RGB-Kanäle ≥ 248 → wird als „keine Füllung" behandelt |
| `_build_coordinates(sheet)` | Erkennt DPM-Koordinaten (Scan bis Zeile 15 / Spalte 5); setzt `cell.coordinate` für Eingabezellen sowie `cell.label_key` für Zeilen-/Spaltenköpfe |
| `CellData` | Datenklasse pro Zelle: Wert, Anzeigetext, Style, Rowspan/Colspan, `coordinate`, `label_key`, `annotation` |
| `SheetData` | Datenklasse pro Blatt: Zellen-Matrix, Spaltenbreiten, Zeilenhöhen |

### `renderer.py`
Wandelt `SheetData`-Objekte in HTML-`<table>`-Strings um. Unterscheidet drei Zelltypen beim Rendering:

| Typ | Bedingung | Rendering |
|---|---|---|
| Eingabezelle | `cell.coordinate` gesetzt | Hellblauer Hintergrund, Koordinaten-Label, Tooltip |
| Beschriftungszelle | `cell.label_key` gesetzt | Normales Styling + Tooltip bei Hover |
| Standardzelle | keines von beidem | Normales Styling, kein Tooltip |

CSS wird aus `eba_styles.css` geladen und einmalig pro Tabelle als `<style>`-Block eingefügt. Fällt die CSS-Datei weg, bleibt der Viewer voll funktionsfähig (nur ohne Tooltip-Styling).

### `export_coordinates.py`
Iteriert über alle geparsten Blätter und Zellen, sammelt Eingabe-Koordinaten (`cell.coordinate`) und Label-Keys (`cell.label_key`) und schreibt sie als CSV. Bestehende Annotationen werden vor dem Überschreiben eingelesen und in die neue Datei übertragen.

### `eba_styles.css`
Zentrale Stylesheet-Datei. Jede Regel ist mit Kommentaren versehen, die erklären, welchen visuellen Parameter sie steuert. Kann ohne Python-Kenntnisse angepasst werden.

---

## Bekannte Einschränkungen

- Diagramme und eingebettete Bilder werden nicht dargestellt
- Bedingte Formatierungen werden nicht ausgewertet
- Passwortgeschützte Dateien werden nicht unterstützt
- Formeln werden nicht neu berechnet (`data_only=True`) — es werden die zuletzt in Excel gespeicherten Werte angezeigt

---

## Performance-Hinweise

- `@st.cache_resource` stellt sicher, dass Parsen, CSV-Export und Annotation-Mapping nur einmal pro Anwendungsstart ausgeführt werden
- Für Dateien über 50 MB: `server.maxUploadSize = 200` in der Streamlit-Konfiguration setzen
- Bei sehr vielen Blättern (> 50) kann eine Filterung der Sidebar nach Kategorie sinnvoll sein