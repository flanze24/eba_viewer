# EBA ITS DPM Viewer

Eine Streamlit-Webanwendung zur originalgetreuen Darstellung von EBA-ITS-Excel-Dateien nach dem Data Point Model (DPM). Leere Eingabezellen werden automatisch mit DPM-Koordinaten beschriftet und als CSV exportiert.

---

## Projektstruktur

```
eba_viewer/
├── app.py                   # Streamlit-Hauptanwendung
├── excel_parser.py          # Excel-Verarbeitung und Koordinaten-Erkennung
├── renderer.py              # HTML-Rendering der Tabellen
├── export_coordinates.py    # CSV-Export aller DPM-Koordinaten
├── requirements.txt         # Python-Abhängigkeiten
├── README.md
└── data/
    ├── eba_template.xlsx    # Excel-Datei hier ablegen
    └── coordinates.csv      # wird automatisch generiert
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
- CSS `!important` auf Wrapper-Ebene verhindert jeden per-Zelle-Override

### DPM-Koordinaten
Leere, ungefärbte Eingabezellen erhalten automatisch eine Koordinatenbezeichnung nach dem DPM-Schema.

**Format:** `<Blattname>_<Zeilencode>_<Spaltencode>`
Beispiel: `C 01.00_0010_0020`

**Erkennungslogik:**
- *Spaltencodes:* Die ersten **15 Zeilen** jedes Blatts werden durchsucht. Die erste Zeile mit mindestens einem vierstelligen numerischen Code (`0010`, `0020` …) gilt als Spalten-Header.
- *Zeilencodes:* Die ersten **5 Spalten** werden verglichen. Die Spalte mit den meisten vierstelligen Codes wird als Zeilen-Header-Spalte verwendet.
- Eine Zelle erhält eine Koordinate **nur wenn** sie im Schnittbereich einer codierten Zeile und Spalte liegt **und** weder Hintergrundfüllung noch Textinhalt hat.
- Zellen ohne Zeilenzuordnung erhalten keine Koordinate.

Die Koordinate erscheint als kleine blaue Beschriftung oben in der Zelle und im Tooltip.

### CSV-Export der Koordinaten
Beim Start der App wird `data/coordinates.csv` automatisch erzeugt (bzw. aktualisiert).

**Speicherort:** `data/coordinates.csv` (relativ zum `app.py`-Verzeichnis)

**Spalten:**

| Spalte | Beschreibung | Beispiel |
|---|---|---|
| `coordinate` | Vollständige Koordinate | `C 01.00_0010_0020` |
| `sheet` | Blattname | `C 01.00` |
| `row_code` | Vierstelliger Zeilencode | `0010` |
| `col_code` | Vierstelliger Spaltencode | `0020` |

**Integration in `app.py`** – der Export wird in Zeile 198 aufgerufen, direkt nach `parse_workbook`:
```python
@st.cache_resource(show_spinner="⏳ Lade Excel-Datei …")
def load_workbook(path: str) -> dict[str, SheetData] | None:
    try:
        sheets = parse_workbook(path)
        from export_coordinates import export_coordinates
        export_coordinates(path)        # ← Zeile 199: CSV-Export
        return sheets
    ...
```

Da `@st.cache_resource` den Block nur einmal pro Anwendungsstart ausführt, wird die CSV ebenfalls nur einmal geschrieben.

**Standalone-Nutzung** (ohne App):
```bash
# Mit Standardpfaden
python export_coordinates.py

# Mit eigenen Pfaden
python export_coordinates.py /pfad/zur/datei.xlsx /pfad/zur/ausgabe.csv
```

**Als Modul:**
```python
from export_coordinates import export_coordinates
n = export_coordinates("data/meine_datei.xlsx", "data/coordinates.csv")
print(f"{n} Koordinaten exportiert")
```

---

## Modulbeschreibung

### `app.py`
Streamlit-Hauptanwendung. Enthält Seitenkonfiguration, globales CSS, Session-State-Routing, Sidebar, Index-Seite und Blatt-Ansicht. Lädt die Workbook-Daten über `@st.cache_resource` (einmaliges Parsen pro Anwendungsstart) und triggert den CSV-Export nach dem Laden.

### `excel_parser.py`
Liest Excel-Dateien mit `openpyxl` und gibt strukturierte `SheetData`-Objekte zurück.

| Funktion / Klasse | Aufgabe |
|---|---|
| `parse_workbook(path)` | Lädt die Arbeitsmappe, filtert ausgeblendete Blätter, parst alle sichtbaren Blätter, ruft `_build_coordinates` auf |
| `_parse_sheet(ws, theme_colors)` | Liest Zellen, Merge-Bereiche, Spaltenbreiten, Zeilenhöhen; entfernt leere Zeilen/Spalten |
| `_extract_style(cell, theme_colors)` | Extrahiert Füllfarbe, Schriftformat, Ausrichtung, Rahmen, Zahlenformat |
| `_resolve_color(color_obj, theme_colors, ignore_alpha)` | Konvertiert ARGB-, Theme- und Indexed-Farben in 6-stellige Hex-Strings; ignoriert Alpha-Byte bei Füllfarben |
| `_is_near_white(hex6, threshold=248)` | Gibt `True` zurück wenn alle RGB-Kanäle ≥ threshold — solche Farben gelten als „keine Füllung" |
| `_build_coordinates(sheet)` | Erkennt das DPM-Koordinatensystem (Scan bis Zeile 15 / Spalte 5) und weist Eingabezellen ihre Koordinate zu |
| `CellData` | Datenklasse pro Zelle: Wert, Anzeigetext, Style, Rowspan/Colspan, Koordinate |
| `SheetData` | Datenklasse pro Blatt: Zellen-Matrix, Spaltenbreiten, Zeilenhöhen |

### `renderer.py`
Wandelt `SheetData`-Objekte in HTML-`<table>`-Strings um. Keine per-Zelle-Schriftgröße oder Schriftart. Eingabezellen mit Koordinate erhalten hellblauen Hintergrund (`#F0F4FF`) mit Koordinatenbeschriftung.

### `export_coordinates.py`
Iteriert über alle geparsten Blätter und Zellen, sammelt alle gesetzten `cell.coordinate`-Werte und schreibt sie als CSV mit den Spalten `coordinate`, `sheet`, `row_code`, `col_code`. Gibt die Anzahl der exportierten Koordinaten zurück. Kann als Skript oder als importiertes Modul verwendet werden.

---

## Bekannte Einschränkungen

- Diagramme und eingebettete Bilder werden nicht dargestellt
- Bedingte Formatierungen werden nicht ausgewertet
- Passwortgeschützte Dateien werden nicht unterstützt
- Formeln werden nicht neu berechnet (`data_only=True`) — es werden die zuletzt in Excel gespeicherten Werte angezeigt

---

## Performance-Hinweise

- `@st.cache_resource` stellt sicher, dass Parsen und CSV-Export nur einmal pro Anwendungsstart ausgeführt werden
- Für Dateien über 50 MB: `server.maxUploadSize = 200` in der Streamlit-Konfiguration setzen
- Bei sehr vielen Blättern (> 50) kann eine Filterung der Sidebar nach Kategorie sinnvoll sein