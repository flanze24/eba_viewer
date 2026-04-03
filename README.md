# EBA ITS DPM Viewer

Eine Streamlit-Webanwendung zur originalgetreuen Darstellung von EBA-ITS-Excel-Dateien nach dem Data Point Model (DPM).

---

## Projektstruktur

```
eba_viewer/
├── app.py              # Streamlit-Hauptanwendung
├── excel_parser.py     # Excel-Verarbeitung und Koordinaten-Erkennung
├── renderer.py         # HTML-Rendering der Tabellen
├── requirements.txt    # Python-Abhängigkeiten
├── README.md
└── data/
    └── eba_template.xlsx   # Excel-Datei hier ablegen
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

---

## Konfiguration

Der Pfad zur Excel-Datei kann auf drei Wegen gesetzt werden, in absteigender Priorität:

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

Falls die Datei beim Start nicht gefunden wird, erscheint eine Fehlerseite mit einem temporären Upload-Formular.

Der Name des Index-Blatts ist ebenfalls konfigurierbar:
```python
INDEX_SHEET = "Index"   # Standardwert in app.py
```

---

## Features

### Navigation
- **Sidebar** mit allen sichtbaren Tabellenblättern — ausgeblendete (`hidden`, `veryHidden`) werden automatisch gefiltert
- **Index-Seite** als Startseite mit Tabellenansicht des Index-Blatts und Direktlink-Karten zu allen Blättern
- **Rücknavigation** von jedem Blatt zum Index über einen Button oben rechts
- **Tab-Leiste** für schnellen Wechsel zwischen Blättern

### Darstellung
- Originalgetreue Wiedergabe aller Zellformate: Hintergrundfarben, Schriftfarbe, Fett/Kursiv, Textausrichtung, Rahmen, Merge-Bereiche (rowspan/colspan)
- Korrekte Auflösung von Excel-Farbformaten: ARGB-RGB, Theme-Farben mit Tint/Shade, Indexed Colors
- Weiße und schwarze Füllungen werden als „keine Füllung" behandelt (visuell identisch mit leerem Hintergrund)
- Vollständig leere Zeilen und Spalten werden automatisch entfernt

### Einheitliche Typografie
- Eine einzige Schriftgröße (`10pt`) für alle Tabelleninhalte — Excel-Schriftgrößen werden nicht übernommen
- Ein einziger Font-Stack für die gesamte Anwendung: `'Segoe UI', 'Inter', 'Calibri', system-ui, sans-serif`
- Kein per-Zelle-Override möglich (CSS `!important` auf Wrapper-Ebene als Absicherung)

### DPM-Koordinaten
Leere, ungefärbte Eingabezellen in Tabellenblättern mit erkennbarer DPM-Struktur erhalten automatisch eine Koordinatenbezeichnung.

**Format:** `<Blattname>_<Zeilencode>_<Spaltencode>`
Beispiel: `C 01.00_0010_0020`

**Erkennungslogik:**
- *Spaltencodes:* Die ersten **15 Zeilen** jedes Blatts werden nach einer Zeile durchsucht, die mindestens einen vierstelligen numerischen Code (z.B. `0010`, `0020`) enthält. Die erste solche Zeile gilt als Spalten-Header.
- *Zeilencodes:* Die ersten **5 Spalten** werden verglichen; die Spalte mit den meisten vierstelligen Codes wird als Zeilen-Header-Spalte verwendet.
- Eine Zelle bekommt eine Koordinate **nur wenn** sie sich im Schnittbereich einer Zeile mit Zeilencode und einer Spalte mit Spaltencode befindet **und** weder eine Hintergrundfüllung noch Textinhalt hat.
- Zellen ohne Zeilenzuordnung bekommen keine Koordinate.

Die Koordinate wird als kleine blaue Beschriftung oben in der Zelle angezeigt und ist auch im Tooltip sichtbar.

---

## Modulbeschreibung

### `excel_parser.py`

Liest Excel-Dateien mit `openpyxl` und gibt strukturierte `SheetData`-Objekte zurück.

| Funktion / Klasse | Aufgabe |
|---|---|
| `parse_workbook(path)` | Lädt die Arbeitsmappe, filtert ausgeblendete Blätter, parst alle sichtbaren Blätter und ruft `_build_coordinates` auf |
| `_parse_sheet(ws, theme_colors)` | Liest Zellen, Merge-Bereiche, Spaltenbreiten und Zeilenhöhen; entfernt leere Zeilen/Spalten |
| `_extract_style(cell, theme_colors)` | Extrahiert Füllfarbe, Schriftformat, Ausrichtung, Rahmen und Zahlenformat |
| `_resolve_color(color_obj, theme_colors, ignore_alpha)` | Konvertiert ARGB-, Theme- und Indexed-Farben in 6-stellige Hex-Strings; ignoriert bei Füllfarben das Alpha-Byte (Excel-Verhalten) |
| `_build_coordinates(sheet)` | Erkennt das DPM-Koordinatensystem und weist Eingabezellen ihre Koordinate zu |
| `CellData` | Datenklasse pro Zelle: Wert, Anzeigetext, Style, Rowspan/Colspan, Koordinate |
| `SheetData` | Datenklasse pro Blatt: Zellen-Matrix, Spaltenbreiten, Zeilenhöhen |

### `renderer.py`

Wandelt `SheetData`-Objekte in HTML-`<table>`-Strings um.

- Keine per-Zelle-Schriftgröße oder Schriftart — einheitliche Basis auf `<table>`-Ebene
- Eingabezellen mit Koordinate erhalten hellblauen Hintergrund (`#F0F4FF`) und die Koordinate als kleine Beschriftung
- Unterstützt einen optionalen `link_resolver` für klickbare Index-Links

### `app.py`

Streamlit-Hauptanwendung mit:
- `@st.cache_resource` für performantes Laden — die Datei wird nur einmal geparst, auch bei mehreren gleichzeitigen Sitzungen
- Session-State-basiertes Routing zwischen Index und Blatt-Ansichten
- Globaler CSS-Override zur Durchsetzung der einheitlichen Typografie
- Fallback-Upload wenn die konfigurierte Datei nicht gefunden wird

---

## Bekannte Einschränkungen

- Diagramme und eingebettete Bilder werden nicht dargestellt
- Bedingte Formatierungen werden nicht ausgewertet
- Passwortgeschützte Dateien werden nicht unterstützt
- Formeln werden nicht neu berechnet (`data_only=True`) — es werden die zuletzt in Excel gespeicherten Werte angezeigt

---

## Performance-Hinweise

- Durch `@st.cache_resource` wird die Datei nur beim ersten Aufruf geparst; alle weiteren Sitzungen greifen auf den Cache zu
- Für Dateien über 50 MB empfiehlt sich `server.maxUploadSize = 200` in der Streamlit-Konfiguration
- Bei sehr vielen Blättern (> 50) kann die Sidebar-Darstellung lang werden — ggf. Filterung nach Kategorie ergänzen