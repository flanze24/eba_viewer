# EBA ITS DPM Viewer

Eine Streamlit-Webanwendung zur originalgetreuen Darstellung von EBA-ITS-Excel-Dateien nach dem Data Point Model (DPM). Leere Eingabezellen werden automatisch mit DPM-Koordinaten beschriftet. Zeilen- und Spaltenbeschriftungen sowie Eingabezellen können manuell mit Freitextannotationen versehen werden, die als interaktive Tooltips beim Mouse-over erscheinen. Erläuterungstexte aus Word-Dokumenten werden automatisch als aufklappbare Bereiche über den Tabellen angezeigt. Das Design folgt der Corporate Identity der Sparkassen Rating und Risikosysteme GmbH (SR).

---

## Projektstruktur

```
eba_viewer/
├── app.py                   # Streamlit-Hauptanwendung
├── excel_parser.py          # Excel-Verarbeitung und Koordinaten-Erkennung
├── renderer.py              # HTML-Rendering der Tabellen
├── export_coordinates.py    # CSV-Export aller DPM-Koordinaten und Labels
├── docx_annotations.py      # DOCX/Text-Erläuterungen → Sheet-Mapping + Rendering
├── eba_styles.css           # Alle CSS-Regeln (zentral anpassbar, SR-Corporate-Design)
├── requirements.txt         # Python-Abhängigkeiten
├── README.md
└── data/
    ├── eba_template.xlsx    # Excel-Datei hier ablegen
    ├── coordinates.csv      # wird automatisch generiert / manuell gepflegt
    └── *.DOCX               # Optional: Erläuterungstexte (DOCX oder plain-text)
```

---

## Schnellstart

```bash
# 1. Abhängigkeiten installieren
pip install -r requirements.txt

# 2. Excel-Datei ablegen
mkdir -p data
cp /pfad/zur/datei.xlsx data/eba_template.xlsx

# 3. Optional: Erläuterungstexte ablegen
cp /pfad/zur/erlaeuterung.docx data/

# 4. App starten
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
Die App sucht standardmäßig nach `data/C_2024_8389_F1_ANNEX_DE_V1_P1_3682615.XLSX` relativ zum `app.py`-Verzeichnis. Den Dateinamen direkt in `app.py` (Zeile `EXCEL_PATH`) anpassen.

**Option C – Direktanpassung in `app.py`**
```python
EXCEL_PATH = "/vollständiger/pfad/zur/datei.xlsx"
```

Der Name des Index-Blatts ist ebenfalls konfigurierbar:
```python
INDEX_SHEET = "Index"
```

Falls die Datei beim Start nicht gefunden wird, erscheint eine Fehlerseite mit einem temporären Upload-Formular.

---

## Features

### Navigation
- **Sidebar** mit gruppierten Tabellenblättern (Gruppen aus dem Index-Blatt: COREP, FINREP, Leverage Ratio, etc.)
- **Index-Seite** als Startseite mit Tabellenansicht des Index-Blatts und Direktlink-Karten nach Themenbereich
- **Rücknavigation** von jedem Blatt zum Index über einen Button oben links
- **Breadcrumb** zeigt den Themenbereich des aktuellen Blatts

### Darstellung
- Originalgetreue Wiedergabe aller Zellformate: Hintergrundfarben, Schriftfarbe, Fett/Kursiv, Textausrichtung, Rahmen, Merge-Bereiche (rowspan/colspan)
- Korrekte Auflösung aller Excel-Farbformate: ARGB-RGB, Theme-Farben mit Tint/Shade, Indexed Colors (Legacy)
- Weiße, schwarze und near-white Füllungen (alle RGB-Kanäle ≥ 248) werden als „keine Füllung" behandelt
- Vollständig leere Zeilen und Spalten werden automatisch entfernt

### Einheitliche Typografie
- Eine einzige Schriftgröße (`10pt`) für alle Tabelleninhalte
- Font-Stack: `'Sparkasse Head', 'Segoe UI', 'Inter', 'Calibri', system-ui, sans-serif`

### DPM-Koordinaten

Leere, ungefärbte Eingabezellen erhalten automatisch eine Koordinatenbezeichnung nach dem DPM-Schema.

**Format:** `<Blattname>_<Zeilencode>_<Spaltencode>`  
Beispiel: `C 01.00_0010_0020`

**Erkennungslogik:**
- *Spaltencodes:* Dynamischer Scan über alle Zeilen. Zeilen mit mindestens zwei vierstelligen Codes (`0010`, `0020` …) gelten als Column-Header.
- *Zeilencodes:* Die ersten 5 Spalten werden verglichen. Die Spalte mit den meisten vierstelligen Codes wird als Zeilen-Header-Spalte verwendet.
- *Fallback:* Werden keine Column-Header gefunden, werden synthetische Codes `0010`, `0020`, … vergeben.
- Eine Zelle erhält eine Koordinate **nur wenn** sie weder Hintergrundfüllung noch Textinhalt hat.

Die Koordinate erscheint als kleines Rot-Label (`#B03030`) oben in der Zelle sowie im Tooltip.

### Annotationen & Tooltips

Jede Eingabezelle sowie jede Zeilen- und Spaltenbeschriftung kann mit einem Freitext-Kommentar versehen werden. Die Annotation erscheint beim Mouse-over als Sprechblase.

**Visueller Hinweis:** Zellen mit Annotation tragen einen kleinen roten Punkt (●) in der oberen rechten Ecke.

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

### Erläuterungstexte (DOCX-Annotationen)

Legt man `.DOCX`-Dateien in den `data/`-Ordner, werden die enthaltenen Erläuterungstexte automatisch den passenden Tabellenblättern zugeordnet und als aufklappbarer Bereich **„📄 Erläuterungen"** direkt über der Tabelle angezeigt.

**Unterstützte Formate:**
- Echte Word-Dokumente (`.docx`) – Paragraphen und Tabellen werden ausgewertet
- Als `.DOCX` gespeicherte Plain-Text/Markdown-Dateien (wie FISMA-Quelldokumente)

**Zuordnung:** Abschnitte im Dokument werden anhand der Template-Codes (`C 34.01`, `C 36.00` etc.) erkannt und über das Index-Blatt den Tabellenblattnamen zugeordnet.

**Graceful Degradation:** Sind keine DOCX-Dateien vorhanden, bleibt die App vollständig funktionsfähig – kein Expander, kein Fehler.

### CSV-Export der Koordinaten

Beim Start der App wird `data/coordinates.csv` automatisch erzeugt bzw. aktualisiert.

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

---

## Styling anpassen (`eba_styles.css`)

Alle visuellen Parameter der Tooltips und annotierbaren Zellen sind in `eba_styles.css` zentral zusammengefasst und kommentiert. Die Datei wird beim Rendern jedes Blatts eingelesen – Änderungen wirken sich nach einem App-Neustart sofort aus.

| CSS-Klasse | Steuert |
|---|---|
| `.eba-coord-cell` | Eingabezellen mit DPM-Koordinate (position, overflow, Hintergrund `#FFF8F8`) |
| `.eba-coord-label` | Koordinatentext-Label in Eingabezellen (Farbe, Schriftgröße) |
| `.eba-label-cell` | Zeilen-/Spaltenbeschriftungen mit 4-stelligem Code |
| `.eba-badge` | Roter Hinweis-Punkt bei vorhandener Annotation (`#CC0000`) |
| `.eba-tooltip` | Sprechblase: Position, Größe, Farbe, Schatten, Schrift |
| `.eba-tooltip-coord` | Technischer Key im Tooltip (hellrot `#F08080`) |
| `.eba-tooltip-divider` | Trennlinie zwischen Key und Annotationstext |
| `.eba-tooltip-text` | Annotationstext im Tooltip (fast-weiß `#F0F0F0`) |

---

## Modulbeschreibung

### `app.py`
Streamlit-Hauptanwendung. Enthält Seitenkonfiguration, globales CSS (SR-Corporate-Design), Session-State-Routing, Sidebar mit gruppierten Einträgen, Index-Seite und Blatt-Ansicht. Lädt Workbook-Daten über `@st.cache_resource` (einmaliges Parsen pro Anwendungsstart), triggert den CSV-Export, wendet Annotationen aus der CSV auf die geparsten Zellen an (`_apply_annotations`) und lädt DOCX-Erläuterungen (`get_sheet_annotations`).

```python
sheets, groups, sheet_annotations = load_workbook(EXCEL_PATH)
```

### `excel_parser.py`
Liest Excel-Dateien mit `openpyxl` und gibt strukturierte `SheetData`-Objekte zurück.

| Funktion / Klasse | Aufgabe |
|---|---|
| `parse_workbook(path)` | Lädt die Arbeitsmappe, filtert ausgeblendete Blätter, parst alle sichtbaren Blätter |
| `_parse_sheet(ws, theme_colors)` | Liest Zellen, Merge-Bereiche, Spaltenbreiten, Zeilenhöhen; entfernt leere Zeilen/Spalten |
| `_extract_style(cell, theme_colors)` | Extrahiert Füllfarbe, Schriftformat, Ausrichtung, Rahmen, Zahlenformat |
| `_resolve_color(...)` | Konvertiert ARGB-, Theme- und Indexed-Farben in 6-stellige Hex-Strings |
| `_is_near_white(hex6)` | `True` wenn alle RGB-Kanäle ≥ 248 → wird als „keine Füllung" behandelt |
| `_build_coordinates(sheet)` | Erkennt DPM-Koordinaten; setzt `cell.coordinate` für Eingabezellen und `cell.label_key` für Zeilen-/Spaltenköpfe |
| `CellData` | Datenklasse pro Zelle: Wert, Anzeigetext, Style, Rowspan/Colspan, `coordinate`, `label_key`, `annotation` |
| `SheetData` | Datenklasse pro Blatt: Zellen-Matrix, Spaltenbreiten, Zeilenhöhen |

### `renderer.py`
Wandelt `SheetData`-Objekte in HTML-`<table>`-Strings um. Unterscheidet drei Zelltypen beim Rendering:

| Typ | Bedingung | Rendering |
|---|---|---|
| Eingabezelle | `cell.coordinate` gesetzt | CSS-Klassen `eba-coord-cell` + `eba-coord-label`, Tooltip |
| Beschriftungszelle | `cell.label_key` gesetzt | CSS-Klasse `eba-label-cell`, Tooltip bei Hover |
| Standardzelle | keines von beidem | Normales Styling, kein Tooltip |

CSS wird aus `eba_styles.css` geladen und einmalig pro Tabelle als `<style>`-Block eingefügt. Fällt die CSS-Datei weg, bleibt der Viewer voll funktionsfähig.

### `docx_annotations.py`
Liest `.DOCX`-Dateien aus dem `data/`-Ordner und ordnet Textabschnitte den Tabellenblättern zu.

| Funktion | Aufgabe |
|---|---|
| `get_sheet_annotations(excel_path)` | Durchsucht `data/` nach DOCX-Dateien, parst Abschnitte anhand Template-Codes, gibt `{sheet_name: html_text}` zurück |
| `render_annotation(html)` | Rendert den Erläuterungstext als aufklappbaren Expander in Streamlit |
| `_is_text_file(path)` | Magic-Bytes-Check: unterscheidet echte DOCX (`PK`-Header) von Plain-Text |
| `_parse_sections(content)` | Extrahiert Abschnitte anhand Template-Code-Regex (z. B. `C 34.01`) |
| `_build_code_to_sheet_map(excel_path)` | Liest Index-Blatt und erstellt Mapping Template-Code → Sheet-Name |
| `_markdown_to_html(md)` | Konvertiert Markdown-Text (Überschriften, Listen, Tabellen) in HTML |

### `export_coordinates.py`
Iteriert über alle geparsten Blätter und Zellen, sammelt Eingabe-Koordinaten und Label-Keys und schreibt sie als CSV. Bestehende Annotationen werden vor dem Überschreiben eingelesen und in die neue Datei übertragen.

### `eba_styles.css`
Zentrale Stylesheet-Datei im SR-Corporate-Design (Sparkassen-Rot `#CC0000`). Jede Regel ist kommentiert. Kann ohne Python-Kenntnisse angepasst werden.

---

## Bekannte Einschränkungen

- Diagramme und eingebettete Bilder werden nicht dargestellt
- Bedingte Formatierungen werden nicht ausgewertet
- Passwortgeschützte Dateien werden nicht unterstützt
- Formeln werden nicht neu berechnet (`data_only=True`) — es werden die zuletzt in Excel gespeicherten Werte angezeigt

---

## Performance-Hinweise

- `@st.cache_resource` stellt sicher, dass Parsen, CSV-Export, Annotation-Mapping und DOCX-Parsing nur einmal pro Anwendungsstart ausgeführt werden
- Für Dateien über 50 MB: `server.maxUploadSize = 200` in der Streamlit-Konfiguration setzen
- Bei sehr vielen Blättern (> 50) kann eine Filterung der Sidebar nach Kategorie sinnvoll sein
