# EBA ITS DPM Viewer – Projektkontext & Neustart-Prompt

> Diese Datei dient als vollständiger Einstiegspunkt für einen neuen Chat oder einen neuen Entwicklungsansatz.  
> Sie enthält den gesamten bisherigen Entwicklungsstand als strukturierten Prompt.

---

## PROMPT FÜR NEUEN CHAT

---

Du unterstützt mich bei der Weiterentwicklung einer bestehenden Python-Webanwendung: dem **EBA ITS DPM Viewer**.  
Die App ist bereits funktionsfähig und in Produktion-nahem Zustand. Ich beschreibe dir unten den vollständigen Stand.

---

## 1. Projektziel

Eine **Streamlit-Webanwendung**, die EBA-ITS-Excel-Dateien (European Banking Authority – Implementing Technical Standards) originalgetreu im Browser darstellt. Die App richtet sich an Mitarbeitende in der deutschen Bankenregulierung und Meldewesen-Teams (COREP, FINREP, Leverage Ratio, etc.).

**Kernfunktionen:**
- Excel-Blätter werden pixeltreu als HTML-Tabellen gerendert (Farben, Borders, Merges, Schrift)
- Leere Eingabezellen werden automatisch mit **DPM-Koordinaten** beschriftet (z. B. `C 01.00_0010_0020`)
- Zeilen- und Spaltenbeschriftungen erhalten ebenfalls Label-Keys (z. B. `C 01.00_row_0010`)
- Koordinaten + manuelle Annotationstexte werden in einer **`coordinates.csv`** verwaltet
- Annotationen erscheinen als **CSS-only-Tooltips** beim Mouse-over (kein JavaScript nötig)
- Navigation über eine **Sidebar** mit gruppierten Einträgen aus dem Excel-Index-Blatt
- DOCX-Dateien mit Erläuterungstexten werden automatisch geladen und als aufklappbare Expander über den Tabellen angezeigt (Graceful Degradation: ohne DOCX-Dateien bleibt der Status-Quo erhalten)

---

## 2. Technischer Stack

| Komponente     | Technologie                              |
|----------------|------------------------------------------|
| UI-Framework   | **Streamlit** (Python)                   |
| Excel-Parsing  | **openpyxl**                             |
| DOCX-Parsing   | **python-docx** (oder plain-text Fallback) |
| Styling        | Externes CSS (`eba_styles.css`), inline HTML in Streamlit |
| Datenhaltung   | `coordinates.csv` (manuell editierbar)   |
| Deployment     | Ziel: webgo (DE), Alternativ: Streamlit Community Cloud |

**`requirements.txt`:**
```
streamlit>=1.32
openpyxl>=3.1
python-docx>=1.1
```

---

## 3. Modulstruktur

```
eba_viewer/
├── app.py                  # Streamlit-Hauptanwendung
├── excel_parser.py         # Excel-Parsing + DPM-Koordinatenerkennung
├── renderer.py             # HTML-Tabellenrendering
├── export_coordinates.py   # CSV-Export aller Koordinaten
├── docx_annotations.py     # DOCX/Text-Erläuterungen → Sheet-Mapping + Streamlit-Rendering
├── eba_styles.css          # Zentrale CSS-Regeln (Tooltips, Zellklassen, SR-Corporate-Design)
├── requirements.txt
├── README.md
└── data/
    ├── *.XLSX              # EBA-Excel-Datei (z. B. C_2024_8389_F1_ANNEX_DE_V1_P1.XLSX)
    ├── coordinates.csv     # Auto-generiert + manuell pflegbar
    └── *.DOCX              # Optional: Erläuterungstexte (können auch plain-text sein)
```

---

## 4. Datenmodell (Datenklassen in `excel_parser.py`)

```python
@dataclass
class CellStyle:
    bg_color: str | None      # 6-stelliger Hex ohne #
    fg_color: str | None
    bold: bool
    italic: bool
    h_align: str              # "left" | "center" | "right"
    v_align: str              # "top" | "middle" | "bottom"
    wrap_text: bool
    border_top/bottom/left/right: str | None   # CSS-String, z. B. "1px solid #BDBDBD"
    number_format: str | None

@dataclass
class CellData:
    value: Any
    display_value: str
    style: CellStyle
    rowspan: int              # für merged cells
    colspan: int
    is_merged_hidden: bool    # versteckte Merge-Folge-Zellen
    coordinate: str | None   # z. B. "C 01.00_0010_0020"
    annotation: str | None   # aus coordinates.csv geladen
    label_key: str | None     # z. B. "C 01.00_row_0010" oder "C 01.00_col_0020"

@dataclass
class SheetData:
    name: str
    rows: list[list[CellData]]
    col_widths: list[float]
    row_heights: list[float]
```

---

## 5. DPM-Koordinatensystem

**Koordinatenformat:** `{SheetName}_{RowCode}_{ColCode}`  
Beispiel: `C 01.00_0010_0020`

**Label-Keys:**
- Spaltenheader: `C 01.00_col_0020`
- Zeilenheader: `C 01.00_row_0010`

**Erkennungslogik (in `_build_coordinates` in `excel_parser.py`):**
1. **Row-Code-Spalte**: Scan der ersten 5 Spalten → die Spalte mit den meisten 4-stelligen Zahlencodes (`^\d{4}$`) wird als Row-Code-Spalte gewählt
2. **Col-Header-Zeilen**: Dynamischer Scan über alle Zeilen (nicht nur erste N) – Zeilen mit ≥ 2 Zellen mit 4-stelligen Codes gelten als Column-Header
3. **Fallback**: Wenn keine Col-Header gefunden werden, werden synthetische Codes `0010`, `0020`, … vergeben
4. **Input-Zellen-Erkennung**: Eine Zelle gilt als Eingabezelle wenn: kein Hintergrund UND kein Textinhalt UND nicht `is_merged_hidden`

**Besonderheiten:**
- Excel ARGB-Farben mit Alpha=`00` müssen trotzdem gerendert werden → `ignore_alpha=True` beim Parsen
- Near-white-Farben (alle RGB-Kanäle ≥ 248, z. B. `FCFCFC`) werden als "keine Füllung" behandelt
- Theme-Farben werden aus dem eingebetteten XML aufgelöst (mit Tint/Shade-Berechnung)

---

## 6. CSV-Format (`coordinates.csv`)

```
key,type,sheet,row_code,col_code,annotation
C 01.00_0010_0020,cell,C 01.00,0010,0020,
C 01.00_col_0020,col_label,C 01.00,,0020,Eigenmittelquote nach CRR
C 01.00_row_0010,row_label,C 01.00,0010,,Hartes Kernkapital
```

Felder:
- `key` – eindeutiger Schlüssel (Koordinate oder Label-Key)
- `type` – `"cell"` | `"col_label"` | `"row_label"`
- `sheet` – Blattname
- `row_code` / `col_code` – 4-stellige Codes (leer wenn nicht zutreffend)
- `annotation` – manuell pflegbarer Freitext; wird beim Re-Export erhalten

**Wichtig:** Beim Neustart der App wird `coordinates.csv` automatisch neu generiert, wobei bestehende Annotationen erhalten bleiben (Merge via `existing_annotations`-Dict).

---

## 7. HTML-Rendering (`renderer.py`)

Drei Zelltypen werden unterschiedlich gerendert:

**A) Eingabezellen** (mit `coordinate`):
- Sehr heller Cremeweiß-Hintergrund (`#FFF8F8`) – gesteuert über CSS-Klasse `.eba-coord-cell`
- Koordinate als kleines Label oben in der Zelle – gesteuert über CSS-Klasse `.eba-coord-label` (gedämpftes Rot `#B03030`, 7.5pt)
- CSS-Tooltip mit Koordinate + optionalem Annotationstext
- Roter Punkt-Badge wenn Annotation vorhanden

**B) Label-Zellen** (mit `label_key`):
- Originaler Zellinhalt + Tooltip
- Badge wenn Annotation vorhanden

**C) Normale Zellen:**
- Nur `title`-Attribut mit Zelltext

**CSS-Klassen in `eba_styles.css`:**
- `.eba-coord-cell` – Eingabezelle (position:relative, overflow:visible, background:#FFF8F8)
- `.eba-coord-label` – Koordinatentext-Label in Eingabezellen (color:#B03030, font-size:7.5pt)
- `.eba-label-cell` – Beschriftungszelle
- `.eba-badge` – roter Punkt (7×7px, position:absolute, top-right, #CC0000)
- `.eba-tooltip` – Tooltip-Container (display:none, wird per :hover eingeblendet)
- `.eba-tooltip-coord` – Koordinatenzeile im Tooltip (hellrot #F08080, 8pt)
- `.eba-tooltip-divider` – Trennlinie
- `.eba-tooltip-text` – Annotationstext (fast-weiß #F0F0F0, 9pt)

**Wichtig:** Farben und Label-Styling sind vollständig in `eba_styles.css` definiert – kein Inline-Styling mehr für Koordinatenfarben in `renderer.py`.

---

## 8. Navigation & Sidebar (`app.py`)

**Indexstruktur:**
- Das Excel-Blatt `"Index"` wird geparst, um Gruppen und Template-Einträge zu erkennen
- Section-Erkennung via `SECTION_MAP` (Keywords: COREP, FINREP, LEVERA, LIQUID, IRRBB, …)
- Zusätzlich `ANNEX_KEYWORDS` für deutsche Annex-Abschnitte (IP-VERLUSTE, VERSCHULDUNG)
- Subgruppen innerhalb von Abschnitten werden erkannt und als Gruppe angezeigt
- `_find_sheet()` normalisiert Non-Breaking-Spaces und macht Substring-Matching für kombinierte Blätter (z. B. `LR6.1` → `LR6.1, LR6.2`)
- Duplikate innerhalb einer Gruppe werden herausgefiltert

**Session State:**
- `st.session_state.current_sheet` speichert das aktive Blatt
- Navigation via `go_to(sheet_name)` → löst Streamlit-Rerun aus

**Sidebar-Layout:**
- Dunkelgrau-Anthrazit (`#2B2B2B`), Gruppen-Labels in Hellrot (`#E84040`)
- Aktives Blatt: linker roter Balken (`#CC0000`), Hintergrund-Tint `rgba(204,0,0,0.15)`
- Navigation-Labels zeigen **nur den Sheet-Namen** (kein Template-Code-Prefix)

**`load_workbook()` gibt drei Werte zurück:**
```python
sheets, groups, sheet_annotations = load_workbook(EXCEL_PATH)
```

---

## 9. DOCX-Annotationsmodul (`docx_annotations.py`) – vollständig integriert

**Öffentliche API:**

```python
get_sheet_annotations(excel_path: str) -> dict[str, str]
```
- Durchsucht den Ordner der Excel-Datei nach `*.DOCX`-Dateien
- Erkennt automatisch ob es sich um Text/Markdown (wie FISMA-Quelldateien) oder echte DOCX-Dateien handelt (Magic-Bytes-Check: `PK`-Header)
- Parst Abschnitte anhand der Template-Codes (z. B. `C 34.01`, `C 36.00`) via Regex
- Ordnet Abschnitte über das Index-Blatt den richtigen Tabellenblattnamen zu
- Gibt `{}` zurück wenn keine Dateien vorhanden → Status-Quo bleibt vollständig erhalten

```python
render_annotation(annotation_html: str) -> None
```
- Zeigt den Erläuterungstext als aufklappbaren Expander (`📄 Erläuterungen`) direkt **vor** der Tabelle
- CSS wird einmalig pro Session injiziert (via `st.session_state["_doc_annotation_css_injected"]`)
- Leerer String → kein Output (Graceful Degradation)

**Integration in `app.py`:**
```python
from docx_annotations import get_sheet_annotations, render_annotation
# in load_workbook(): sheet_annotations = get_sheet_annotations(path)
# in render_sheet(): render_annotation(sheet_annotations.get(sheet.name, ""))
```

**CSS-Klassen (SR-Design):**
- `.doc-annotation-container` – Container mit rotem linken Rand (`#CC0000`), Cremeweiß-Hintergrund
- `.doc-annotation-h3` / `.doc-annotation-h4` – Überschriften in SR-Farben
- `.doc-annotation-numbered` – nummerierte Absätze (Nummer in `#CC0000`)
- `.doc-annotation-table` – Tabellen mit SR-Tabellenheader (`#FDECEA`)

---

## 10. Bekannte gelöste Probleme (lessons learned)

| Problem | Lösung |
|---|---|
| Alpha=00 in Excel ARGB → Farbe wurde ignoriert | `ignore_alpha=True` in `_resolve_color()` |
| Near-white Theme-Tints (`FCFCFC`) als "gefüllt" erkannt | `_is_near_white()` mit Threshold 248 |
| Col-Header nur in ersten N Zeilen gesucht | Dynamischer Scan über alle Zeilen in `_build_coordinates` |
| Deutsche Annex-Abschnitte nicht erkannt | `ANNEX_KEYWORDS`-Dict + Substring-Matching |
| Non-Breaking-Spaces in Sheet-Namen | Normalisierung in `_find_sheet()` |
| Template-Code im Sidebar-Label | Entfernt; nur `short_name` wird angezeigt |
| Sheets mit Row-Codes aber ohne Col-Header | Synthetische Codes `0010`, `0020`, … |
| `.DOCX`-Dateien als plain-text getarnt | Magic-Bytes-Check (`PK`-Header) |
| Koordinaten-Duplikate im CSV | `seen`-Set in `export_coordinates()` |
| Koordinatenfarbe + Zell-Hintergrund inline in renderer.py | Ausgelagert in CSS-Klassen `.eba-coord-cell` und `.eba-coord-label` |

---

## 11. Offene Punkte / nächste Schritte

- [ ] **Deployment auf webgo**: VPS mit systemd-Service + nginx Reverse Proxy, oder Streamlit Community Cloud
- [ ] **Annotations-Editor**: Direktbearbeitung von `coordinates.csv` im Browser
- [ ] **Suchfunktion**: Suche über Koordinaten oder Annotationstexte
- [ ] **Performance**: Bei großen Dateien ggf. Sheet-lazy-loading prüfen

---

## 12. Typische Entwicklungsworkflow

1. Diagnose-Skript schreiben (plain Python, kein Streamlit) um ein Problem zu isolieren
2. Fix in der zuständigen Datei implementieren (`excel_parser.py`, `renderer.py`, `app.py`)
3. Mit Diagnose-Skript verifizieren
4. Dann erst in Streamlit testen

**Prinzip:** Graceful Degradation – neue Features dürfen existierendes Verhalten nicht brechen.

---

## 13. Konventionen

- Sprache: Kommentare und Variablennamen auf **Englisch**, Kommunikation mit mir auf **Deutsch**
- Immer erst Diagnose → dann Fix (keine blinden Änderungen)
- CSS zentral in `eba_styles.css`, nicht inline im Python-Code (Ausnahme: einmalige Streamlit-Overrides in `app.py`)
- Keine JavaScript-Abhängigkeiten – alles CSS-only wo möglich

---

## 14. Design – Corporate Identity Sparkassen Rating und Risikosysteme GmbH (SR)

**Status: vollständig implementiert** (Stand 2026-04-19)

Die App verwendet das SR-Corporate-Design mit **Sparkassen-Dunkelrot** (`#CC0000`) als Primärfarbe anstelle des früheren Blaus. Das Grundprinzip (dunkle Sidebar, helle Hauptfläche, farbige Akzente für Koordinaten) bleibt erhalten.

---

### 14.1 Markenkontext & Tonalität

| Merkmal | Beschreibung |
|---|---|
| **Unternehmen** | Sparkassen Rating und Risikosysteme GmbH (SR), Berlin |
| **Zielgruppe** | Fachexperten im Meldewesen, Risikomanagement, Banksteuerung |
| **Tonalität** | Professionell, präzise, sachlich – kein Consumer-Design |
| **Markenfamilie** | Sparkassen-Finanzgruppe (größte Finanzgruppe Europas) |
| **Claim** | „Wir machen Zukunft berechenbar." |

---

### 14.2 Farbpalette (implementiert)

#### Primärfarben
| Rolle | Hex | Verwendung |
|---|---|---|
| **Brand Primary** | `#CC0000` | Hauptakzent, aktive Nav-Elemente, Tooltip-Rahmen, Badges |
| **Brand Dark** | `#990000` | Hover auf Primärfarbe, Stat-Chip-Text |
| **Brand Light** | `#FDECEA` | Stat-Chip-Hintergrund, Tabellen-Header in DOCX-Annotations |

#### Neutralfarben
| Rolle | Hex | Verwendung |
|---|---|---|
| **Sidebar BG** | `#2B2B2B` | Sidebar-Hintergrund |
| **Sidebar Text** | `#E8E8E8` | Navigationstext |
| **Sidebar Akzent** | `#E84040` | Gruppen-Labels |
| **Content BG** | `#FFFFFF` | Hauptinhaltsfläche |
| **Seiten-BG** | `#F5F5F5` | Sheet-Titelleiste |
| **Border** | `#DDDDDD` | Tabellenrahmen, Trennlinien |
| **Text Primary** | `#1A1A1A` | Fließtext, Überschriften |
| **Text Secondary** | `#666666` | Breadcrumbs, Metaangaben |

#### Akzentfarben (funktional)
| Rolle | Hex | Verwendung |
|---|---|---|
| **Koordinaten-Zelle BG** | `#FFF8F8` | Eingabezellen (`.eba-coord-cell`) |
| **Koordinaten-Label** | `#B03030` | Koordinatentext (`.eba-coord-label`) |
| **Tooltip BG** | `#2B2B2B` | Tooltip-Hintergrund |
| **Tooltip Border** | `#CC0000` | Tooltip-Rahmen |
| **Tooltip Coord** | `#F08080` | Koordinatenzeile im Tooltip |
| **Annotation Badge** | `#CC0000` | Punkt-Badge bei vorhandener Annotation |
| **Annotation Text** | `#F0F0F0` | Annotationstext im Tooltip |

---

### 14.3 Typografie

| Element | Font-Stack | Größe | Gewicht |
|---|---|---|---|
| App-weite Basis | `'Sparkasse Head', 'Segoe UI', 'Inter', system-ui, sans-serif` | — | — |
| Tabellen-Inhalt | `'Segoe UI', 'Calibri', system-ui, sans-serif` | `10pt` | normal |
| Sidebar-Gruppe | Gleicher Stack | `0.7rem` | `700`, Großbuchstaben |
| Sidebar-Nav | Gleicher Stack | `0.82rem` | normal |
| Koordinaten-Label | Gleicher Stack | `7.5pt` | normal |
| Tooltip-Coord | Gleicher Stack | `8pt` | `600` |
| Tooltip-Text | Gleicher Stack | `9pt` | normal |

> „Sparkasse Head" ist die offizielle Hausschrift der Sparkassen-Finanzgruppe. Sie ist als erste Option im Font-Stack genannt; `'Segoe UI'` dient als systemweiter Fallback.

---

### 14.4 Was sich NICHT ändert

Folgende Aspekte bleiben vom Design-Wechsel unberührt:

- Die **Koordinaten-Logik** (DPM-Koordinatenformat, CSV-Struktur, Parsing)
- Das **HTML-Tabellenrendering** (Struktur, rowspan/colspan, Merge-Handling)
- Die **CSS-Klassen-Namen** (`.eba-coord-cell`, `.eba-label-cell`, `.eba-tooltip`, etc.)
- Die **Streamlit-Architektur** (`session_state`, `@st.cache_resource`, etc.)
- Die **Funktionslogik** aller Python-Module

---

*Erstellt: 2026-04-19 | Zuletzt aktualisiert: 2026-04-19 | Projektstand: SR-Design implementiert, DOCX-Integration aktiv*
