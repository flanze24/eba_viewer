# EBA ITS DPM Viewer

Eine Streamlit-Webanwendung zur originalgetreuen Darstellung von EBA-ITS-Excel-Dateien (DPM – Data Point Model).

## Projektstruktur

```
eba_viewer/
├── app.py              # Streamlit-Hauptanwendung (UI, Navigation, Routing)
├── excel_parser.py     # Excel-Verarbeitung: Daten + Formatierung (openpyxl)
├── renderer.py         # HTML-Rendering der formatierten Tabellen
├── requirements.txt    # Python-Abhängigkeiten
├── README.md
└── data/
    └── eba_template.xlsx   # ← Excel-Datei hier ablegen
```

## Wichtigste Komponenten

| Modul | Aufgabe |
|-------|---------|
| `excel_parser.py` | Liest alle Tabellenblätter, extrahiert Zellinhalte, Farben (inkl. Theme-Tints), Schriftarten, Rahmen, Merge-Bereiche und Spaltenbreiten. Leere Zeilen/Spalten werden entfernt. |
| `renderer.py` | Wandelt `SheetData`-Objekte in formatierte HTML-`<table>`-Strings um. Unterstützt Rowspan/Colspan, Hintergrundfarben, Schriftgrößen, Textausrichtung und Rahmen. |
| `app.py` | Streamlit-App mit Sidebar-Navigation, Index-Seite, Blatt-Ansicht und Fehlerbehandlung. Verwendet `@st.cache_resource` für performantes Laden großer Dateien. |

## Installation

```bash
# 1. Abhängigkeiten installieren
pip install -r requirements.txt

# 2. Excel-Datei ablegen
mkdir -p data
cp /pfad/zur/eba_datei.xlsx data/eba_template.xlsx

# 3. App starten
streamlit run app.py
```

## Konfiguration

### Option A – Datei im `data/`-Ordner ablegen
Standardmäßig sucht die App nach `data/eba_template.xlsx` relativ zum `app.py`-Verzeichnis.

### Option B – Umgebungsvariable
```bash
export EBA_EXCEL_PATH="/vollständiger/pfad/zur/datei.xlsx"
streamlit run app.py
```

### Option C – Direkt in `app.py` anpassen
```python
EXCEL_PATH = "/vollständiger/pfad/zur/datei.xlsx"
```

### Blattname des Index
Falls das Index-Blatt anders heißt:
```python
INDEX_SHEET = "Index"   # ← hier anpassen
```

## Features

- **Automatisches Laden** der Excel-Datei beim Start (kein Upload nötig)
- **Index-Seite** mit Übersicht und Direktlinks zu allen Tabellenblättern
- **Originalgetreue Darstellung**: Hintergrundfarben, Theme-Tints, Schriftarten, -größen, -farben, Rahmen, Merge-Bereiche
- **Sidebar-Navigation** mit allen Tabellenblättern
- **Rücklink** von jedem Blatt zum Index
- **Leere Zeilen/Spalten** werden automatisch entfernt
- **Scrollbare Tabellen** mit fixierter Höhe
- **Performance-Caching** via `@st.cache_resource` – Datei wird nur einmal geparst
- **Fallback-Upload** wenn Datei nicht am konfigurierten Pfad gefunden wird

## Performance-Hinweise für sehr große Dateien

- `data_only=True` beim Laden verhindert die Auswertung von Formeln (schneller)
- Der `@st.cache_resource`-Decorator speichert das geparste Ergebnis über alle Sitzungen
- Für Dateien >50 MB: Streamlit-Config anpassen: `server.maxUploadSize = 200`
- Bei >100 Tabellenblättern empfiehlt sich lazy loading (nur aktuell sichtbares Blatt parsen)

## Bekannte Einschränkungen

- Diagramme und Bilder innerhalb der Excel-Datei werden nicht dargestellt
- Sehr komplexe bedingte Formatierungen werden ggf. nicht vollständig abgebildet
- Passwortgeschützte Dateien werden nicht unterstützt
