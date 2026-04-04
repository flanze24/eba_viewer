"""
app.py – EBA ITS DPM Viewer (Streamlit)
========================================
Lädt eine Excel-Datei automatisch, stellt alle Tabellenblätter originalgetreu dar
und bietet eine Navigation über das "Index"-Blatt.

Konfiguration:
  Setze EXCEL_PATH auf den vollständigen Pfad zur Excel-Datei. 
"""

from __future__ import annotations
import csv
import os
import sys
from pathlib import Path

import streamlit as st

# ── Konfiguration ──────────────────────────────────────────────────────────────
# Pfad zur Excel-Datei – kann auch als Umgebungsvariable gesetzt werden:
#   export EBA_EXCEL_PATH="/pfad/zur/datei.xlsx"
EXCEL_PATH = os.environ.get(
    "EBA_EXCEL_PATH",
    str(Path(__file__).parent / "data" / "C_2024_8389_1_ANNEX_EN_V3_P1_3682615.XLSX"),
)
INDEX_SHEET = "Index"   # Name des Index-Blatts (case-sensitive)
APP_TITLE = "EBA ITS DPM Viewer"
# ──────────────────────────────────────────────────────────────────────────────

# Add current dir to path so modules resolve correctly
sys.path.insert(0, str(Path(__file__).parent))

from excel_parser import parse_workbook, SheetData
from renderer import render_sheet_html


# ── Page setup ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Global CSS
st.markdown("""
<style>
/* App-wide font – single unified stack for all UI and table text */
html, body, [class*="css"], table, td, th {
    font-family: 'Segoe UI', 'Inter', 'Calibri', system-ui, sans-serif !important;
}

/* Single authoritative font-size for ALL table content – no exceptions */
table, table td, table th, table * {
    font-size: 10pt !important;
    line-height: 1.45 !important;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #1B2A4A;
    color: #E8EFF8;
}
section[data-testid="stSidebar"] .stButton button {
    width: 100%;
    text-align: left;
    background: transparent;
    color: #C8D8F0;
    border: none;
    border-left: 3px solid transparent;
    border-radius: 0;
    padding: 6px 12px;
    font-size: 0.85rem;
    margin: 1px 0;
    transition: all 0.15s;
}
section[data-testid="stSidebar"] .stButton button:hover {
    background: rgba(255,255,255,0.08);
    border-left-color: #4D9FEC;
    color: #ffffff;
}
section[data-testid="stSidebar"] h1 {
    color: #E8EFF8;
    font-size: 1.1rem;
}

/* Main header */
.eba-header {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 10px 0 16px 0;
    border-bottom: 2px solid #2B5FA8;
    margin-bottom: 16px;
}
.eba-header h1 {
    margin: 0;
    font-size: 1.5rem;
    color: #1B2A4A;
    font-weight: 700;
}
.eba-header .badge {
    background: #2B5FA8;
    color: white;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 0.5px;
}

/* Sheet title bar */
.sheet-title-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #F0F4FA;
    border: 1px solid #C5D5E8;
    border-radius: 6px;
    padding: 8px 16px;
    margin-bottom: 12px;
}
.sheet-title-bar h2 {
    margin: 0;
    font-size: 1.1rem;
    color: #1B2A4A;
}
.back-btn {
    background: #2B5FA8 !important;
    color: white !important;
    border: none !important;
    border-radius: 4px !important;
    padding: 4px 14px !important;
    font-size: 0.82rem !important;
    cursor: pointer !important;
    transition: background 0.15s !important;
}
.back-btn:hover {
    background: #1B4A8A !important;
}

/* Scrollable table wrapper */
.table-scroll-wrapper {
    overflow: auto;
    max-height: 78vh;
    border: 1px solid #D0D8E8;
    border-radius: 4px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.06);
}

/* Index cards */
.index-card {
    background: white;
    border: 1px solid #C5D5E8;
    border-left: 4px solid #2B5FA8;
    border-radius: 6px;
    padding: 10px 16px;
    margin-bottom: 8px;
    cursor: pointer;
    transition: box-shadow 0.15s, border-left-color 0.15s;
}
.index-card:hover {
    box-shadow: 0 2px 8px rgba(43,95,168,0.2);
    border-left-color: #4D9FEC;
}

/* Error box */
.error-box {
    background: #FFF0F0;
    border: 1px solid #E0A0A0;
    border-radius: 6px;
    padding: 16px;
    color: #8B0000;
}

/* Stat chips */
.stat-chip {
    display:inline-block;
    background:#E8F0FC;
    color:#1B2A4A;
    border-radius:12px;
    padding:2px 10px;
    font-size:0.78rem;
    margin-right:6px;
}
</style>
""", unsafe_allow_html=True)


# ── Session state ──────────────────────────────────────────────────────────────
if "current_sheet" not in st.session_state:
    st.session_state.current_sheet = INDEX_SHEET


# ── Data loading (cached) ──────────────────────────────────────────────────────
@st.cache_resource(show_spinner="⏳ Lade Excel-Datei …")
def load_workbook(path: str) -> dict[str, SheetData] | None:
    try:
        from export_coordinates import export_coordinates
        csv_path = Path(path).parent / "coordinates.csv"
        export_coordinates(path, csv_path)
        sheets = parse_workbook(path)
        _apply_annotations(sheets, csv_path)
        return sheets
    except FileNotFoundError:
        return None
    except Exception as exc:
        st.error(f"Fehler beim Laden der Datei: {exc}")
        return None


def _apply_annotations(sheets: dict[str, SheetData], csv_path: Path) -> None:
    """Read annotations from coordinates.csv and attach them to matching cells.
    Supports both full coordinates (input cells) and label keys (row/col headers).
    """
    if not csv_path.exists():
        return
    annotations: dict[str, str] = {}
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames or "annotation" not in reader.fieldnames:
            return
        # Support both old ("coordinate") and new ("key") CSV format
        key_field = "key" if "key" in reader.fieldnames else "coordinate"
        for row in reader:
            note = row.get("annotation", "").strip()
            if note:
                annotations[row.get(key_field, "")] = note
    if not annotations:
        return
    for sheet in sheets.values():
        for row in sheet.rows:
            for cell in row:
                if cell.coordinate and cell.coordinate in annotations:
                    cell.annotation = annotations[cell.coordinate]
                lk = getattr(cell, "label_key", None)
                if lk and lk in annotations:
                    cell.annotation = annotations[lk]

# ── Navigation helper ──────────────────────────────────────────────────────────
def go_to(sheet_name: str) -> None:
    st.session_state.current_sheet = sheet_name


# ── Sidebar ────────────────────────────────────────────────────────────────────
def render_sidebar(sheets: dict[str, SheetData]) -> None:
    with st.sidebar:
        st.markdown("## 📊 EBA DPM Viewer")
        st.markdown("---")
        st.markdown("**Tabellenblätter**")
        for name in sheets:
            icon = "🏠" if name == INDEX_SHEET else "📋"
            active = " ← aktiv" if name == st.session_state.current_sheet else ""
            if st.button(f"{icon} {name}{active}", key=f"nav_{name}"):
                go_to(name)
        st.markdown("---")
        st.markdown(
            f"<small style='color:#8899BB'>Datei:<br><code style='font-size:0.7rem'>"
            f"{Path(EXCEL_PATH).name}</code></small>",
            unsafe_allow_html=True,
        )


# ── Main header ────────────────────────────────────────────────────────────────
def render_header() -> None:
    st.markdown(
        f"""
        <div class="eba-header">
            <span style="font-size:2rem">📊</span>
            <h1>{APP_TITLE}</h1>
            <span class="badge">EBA ITS DPM</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ── Index page ─────────────────────────────────────────────────────────────────
def render_index(sheets: dict[str, SheetData]) -> None:
    render_header()

    index_data = sheets.get(INDEX_SHEET)
    if index_data is None:
        st.warning(f"Kein Tabellenblatt mit dem Namen „{INDEX_SHEET}“ gefunden.")
        _render_fallback_index(sheets)
        return

    # Stat bar
    n_sheets = len(sheets)
    st.markdown(
        f'<span class="stat-chip">📄 {n_sheets} Tabellenblätter</span>'
        f'<span class="stat-chip">📁 {Path(EXCEL_PATH).name}</span>',
        unsafe_allow_html=True,
    )
    st.markdown("")

    # Determine which cell values correspond to sheet names → build link map
    sheet_names_lower = {s.lower(): s for s in sheets}

    def index_link_resolver(sheet_name: str, cell_text: str) -> str | None:
        key = cell_text.strip().lower()
        target = sheet_names_lower.get(key)
        return f"#sheet_{target}" if target and target != INDEX_SHEET else None

    # Render the index sheet as a table first
    html_table = render_sheet_html(index_data, link_resolver=None)
    st.markdown('<div class="table-scroll-wrapper">', unsafe_allow_html=True)
    st.markdown(html_table, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 🔗 Direktlinks zu allen Tabellenblättern")

    cols = st.columns(3)
    for i, name in enumerate(sheets):
        if name == INDEX_SHEET:
            continue
        with cols[i % 3]:
            if st.button(f"📋 {name}", key=f"idx_link_{name}", use_container_width=True):
                go_to(name)


def _render_fallback_index(sheets: dict[str, SheetData]) -> None:
    """Fallback wenn kein 'Index' Blatt vorhanden ist."""
    st.markdown("### Alle verfügbaren Tabellenblätter")
    cols = st.columns(3)
    for i, name in enumerate(sheets):
        with cols[i % 3]:
            if st.button(f"📋 {name}", key=f"fb_{name}", use_container_width=True):
                go_to(name)


# ── Sheet page ─────────────────────────────────────────────────────────────────
def render_sheet(sheet: SheetData) -> None:
    # Title bar with back button
    col_back, col_title = st.columns([1, 6])
    with col_back:
        if st.button("⬅ Index", key="back_to_index", help="Zurück zur Index-Seite"):
            go_to(INDEX_SHEET)
    with col_title:
        n_rows = len(sheet.rows)
        n_cols = len(sheet.col_widths)
        st.markdown(
            f"""
            <div style="display:flex;align-items:center;gap:12px;margin-top:4px">
                <h2 style="margin:0;color:#1B2A4A;font-size:1.25rem">📋 {sheet.name}</h2>
                <span class="stat-chip">{n_rows} Zeilen</span>
                <span class="stat-chip">{n_cols} Spalten</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("")

    # Render table
    html_table = render_sheet_html(sheet)
    st.markdown('<div class="table-scroll-wrapper">', unsafe_allow_html=True)
    st.markdown(html_table, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


# ── Error page ─────────────────────────────────────────────────────────────────
def render_error_page() -> None:
    render_header()
    st.markdown(
        f"""
        <div class="error-box">
            <h3>⚠️ Excel-Datei nicht gefunden</h3>
            <p>Die Datei konnte unter folgendem Pfad nicht geladen werden:</p>
            <code>{EXCEL_PATH}</code>
            <hr>
            <p><strong>Lösungen:</strong></p>
            <ul>
                <li>Kopiere deine Excel-Datei nach <code>{Path(EXCEL_PATH).parent}</code>
                    und benenne sie <code>{Path(EXCEL_PATH).name}</code>.</li>
                <li>Oder setze die Umgebungsvariable:<br>
                    <code>export EBA_EXCEL_PATH="/vollständiger/pfad/zur/datei.xlsx"</code></li>
                <li>Oder passe <code>EXCEL_PATH</code> direkt in <code>app.py</code> an.</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Upload fallback
    st.markdown("---")
    st.markdown("### 📤 Oder Datei jetzt hochladen (temporär)")
    uploaded = st.file_uploader("Excel-Datei hochladen", type=["xlsx", "xls"])
    if uploaded:
        import tempfile, shutil
        tmp = tempfile.mktemp(suffix=".xlsx")
        with open(tmp, "wb") as f:
            f.write(uploaded.read())
        # Clear cache and reload with uploaded file
        load_workbook.clear()
        os.environ["EBA_EXCEL_PATH"] = tmp
        st.rerun()


# ── Entry point ────────────────────────────────────────────────────────────────
def main() -> None:
    sheets = load_workbook(EXCEL_PATH)

    if sheets is None:
        render_error_page()
        return

    render_sidebar(sheets)

    current = st.session_state.current_sheet

    # Guard: if stored sheet no longer exists, fall back to index
    if current not in sheets:
        current = INDEX_SHEET
        st.session_state.current_sheet = current

    if current == INDEX_SHEET or current not in sheets:
        render_index(sheets)
    else:
        render_sheet(sheets[current])


if __name__ == "__main__":
    main()