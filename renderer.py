"""
renderer.py
Converts SheetData objects into styled HTML table strings.
"""

from __future__ import annotations
import html
from typing import Callable

from excel_parser import SheetData, CellData, CellStyle

BASE_FONT_SIZE = "10pt"
FONT_STACK = "'Segoe UI','Inter','Calibri',system-ui,sans-serif"

# ---------------------------------------------------------------------------
# CSS helpers
# ---------------------------------------------------------------------------

def _style_to_css(style: CellStyle) -> str:
    parts: list[str] = []

    if style.bg_color:
        parts.append(f"background-color:#{style.bg_color}")

    color = style.fg_color or "000000"
    parts.append(f"color:#{color}")

    if style.bold:
        parts.append("font-weight:600")
    if style.italic:
        parts.append("font-style:italic")

    # No per-cell font-size or font-family – unified on <table>

    h_align = style.h_align or "left"
    parts.append(f"text-align:{h_align}")

    v_align = style.v_align or "top"
    v_map = {"top": "top", "center": "middle", "bottom": "bottom"}
    parts.append(f"vertical-align:{v_map.get(v_align, 'top')}")

    if style.wrap_text:
        parts.append("white-space:pre-wrap")
        parts.append("word-break:break-word")
    else:
        parts.append("white-space:nowrap")
        parts.append("overflow:hidden")
        parts.append("text-overflow:ellipsis")

    for attr, val in [
        ("border-top",    style.border_top),
        ("border-bottom", style.border_bottom),
        ("border-left",   style.border_left),
        ("border-right",  style.border_right),
    ]:
        if val:
            parts.append(f"{attr}:{val}")

    parts.append("padding:2px 4px")
    parts.append("max-width:320px")

    return ";".join(parts)


# ---------------------------------------------------------------------------
# Table renderer
# ---------------------------------------------------------------------------

def render_sheet_html(
    sheet: SheetData,
    link_resolver: Callable[[str, str], str] | None = None,
) -> str:
    col_widths = sheet.col_widths
    rows = sheet.rows

    colgroup_parts = ["<colgroup>"]
    for w in col_widths:
        px = max(int(w * 7.5), 40)
        colgroup_parts.append(f'<col style="width:{px}px;min-width:{px}px">')
    colgroup_parts.append("</colgroup>")
    colgroup = "\n".join(colgroup_parts)

    tbody_parts: list[str] = []
    for r_idx, row_cells in enumerate(rows):
        row_height = sheet.row_heights[r_idx] if r_idx < len(sheet.row_heights) else 15.0
        row_h_px = max(row_height * 1.33, 20)
        tr_parts: list[str] = [f'<tr style="height:{row_h_px:.0f}px">']

        for cell in row_cells:
            if cell.is_merged_hidden:
                continue

            css = _style_to_css(cell.style)
            text = html.escape(cell.display_value) if cell.display_value else "&nbsp;"

            span_attrs = ""
            if cell.rowspan > 1:
                span_attrs += f' rowspan="{cell.rowspan}"'
            if cell.colspan > 1:
                span_attrs += f' colspan="{cell.colspan}"'

            if link_resolver and cell.display_value:
                href = link_resolver(sheet.name, cell.display_value)
                if href:
                    text = f'<a href="{href}" style="color:inherit;text-decoration:underline">{text}</a>'

            # ── Coordinate label on input cells ──────────────────────────────
            if cell.coordinate:
                coord_escaped = html.escape(cell.coordinate)
                # Input cell: light background, coordinate shown as small label
                # stacked above an invisible input area
                input_css = (
                    css
                    + ";background-color:#F0F4FF"
                    + ";position:relative"
                    + ";vertical-align:top"
                    + ";padding:0"
                )
                inner = (
                    f'<div style="'
                    f'font-size:7.5pt;color:#5573A8;line-height:1.1;'
                    f'padding:1px 4px 0 4px;white-space:nowrap;overflow:hidden;'
                    f'text-overflow:ellipsis;user-select:none;pointer-events:none'
                    f'">{coord_escaped}</div>'
                    f'<div style="min-height:14px;padding:0 4px 2px 4px">&nbsp;</div>'
                )
                tr_parts.append(
                    f'<td{span_attrs} style="{input_css}" '
                    f'title="{coord_escaped}">'
                    f'{inner}</td>'
                )
            else:
                tr_parts.append(
                    f'<td{span_attrs} style="{css}" title="{html.escape(cell.display_value)}">'
                    f"{text}</td>"
                )

        tr_parts.append("</tr>")
        tbody_parts.append("\n".join(tr_parts))

    tbody = "\n".join(tbody_parts)

    return f"""
<table style="
    border-collapse:collapse;
    font-family:{FONT_STACK};
    font-size:{BASE_FONT_SIZE};
    line-height:1.45;
    table-layout:fixed;
    border:1px solid #ccc;
">
{colgroup}
<tbody>
{tbody}
</tbody>
</table>
"""