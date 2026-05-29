"""
Pivot table utilities: read from Excel sheet and render as HTML.
"""

def _int_to_hex(color_int: int) -> str:
    """Convert Excel BGR integer to HTML #RRGGBB hex string."""
    b = (color_int >> 16) & 0xFF
    g = (color_int >> 8) & 0xFF
    r = color_int & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


def read_pivot(wb, sheet_name: str = "Pivot", source_range: str = "'Daily Manifest'!$A:$S") -> dict:
    """Refresh pivot tables and return values + per-cell styles from the used range."""
    pivot_sheet = wb.sheets[sheet_name]
    for pt in pivot_sheet.api.PivotTables():
        pt.ChangePivotCache(
            wb.api.PivotCaches().Create(
                SourceType=1,  # xlDatabase
                SourceData=source_range,
            )
        )
        pt.RefreshTable()
        pt.SaveData = True

    used = pivot_sheet.used_range
    values = used.value
    if not values or not isinstance(values[0], list):
        values = [values]

    num_cols = max(len(row) for row in values)

    # Pad all rows to the same length so styles and values always align
    for row in values:
        while len(row) < num_cols:
            row.append(None)

    styles = []
    for i, row in enumerate(values):
        row_styles = []
        for j in range(num_cols):
            cell = used[i, j].api
            interior = cell.Interior.Color
            font_color = cell.Font.Color
            bold = bool(cell.Font.Bold)
            row_styles.append({
                "bg": _int_to_hex(int(interior)),
                "color": _int_to_hex(int(font_color)),
                "bold": bold,
            })
        styles.append(row_styles)

    return {"values": values, "styles": styles}


def to_html(pivot: dict) -> str:
    """Render pivot dict (from read_pivot) as an HTML table using actual Excel cell styles."""
    if not pivot:
        return ""
    values = pivot["values"]
    styles = pivot["styles"]

    table = '<table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-size:10pt;font-family:Calibri;border:1px solid #000000">'
    for i, row in enumerate(values):
        # Check if this row is a header (first row) or a Total row
        first_cell = str(row[0]) if row[0] is not None else ""
        is_header = i == 0
        is_total = "total" in first_cell.lower()

        table += "<tr>"
        for j, cell in enumerate(row):
            val = "" if cell is None else (
                int(cell) if isinstance(cell, float) and cell == int(cell) else cell
            )
            s = styles[i][j]
            fw = "bold" if s["bold"] else "normal"

            if is_header or is_total:
                bg = "#9DC3E6"
                color = "#000000"
                fw = "bold"
            else:
                bg = s["bg"]
                color = s["color"]

            align = "right" if isinstance(cell, (int, float)) else "left"
            style = (
                f"background:{bg};color:{color};font-weight:{fw};"
                f"padding:4px 8px;white-space:nowrap;"
                f"border:1px solid #000000;text-align:{align}"
            )
            table += f'<td align="{align}" nowrap bgcolor="{bg}" style="{style}">{val}</td>'
        table += "</tr>"
    table += "</table>"
    return table
