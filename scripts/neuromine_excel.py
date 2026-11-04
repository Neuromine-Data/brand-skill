import io
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime
from xml.sax.saxutils import escape

from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Font as DrawingFont, Paragraph, ParagraphProperties
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import to_excel
from openpyxl.worksheet.properties import PageSetupProperties

HEAD_FONT = "Helvetica Neue"
BODY_FONT = "Helvetica Neue"
FALLBACK_FONT = "Arial"
MONO_FONT = "Courier New"

BRAND = "1F4FAE"
BRAND_DARK = "163D8C"
INK = "1A2332"
TEXT = "2C3E50"
TEXT_2 = "374151"
CAPTION = "4B5563"
PLACEHOLDER = "6E7581"
WHITE = "FFFFFF"
BORDER = "8A9097"
RULE = "E5E7EB"
TABLE_BORDER = "DFE3EB"
CHART_AXIS = "6A7683"
CHART_GRID = "D9E1EA"
CHART = ["008300", "E87BA4", "4A3AA7", "1BAF7A", "EDA100", "882255"]
SEQUENTIAL = ["EEF4FF", "DCE8FF", "C4D9FF", "4070D9", "1F4FAE", "163D8C"]
DIVERGING = ["163D8C", "1F4FAE", "C4D9FF", "F8D7DA", "DC3545", "721C24"]
SINGLE = ["1F4FAE"]
PALETTES = {"categorical": CHART, "sequential": SEQUENTIAL, "diverging": DIVERGING, "single": SINGLE}
CHART_KINDS = ("column", "bar", "line")

MARK = "Neuromine"
EYEBROW = "NEUROMINE"
STYLE_PREFIX = "Neuromine "

FMT_INT = "#,##0;-#,##0;0"
FMT_DEC = "#,##0.0;-#,##0.0;0.0"
FMT_PCT = "0.0%"
FMT_DATE = "yyyy-mm-dd"
FMT_LONG_DATE = "d mmmm yyyy"

SIZE_CAPTION = 9
SIZE_BODY = 11
SIZE_SUBTITLE = 14
SIZE_TITLE = 20

TITLE_ROWS = 4
TITLE_ROW = 6
HEADER_ROW = 7
HEADER_HEIGHT = 22
HEADER_WRAP_HEIGHT = 30
LINE_WIDTH = 28575
AXIS_WIDTH = 9525

HAIR = Side(style="thin", color="FF" + TABLE_BORDER)
TOTAL_RULE = Side(style="thin", color="FF" + BORDER)

TOTAL_STYLES = {
    "Neuromine Integer": "Neuromine Total Integer",
    "Neuromine Decimal": "Neuromine Total Decimal",
    "Neuromine Percent": "Neuromine Total Percent",
}


def argb(hexv):
    return "FF" + hexv


def font(name, size, color, bold=False):
    return Font(name=name, size=size, bold=bold, color=argb(color), family=3 if name == MONO_FONT else 2)


def style(name, fnt, fmt="General", fill=None, border=None, align=None):
    s = NamedStyle(name=name)
    s.font = fnt
    s.number_format = fmt
    if fill:
        s.fill = PatternFill("solid", fgColor=argb(fill))
    if border:
        s.border = border
    s.alignment = align or Alignment(vertical="center")
    return s


def styles():
    body_rule = Border(bottom=HAIR)
    total_rule = Border(top=TOTAL_RULE)
    left = Alignment(horizontal="left", vertical="center")
    right = Alignment(horizontal="right", vertical="center")
    p = STYLE_PREFIX
    return [
        style(p + "Eyebrow", font(MONO_FONT, SIZE_CAPTION, CAPTION, bold=True), align=left),
        style(p + "Title", font(HEAD_FONT, SIZE_TITLE, BRAND, bold=True), align=Alignment(horizontal="left", vertical="bottom")),
        style(p + "Subtitle", font(BODY_FONT, SIZE_SUBTITLE, TEXT_2), align=left),
        style(p + "Date", font(MONO_FONT, SIZE_CAPTION, CAPTION), FMT_LONG_DATE, align=left),
        style(p + "Header", font(BODY_FONT, SIZE_BODY, WHITE, bold=True), fill=BRAND, align=Alignment(horizontal="left", vertical="center", wrap_text=True)),
        style(p + "Text", font(BODY_FONT, SIZE_BODY, TEXT_2), border=body_rule, align=left),
        style(p + "Text Wrap", font(BODY_FONT, SIZE_BODY, TEXT_2), border=body_rule, align=Alignment(horizontal="left", vertical="top", wrap_text=True)),
        style(p + "Integer", font(BODY_FONT, SIZE_BODY, TEXT_2), FMT_INT, border=body_rule, align=right),
        style(p + "Decimal", font(BODY_FONT, SIZE_BODY, TEXT_2), FMT_DEC, border=body_rule, align=right),
        style(p + "Percent", font(BODY_FONT, SIZE_BODY, TEXT_2), FMT_PCT, border=body_rule, align=right),
        style(p + "Date Cell", font(BODY_FONT, SIZE_BODY, TEXT_2), FMT_DATE, border=body_rule, align=right),
        style(p + "Total Label", font(BODY_FONT, SIZE_BODY, INK, bold=True), border=total_rule, align=left),
        style(p + "Total Integer", font(BODY_FONT, SIZE_BODY, INK, bold=True), FMT_INT, border=total_rule, align=right),
        style(p + "Total Decimal", font(BODY_FONT, SIZE_BODY, INK, bold=True), FMT_DEC, border=total_rule, align=right),
        style(p + "Total Percent", font(BODY_FONT, SIZE_BODY, INK, bold=True), FMT_PCT, border=total_rule, align=right),
        style(p + "Table Title", font(HEAD_FONT, SIZE_BODY, INK, bold=True), align=left),
        style(p + "Caption", font(MONO_FONT, SIZE_CAPTION, CAPTION), align=left),
        style(p + "Note", font(BODY_FONT, SIZE_CAPTION, CAPTION), align=left),
    ]


def setup_sheet(ws, name, title, subtitle, when, landscape=False):
    ws.title = name
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    ws["A1"] = EYEBROW
    ws["A1"].style = "Neuromine Eyebrow"
    ws["A2"] = title
    ws["A2"].style = "Neuromine Title"
    ws["A3"] = subtitle
    ws["A3"].style = "Neuromine Subtitle"
    ws["A4"] = when
    ws["A4"].style = "Neuromine Date"
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[HEADER_ROW].height = HEADER_HEIGHT
    ws.freeze_panes = ws.cell(row=HEADER_ROW + 1, column=1)
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins.left = ws.page_margins.right = 1.0
    ws.page_margins.top = ws.page_margins.bottom = 1.0
    ws.page_margins.header = ws.page_margins.footer = 0.5
    ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"
    for part, text, color in ((ws.oddHeader.left, MARK, CAPTION), (ws.oddHeader.right, "&A", CAPTION), (ws.oddFooter.right, "&P", PLACEHOLDER)):
        part.text = text
        part.font = f"{MONO_FONT},Regular"
        part.size = SIZE_CAPTION
        part.color = color


def new_sheet(wb, name, title, subtitle, when, landscape=False):
    ws = wb.create_sheet(name)
    setup_sheet(ws, name, title, subtitle, when, landscape=landscape)
    return ws


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def fill_formula(value, row, first, last):
    if isinstance(value, str) and value.startswith("="):
        return value.format(row=row, first=first, last=last)
    return value


def is_pair(value):
    return isinstance(value, tuple) and len(value) == 2


def cached(wb):
    if not hasattr(wb, "_neuromine_cached"):
        wb._neuromine_cached = {}
    return wb._neuromine_cached


def put(ws, row, column, value, first, last):
    store = cached(ws.parent).setdefault(ws.title, {})
    if is_pair(value):
        formula, result = value
        cell = ws.cell(row=row, column=column, value=fill_formula(formula, row, first, last))
        if result is None:
            store.pop(cell.coordinate, None)
        else:
            store[cell.coordinate] = result
        return cell
    cell = ws.cell(row=row, column=column, value=fill_formula(value, row, first, last))
    store.pop(cell.coordinate, None)
    return cell


def column_sum(rows, index):
    total = 0
    for row in rows:
        v = row[index] if index < len(row) else None
        if is_pair(v):
            v = v[1]
        if v is None or v == "":
            continue
        if isinstance(v, bool) or not isinstance(v, (int, float)):
            return None
        total += v
    return total


def header_wraps(ws, heads):
    for c, h in enumerate(heads, 1):
        text = str(h)
        width = ws.column_dimensions[get_column_letter(c)].width or 8.43
        if "\n" in text or len(text) > width:
            return True
    return False


def add_table(ws, heads, rows, styles, first_row=HEADER_ROW, total=None, caption=None, autofilter=False, title=None):
    if len(styles) != len(heads):
        raise ValueError("styles needs one style name per column")
    above = ws.cell(row=first_row - 1, column=1)
    if title:
        above.value = title
        above.style = "Neuromine Table Title"
    elif above.style == "Neuromine Table Title":
        above.value = None
        above.style = "Normal"
    for c, h in enumerate(heads, 1):
        ws.cell(row=first_row, column=c, value=h).style = "Neuromine Header"
    c = len(heads) + 1
    while ws.cell(row=first_row, column=c).style == "Neuromine Header":
        ws.cell(row=first_row, column=c).value = None
        ws.cell(row=first_row, column=c).style = "Normal"
        c += 1
    ws.row_dimensions[first_row].height = HEADER_WRAP_HEIGHT if header_wraps(ws, heads) else HEADER_HEIGHT
    first = first_row + 1
    last = first_row + len(rows)
    for r, row in enumerate(rows, first):
        for c, v in enumerate(row, 1):
            put(ws, r, c, v, first, last).style = styles[c - 1]
    nxt = last + 1
    if total is not None:
        cells = [total] if isinstance(total, str) else list(total)
        cells += [None] * (len(heads) - len(cells))
        for c, v in enumerate(cells, 1):
            col = get_column_letter(c)
            body = styles[c - 1]
            if c == 1:
                put(ws, nxt, c, v, first, last).style = "Neuromine Total Label"
                continue
            if v is None and body in ("Neuromine Integer", "Neuromine Decimal"):
                v = (f"=SUM({col}{first}:{col}{last})", column_sum(rows, c - 1))
            cell = put(ws, nxt, c, v if v != "" else None, first, last)
            cell.style = TOTAL_STYLES.get(body, "Neuromine Total Label")
        nxt += 1
    if caption:
        ws.cell(row=nxt, column=1, value=caption).style = "Neuromine Caption"
        nxt += 1
    if autofilter:
        ws.auto_filter.ref = f"A{first_row}:{get_column_letter(len(heads))}{last}"
    return first, last, nxt


def chart_text(size, color, bold=False):
    cp = CharacterProperties(sz=size * 100, b=bold, solidFill=color, latin=DrawingFont(typeface=BODY_FONT))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def series_colors(palette, n):
    colors = PALETTES[palette]
    if palette == "categorical" and n == 1:
        return list(SINGLE)
    if palette in ("categorical", "single") or n >= len(colors):
        return [colors[i % len(colors)] for i in range(n)]
    if palette == "sequential":
        return colors[len(colors) - n:]
    if n == 1:
        return [colors[0]]
    return [colors[round(i * (len(colors) - 1) / (n - 1))] for i in range(n)]


def add_chart(ws, kind, first, last, ncols, anchor, palette="categorical", label_col=1):
    if kind not in CHART_KINDS:
        raise ValueError(f"unknown chart kind {kind!r}; use column, bar or line")
    if palette not in PALETTES:
        raise ValueError(f"unknown palette {palette!r}; use categorical, sequential, diverging or single")
    if kind == "line":
        ch = LineChart()
        ch.grouping = "standard"
    else:
        ch = BarChart()
        ch.type = "col" if kind == "column" else "bar"
        ch.grouping = "clustered"
        ch.overlap = -10
        ch.gapWidth = 80
    ch.title = None
    ch.style = 2
    data = Reference(ws, min_col=label_col + 1, max_col=label_col + ncols - 1, min_row=first - 1, max_row=last)
    cats = Reference(ws, min_col=label_col, min_row=first, max_row=last)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    for s, color in zip(ch.series, series_colors(palette, len(ch.series))):
        if kind == "line":
            s.graphicalProperties.line.solidFill = color
            s.graphicalProperties.line.width = LINE_WIDTH
            s.marker.symbol = "none"
            s.smooth = False
        else:
            s.graphicalProperties.solidFill = color
            s.graphicalProperties.line.noFill = True
    fmt = ws.cell(row=first, column=label_col + 1).number_format
    fmt = fmt if fmt in (FMT_INT, FMT_DEC, FMT_PCT) else FMT_INT
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=CHART_GRID, w=AXIS_WIDTH))
    ch.y_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=CHART_GRID, w=AXIS_WIDTH))
    ch.y_axis.numFmt = fmt
    ch.y_axis.number_format = fmt
    if palette == "diverging":
        ch.x_axis.crosses = None
        ch.x_axis.crossesAt = 0
        ch.x_axis.tickLblPos = "low"
    ch.x_axis.txPr = chart_text(SIZE_CAPTION, CHART_AXIS)
    ch.y_axis.txPr = chart_text(SIZE_CAPTION, CHART_AXIS)
    ch.legend.position = "t"
    ch.legend.overlay = False
    ch.legend.txPr = chart_text(SIZE_CAPTION, TEXT_2)
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.width = 18
    ch.height = 8.5
    ws.add_chart(ch, anchor)
    return ch


def finish(wb):
    wb.calculation.fullCalcOnLoad = True
    for ws in list(wb.worksheets):
        if len(wb.worksheets) == 1:
            break
        if ws.title in ("Sheet", "Sheet1") and not any(c.value is not None for row in ws.iter_rows() for c in row):
            wb.remove(ws)
    return wb


CELL_RE = re.compile(r'<c r="([A-Z]+[0-9]+)"([^>]*)>(<f[ >].*?</f>|<f[^>]*/>)(?:<v>[^<]*</v>|<v/>)?</c>', re.S)
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def sheet_parts(zin):
    rels = {}
    for r in ET.fromstring(zin.read("xl/_rels/workbook.xml.rels")):
        target = r.get("Target")
        rels[r.get("Id")] = target.lstrip("/") if target.startswith("/") else "xl/" + target
    book = ET.fromstring(zin.read("xl/workbook.xml"))
    return {s.get("name"): rels.get(s.get(REL_NS)) for s in book.iter(MAIN_NS + "sheet")}


def value_xml(attrs, value):
    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
    if isinstance(value, bool):
        return attrs + ' t="b"', "1" if value else "0"
    if isinstance(value, (datetime, date)):
        value = to_excel(value)
    if isinstance(value, (int, float)):
        return attrs, repr(value) if isinstance(value, float) else str(value)
    return attrs + ' t="str"', escape(str(value))


def fill_cached(xml, values):
    def sub(m):
        if m.group(1) not in values:
            return m.group(0)
        attrs, text = value_xml(m.group(2), values[m.group(1)])
        return f'<c r="{m.group(1)}"{attrs}>{m.group(3)}<v>{text}</v></c>'
    return CELL_RE.sub(sub, xml)


def save(wb, path):
    finish(wb)
    buf = io.BytesIO()
    wb.save(buf)
    store = cached(wb)
    with zipfile.ZipFile(io.BytesIO(buf.getvalue())) as zin:
        parts = {part: store.get(name, {}) for name, part in sheet_parts(zin).items() if part}
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if parts.get(info.filename):
                    data = fill_cached(data.decode("utf-8"), parts[info.filename]).encode("utf-8")
                zout.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
    return path
