import io
import re
import subprocess
import sys
import zipfile
from datetime import date, datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

from openpyxl.styles import Font
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.writer.theme import theme_xml

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
BOARD = HERE.parents[2]
sys.path.insert(0, str(HERE))

from neuromine_excel import (
    BODY_FONT, BRAND, BRAND_DARK, CHART, HEAD_FONT, HEADER_ROW, INK, RULE, SIZE_BODY, TEXT, TITLE_ROW,
    add_chart, add_table, argb, new_sheet, save, set_widths, setup_sheet, styles,
)

TEMPLATE_NAME = "neuromine-template.xlsx"
SAMPLE_NAME = "neuromine-sample.xlsx"
AUTHOR = "Neuromine data team"
FIXED_DATE = datetime(2026, 1, 1, 0, 0, 0)
ZIP_DATE = (2026, 1, 1, 0, 0, 0)
SAMPLE_DATE = date(2026, 11, 4)
SOURCE = "Source: example numbers, not release data."


def brand_theme():
    t = theme_xml
    for slot, hexv in zip(["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"], CHART):
        t = re.sub(rf"(<a:{slot}>\s*<a:srgbClr val=\")[0-9A-F]{{6}}", rf"\g<1>{hexv}", t)
    t = re.sub(r"(<a:dk2>\s*<a:srgbClr val=\")[0-9A-F]{6}", rf"\g<1>{INK}", t)
    t = re.sub(r"(<a:lt2>\s*<a:srgbClr val=\")[0-9A-F]{6}", rf"\g<1>{RULE}", t)
    t = re.sub(r"(<a:hlink>\s*<a:srgbClr val=\")[0-9A-F]{6}", rf"\g<1>{BRAND}", t)
    t = re.sub(r"(<a:folHlink>\s*<a:srgbClr val=\")[0-9A-F]{6}", rf"\g<1>{BRAND_DARK}", t)
    t = t.replace('<a:clrScheme name="Office">', '<a:clrScheme name="Neuromine">')
    t = t.replace('<a:fontScheme name="Office">', '<a:fontScheme name="Neuromine">')
    t = t.replace('<a:latin typeface="Cambria"', f'<a:latin typeface="{HEAD_FONT}"')
    t = t.replace('<a:latin typeface="Calibri"', f'<a:latin typeface="{BODY_FONT}"')
    return t.encode("utf-8")


def new_workbook(title):
    wb = openpyxl.Workbook()
    base = Font(name=BODY_FONT, size=SIZE_BODY, color=argb(TEXT), family=2, scheme="minor")
    wb._fonts = IndexedList([base])
    wb._named_styles["Normal"].font = base
    for s in styles():
        wb.add_named_style(s)
    wb.loaded_theme = brand_theme()
    wb.properties.title = title
    wb.properties.creator = AUTHOR
    wb.properties.lastModifiedBy = AUTHOR
    wb.properties.created = FIXED_DATE
    wb.properties.modified = FIXED_DATE
    return wb


def build_template(path):
    wb = new_workbook("Neuromine workbook template")
    ws = wb.active
    setup_sheet(ws, "Sheet1", "Workbook title", "One line on what this sheet holds", date(2026, 1, 1))
    heads = ["Item", "Value 1", "Value 2", "Value 3", "Value 4"]
    ws.cell(row=TITLE_ROW, column=1, value="Table title").style = "Neuromine Table Title"
    for c, h in enumerate(heads, 1):
        ws.cell(row=HEADER_ROW, column=c, value=h).style = "Neuromine Header"
    set_widths(ws, [24, 14, 14, 14, 14])
    save_stable(wb, path)


def build_visits(wb):
    ws = wb.active
    setup_sheet(ws, "Visits", "Visits by site", "Example participant counts by site and visit; the numbers come from no release", SAMPLE_DATE)
    heads = ["Site", "Screening", "Month 3", "Month 6", "Kept to month 6"]
    rows = [
        ["Site A", 120, 108, 97],
        ["Site B", 95, 91, 86],
        ["Site C", 64, 58, 51],
        ["Site D", 41, 39, 37],
    ]
    rows = [r + [("=D{row}/B{row}", r[3] / r[1])] for r in rows]
    screen = sum(r[1] for r in rows)
    month6 = sum(r[3] for r in rows)
    set_widths(ws, [24, 14, 14, 14, 16])
    first, last, nxt = add_table(
        ws, heads, rows,
        ["Neuromine Text", "Neuromine Integer", "Neuromine Integer", "Neuromine Integer", "Neuromine Percent"],
        total=["All sites", None, None, None, ("=D{row}/B{row}", month6 / screen)],
        caption="Table 1. Example participants by site and visit",
        title="Participants by site and visit",
    )
    ws.cell(row=nxt, column=1, value=SOURCE).style = "Neuromine Note"
    add_chart(ws, "column", first, last, 4, f"A{nxt + 2}")
    ws.cell(row=nxt + 20, column=1, value="Figure 1. Example participants by site and visit").style = "Neuromine Caption"
    ws.cell(row=nxt + 21, column=1, value=SOURCE).style = "Neuromine Note"


def build_trend(wb):
    ws = new_sheet(wb, "Trend", "Visits by month", "Example visit counts by month, January to June 2026, two sites", SAMPLE_DATE)
    heads = ["Month", "Site A", "Site B"]
    rows = [
        ["January", 18, 12],
        ["February", 21, 14],
        ["March", 24, 15],
        ["April", 22, 17],
        ["May", 27, 19],
        ["June", 30, 21],
    ]
    set_widths(ws, [24, 14, 14])
    first, last, nxt = add_table(
        ws, heads, rows,
        ["Neuromine Text", "Neuromine Integer", "Neuromine Integer"],
        total="Total",
        caption="Table 2. Example visits by month and site",
        title="Visits by month",
    )
    ws.cell(row=nxt, column=1, value=SOURCE).style = "Neuromine Note"
    add_chart(ws, "line", first, last, 3, f"A{nxt + 2}")
    ws.cell(row=nxt + 20, column=1, value="Figure 2. Example visits by month, January to June 2026").style = "Neuromine Caption"
    ws.cell(row=nxt + 21, column=1, value=SOURCE).style = "Neuromine Note"


def build_notes(wb):
    ws = new_sheet(wb, "Notes", "Notes on the data", "Sources, definitions and known gaps for this workbook", SAMPLE_DATE)
    set_widths(ws, [24, 60])
    add_table(
        ws, ["Label", "Note"],
        [
            ["Key point", "Site A kept 80.8 percent of its participants to month 6 in this example."],
            ["Source", "Example numbers written for this sample. A real workbook names the release in A3 and in each source line."],
            ["Kept", "Kept to month 6 is the month 6 count over the screening count, stored as a fraction."],
            ["Blank cells", "A blank cell means missing; 0 means zero."],
            ["Step 1", "Take the counts from the release index."],
            ["Step 2", "Fill the template with scripts/neuromine_excel.py and save with save(wb, path)."],
            ["Step 3", "Run scripts/check_xlsx.py on the file and fix every FAIL."],
            ["Warning", "Never put participant level rows in a workbook that leaves the team."],
        ],
        ["Neuromine Text", "Neuromine Text Wrap"],
        caption="Table 3. Notes",
        title="Notes",
    )


def build_sample(path):
    wb = new_workbook("Neuromine sample workbook")
    build_visits(wb)
    build_trend(wb)
    build_notes(wb)
    save_stable(wb, path, lambda buf: save(wb, buf))


def save_stable(wb, path, write=None):
    path.parent.mkdir(parents=True, exist_ok=True)
    buf = io.BytesIO()
    (write or wb.save)(buf)
    with zipfile.ZipFile(io.BytesIO(buf.getvalue())) as zin:
        items = [(info.filename, zin.read(info.filename)) for info in zin.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in items:
            if name == "docProps/core.xml":
                data = re.sub(rb"[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z", b"2026-01-01T00:00:00Z", data)
            zout.writestr(zipfile.ZipInfo(name, ZIP_DATE), data, compress_type=zipfile.ZIP_DEFLATED)


def default_out():
    if (BOARD / "board_check.py").is_file():
        sys.path.insert(0, str(BOARD))
        import board_check
        return board_check.data_path("") / ROOT.relative_to(BOARD) / "v1" / "assets"
    return ROOT / "assets"


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else default_out()
    template = out / TEMPLATE_NAME
    sample = out / SAMPLE_NAME
    build_template(template)
    build_sample(sample)
    for p in (template, sample):
        print(f"wrote {p.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
