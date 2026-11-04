import io
import json
import re
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path

try:
    import docx
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
    import docx

try:
    import matplotlib
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "matplotlib"])
    import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from PIL import Image, ImageDraw, ImageFont

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from check_docx import NUMBER_RE, PLACEHOLDER_SECTION, TITLE_RIGHT_IND

ROOT = HERE.parent
BOARD = HERE.parents[2]
TOKENS = json.loads((ROOT / "design" / "tokens.json").read_text(encoding="ascii"))
LOGO = ROOT / "design" / "assets" / "neuromine-logo-light.png"
TEMPLATE_NAME = "neuromine-template.docx"
SAMPLE_NAME = "neuromine-sample.docx"
FIGURE_NAME = "neuromine-sample-figure-1.png"

C = {k: v.upper() for k, v in TOKENS["color"].items()}
PT = TOKENS["size"]["document_pt"]
HEAD_FONT = TOKENS["font"]["heading"]
BODY_FONT = TOKENS["font"]["body"]
FALLBACK = TOKENS["font"]["fallback"]
CHART = [C[f"chart_{i}"] for i in range(1, 7)]
MARK = "Neuromine"
WORDMARK = TOKENS["font"]["wordmark"]

PAGE_W = 12240
PAGE_H = 15840
MARGIN = 1440
TEXT_W = PAGE_W - 2 * MARGIN
BAND_IN = (8.5, 1.1)
BAND_DPI = 200
LOGO_IN = 0.6
LOGO_LEFT_IN = 1.0
PROGRAM_PT = 11
PROGRAM_SPACING = 10
PROGRAM_W = 5040
PROGRAM_LINE = 300
PROGRAM_Y = round(BAND_IN[1] * 1440 / 2 - PROGRAM_LINE / 2)
PROGRAM_MAX = 20
EYEBROW_BEFORE = 720

BULLET_NUM = 90
NUMBER_NUM = 91
BULLET_CHAR = chr(0xF0B7)

SAMPLE_PROGRAM = "ALS TDI"

FIXED_DATE = datetime(2026, 1, 1, 0, 0, 0)
ZIP_DATE = (2026, 1, 1, 0, 0, 0)


def default_out():
    if (BOARD / "board_check.py").is_file():
        sys.path.insert(0, str(BOARD))
        import board_check
        return board_check.data_path("") / ROOT.relative_to(BOARD) / "v1" / "assets"
    return ROOT / "assets"


def rpr(font=None, size=None, color=None, bold=False, caps=False, spacing=None):
    parts = []
    if font:
        parts.append(f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>')
    if bold:
        parts.append("<w:b/><w:bCs/>")
    if caps:
        parts.append("<w:caps/>")
    if color:
        parts.append(f'<w:color w:val="{color}"/>')
    if spacing is not None:
        parts.append(f'<w:spacing w:val="{spacing}"/>')
    if size:
        half = int(round(size * 2))
        parts.append(f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/>')
    return "<w:rPr>" + "".join(parts) + "</w:rPr>"


def ppr(keep=False, num=None, border=None, shd=None, tabs=None, before=None, after=None, line=None, ind=None, jc=None, outline=None):
    parts = []
    if keep:
        parts.append("<w:keepNext/><w:keepLines/>")
    if num is not None:
        parts.append(f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num}"/></w:numPr>')
    if border:
        side, sz, space, color = border
        parts.append(f'<w:pBdr><w:{side} w:val="single" w:sz="{sz}" w:space="{space}" w:color="{color}"/></w:pBdr>')
    if shd:
        parts.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{shd}"/>')
    if tabs:
        parts.append("<w:tabs>" + "".join(f'<w:tab w:val="{v}" w:pos="{p}"/>' for v, p in tabs) + "</w:tabs>")
    attrs = []
    if before is not None:
        attrs.append(f'w:before="{before}"')
    if after is not None:
        attrs.append(f'w:after="{after}"')
    if line is not None:
        attrs.append(f'w:line="{line}" w:lineRule="auto"')
    if attrs:
        parts.append("<w:spacing " + " ".join(attrs) + "/>")
    if ind:
        parts.append("<w:ind " + " ".join(f'w:{k}="{v}"' for k, v in ind.items()) + "/>")
    if jc:
        parts.append(f'<w:jc w:val="{jc}"/>')
    if outline is not None:
        parts.append(f'<w:outlineLvl w:val="{outline}"/>')
    return "<w:pPr>" + "".join(parts) + "</w:pPr>"


def put_style(styles_el, kind, style_id, name, body, based="Normal", nxt=None, ui=None, custom=False):
    for old in styles_el.findall(qn("w:style")):
        if old.get(qn("w:styleId")) == style_id:
            styles_el.remove(old)
    head = f'<w:name w:val="{name}"/>'
    if based:
        head += f'<w:basedOn w:val="{based}"/>'
    if nxt:
        head += f'<w:next w:val="{nxt}"/>'
    if ui is not None:
        head += f'<w:uiPriority w:val="{ui}"/>'
    head += "<w:qFormat/>"
    flag = ' w:customStyle="1"' if custom else ""
    default = ' w:default="1"' if style_id == "Normal" else ""
    xml = f'<w:style {nsdecls("w")} w:type="{kind}"{default}{flag} w:styleId="{style_id}">{head}{body}</w:style>'
    styles_el.append(parse_xml(xml))


def build_styles(doc):
    styles_el = doc.styles.element
    styles_el.remove(styles_el.find(qn("w:docDefaults")))
    styles_el.insert(0, parse_xml(
        f'<w:docDefaults {nsdecls("w")}>'
        f'<w:rPrDefault>{rpr(BODY_FONT, PT["body"], C["text"])}</w:rPrDefault>'
        f'<w:pPrDefault>{ppr(after=132, line=276)}</w:pPrDefault>'
        "</w:docDefaults>"
    ))
    put_style(styles_el, "paragraph", "Normal", "Normal", "", based=None)
    heads = [
        ("Heading1", "heading 1", PT["heading"], C["ink"], 360, 120, 0),
        ("Heading2", "heading 2", PT["body"], C["ink"], 240, 80, 1),
        ("Heading3", "heading 3", PT["body"], C["text_2"], 200, 60, 2),
    ]
    for sid, name, size, color, before, after, level in heads:
        put_style(styles_el, "paragraph", sid, name,
                  ppr(keep=True, before=before, after=after, line=276, outline=level)
                  + rpr(HEAD_FONT, size, color, bold=True),
                  nxt="Normal", ui=9)
    put_style(styles_el, "paragraph", "Title", "Title",
              ppr(before=0, after=160, line=240, ind={"right": TITLE_RIGHT_IND})
              + rpr(HEAD_FONT, PT["title"], C["brand"], bold=True),
              nxt="Subtitle", ui=10)
    put_style(styles_el, "paragraph", "Subtitle", "Subtitle",
              ppr(before=0, after=0, line=276, ind={"right": TITLE_RIGHT_IND}) + rpr(BODY_FONT, PT["heading"], C["text_2"]),
              ui=11)
    put_style(styles_el, "paragraph", "CoverEyebrow", "Cover Eyebrow",
              ppr(before=EYEBROW_BEFORE, after=160) + rpr(BODY_FONT, PT["caption"], C["caption"], bold=True, caps=True, spacing=10),
              nxt="Title", custom=True)
    put_style(styles_el, "paragraph", "CoverDate", "Cover Date",
              ppr(before=0, after=0, line=240) + rpr(BODY_FONT, PT["caption"], C["caption"]),
              custom=True)
    put_style(styles_el, "paragraph", "CoverProgram", "Cover Program",
              f'<w:pPr><w:spacing w:before="0" w:after="0" w:line="{PROGRAM_LINE}" w:lineRule="exact"/><w:jc w:val="right"/></w:pPr>'
              + rpr(BODY_FONT, PROGRAM_PT, C["tint_3"], bold=True, caps=True, spacing=PROGRAM_SPACING),
              custom=True)
    put_style(styles_el, "paragraph", "Header", "header",
              ppr(border=("bottom", 6, 4, C["rule"]), tabs=[("right", TEXT_W)], before=0, after=0, line=240)
              + rpr(BODY_FONT, PT["caption"], C["caption"]))
    put_style(styles_el, "paragraph", "Footer", "footer",
              ppr(border=("top", 6, 4, C["rule"]), before=0, after=0, line=240, jc="right")
              + rpr(BODY_FONT, PT["caption"], C["placeholder"]))
    put_style(styles_el, "paragraph", "ListBullet", "List Bullet", ppr(num=BULLET_NUM, after=90), ui=99)
    put_style(styles_el, "paragraph", "ListNumber", "List Number", ppr(num=NUMBER_NUM, after=90), ui=99)
    put_style(styles_el, "paragraph", "Quote", "Quote",
              ppr(before=120, after=200, ind={"left": 440, "right": 440}) + rpr(BODY_FONT, PT["body"], C["text_2"]),
              nxt="Normal", ui=29)
    put_style(styles_el, "paragraph", "Callout", "Callout",
              ppr(border=("left", 18, 8, C["brand"]), shd=C["tint_1"], before=200, after=200, ind={"left": 200, "right": 120})
              + rpr(BODY_FONT, PT["body"], C["text"]),
              nxt="Normal", custom=True)
    put_style(styles_el, "paragraph", "Caption", "caption",
              ppr(before=80, after=200, line=240) + rpr(BODY_FONT, PT["caption"], C["caption"]),
              nxt="Normal", ui=35)
    edge = f'w:val="single" w:sz="6" w:space="0" w:color="{C["table_border"]}"'
    put_style(styles_el, "table", "NeuromineTable", "Neuromine Table",
              ppr(before=0, after=0, line=240)
              + rpr(color=C["text_2"])
              + "<w:tblPr><w:tblBorders>"
              + f"<w:top {edge}/><w:left {edge}/><w:bottom {edge}/><w:right {edge}/>"
              + f'<w:insideH {edge}/><w:insideV w:val="nil"/></w:tblBorders>'
              + '<w:tblCellMar><w:top w:w="90" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
              + '<w:bottom w:w="90" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar>'
              + "</w:tblPr>"
              + '<w:tblStylePr w:type="firstRow">' + rpr(color=C["ink"], bold=True)
              + '<w:tblPr/><w:trPr><w:tblHeader/></w:trPr>'
              + f'<w:tcPr><w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="{C["table_head_rule"]}"/></w:tcBorders>'
              + f'<w:shd w:val="clear" w:color="auto" w:fill="{C["table_head"]}"/></w:tcPr>'
              + "</w:tblStylePr>",
              based="TableNormal", custom=True)


def build_numbering(doc):
    numbering = doc.part.numbering_part.element
    bullet = (
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{BULLET_NUM}">'
        '<w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        f'<w:lvlText w:val="{BULLET_CHAR}"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>'
        f'<w:color w:val="{C["brand"]}"/></w:rPr></w:lvl></w:abstractNum>'
    )
    number = (
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{NUMBER_NUM}">'
        '<w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>'
        f'<w:rPr><w:color w:val="{C["brand"]}"/></w:rPr></w:lvl></w:abstractNum>'
    )
    first_num = numbering.find(qn("w:num"))
    for xml in (bullet, number):
        el = parse_xml(xml)
        if first_num is not None:
            first_num.addprevious(el)
        else:
            numbering.append(el)
    for num_id in (BULLET_NUM, NUMBER_NUM):
        numbering.append(parse_xml(
            f'<w:num {nsdecls("w")} w:numId="{num_id}"><w:abstractNumId w:val="{num_id}"/></w:num>'
        ))


def part_by_name(doc, name):
    for rel in doc.part.rels.values():
        if not rel.is_external and str(rel.target_part.partname) == name:
            return rel.target_part
    raise KeyError(name)


def build_fonts(doc):
    part = part_by_name(doc, "/word/fontTable.xml")
    text = part.blob.decode("utf-8")
    entry = (
        f'<w:font w:name="{HEAD_FONT}"><w:altName w:val="{FALLBACK}"/><w:charset w:val="00"/>'
        '<w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>'
    )
    part._blob = text.replace("</w:fonts>", entry + "</w:fonts>").encode("utf-8")


def build_theme(doc):
    part = part_by_name(doc, "/word/theme/theme1.xml")
    text = part.blob.decode("utf-8")
    colours = {"dk1": C["ink"], "lt1": C["surface"], "dk2": C["text_2"], "lt2": C["page"], "hlink": C["brand"], "folHlink": C["brand_dark"]}
    for i, hexval in enumerate(CHART, 1):
        colours[f"accent{i}"] = hexval
    for slot, hexval in colours.items():
        text = re.sub(rf"<a:{slot}>.*?</a:{slot}>", f'<a:{slot}><a:srgbClr val="{hexval}"/></a:{slot}>', text, flags=re.S)
    text = re.sub(r'(<a:majorFont>\s*<a:latin typeface=")[^"]*', rf"\g<1>{HEAD_FONT}", text)
    text = re.sub(r'(<a:minorFont>\s*<a:latin typeface=")[^"]*', rf"\g<1>{BODY_FONT}", text)
    part._blob = text.encode("utf-8")


def band_png():
    w, h = int(BAND_IN[0] * BAND_DPI), int(BAND_IN[1] * BAND_DPI)
    band = Image.new("RGB", (w, h), "#" + C["brand"])
    side = int(LOGO_IN * BAND_DPI)
    logo = Image.open(LOGO).convert("RGBA").resize((side, side), Image.LANCZOS)
    top = (h - side) // 2
    band.paste(logo, (round(LOGO_LEFT_IN * BAND_DPI), top), logo)
    face = font_manager.FontProperties(family=[WORDMARK["family"], WORDMARK["fallback"]], weight=WORDMARK["weight"])
    font = ImageFont.truetype(font_manager.findfont(face, fallback_to_default=False), round(WORDMARK["document_pt"] / 72 * BAND_DPI))
    x = round(LOGO_LEFT_IN * BAND_DPI) + side + round(WORDMARK["gap_ratio"] * side)
    cap = -font.getbbox(WORDMARK["text"], anchor="ls")[1]
    ImageDraw.Draw(band).text((x, top + (side + cap) // 2), WORDMARK["text"], font=font, fill="#" + C[WORDMARK["color"]], anchor="ls")
    buf = io.BytesIO()
    band.save(buf, format="PNG")
    buf.seek(0)
    return buf


def anchor_behind(run):
    inline = run._r.find(".//" + qn("wp:inline"))
    extent = inline.find(qn("wp:extent"))
    anchor = parse_xml(
        f'<wp:anchor {nsdecls("wp")} distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" '
        'behindDoc="1" locked="1" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{extent.get("cx")}" cy="{extent.get("cy")}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/></wp:anchor>'
    )
    for tag in ("wp:docPr", "wp:cNvGraphicFramePr", "a:graphic"):
        el = inline.find(qn(tag))
        if el is not None:
            anchor.append(el)
    inline.getparent().replace(inline, anchor)


def field(instr, cached):
    instr = instr.replace('"', "&quot;")
    return parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr=" {instr} "><w:r><w:t>{cached}</w:t></w:r></w:fldSimple>')


def clear_paragraph(paragraph):
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)


def header_mark(program):
    return f"{MARK} | {program}" if program else MARK


def add_program(header):
    p = header.add_paragraph(style="Cover Program")
    p_pr = p._p.get_or_add_pPr()
    p_pr.insert(1, parse_xml(
        f'<w:framePr {nsdecls("w")} w:w="{PROGRAM_W}" w:h="{PROGRAM_LINE}" w:hRule="exact" w:wrap="around" '
        f'w:hAnchor="margin" w:vAnchor="page" w:xAlign="right" w:y="{PROGRAM_Y}"/>'
    ))
    p.add_run("")
    return p


def set_program(doc, program):
    if len(program) > PROGRAM_MAX:
        raise ValueError(f"program name over {PROGRAM_MAX} characters: {program}")
    sec = doc.sections[0]
    sec.first_page_header.paragraphs[1].runs[0].text = program
    sec.header.paragraphs[0].runs[0].text = header_mark(program)


def set_header_footer(doc, section_text, date_text):
    sec = doc.sections[0]
    sec.different_first_page_header_footer = True
    head = sec.header.paragraphs[0]
    clear_paragraph(head)
    head.style = doc.styles["Header"]
    head.add_run(MARK)
    head.add_run("\t")
    head._p.append(field('STYLEREF "Heading 1" \\* MERGEFORMAT', section_text))
    foot = sec.footer.paragraphs[0]
    clear_paragraph(foot)
    foot.style = doc.styles["Footer"]
    foot._p.append(field("PAGE \\* MERGEFORMAT", "2"))
    first_head = sec.first_page_header.paragraphs[0]
    clear_paragraph(first_head)
    first_head.style = doc.styles["Normal"]
    first_head.paragraph_format.space_after = 0
    anchor_behind(add_band(first_head))
    add_program(sec.first_page_header)
    first_foot = sec.first_page_footer.paragraphs[0]
    first_foot.style = doc.styles["Cover Date"]
    first_foot.text = date_text


def add_band(paragraph):
    run = paragraph.add_run()
    run.add_picture(band_png(), width=Inches(BAND_IN[0]), height=Inches(BAND_IN[1]))
    return run


def set_page(doc):
    sec = doc.sections[0]
    sec.page_width = PAGE_W * 635
    sec.page_height = PAGE_H * 635
    for side in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(sec, side, MARGIN * 635)
    sec.header_distance = 720 * 635
    sec.footer_distance = 720 * 635


def set_settings(doc):
    settings = doc.settings.element
    for old in settings.findall(qn("w:updateFields")):
        settings.remove(old)
    el = parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>')
    anchor = None
    for tag in ("w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars", "w:rsids"):
        anchor = settings.find(qn(tag))
        if anchor is not None:
            break
    if anchor is not None:
        anchor.addprevious(el)
    else:
        settings.append(el)


def set_core(doc, title):
    core = doc.core_properties
    core.author = "Neuromine data team"
    core.last_modified_by = "Neuromine data team"
    core.title = title
    core.comments = ""
    core.created = FIXED_DATE
    core.modified = FIXED_DATE
    core.last_printed = FIXED_DATE
    core.revision = 1


def save_stable(doc, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))
    with zipfile.ZipFile(path) as zin:
        items = [(info.filename, zin.read(info.filename)) for info in zin.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in items:
            zout.writestr(zipfile.ZipInfo(name, ZIP_DATE), data, compress_type=zipfile.ZIP_DEFLATED)


def clear_body(doc):
    body = doc.element.body
    for p in body.findall(qn("w:p")):
        body.remove(p)


def build_template(path):
    doc = docx.Document()
    build_styles(doc)
    build_numbering(doc)
    build_fonts(doc)
    build_theme(doc)
    set_page(doc)
    set_header_footer(doc, PLACEHOLDER_SECTION, "Month Year")
    set_settings(doc)
    clear_body(doc)
    doc.add_paragraph("")
    set_core(doc, "Neuromine template")
    save_stable(doc, path)


def add_table(doc, rows):
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = doc.styles["Neuromine Table"]
    numeric = {c for c in range(len(rows[0])) if all(NUMBER_RE.match(row[c].strip()) for row in rows[1:])}
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.text = value
            if c in numeric:
                cell.paragraphs[0].alignment = 2
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:tblHeader")) is None:
        tr_pr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    return table


def set_section_text(doc, text):
    doc.sections[0].header.paragraphs[0]._p.xpath(".//w:fldSimple//w:t")[0].text = text


def chart_png(labels, series, names):
    axis = "#" + C["chart_axis"]
    grid = "#" + C["chart_grid"]
    plt.rcParams["font.family"] = [FALLBACK, "sans-serif"]
    fig, ax = plt.subplots(figsize=(6.5, 3.25), dpi=200)
    fig.patch.set_alpha(0)
    ax.patch.set_alpha(0)
    for values, name, colour in zip(series, names, CHART):
        ax.plot(labels, values, color="#" + colour, linewidth=2, label=name)
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_color(grid)
        ax.spines[side].set_linewidth(0.75)
    ax.grid(axis="y", color=grid, linewidth=0.5)
    ax.set_axisbelow(True)
    ax.tick_params(colors=axis, labelsize=9, width=0.75)
    ax.set_xlabel("Month", color=axis, fontsize=9)
    ax.set_ylabel("Visits (count)", color=axis, fontsize=9)
    ax.legend(frameon=False, fontsize=9, labelcolor=axis)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, transparent=True, metadata={"Software": None})
    plt.close(fig)
    return buf.getvalue()


def figure_png():
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
    site_a = [18, 20, 23, 22, 26, 29]
    site_b = [12, 15, 15, 18, 19, 21]
    return chart_png(months, [site_a, site_b], ["Site A", "Site B"])


def build_sample(template, path, figure):
    doc = docx.Document(str(template))
    clear_body(doc)
    doc.sections[0].first_page_footer.paragraphs[0].text = "November 2026"
    set_section_text(doc, "1. Summary")
    set_program(doc, SAMPLE_PROGRAM)
    doc.add_paragraph("Sample report", style="Cover Eyebrow")
    doc.add_paragraph("Visit review", style="Title")
    doc.add_paragraph("Every recipe in the Word template, filled with example numbers", style="Subtitle")
    h1 = doc.add_paragraph("1. Summary", style="Heading 1")
    h1.paragraph_format.page_break_before = True
    doc.add_paragraph(
        "This sample shows each part of the Neuromine Word template. "
        "The numbers are made up for the example and come from no release."
    )
    doc.add_paragraph("Counts by site", style="Heading 2")
    doc.add_paragraph("The table compares two example sites across two visits.")
    add_table(doc, [
        ["Site", "Screening", "Month 3", "Change"],
        ["Site A", "120", "108", "-10%"],
        ["Site B", "95", "91", "-4%"],
        ["Both sites", "215", "199", "-7%"],
    ])
    doc.add_paragraph("Table 1. Example visit counts by site. Source: example numbers, not release data.", style="Caption")
    doc.add_paragraph("What the table shows", style="Heading 3")
    for text in ["Both sites kept most participants to month 3.", "Site B lost fewer participants than Site A."]:
        doc.add_paragraph(text, style="List Bullet")
    doc.add_paragraph("To repeat the table for a real release, follow these steps.")
    for text in ["Take the counts from the release index.", "Name the release in the caption.", "Run the checker on the file."]:
        doc.add_paragraph(text, style="List Number")
    callout = doc.add_paragraph(style="Callout")
    callout.add_run("Key point. ").bold = True
    callout.add_run("Every number in a real document comes from a named release.")
    doc.add_paragraph("2. Trend", style="Heading 1")
    doc.add_paragraph("Visits at both example sites rose through the first half of the year.")
    png = figure_png()
    figure.parent.mkdir(parents=True, exist_ok=True)
    figure.write_bytes(png)
    doc.add_picture(io.BytesIO(png), width=Inches(6.5))
    doc.add_paragraph("Figure 1. Example visits by month. Source: example numbers, not release data.", style="Caption")
    set_core(doc, "Visit review")
    save_stable(doc, path)


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else default_out()
    template = out / TEMPLATE_NAME
    sample = out / SAMPLE_NAME
    figure = out / FIGURE_NAME
    build_template(template)
    build_sample(template, sample, figure)
    for path in (template, sample, figure):
        print(f"wrote {path.name}")
    pixel = Image.open(band_png()).convert("RGB").getpixel((10, 10))
    print(f"band pixel {pixel}, brand #{C['brand']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
