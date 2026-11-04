import io
import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path

try:
    import pptx
    import xlsxwriter
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-pptx", "XlsxWriter"])
    import pptx

try:
    import matplotlib
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "matplotlib"])
    import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

from lxml import etree
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Emu

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from neuromine_pptx import (NS, SLIDE_NUM_TEXT, TABLE_STYLE_ID, TABLE_STYLE_NAME, add_callout, add_chart, add_hf,
                       add_image, add_table, layout_named, set_core, set_lines)

ROOT = HERE.parent
BOARD = HERE.parents[2]
TOKENS = json.loads((ROOT / "design" / "tokens.json").read_text(encoding="ascii"))
LOGO = ROOT / "design" / "assets" / "neuromine-logo-light.png"
TEMPLATE_NAME = "neuromine-template.pptx"
SAMPLE_NAME = "neuromine-sample.pptx"

C = {k: v.upper() for k, v in TOKENS["color"].items()}
PT = TOKENS["size"]["slide_pt"]
SIZE_IN = TOKENS["size"]["slide_in"]
HEAD_FONT = TOKENS["font"]["heading"]
BODY_FONT = TOKENS["font"]["body"]
MONO_FONT = TOKENS["font"]["mono"]
EYEBROW_SPC = TOKENS["font"]["eyebrow_tracking_slide_pt"] * 100
CHART = [C[f"chart_{i}"] for i in range(1, 7)]

SLIDE_W = 12192000
SLIDE_H = 6858000
IN = 914400
LEFT = int(0.9 * IN)
TEXT_W = SLIDE_W - 2 * LEFT
COVER_LEFT = int(0.75 * IN)
COVER_W = SLIDE_W - 2 * COVER_LEFT
STRIP_W = int(round(SIZE_IN["edge_strip"] * IN))
LOGO_H = int(round(SIZE_IN["logo_height"] * IN))
WORDMARK = TOKENS["font"]["wordmark"]
WORDMARK_X = COVER_LEFT + LOGO_H + int(round(WORDMARK["gap_ratio"] * LOGO_H))

ZIP_DATE = (2026, 1, 1, 0, 0, 0)
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
BULLET = chr(0x2022)

LAYOUTS = ["Title", "Section", "Title and Content", "Two Content", "Table", "Chart", "Picture", "Closing"]


def inch(v):
    return int(round(v * IN))


def fill(color):
    if color in ("bg1", "tx1"):
        return f'<a:solidFill><a:schemeClr val="{color}"/></a:solidFill>'
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'


def font_tags(font):
    if font == "major":
        return '<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/>'
    if font == "minor":
        return '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>'
    return f'<a:latin typeface="{font}" pitchFamily="49" charset="0"/><a:cs typeface="{font}" pitchFamily="49" charset="0"/>'


def rpr(tag, size, color, font, bold=False, caps=False, spc=None):
    attrs = f'lang="en-US" sz="{int(size * 100)}"'
    if bold:
        attrs += ' b="1"'
    if caps:
        attrs += ' cap="all"'
    if spc is not None:
        attrs += f' spc="{spc}"'
    return f"<a:{tag} {attrs}>{fill(color)}{font_tags(font)}</a:{tag}>"


def lvl(n, size, color, font, algn="l", bullet=None, mar=0, indent=0, after=0, line=None, bold=False, caps=False, spc=None):
    parts = []
    if line:
        parts.append(f'<a:lnSpc><a:spcPct val="{line}"/></a:lnSpc>')
    parts.append('<a:spcBef><a:spcPts val="0"/></a:spcBef>')
    parts.append(f'<a:spcAft><a:spcPts val="{after}"/></a:spcAft>')
    if bullet:
        parts.append(fill(bullet).replace("solidFill", "buClr"))
        parts.append('<a:buSzPct val="100000"/><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>')
        parts.append(f'<a:buChar char="{BULLET}"/>')
    else:
        parts.append("<a:buNone/>")
    parts.append(rpr("defRPr", size, color, font, bold, caps, spc))
    return f'<a:lvl{n}pPr marL="{mar}" indent="{indent}" algn="{algn}" defTabSz="914400" rtl="0">' + "".join(parts) + f"</a:lvl{n}pPr>"


def body_levels():
    out = []
    for n in range(1, 10):
        size, color = (PT["subline"], C["text"]) if n == 1 else (PT["command"], C["text_2"])
        mar = 342900 * min(n, 3)
        out.append(lvl(n, size, color, "minor", bullet=C["brand"], mar=mar, indent=-342900, after=600, line=110000))
    return "".join(out)


def plain_levels(size, color, font, algn="l", bold=False, caps=False, spc=None, line=None, after=0):
    return "".join(lvl(n, size, color, font, algn=algn, bold=bold, caps=caps, spc=spc, line=line, after=after) for n in range(1, 10))


def xfrm(x, y, w, h):
    return f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'


def text_p(text):
    if text is None:
        return '<a:p><a:endParaRPr lang="en-US"/></a:p>'
    return f'<a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p>'


def ph(sid, name, ph_attrs, box=None, lst="", text="", anchor="t"):
    sppr = f'<p:spPr>{xfrm(*box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' if box else "<p:spPr/>"
    paras = text if text and text.startswith("<a:p") else text_p(text or None)
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr><p:ph {ph_attrs}/></p:nvPr></p:nvSpPr>{sppr}'
            f'<p:txBody><a:bodyPr vert="horz" lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
            f'<a:lstStyle>{lst}</a:lstStyle>{paras}</p:txBody></p:sp>')


def rect(sid, name, box, color):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr>{xfrm(*box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill(color)}<a:ln><a:noFill/></a:ln></p:spPr></p:sp>')


def rule_line(sid, name, x, y, w, color, width=12700):
    return (f'<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvCxnSpPr/><p:nvPr userDrawn="1"/></p:nvCxnSpPr>'
            f'<p:spPr>{xfrm(x, y, w, 0)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
            f'<a:ln w="{width}">{fill(color)}</a:ln></p:spPr></p:cxnSp>')


def logo_pic(sid, rid):
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{sid}" name="Logo" descr="Neuromine"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
            f'<p:nvPr userDrawn="1"/></p:nvPicPr><p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr>{xfrm(COVER_LEFT, COVER_LEFT, LOGO_H, LOGO_H)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')


def wordmark(sid):
    run = rpr("rPr", WORDMARK["slide_pt"], C[WORDMARK["color"]], "major", bold=WORDMARK["weight"] >= 700, spc=WORDMARK["letter_spacing"])
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="Wordmark"/><p:cNvSpPr txBox="1"/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr>{xfrm(WORDMARK_X, COVER_LEFT, inch(4), LOGO_H)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            '<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
            f'<a:p><a:pPr algn="l"><a:buNone/></a:pPr><a:r>{run}<a:t>{WORDMARK["text"]}</a:t></a:r></a:p></p:txBody></p:sp>')


def tree(shapes):
    return ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + "".join(shapes) + "</p:spTree>")


FOOT_Y = inch(6.88)
FOOT_H = inch(0.32)
NUM_W = inch(1.5)
FOOT_LST = plain_levels(PT["page_number"], C["caption"], MONO_FONT)
NUM_LST = plain_levels(PT["page_number"], C["placeholder"], MONO_FONT, algn="r")
COVER_LST = plain_levels(PT["eyebrow"], C["tint_3"], MONO_FONT)

TITLE_BOX = (LEFT, inch(0.45), TEXT_W, inch(1.35))
BODY_BOX = (LEFT, inch(2.05), TEXT_W, inch(4.45))
FOOT_BOX = (LEFT, FOOT_Y, TEXT_W - NUM_W, FOOT_H)
NUM_BOX = (LEFT + TEXT_W - NUM_W, FOOT_Y, NUM_W, FOOT_H)
DATE_BOX = (COVER_LEFT, inch(6.6), inch(6), FOOT_H)


def footer_phs():
    return [ph(90, "Footer", 'type="ftr" sz="quarter" idx="11"', FOOT_BOX, anchor="ctr"),
            ph(91, "Slide Number", 'type="sldNum" sz="quarter" idx="12"', NUM_BOX, anchor="ctr", text=SLIDE_NUM_TEXT)]


def master_xml(layout_rids):
    shapes = [
        rect(2, "Brand Strip", (0, 0, STRIP_W, SLIDE_H), C["brand"]),
        ph(3, "Title", 'type="title"', TITLE_BOX, anchor="b", text="Slide title"),
        ph(4, "Text", 'type="body" idx="1"', BODY_BOX, text="First level"),
        ph(5, "Date", 'type="dt" sz="half" idx="10"', DATE_BOX, lst=COVER_LST, anchor="ctr", text="Month Year"),
        ph(6, "Footer", 'type="ftr" sz="quarter" idx="11"', FOOT_BOX, lst=FOOT_LST, anchor="ctr", text="Deck name"),
        ph(7, "Slide Number", 'type="sldNum" sz="quarter" idx="12"', NUM_BOX, lst=NUM_LST, anchor="ctr", text=SLIDE_NUM_TEXT),
        rule_line(8, "Footer Rule", LEFT, inch(6.78), TEXT_W, C["rule"]),
    ]
    ids = "".join(f'<p:sldLayoutId id="{2147483649 + i}" r:id="{rid}"/>' for i, rid in enumerate(layout_rids))
    title_style = lvl(1, PT["headline"], C["brand"], "major", bold=True, line=95000)
    other = plain_levels(PT["command"], C["text"], "minor")
    return (f'<p:sldMaster {NS}><p:cSld><p:bg><p:bgPr>{fill("bg1")}<a:effectLst/></p:bgPr></p:bg>{tree(shapes)}</p:cSld>'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
            'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
            f'<p:sldLayoutIdLst>{ids}</p:sldLayoutIdLst>'
            '<p:hf hdr="0" dt="0"/>'
            f'<p:txStyles><p:titleStyle>{title_style}</p:titleStyle><p:bodyStyle>{body_levels()}</p:bodyStyle>'
            f'<p:otherStyle>{other}</p:otherStyle></p:txStyles></p:sldMaster>')


def layout_xml(name, ltype, shapes, brand_bg=False):
    type_attr = f' type="{ltype}"' if ltype else ""
    show = ' showMasterSp="0"' if brand_bg else ""
    bg = f'<p:bg><p:bgPr>{fill(C["brand"])}<a:effectLst/></p:bgPr></p:bg>' if brand_bg else ""
    return (f'<p:sldLayout {NS}{type_attr} preserve="1"{show}><p:cSld name="{name}">{bg}{tree(shapes)}</p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>')


def cover(logo_rid, title_type, title_text, sub_text, with_date):
    eyebrow_lst = plain_levels(PT["eyebrow"], C["tint_3"], "major", bold=True, caps=True, spc=EYEBROW_SPC)
    title_lst = lvl(1, PT["title_headline"], C["surface"], "major", bold=True, line=95000)
    sub_lst = plain_levels(PT["subline"], C["tint_3"], "minor", line=110000)
    shapes = [
        logo_pic(2, logo_rid),
        wordmark(8),
        ph(3, "Eyebrow", 'type="body" sz="quarter" idx="15"', (COVER_LEFT, inch(2.0), COVER_W, inch(0.4)), lst=eyebrow_lst, anchor="b", text="Program and event"),
        rect(4, "Cover Bar", (COVER_LEFT, inch(2.58), inch(0.9), inch(0.07)), C["tint_3"]),
        ph(5, "Title", f'type="{title_type}"', (COVER_LEFT, inch(2.85), COVER_W, inch(2.3)), lst=title_lst, text=title_text),
        ph(6, "Subtitle", ('type="subTitle" idx="1"' if with_date else 'type="body" idx="1"'), (COVER_LEFT, inch(5.3), COVER_W, inch(0.95)), lst=sub_lst, text=sub_text),
    ]
    if with_date:
        shapes.append(ph(7, "Date", 'type="dt" sz="half" idx="10"', DATE_BOX, lst=COVER_LST, anchor="ctr", text="Month Year"))
    return shapes


def layouts(logo_rid):
    sub_lst = plain_levels(PT["subline"], C["text_2"], "minor", line=110000)
    cap_lst = plain_levels(PT["eyebrow"], C["caption"], "minor")
    title = ph(2, "Title", 'type="title"', TITLE_BOX, anchor="b", text="Slide title")
    section = [
        rect(2, "Section Bar", (LEFT, inch(2.45), inch(0.9), inch(0.07)), C["brand"]),
        ph(3, "Title", 'type="title"', (LEFT, inch(2.7), TEXT_W, inch(1.35)), anchor="t", text="1. Section name"),
        ph(4, "Text", 'type="body" idx="1"', (LEFT, inch(4.2), inch(10), inch(1.1)), lst=sub_lst, text="One line on what the section covers"),
    ] + footer_phs()
    content = [title, ph(3, "Content", 'idx="1"', BODY_BOX, text="First level")] + footer_phs()
    half = (TEXT_W - inch(0.4)) // 2
    two = [
        title,
        ph(3, "Left Content", 'sz="half" idx="1"', (LEFT, BODY_BOX[1], half, BODY_BOX[3]), text="First level"),
        ph(4, "Right Content", 'sz="half" idx="2"', (LEFT + half + inch(0.4), BODY_BOX[1], half, BODY_BOX[3]), text="First level"),
    ] + footer_phs()
    table = [
        title,
        ph(3, "Caption", 'type="body" sz="quarter" idx="14"', (LEFT, BODY_BOX[1], TEXT_W, inch(0.4)), lst=cap_lst, anchor="ctr", text="Table 1. What the table shows"),
        ph(4, "Table", 'type="tbl" idx="13"', (LEFT, inch(2.6), TEXT_W, inch(3.9))),
    ] + footer_phs()
    figure_box = (LEFT, BODY_BOX[1], TEXT_W, inch(3.95))
    cap_box = (LEFT, inch(6.1), TEXT_W, inch(0.4))
    chart = [
        title,
        ph(3, "Chart", 'type="chart" idx="13"', figure_box),
        ph(4, "Caption", 'type="body" sz="quarter" idx="14"', cap_box, lst=cap_lst, anchor="ctr", text="Figure 1. What the chart shows. Source: the release"),
    ] + footer_phs()
    picture = [
        title,
        ph(3, "Picture", 'type="pic" idx="13"', figure_box),
        ph(4, "Caption", 'type="body" sz="quarter" idx="14"', cap_box, lst=cap_lst, anchor="ctr", text="Figure 1. What the picture shows and where it came from"),
    ] + footer_phs()
    return [
        layout_xml("Title", "title", cover(logo_rid, "ctrTitle", "Deck title", "One line on what the deck is", True), brand_bg=True),
        layout_xml("Section", "secHead", section),
        layout_xml("Title and Content", "obj", content),
        layout_xml("Two Content", "twoObj", two),
        layout_xml("Table", "tbl", table),
        layout_xml("Chart", "chart", chart),
        layout_xml("Picture", "picTx", picture),
        layout_xml("Closing", None, cover(logo_rid, "title", "Thank you", "Contact or next step", False), brand_bg=True),
    ]


def theme_blob(blob):
    root = etree.fromstring(blob)
    a = "{%s}" % A_NS
    root.set("name", "Neuromine")
    elems = root.find(f"{a}themeElements")
    clr = elems.find(f"{a}clrScheme")
    for child in list(clr):
        clr.remove(child)
    clr.set("name", "Neuromine")
    colors = [("dk1", C["ink"]), ("lt1", C["surface"]), ("dk2", C["text"]), ("lt2", C["page"])]
    colors += [(f"accent{i + 1}", c) for i, c in enumerate(CHART)]
    colors += [("hlink", C["brand"]), ("folHlink", C["brand_dark"])]
    for tag, value in colors:
        el = etree.SubElement(clr, f"{a}{tag}")
        etree.SubElement(el, f"{a}srgbClr", val=value)
    fonts = elems.find(f"{a}fontScheme")
    fonts.set("name", "Neuromine")
    for tag, face in (("majorFont", HEAD_FONT), ("minorFont", BODY_FONT)):
        group = fonts.find(f"{a}{tag}")
        for child in list(group):
            group.remove(child)
        etree.SubElement(group, f"{a}latin", typeface=face, pitchFamily="34", charset="0")
        etree.SubElement(group, f"{a}ea", typeface="")
        etree.SubElement(group, f"{a}cs", typeface="")
    elems.find(f"{a}fmtScheme").set("name", "Neuromine")
    return flat_effects(etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True))


def flat_effects(blob):
    root = etree.fromstring(blob)
    a = "{%s}" % A_NS
    fmt = root.find(f"{a}themeElements").find(f"{a}fmtScheme")
    for grad in list(fmt.iter(f"{a}gradFill")):
        solid = etree.Element(f"{a}solidFill")
        etree.SubElement(solid, f"{a}schemeClr", val="phClr")
        grad.getparent().replace(grad, solid)
    for effect in fmt.iter(f"{a}effectStyle"):
        for child in list(effect):
            effect.remove(child)
        etree.SubElement(effect, f"{a}effectLst")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def table_styles_blob():
    none = "<a:ln><a:noFill/></a:ln>"
    border = f'<a:ln w="12700">{fill(C["table_border"])}</a:ln>'
    head_rule = f'<a:ln w="12700">{fill(C["table_head_rule"])}</a:ln>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:tblStyleLst xmlns:a="{A_NS}" def="{TABLE_STYLE_ID}">'
            f'<a:tblStyle styleId="{TABLE_STYLE_ID}" styleName="{TABLE_STYLE_NAME}">'
            f'<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:srgbClr val="{C["text_2"]}"/></a:tcTxStyle>'
            f'<a:tcStyle><a:tcBdr><a:left>{border}</a:left><a:right>{border}</a:right><a:top>{border}</a:top>'
            f'<a:bottom>{border}</a:bottom><a:insideH>{border}</a:insideH><a:insideV>{none}</a:insideV></a:tcBdr>'
            f'<a:fill>{fill(C["surface"])}</a:fill></a:tcStyle></a:wholeTbl>'
            f'<a:firstRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:srgbClr val="{C["ink"]}"/></a:tcTxStyle>'
            f'<a:tcStyle><a:tcBdr><a:bottom>{head_rule}</a:bottom></a:tcBdr><a:fill>{fill(C["table_head"])}</a:fill></a:tcStyle></a:firstRow>'
            '</a:tblStyle></a:tblStyleLst>').encode("utf-8")


def stable_xlsx(data):
    with zipfile.ZipFile(io.BytesIO(data)) as zin:
        items = [(info.filename, zin.read(info.filename)) for info in zin.infolist()]
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, body in items:
            if name == "docProps/core.xml":
                body = re.sub(rb"[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z", b"2026-01-01T00:00:00Z", body)
            zout.writestr(zipfile.ZipInfo(name, ZIP_DATE), body, compress_type=zipfile.ZIP_DEFLATED)
    return out.getvalue()


def save_stable(prs, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    with zipfile.ZipFile(path) as zin:
        items = [(info.filename, zin.read(info.filename)) for info in zin.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in items:
            if name.startswith("ppt/theme/") and name != "ppt/theme/theme1.xml":
                data = flat_effects(data)
            if name.endswith(".xlsx"):
                data = stable_xlsx(data)
            if name == "docProps/app.xml":
                data = data.replace(b"On-screen Show (4:3)", b"Widescreen").replace(b"Office Theme", b"Neuromine")
            zout.writestr(zipfile.ZipInfo(name, ZIP_DATE), data, compress_type=zipfile.ZIP_DEFLATED)


def drop_slides(prs):
    ids = prs.slides._sldIdLst
    for sld in list(ids):
        prs.part.drop_rel(sld.rId)
        ids.remove(sld)


def add(prs, name, title, foot, notes=None):
    slide = prs.slides.add_slide(layout_named(prs, name))
    slide.shapes.title.text = title
    if name not in ("Title", "Closing"):
        add_hf(slide, {"ftr", "sldNum"}, foot)
    if notes:
        slide.notes_slide.notes_text_frame.text = notes
    return slide


def master_frame():
    prs = pptx.Presentation()
    prs.slide_width = Emu(SLIDE_W)
    prs.slide_height = Emu(SLIDE_H)
    prs.element.find(qn("p:sldSz")).attrib.pop("type", None)
    pres_part = prs.part
    master_rid = prs.element.find(qn("p:sldMasterIdLst"))[0].get(qn("r:id"))
    master_part = pres_part.related_part(master_rid)
    old = master_part._element.find(qn("p:sldLayoutIdLst"))
    rids = [el.get(qn("r:id")) for el in old]
    keep = rids[:len(LAYOUTS)]
    for rid in rids[len(LAYOUTS):]:
        master_part.drop_rel(rid)
    parts = [master_part.related_part(rid) for rid in keep]
    logo_rids = {}
    for name, part in zip(LAYOUTS, parts):
        if name in ("Title", "Closing"):
            logo_rids[name] = part.get_or_add_image_part(str(LOGO))[1]
    xmls = layouts("LOGO_RID")
    for i, (name, part, xml) in enumerate(zip(LAYOUTS, parts, xmls), 1):
        part.partname = PackURI(f"/ppt/slideLayouts/slideLayout{i}.xml")
        part._element = parse_xml(xml.replace("LOGO_RID", logo_rids.get(name, "")))
    master_part._element = parse_xml(master_xml(keep))
    theme = master_part.part_related_by(RT.THEME)
    theme._blob = theme_blob(theme.blob)
    pres_part.part_related_by(RT.TABLE_STYLES)._blob = table_styles_blob()
    return prs


def build_template(path):
    prs = master_frame()
    foot = {"ftr": "Deck name", "dt": "Month Year"}
    s = add(prs, "Title", "Deck title: the point in under ten words", foot,
            "Title layout, always the first slide. Brand blue with the white logo top left; do not move or add a logo. "
            "Fill the eyebrow (idx 15) with the program and event, the title, the subline (idx 1) and the date with add_hf(slide, {'dt'}, values).")
    s.placeholders[15].text = "Program and event, two to four words"
    s.placeholders[1].text = "Subline: who, what and when, in one line"
    add_hf(s, {"dt"}, foot)
    add(prs, "Section", "1. Section name", foot,
        "Section layout: starts each section. Title 'N. Name' and one line in idx 1. Use it for a one line statement too.").placeholders[1].text = "One line on what the section covers"
    s = add(prs, "Title and Content", "Headline states the point in one sentence", foot,
            "Title and Content: most slides. Fill idx 1 with set_lines; at least three items, at most six, two levels. "
            "Use set_lines(frame, lines, numbered=True) for steps; never type the numbers.")
    set_lines(s.placeholders[1].text_frame, [
        (0, "Fill this box with set_lines, one item per line"),
        (1, "Level two sits under its level one item"),
        (0, "At most six items and two levels"),
        (0, "About ten words per item, no full stop"),
    ])
    s = add(prs, "Two Content", "Two boxes side by side", foot,
            "Two Content: idx 1 left and idx 2 right. Put a table or chart beside its points, or points beside a callout. "
            "add_callout(slide, placeholder, text, align='middle') sits level with a table or chart.")
    set_lines(s.placeholders[1].text_frame, [(0, "Left box: points, a table or a chart"), (0, "Keep both sides about the same length"), (0, "One idea per side")])
    add_callout(s, s.placeholders[2], "Right box: one callout for the point the reader must take away.")
    s = add(prs, "Table", "Headline says what the table shows", foot,
            "Table layout: caption above (idx 14), 'Table N. What it shows.' Fill the table box (idx 13) with add_table: "
            "at most 7 rows and 5 columns, plain word headers, no merged cells. Numbers and their headers are right aligned.")
    s.placeholders[14].text = "Table 1. What each column holds. Source: name the release"
    add_table(s, s.placeholders[13], [
        ["Column", "What goes here", "Words"],
        ["Label", "Plain words, never a raw column name", "3"],
        ["Count", "Numbers in one unit per column", "1"],
        ["Release", "The release name, such as Release 8", "2"],
    ])
    s = add(prs, "Chart", "Headline states what the chart shows", foot,
            "Chart layout: one native chart per slide with add_chart(slide.placeholders[13], chart_type, data, palette). "
            "Set the alt text to the caption. Caption in idx 14 names the figure and the release. Never paste a chart image.")
    s.placeholders[14].text = "Figure 1. What the chart shows. Source: name the release"
    s = add(prs, "Picture", "Headline says why the picture matters", foot,
            "Picture layout: one PNG or JPG with add_image(slide, path, caption); it fits the box, keeps the aspect ratio "
            "and sets the alt text and the caption. No borders or shadows. Never a logo on a white slide.")
    s.placeholders[14].text = "Figure 1. What the picture shows and where it came from"
    s = add(prs, "Closing", "Closing: one next step for the audience", foot,
            "Closing layout, always the last slide. Brand blue with the white logo top left. Eyebrow (idx 15), "
            "title and one line (idx 1) with the contact or next step. No footer.")
    s.placeholders[15].text = "Program and event"
    s.placeholders[1].text = "Who to contact and how, or the next step"
    set_core(prs, "Neuromine template")
    save_stable(prs, path)


def swatch_png():
    plt.rcParams["font.family"] = [TOKENS["font"]["fallback"], "sans-serif"]
    fig, ax = plt.subplots(figsize=(10, 3.95), dpi=200)
    fig.patch.set_alpha(0)
    ax.set_xlim(0, 6)
    ax.set_ylim(0, 1)
    ax.axis("off")
    for i, color in enumerate(CHART):
        ax.add_patch(matplotlib.patches.Rectangle((i + 0.08, 0.3), 0.84, 0.62, color="#" + color, linewidth=0))
        ax.text(i + 0.5, 0.17, f"{i + 1}", ha="center", va="center", fontsize=16, color="#" + C["chart_axis"])
        ax.text(i + 0.5, 0.04, "#" + color, ha="center", va="center", fontsize=12, color="#" + C["chart_axis"])
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, transparent=True, metadata={"Software": None})
    plt.close(fig)
    buf.seek(0)
    return buf


def build_sample(template, path):
    prs = pptx.Presentation(str(template))
    drop_slides(prs)
    deck = "Building a Neuromine deck"
    foot = {"ftr": deck, "dt": "November 2026"}

    s = add(prs, "Title", deck, foot)
    s.placeholders[15].text = "Neuromine brand skill"
    s.placeholders[1].text = "The template, the helpers and the checker"
    add_hf(s, {"dt"}, foot)

    s = add(prs, "Section", "1. Start from the template", foot)
    s.placeholders[1].text = "Every deck opens neuromine-template.pptx and adds slides by layout."

    s = add(prs, "Title and Content", "Four steps take a deck from template to send", foot)
    set_lines(s.placeholders[1].text_frame, [
        (0, "Open the template and drop its guide slides"),
        (0, "Add slides from the eight layouts by name"),
        (1, "Title first, Closing last"),
        (0, "Fill text, tables, charts and pictures with the helpers"),
        (0, "Run check_deck.py and fix every FAIL"),
    ], numbered=True)

    s = add(prs, "Table", "Slide text comes in six sizes", foot)
    s.placeholders[14].text = "Table 1. Text sizes on a slide. Source: references/brand.md"
    add_table(s, s.placeholders[13], [
        ["Text", "Font", "Size (pt)"],
        ["Title slide headline", "Helvetica Neue bold", str(PT["title_headline"])],
        ["Content slide headline", "Helvetica Neue bold", str(PT["headline"])],
        ["Subline and first level text", "Helvetica Neue", str(PT["subline"])],
        ["Second level, tables and chart labels", "Helvetica Neue", str(PT["command"])],
        ["Eyebrow and captions", "Helvetica Neue", str(PT["eyebrow"])],
        ["Footer and page number", "Courier New", str(PT["page_number"])],
    ])

    s = add(prs, "Two Content", "Blue carries the brand on every slide", foot)
    set_lines(s.placeholders[1].text_frame, [
        (0, "Title and closing slides are brand blue all over"),
        (0, "Content slides are white with a blue strip at left"),
        (0, "Headlines are brand blue at 40 pt"),
    ])
    add_callout(s, s.placeholders[2], "The white logo sits only on brand blue, top left, at 0.75 in high.", align="middle")

    s = add(prs, "Chart", "Title headlines are over twice the subline size", foot)
    data = CategoryChartData()
    roles = ["Title headline", "Headline", "Subline", "Labels", "Eyebrow", "Page number"]
    keys = ["title_headline", "headline", "subline", "command", "eyebrow", "page_number"]
    data.categories = roles
    data.add_series("Size (pt)", [PT[k] for k in keys])
    caption = "Figure 1. Slide text sizes by role, in points. Source: references/brand.md"
    frame = add_chart(s.placeholders[13], XL_CHART_TYPE.COLUMN_CLUSTERED, data, palette="single")
    frame._element.nvGraphicFramePr.cNvPr.set("descr", caption)
    s.placeholders[14].text = caption

    s = add(prs, "Picture", "Six chart colors, always in this order", foot)
    add_image(s, swatch_png(), "Figure 2. The six categorical chart colors in order. Source: references/brand.md")

    s = add(prs, "Closing", "Run check_deck.py before you send", foot)
    s.placeholders[15].text = "Neuromine brand skill"
    s.placeholders[1].text = "Rules: references/powerpoint.md"

    set_core(prs, deck)
    save_stable(prs, path)


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
    build_sample(template, sample)
    for p in (template, sample):
        print(f"wrote {p.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
