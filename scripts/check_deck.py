import re
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import prose

try:
    import pptx
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-pptx"])
    import pptx

from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

TABLE_STYLE_ID = "{1F4FAE00-A1A5-4A15-8000-000000000001}"
TABLE_STYLE_NAME = "Neuromine Table"
LAYOUTS = ["Title", "Section", "Title and Content", "Two Content", "Table", "Chart", "Picture", "Closing"]
NO_FOOTER = ("Title", "Closing")
NUMBER_RE = re.compile(r"^[+-]?\$?[+-]?\d[\d,]*(\.\d+)?\s?(%|pts|pt)?$")
LETTER_RE = re.compile(r"[A-Za-z]")
EFFECT_TAGS = ("a:gradFill", "a:outerShdw", "a:innerShdw", "a:prstShdw", "a:reflection", "a:glow", "a:softEdge")
EXEMPT_PH = (PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE)
EYEBROW_IDX = 15
MAX_ROWS = 7
MAX_COLS = 5
MAX_BULLETS = 6
MAX_TITLE = 55
MIN_FILL = 3
FILL_LAYOUTS = ("Title and Content", "Two Content")
BRAND = "1F4FAE"
TYPED_NUM_RE = re.compile(r"^\s*\d+[.)]\s")
ROW_MAX = 685800
ROW_TOL = 12700


def layout_name(slide):
    return slide.slide_layout.name


def ph_type(shape):
    return shape.placeholder_format.type if shape.is_placeholder else None


def short(text):
    return text if len(text) <= 30 else text[:27] + "..."


def check_style(prs):
    blob = prs.part.part_related_by(RT.TABLE_STYLES).blob.decode("utf-8", "replace")
    if f'styleId="{TABLE_STYLE_ID}"' not in blob or f'styleName="{TABLE_STYLE_NAME}"' not in blob:
        return False, "tableStyles.xml lacks the Neuromine Table style; start from assets/neuromine-template.pptx"
    return True, "tableStyles.xml defines Neuromine Table"


def check_layouts(prs):
    bad = [f"slide {i} '{layout_name(s)}'" for i, s in enumerate(prs.slides, 1) if layout_name(s) not in LAYOUTS]
    if bad:
        return False, "layouts outside the template: " + "; ".join(bad)
    return True, "every slide uses a skill layout"


def check_ends(prs):
    slides = list(prs.slides)
    if not slides:
        return False, "deck has no slides"
    first, last = layout_name(slides[0]), layout_name(slides[-1])
    if first != "Title" or last != "Closing":
        return False, f"first slide is '{first}', last is '{last}'; want Title and Closing"
    return True, "first slide is Title, last is Closing"


def check_footer(prs):
    bad = []
    for i, slide in enumerate(prs.slides, 1):
        if layout_name(slide) in NO_FOOTER:
            continue
        types = {ph_type(s): s for s in slide.placeholders}
        foot = types.get(PP_PLACEHOLDER.FOOTER)
        if foot is None or not foot.text_frame.text.strip():
            bad.append(f"slide {i} footer")
        if PP_PLACEHOLDER.SLIDE_NUMBER not in types:
            bad.append(f"slide {i} slide number")
    if bad:
        return False, "missing: " + "; ".join(bad)
    return True, "every content slide has a footer and slide number"


def check_label_fonts(prs):
    bad = []
    owners = [("master", prs.slide_master)] + [(f"layout '{l.name}'", l) for l in prs.slide_layouts]
    owners += [(f"slide {i}", s) for i, s in enumerate(prs.slides, 1)]
    for where, owner in owners:
        for shape in owner.placeholders:
            pf = shape.placeholder_format
            if pf.type not in EXEMPT_PH and pf.idx != EYEBROW_IDX:
                continue
            if any("courier" in (el.get("typeface") or "").lower() for el in shape._element.iter()):
                bad.append(f"{where} {shape.name}")
    if bad:
        return False, "Courier New in " + "; ".join(bad[:5]) + "; the footer, page number, date and eyebrow use the body family"
    return True, "no footer, page number, date or eyebrow placeholder names Courier New"


def check_direct(prs):
    bad = []
    for i, slide in enumerate(prs.slides, 1):
        for shape in slide.placeholders:
            if ph_type(shape) in EXEMPT_PH or not shape.has_text_frame:
                continue
            for r_pr in shape.text_frame._txBody.iter(qn("a:rPr")):
                tags = []
                if r_pr.get("sz") is not None:
                    tags.append("sz")
                if r_pr.find(qn("a:latin")) is not None:
                    tags.append("latin")
                if r_pr.find(qn("a:solidFill")) is not None:
                    tags.append("solidFill")
                if tags:
                    text = "".join(t.text or "" for t in r_pr.getparent().iter(qn("a:t")))
                    bad.append(f"slide {i} '{short(text)}' ({', '.join(tags)})")
    if bad:
        return False, "direct run formatting: " + "; ".join(bad[:5])
    return True, "no placeholder run has direct latin, sz or solidFill"


def tables(prs):
    for i, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if shape.has_table:
                yield i, shape.table


def content_box(slide, shape):
    boxes = [p for p in slide.slide_layout.placeholders if p.placeholder_format.idx in (1, 2)]
    if not boxes:
        return None
    return min(boxes, key=lambda p: abs(p.left - shape.left) + abs(p.top - shape.top))


def short_rows(prs):
    bad = []
    for i, slide in enumerate(prs.slides, 1):
        if layout_name(slide) not in FILL_LAYOUTS:
            continue
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            rows = shape.table.rows
            box = content_box(slide, shape)
            if box is None or len(rows) >= MAX_ROWS:
                continue
            want = min(ROW_MAX, box.height // len(rows))
            low = [r.height for r in rows if r.height < want - ROW_TOL]
            if low:
                bad.append(f"slide {i} rows at {low[0] / 914400:.2f} in, want {want / 914400:.2f} in to fill the box")
    return bad


def check_tables(prs):
    bad = []
    count = 0
    for i, table in tables(prs):
        count += 1
        tbl = table._tbl
        style = tbl.tblPr.find(qn("a:tableStyleId"))
        if style is None or style.text != TABLE_STYLE_ID:
            bad.append(f"slide {i} not Neuromine Table")
        rows, cols = len(table.rows), len(table.columns)
        if rows > MAX_ROWS or cols > MAX_COLS:
            bad.append(f"slide {i} is {rows}x{cols}")
        for tc in tbl.iter(qn("a:tc")):
            if any(tc.get(a) for a in ("gridSpan", "rowSpan", "hMerge", "vMerge")):
                bad.append(f"slide {i} has merged cells")
                break
        for r in range(1, rows):
            for c in range(1, cols):
                cell = table.cell(r, c)
                if not NUMBER_RE.match(cell.text.strip()):
                    continue
                if any(p.alignment != PP_ALIGN.RIGHT for p in cell.text_frame.paragraphs):
                    bad.append(f"slide {i} '{cell.text}' not right aligned")
        for c in range(1, cols):
            body = [table.cell(r, c).text.strip() for r in range(1, rows)]
            if not body or not all(NUMBER_RE.match(t) for t in body):
                continue
            head = table.cell(0, c)
            if any(p.alignment != PP_ALIGN.RIGHT for p in head.text_frame.paragraphs):
                bad.append(f"slide {i} header cell '{short(head.text)}' over a number column not right aligned")
    bad.extend(short_rows(prs))
    if bad:
        return False, "; ".join(bad[:5])
    return True, f"{count} tables use Neuromine Table, fit 7x5, no merges, numbers and their header cells right aligned, short tables fill the box"


def check_bullets(prs):
    bad = []
    for i, slide in enumerate(prs.slides, 1):
        for shape in slide.placeholders:
            if ph_type(shape) != PP_PLACEHOLDER.OBJECT or not shape.has_text_frame:
                continue
            paras = [p for p in shape.text_frame.paragraphs if p.text.strip()]
            if len(paras) > MAX_BULLETS:
                bad.append(f"slide {i} has {len(paras)} bullets")
            deep = [p for p in paras if p.level >= 2]
            if deep:
                bad.append(f"slide {i} '{short(deep[0].text)}' at level {deep[0].level}")
    if bad:
        return False, "; ".join(bad)
    return True, "bullets keep to six items and two levels"


def check_numbering(prs):
    bad = []
    for i, slide in enumerate(prs.slides, 1):
        for shape in slide.placeholders:
            if ph_type(shape) in EXEMPT_PH or not shape.has_text_frame:
                continue
            section_title = slide.slide_layout.name == "Section" and shape == slide.shapes.title
            for p in shape.text_frame.paragraphs:
                ppr = p._p.pPr
                auto = ppr is not None and ppr.find(qn("a:buAutoNum")) is not None
                if auto:
                    clr = ppr.find(qn("a:buClr"))
                    srgb = clr.find(qn("a:srgbClr")) if clr is not None else None
                    if srgb is None or (srgb.get("val") or "").upper() != BRAND:
                        bad.append(f"slide {i} '{short(p.text)}' numbered without #1F4FAE buClr")
                if p.level == 0 and TYPED_NUM_RE.match(p.text) and not section_title:
                    bad.append(f"slide {i} '{short(p.text)}' has a typed number")
    if bad:
        return False, "; ".join(bad[:5])
    return True, "numbered lists use buAutoNum in #1F4FAE and no typed numbers"


def check_titles(prs):
    bad = []
    for i, slide in enumerate(prs.slides, 1):
        title = slide.shapes.title
        if title is not None and len(title.text_frame.text) > MAX_TITLE:
            bad.append(f"slide {i} is {len(title.text_frame.text)} characters")
    if bad:
        return False, f"titles over {MAX_TITLE} characters: " + "; ".join(bad)
    return True, f"every title is at most {MAX_TITLE} characters"


def caption_text(slide):
    for shape in slide.placeholders:
        if shape.placeholder_format.idx == 14 and shape.has_text_frame:
            return shape.text_frame.text.strip()
    return ""


def check_figures(prs):
    bad = []
    for i, slide in enumerate(prs.slides, 1):
        name = layout_name(slide)
        if name == "Chart":
            frames = [s for s in slide.shapes if getattr(s, "has_chart", False) and s.has_chart]
            if not frames:
                bad.append(f"slide {i} has no native chart")
            elif not all(s._element.nvGraphicFramePr.cNvPr.get("descr", "").strip() for s in frames):
                bad.append(f"slide {i} chart has no alt text")
        elif name == "Table":
            if not any(getattr(s, "has_table", False) and s.has_table for s in slide.shapes):
                bad.append(f"slide {i} has no table")
        elif name == "Picture":
            pics = [s for s in slide.shapes if s._element.tag == qn("p:pic")]
            if not pics:
                bad.append(f"slide {i} has no picture")
            elif not all(s._element.nvPicPr.cNvPr.get("descr", "").strip() for s in pics):
                bad.append(f"slide {i} picture has no alt text")
        else:
            continue
        if not caption_text(slide):
            bad.append(f"slide {i} has no caption")
    if bad:
        return False, "; ".join(bad)
    return True, "chart and picture slides hold their object with alt text, table slides hold a table, all three have captions"


def check_fill(prs):
    bad = []
    for i, slide in enumerate(prs.slides, 1):
        if layout_name(slide) not in FILL_LAYOUTS:
            continue
        paras = 0
        visual = False
        for shape in slide.shapes:
            if shape == slide.shapes.title or ph_type(shape) in EXEMPT_PH:
                continue
            if shape._element.tag == qn("p:pic") or getattr(shape, "has_table", False) and shape.has_table:
                visual = True
            elif getattr(shape, "has_chart", False) and shape.has_chart:
                visual = True
            elif shape.has_text_frame:
                paras += sum(1 for p in shape.text_frame.paragraphs if p.text.strip())
        if not visual and paras < MIN_FILL:
            bad.append(f"slide {i} '{short(slide.shapes.title.text if slide.shapes.title else '')}' holds {paras} paragraph{'' if paras == 1 else 's'} and no table, chart or picture")
    if bad:
        return False, "; ".join(bad)
    return True, "every Title and Content or Two Content slide holds three paragraphs, a table, a chart or a picture"


def check_effects(prs):
    bad = []
    parts = [prs.slide_master.part, prs.slide_master.part.part_related_by(RT.THEME)]
    parts += [l.part for l in prs.slide_layouts] + [s.part for s in prs.slides]
    for part in parts:
        blob = part.blob.decode("utf-8", "replace") if hasattr(part, "_blob") else None
        root = part._element if blob is None else None
        if root is not None:
            found = [t for t in EFFECT_TAGS if root.find(".//" + qn(t)) is not None]
        else:
            found = [t for t in EFFECT_TAGS if "<" + t in blob]
        if found:
            bad.append(f"{part.partname} has {', '.join(found)}")
    if bad:
        return False, "; ".join(bad[:5])
    return True, "no gradient, shadow, glow or reflection in the theme, master, layouts or slides"


def check_name(path):
    if " " in path.name:
        return False, f"file name '{path.name}' has spaces"
    return True, f"file name '{path.name}' has no spaces"


def prose_shapes(shapes):
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from prose_shapes(s.shapes)
        elif s.has_text_frame:
            yield s.text_frame.text
        elif getattr(s, "has_table", False) and s.has_table:
            for row in s.table.rows:
                for cell in row.cells:
                    if LETTER_RE.search(cell.text):
                        yield cell.text


def prose_items(prs):
    for i, slide in enumerate(prs.slides, 1):
        for text in prose_shapes(slide.shapes):
            yield f"slide {i}", text
        if slide.has_notes_slide:
            yield f"slide {i} notes", slide.notes_slide.notes_text_frame.text


def main():
    if len(sys.argv) != 2:
        print("usage: python scripts/check_deck.py path/to/file.pptx")
        return 2
    path = Path(sys.argv[1])
    prs = pptx.Presentation(str(path))
    checks = [
        ("template", lambda: check_style(prs)),
        ("layouts", lambda: check_layouts(prs)),
        ("first-last", lambda: check_ends(prs)),
        ("footer", lambda: check_footer(prs)),
        ("label-fonts", lambda: check_label_fonts(prs)),
        ("direct-format", lambda: check_direct(prs)),
        ("tables", lambda: check_tables(prs)),
        ("bullets", lambda: check_bullets(prs)),
        ("numbering", lambda: check_numbering(prs)),
        ("titles", lambda: check_titles(prs)),
        ("charts-pictures", lambda: check_figures(prs)),
        ("fill", lambda: check_fill(prs)),
        ("effects", lambda: check_effects(prs)),
        ("file-name", lambda: check_name(path)),
    ]
    failed = False
    for name, fn in checks:
        ok, detail = fn()
        print(f"{'PASS' if ok else 'FAIL'} {name}: {detail}")
        failed = failed or not ok
    hits = [(where, word, snip) for where, text in prose_items(prs) for word, snip in prose.find_banned(text)]
    for where, word, snip in hits:
        print(f"FAIL prose: {where}: {word} in '{snip}'")
    if not hits:
        print("PASS prose: no em dash, en dash or banned word in slide text, text cells and notes")
    failed = failed or bool(hits)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
