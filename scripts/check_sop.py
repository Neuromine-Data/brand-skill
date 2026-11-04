import hashlib
import re
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import prose

try:
    import docx
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
    import docx

from docx.oxml.ns import qn

REQUIRED_STYLES = ["Normal", "Title", "Header", "Footer", "NeuromineTable", "SOPHeading", "SOPSection", "SOPStep", "SOPTerm", "SOPDetail"]
SECTIONS = ["Purpose", "Scope", "Definitions and Acronyms", "Procedures", "References"]
TITLE_FIELDS = ["Title", "Version Number", "Effective Date"]
REVISION_FIELDS = ["Version Number", "Effective Date", "Description"]
DIRECT_TAGS = ("w:rFonts", "w:sz", "w:color")
TYPED_NUMBER_RE = re.compile(r"^\s*\d+\.(\d+\.?)*\s")
DATE_RE = re.compile(r"^[0-9]{4}-[0-9]{2}-[0-9]{2}$")
PREFIXES = ("AALS-", "ALSTDI-", "NEUROMINE-")
SOP_ID_RE = re.compile(r"^(" + "|".join(PREFIXES) + r")SOP-(NNN|[0-9]{3})-V[0-9]+$")
FILE_RE = re.compile(r"^(" + "|".join(PREFIXES) + r")SOP-(NNN|[0-9]{3})-V[0-9]+\. \S.*\.docx$")
SAMPLE_NAMES = ("neuromine-sop-sample.docx",)
GUIDE = "Guide:"
WORD_MIN = 700
WORD_BUDGET = 1000
WORD_CAP = 1100
IDENTIFIER_RE = re.compile(r"CASE-[A-Z0-9]{6,}|CTRL-[A-Z0-9]{6,}|ARC[0-9]{6}|NEU(?!ROMINE(?![A-Z0-9]))[A-Z0-9]{6,}")
NAME_HASHES = {
    "265ce76730503a448331353d6562c3d2092636a9da448c246d6366abd19bddb7",
    "2a0213c7d9741c04aebf6d2256c6684310acb7315374b510ce3cd908eafdeb12",
    "5a20b06e006d26af90af53e74002af21c576e616ffd65b02d587a29e34240976",
    "6e4fa02c8da103154125c05f55a5e86ca6724c6ea981dddb69e234497aee5c84",
    "cd1591a63d28e6ea8a57aec8d4d5d57494990f5a49fb29d34f15405c5d7f9eb9",
    "ea2d053b245816fe3c133539235d5e9f16bb48b91bfad41967789e1377dd909a",
}


def name_hash(words):
    return hashlib.sha256(" ".join(words).lower().encode("ascii", "replace")).hexdigest()


def para_text(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def style_of(p):
    s = p.find(qn("w:pPr") + "/" + qn("w:pStyle"))
    return "Normal" if s is None else s.get(qn("w:val"))


def body_paras(doc):
    return [p for p in doc.element.body if p.tag == qn("w:p")]


def styles_by_id(doc):
    return {s.get(qn("w:styleId")): s for s in doc.styles.element.findall(qn("w:style"))}


def num_id(doc, p):
    styles = styles_by_id(doc)
    num = p.find(qn("w:pPr") + "/" + qn("w:numPr") + "/" + qn("w:numId"))
    sid = style_of(p)
    seen = set()
    while num is None and sid in styles and sid not in seen:
        seen.add(sid)
        st = styles[sid]
        num = st.find(qn("w:pPr") + "/" + qn("w:numPr") + "/" + qn("w:numId"))
        based = st.find(qn("w:basedOn"))
        sid = based.get(qn("w:val")) if based is not None else None
    if num is None or num.get(qn("w:val")) == "0":
        return None
    return num.get(qn("w:val"))


def all_texts(doc):
    for p in doc.element.body.iter(qn("w:p")):
        yield para_text(p)
    for sec in doc.sections:
        for part in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer):
            if not part.is_linked_to_previous:
                for p in part._element.iter(qn("w:p")):
                    yield para_text(p)


def sections(doc):
    return [(i, para_text(p).strip()) for i, p in enumerate(body_paras(doc)) if style_of(p) == "SOPSection"]


def section_range(doc, name):
    found = sections(doc)
    names = [n for _, n in found]
    if name not in names:
        return []
    k = names.index(name)
    start = found[k][0] + 1
    end = found[k + 1][0] if k + 1 < len(found) else len(body_paras(doc))
    return body_paras(doc)[start:end]


def table_rows(tbl):
    return [[para_text(tc).strip() for tc in tr.findall(qn("w:tc"))] for tr in tbl.findall(qn("w:tr"))]


def check_styles(doc):
    ids = set(styles_by_id(doc))
    missing = [s for s in REQUIRED_STYLES if s not in ids]
    if missing:
        return False, "styles.xml lacks " + ", ".join(missing) + "; start from neuromine-sop-template.docx"
    return True, "styles.xml defines every Neuromine SOP template style"


def check_direct(doc):
    bad = []
    for r in doc.element.body.iter(qn("w:r")):
        r_pr = r.find(qn("w:rPr"))
        if r_pr is None:
            continue
        tags = [t.split(":")[1] for t in DIRECT_TAGS if r_pr.find(qn(t)) is not None]
        if tags:
            bad.append(f"'{para_text(r)[:30]}' ({', '.join(tags)})")
    if bad:
        return False, f"{len(bad)} runs with direct formatting: " + "; ".join(bad[:5])
    return True, "no body run has direct rFonts, sz or color"


def check_title_table(doc):
    tbls = doc.element.body.findall(qn("w:tbl"))
    if not tbls:
        return False, "no title table"
    rows = table_rows(tbls[0])
    if not rows or rows[0][:3] != TITLE_FIELDS or len(rows) < 2:
        return False, f"title table header row is {rows[0] if rows else []}, expected {TITLE_FIELDS} over a value row"
    empty = [TITLE_FIELDS[c] for c, v in enumerate(rows[1][:3]) if not v or v.startswith(GUIDE)]
    if empty:
        return False, "title table cells not filled: " + ", ".join(empty)
    if not DATE_RE.match(rows[1][2]):
        return False, f"title table Effective Date '{rows[1][2]}' is not YYYY-MM-DD"
    return True, f"title table filled: '{rows[1][0]}', version {rows[1][1]}, {rows[1][2]}"


def check_guides(doc):
    left = [t.strip()[:40] for t in all_texts(doc) if t.strip().startswith(GUIDE)]
    if left:
        return False, f"{len(left)} 'Guide:' lines left, first '{left[0]}'"
    return True, "no 'Guide:' line left"


def check_revision(doc):
    tbls = doc.element.body.findall(qn("w:tbl"))
    if len(tbls) < 2:
        return False, "no revision history table"
    rows = table_rows(tbls[1])
    if not rows or rows[0][:3] != REVISION_FIELDS:
        return False, f"revision history header row is {rows[0] if rows else []}, expected {REVISION_FIELDS}"
    col = rows[0].index("Effective Date")
    dated = [r for r in rows[1:] if len(r) > col and DATE_RE.match(r[col]) and all(r[:3])]
    if not dated:
        return False, "revision history has no filled row with an Effective Date as YYYY-MM-DD"
    return True, f"revision history rows with a date: {len(dated)}"


def check_sections(doc):
    found = sections(doc)
    names = [n for _, n in found]
    if names != SECTIONS:
        return False, f"SOP Section lines are {names}, expected {SECTIONS}"
    paras = body_paras(doc)
    unnumbered = [n for i, n in found if num_id(doc, paras[i]) is None]
    if unnumbered:
        return False, "sections without automatic numbering: " + ", ".join(unnumbered)
    return True, "the five sections in order, each automatically numbered"


def check_typed(doc):
    bad = [t.strip()[:30] for t in all_texts(doc) if TYPED_NUMBER_RE.match(t)]
    if bad:
        return False, f"typed numbers at the start of {len(bad)} lines: " + "; ".join(bad[:5])
    return True, "no line starts with a typed number such as 1. or 4.1"


def check_heading_numbers(doc):
    bad = []
    for p in body_paras(doc):
        sid = style_of(p)
        if (sid.startswith("Heading") or sid == "Title") and num_id(doc, p) is not None:
            bad.append(f"'{para_text(p)[:30]}' ({sid})")
    if bad:
        return False, "Heading style on a numbered line: " + "; ".join(bad[:5])
    return True, "no Heading or Title style on a numbered line"


def check_procedures(doc):
    lines = section_range(doc, "Procedures")
    steps = []
    for p in lines:
        if style_of(p) == "SOPStep":
            steps.append([para_text(p).strip(), 0])
        elif steps and style_of(p) == "SOPDetail" and para_text(p).strip().startswith("Checks:"):
            steps[-1][1] += 1
    if not steps:
        return False, "Procedures has no SOP Step line"
    unchecked = [s for s, n in steps if n == 0]
    if unchecked:
        return False, "steps with no Checks line: " + "; ".join(unchecked)
    return True, f"Procedures has {len(steps)} steps, each with a Checks line"


def body_words(doc):
    found = sections(doc)
    start_el = body_paras(doc)[found[0][0]] if found else None
    words = 0
    counting = start_el is None
    for el in doc.element.body:
        if el is start_el:
            counting = True
        if counting and el.tag in (qn("w:p"), qn("w:tbl")):
            words += len(para_text(el).split()) if el.tag == qn("w:p") else sum(len(c.split()) for r in table_rows(el) for c in r)
    return words


def check_words(doc):
    n = body_words(doc)
    if n < WORD_MIN or n > WORD_CAP:
        return False, f"{n:,} body words, outside {WORD_MIN} to {WORD_CAP:,}"
    note = "" if n <= WORD_BUDGET else f", over the {WORD_BUDGET:,} budget, cut where you can"
    return True, f"{n:,} body words, within {WORD_MIN} to {WORD_CAP:,}{note}"


def check_references(doc):
    texts = [para_text(p).strip() for p in section_range(doc, "References")]
    texts = [t for t in texts if t and not t.startswith(GUIDE)]
    if not texts:
        return False, "References is empty; write None. when nothing outside is cited"
    return True, f"References has {len(texts)} lines"


def footer_id(doc):
    for sec in doc.sections:
        for p in sec.footer.paragraphs:
            first = p.text.split("\t")[0].strip()
            if first:
                return first
    return ""


def check_sop_id(doc):
    sop_id = footer_id(doc)
    if not SOP_ID_RE.match(sop_id):
        return False, f"footer SOP id '{sop_id}' is not PREFIX-SOP-NNN-VN with PREFIX one of {', '.join(PREFIXES)}"
    return True, f"footer SOP id {sop_id}"


def check_name(doc, path):
    tbls = doc.element.body.findall(qn("w:tbl"))
    title = table_rows(tbls[0])[1][0] if tbls and len(table_rows(tbls[0])) > 1 else ""
    expected = f"{footer_id(doc)}. {title}.docx"
    if path.name in SAMPLE_NAMES:
        if not FILE_RE.match(expected):
            return False, f"as an SOP this sample would be named '{expected}', which is not PREFIX-SOP-NNN-VN. Title.docx"
        return True, f"skill sample asset; as an SOP it would be named '{expected}'"
    if not FILE_RE.match(path.name):
        return False, f"file name '{path.name}' is not PREFIX-SOP-NNN-VN. Title.docx with PREFIX one of {', '.join(PREFIXES)}"
    if path.name != expected:
        return False, f"file name '{path.name}' does not match the footer id and title, '{expected}'"
    return True, f"file name '{path.name}'"


def check_names(doc):
    texts = list(all_texts(doc)) + [doc.core_properties.author or "", doc.core_properties.last_modified_by or ""]
    hits = 0
    for text in texts:
        words = re.findall(r"[A-Za-z]+", text)
        grams = [[w] for w in words] + [words[i:i + 2] for i in range(len(words) - 1)]
        hits += sum(1 for g in grams if name_hash(g) in NAME_HASHES)
    if hits:
        return False, f"{hits} staff names found; name people by role"
    return True, "no staff name from the guard list, in the text or the file author fields"


def check_identifiers(doc):
    bad = [m.group(0) for t in all_texts(doc) for m in IDENTIFIER_RE.finditer(t)]
    if bad:
        return False, f"{len(bad)} identifier shaped tokens, first {bad[0][:3]}..."
    return True, "no identifier shaped token"


def check_ascii(doc):
    bad = [t.strip()[:30] for t in all_texts(doc) if any(ord(ch) > 127 for ch in t)]
    if bad:
        return False, f"non-ASCII characters in {len(bad)} lines, first '{bad[0]}'"
    return True, "every line is ASCII"


def check_prose(doc):
    hits = [(word, snip) for text in all_texts(doc) for word, snip in prose.find_banned(text)]
    for word, snip in hits:
        print(f"FAIL prose: {word} in '{snip}'")
    if not hits:
        print("PASS prose: no em dash, en dash or banned word")
    return bool(hits)


def main():
    if len(sys.argv) != 2:
        print("usage: python scripts/check_sop.py \"PREFIX-SOP-NNN-V1. Title.docx\"")
        return 2
    path = Path(sys.argv[1])
    doc = docx.Document(str(path))
    checks = [
        ("template-styles", lambda: check_styles(doc)),
        ("no-direct-format", lambda: check_direct(doc)),
        ("title-table", lambda: check_title_table(doc)),
        ("guide-lines", lambda: check_guides(doc)),
        ("revision-history", lambda: check_revision(doc)),
        ("sections", lambda: check_sections(doc)),
        ("typed-numbers", lambda: check_typed(doc)),
        ("heading-numbers", lambda: check_heading_numbers(doc)),
        ("procedures", lambda: check_procedures(doc)),
        ("words", lambda: check_words(doc)),
        ("references", lambda: check_references(doc)),
        ("sop-id", lambda: check_sop_id(doc)),
        ("file-name", lambda: check_name(doc, path)),
        ("staff-names", lambda: check_names(doc)),
        ("identifiers", lambda: check_identifiers(doc)),
        ("ascii", lambda: check_ascii(doc)),
    ]
    failed = False
    for name, fn in checks:
        ok, detail = fn()
        print(f"{'PASS' if ok else 'FAIL'} {name}: {detail}")
        failed = failed or not ok
    failed = check_prose(doc) or failed
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
