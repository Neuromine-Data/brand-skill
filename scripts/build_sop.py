import copy
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import build_word as bw

import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches
from docx.text.paragraph import Paragraph

TEMPLATE_NAME = "neuromine-sop-template.docx"
SAMPLE_NAME = "neuromine-sop-sample.docx"
C = bw.C
PT = bw.PT
SOP_NUM = 92
SECTIONS = ["Purpose", "Scope", "Definitions and Acronyms", "Procedures", "References"]
DROP_STYLES = ("CoverEyebrow", "CoverDate", "Callout")
TITLE_WIDTHS = (3.3, 1.6, 1.6)
REVISION_WIDTHS = (1.6, 1.6, 3.3)


def level_ppr(level, **kw):
    return bw.ppr(num=SOP_NUM, **kw).replace('<w:ilvl w:val="0"/>', f'<w:ilvl w:val="{level}"/>')


def build_sop_styles(doc):
    styles_el = doc.styles.element
    for old in styles_el.findall(qn("w:style")):
        if old.get(qn("w:styleId")) in DROP_STYLES:
            styles_el.remove(old)
    bw.put_style(styles_el, "paragraph", "SOPHeading", "SOP Heading",
                 bw.ppr(keep=True, before=360, after=120, line=276)
                 + bw.rpr(bw.HEAD_FONT, PT["heading"], C["ink"], bold=True),
                 nxt="Normal", custom=True)
    bw.put_style(styles_el, "paragraph", "SOPSection", "SOP Section",
                 level_ppr(0, keep=True, before=360, after=120, line=276, outline=0)
                 + bw.rpr(bw.HEAD_FONT, PT["heading"], C["ink"], bold=True),
                 nxt="Normal", custom=True)
    bw.put_style(styles_el, "paragraph", "SOPStep", "SOP Step",
                 level_ppr(1, keep=True, before=240, after=80, line=276, outline=1)
                 + bw.rpr(bw.HEAD_FONT, PT["body"], C["ink"], bold=True),
                 nxt="SOPDetail", custom=True)
    bw.put_style(styles_el, "paragraph", "SOPTerm", "SOP Term",
                 level_ppr(3, after=90), nxt="SOPTerm", custom=True)
    bw.put_style(styles_el, "paragraph", "SOPDetail", "SOP Detail",
                 level_ppr(2, after=90), nxt="SOPDetail", custom=True)
    bw.put_style(styles_el, "paragraph", "Footer", "footer",
                 bw.ppr(border=("top", 6, 4, C["rule"]), tabs=[("right", bw.TEXT_W)], before=0, after=0, line=240)
                 + bw.rpr(bw.BODY_FONT, PT["caption"], C["placeholder"]))


def build_sop_numbering(doc):
    numbering = doc.part.numbering_part.element
    abstract = parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{SOP_NUM}">'
        '<w:multiLevelType w:val="multilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="576" w:hanging="576"/></w:pPr>'
        f'<w:rPr><w:b/><w:bCs/><w:color w:val="{C["ink"]}"/></w:rPr></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="576" w:hanging="576"/></w:pPr>'
        f'<w:rPr><w:b/><w:bCs/><w:color w:val="{C["brand"]}"/></w:rPr></w:lvl>'
        f'<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="{bw.BULLET_CHAR}"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="816" w:hanging="240"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>'
        f'<w:color w:val="{C["brand"]}"/></w:rPr></w:lvl>'
        '<w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%4"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="576" w:hanging="576"/></w:pPr>'
        f'<w:rPr><w:b/><w:bCs/><w:color w:val="{C["brand"]}"/></w:rPr></w:lvl>'
        "</w:abstractNum>"
    )
    numbering.find(qn("w:num")).addprevious(abstract)
    numbering.append(parse_xml(
        f'<w:num {nsdecls("w")} w:numId="{SOP_NUM}"><w:abstractNumId w:val="{SOP_NUM}"/></w:num>'
    ))


HEAD_BELOW_BAND = 792


def sop_head(paragraph, doc, program):
    bw.clear_paragraph(paragraph)
    paragraph.style = doc.styles["Header"]
    paragraph.add_run(bw.header_mark(program))
    paragraph.add_run("\t")
    paragraph.add_run("Standard Operating Procedure")


def sop_foot(paragraph, doc, sop_id):
    bw.clear_paragraph(paragraph)
    paragraph.style = doc.styles["Footer"]
    paragraph.add_run(sop_id)
    paragraph.add_run("\t")
    paragraph.add_run("Page ")
    paragraph._p.append(bw.field("PAGE \\* MERGEFORMAT", "1"))


def set_sop_header_footer(doc, sop_id, program=""):
    sec = doc.sections[0]
    sec.different_first_page_header_footer = True
    sop_head(sec.header.paragraphs[0], doc, program)
    sop_foot(sec.footer.paragraphs[0], doc, sop_id)
    first = sec.first_page_header.paragraphs[0]
    bw.clear_paragraph(first)
    first.style = doc.styles["Normal"]
    first.paragraph_format.space_after = 0
    bw.anchor_behind(bw.add_band(first))
    bw.add_program(sec.first_page_header)
    line = sec.first_page_header.add_paragraph()
    sop_head(line, doc, program)
    line.paragraph_format.space_before = HEAD_BELOW_BAND * 635
    sop_foot(sec.first_page_footer.paragraphs[0], doc, sop_id)


def set_sop_program(doc, program):
    bw.set_program(doc, program)
    doc.sections[0].first_page_header.paragraphs[2].runs[0].text = bw.header_mark(program)


def set_sop_id(doc, sop_id):
    sec = doc.sections[0]
    for foot in (sec.footer, sec.first_page_footer):
        foot.paragraphs[0].runs[0].text = sop_id


def sop_table(doc, rows, widths):
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = doc.styles["Neuromine Table"]
    table.autofit = False
    for c, width in enumerate(widths):
        table.columns[c].width = Inches(width)
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.width = Inches(widths[c])
            cell.text = value
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    tr_pr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    return table


def build_template(path):
    doc = docx.Document()
    bw.build_styles(doc)
    build_sop_styles(doc)
    bw.build_numbering(doc)
    build_sop_numbering(doc)
    bw.build_fonts(doc)
    bw.build_theme(doc)
    bw.set_page(doc)
    set_sop_header_footer(doc, "PREFIX-SOP-NNN-V1")
    bw.set_settings(doc)
    bw.clear_body(doc)
    doc.add_paragraph("Standard Operating Procedure", style="Title")
    sop_table(doc, [
        ["Title", "Version Number", "Effective Date"],
        ["Guide: the SOP title, the same words as in the file name", "Guide: 1", "Guide: YYYY-MM-DD"],
    ], TITLE_WIDTHS)
    doc.add_paragraph("Revision History", style="SOP Heading")
    sop_table(doc, [
        ["Version Number", "Effective Date", "Description"],
        ["Guide: 1", "Guide: YYYY-MM-DD", "Guide: what changed; the first row reads Initial SOP."],
    ], REVISION_WIDTHS)
    guides = {
        "Purpose": ["Guide: why this SOP exists and what goes wrong without it, in two to four sentences."],
        "Scope": ["Guide: what the SOP covers and what it does not, and who runs it, by role."],
        "Definitions and Acronyms": ["Guide: one SOP Term line per term the reader may not know. Write None. when there are none."],
        "Procedures": ["Guide: one line on when the steps run and in what order."],
        "References": ["Guide: one line per document or tool the SOP cites, by name. Write None. when nothing outside is cited."],
    }
    for i, name in enumerate(SECTIONS):
        section = doc.add_paragraph(name, style="SOP Section")
        if i == 0:
            section.paragraph_format.page_break_before = True
        for text in guides[name]:
            doc.add_paragraph(text, style="Normal")
        if name == "Definitions and Acronyms":
            doc.add_paragraph("Guide: Term: what it means, in one sentence.", style="SOP Term")
        if name == "Procedures":
            doc.add_paragraph("Guide: step name, a verb and an object", style="SOP Step")
            for line in [
                "Guide: Inputs: what the step starts from.",
                "Guide: Steps: what the person does, in order.",
                "Guide: Outputs: what the step leaves behind.",
                "Guide: Checks: Verify what against what, in one sentence.",
            ]:
                doc.add_paragraph(line, style="SOP Detail")
            doc.add_paragraph("Guide: Final validation", style="SOP Step")
            doc.add_paragraph("Guide: Checks: one line per check run before handoff, each starting with Verify.", style="SOP Detail")
    bw.set_core(doc, "Neuromine SOP template")
    bw.save_stable(doc, path)


SAMPLE_ID = "ALSTDI-SOP-NNN-V1"
SAMPLE_PROGRAM = "ALS TDI"
SAMPLE_TITLE = "Monthly Brand Check of Published Documents"
SAMPLE = {
    "Purpose": [
        ("Normal", None, "This SOP sets out how the data team checks, once a month, that the documents it has published follow the Neuromine brand skill. A document that drifts from the brand is harder to read and harder to trust, and drift is cheaper to fix while it is small."),
        ("Normal", None, "The check uses the checkers that ship with the brand skill, so two people who run it on the same files get the same result. The SOP describes the check as the team runs it today."),
    ],
    "Scope": [
        ("Normal", None, "It applies to every Word document, PowerPoint deck, Excel workbook, Markdown page and HTML page that the data team published to a shared site or sent outside the team in the past month, for the ALS TDI program."),
        ("Normal", None, "It does not cover SOPs, data files such as release indexes, or documents that another team owns. A team member on the monthly rota runs the check. The data owner reviews the result and signs off the month."),
    ],
    "Definitions and Acronyms": [
        ("SOPTerm", "Brand skill: ", "the brand rules, templates and checkers that every Neuromine document follows."),
        ("SOPTerm", "Checker: ", "a script in the brand skill that reads one finished file and prints PASS or FAIL for each rule it tests."),
        ("SOPTerm", "Check log: ", "the workbook where each monthly run is recorded, one row per document."),
        ("SOPTerm", "Owner: ", "the person who wrote or last changed a published document, named in the check log by role."),
        ("SOPTerm", "SOP: ", "standard operating procedure."),
    ],
    "Procedures": [
        ("Normal", None, "Run the steps in order in the first week of each month. Each step ends with one check. Do not start the next step until that check holds."),
        ("SOPStep", None, "Build the list of documents"),
        ("SOPDetail", "Inputs: ", "the shared site's list of files changed in the past month and the team's sent items."),
        ("SOPDetail", "Steps: ", "list every document published or sent outside the team since the last run. Give each one a row in the check log with its file name, format, owner role and the date it went out. Leave out drafts that never left the team."),
        ("SOPDetail", "Outputs: ", "the check log for the month, one row per document."),
        ("SOPDetail", "Checks: ", "Verify the number of rows in the check log against the number of files the shared site lists as changed in the month."),
        ("SOPStep", None, "Run the checkers"),
        ("SOPDetail", "Inputs: ", "the check log and the current brand skill, downloaded fresh on the day of the run."),
        ("SOPDetail", "Steps: ", "for each row, run the checker for the file's format: check_docx.py for Word, check_deck.py for decks, check_xlsx.py for workbooks, check_md.py for Markdown and check_html.py for HTML. Paste the PASS and FAIL lines into the row. A format with no checker gets the prose check from prose.py only."),
        ("SOPDetail", "Outputs: ", "a checker result in every row of the check log."),
        ("SOPDetail", "Checks: ", "Verify that every row in the check log has a checker result and that the result names the same file as the row."),
        ("SOPStep", None, "Sort the findings"),
        ("SOPDetail", "Inputs: ", "the rows with at least one FAIL line."),
        ("SOPDetail", "Steps: ", "mark each FAIL line as one of three kinds. A style or layout fault is fixed in the file. A word or punctuation fault is fixed in the text. A checker fault means the rule or the checker is wrong, not the file. Write the kind beside the line."),
        ("SOPDetail", "Outputs: ", "every FAIL line marked with its kind."),
        ("SOPDetail", "Checks: ", "Verify the count of marked FAIL lines against the count of FAIL lines the checkers printed."),
        ("SOPStep", None, "Send the fixes"),
        ("SOPDetail", "Inputs: ", "the sorted rows."),
        ("SOPDetail", "Steps: ", "send each owner the rows for their documents, with the FAIL lines and what to change. Ask for the fixed file within ten working days. Report a checker fault to the brand skill maintainers instead of asking the owner to work around it."),
        ("SOPDetail", "Outputs: ", "one message to each owner and one report for each checker fault."),
        ("SOPDetail", "Checks: ", "Verify the list of owners messaged against the owner roles in the rows that have a FAIL line."),
        ("SOPStep", None, "Close the month"),
        ("SOPDetail", "Inputs: ", "the fixed files from the owners and the check log."),
        ("SOPDetail", "Steps: ", "run the checker again on each fixed file and paste the new result into its row. Mark a row closed when its checker prints only PASS lines. Carry an open row into next month's log with its first date kept, so a late fix stays visible."),
        ("SOPDetail", "Outputs: ", "the check log with every row closed or carried."),
        ("SOPDetail", "Checks: ", "Verify the closed and carried rows against the rows that had a FAIL line."),
        ("SOPStep", None, "Final validation"),
        ("SOPDetail", "Checks: ", "Verify each closed row against a fresh checker run on the published copy, not the copy the owner sent."),
        ("SOPDetail", "Checks: ", "Verify the totals at the top of the check log (documents, PASS, FAIL, closed, carried) against the rows below them."),
        ("SOPDetail", "Checks: ", "Verify that the data owner has signed off the month in the check log before the log is filed."),
    ],
    "References": [
        ("Normal", None, "Neuromine brand skill, references/brand.md: the brand rules every format shares."),
        ("Normal", None, "Neuromine brand skill, scripts/README.md: what each checker tests and how to run it."),
    ],
}


def set_text(paragraph, label, text):
    bw.clear_paragraph(paragraph)
    if label:
        paragraph.add_run(label).bold = True
    paragraph.add_run(text)


def fill_cell(cell, text):
    set_text(cell.paragraphs[0], None, text)


def section_blocks(doc):
    body = doc.element.body
    blocks = {}
    current = None
    for el in list(body):
        if el.tag != qn("w:p"):
            continue
        p = Paragraph(el, doc._body)
        if p.style.style_id == "SOPSection":
            current = p.text
            blocks[current] = (el, [])
        elif current:
            blocks[current][1].append(el)
    return blocks


def fill_section(doc, head, lines, content):
    protos = {}
    for el in lines:
        style = Paragraph(el, doc._body).style.style_id
        protos.setdefault(style, el)
    for el in lines:
        el.getparent().remove(el)
    anchor = head
    for style, label, text in content:
        el = copy.deepcopy(protos[style])
        anchor.addnext(el)
        set_text(Paragraph(el, doc._body), label, text)
        anchor = el


def build_sample(template, path):
    doc = docx.Document(str(template))
    title_table, revision_table = doc.tables[0], doc.tables[1]
    for cell, text in zip(title_table.rows[1].cells, [SAMPLE_TITLE, "1", "2026-11-04"]):
        fill_cell(cell, text)
    for cell, text in zip(revision_table.rows[1].cells, ["1", "2026-11-04", "Initial SOP."]):
        fill_cell(cell, text)
    set_sop_id(doc, SAMPLE_ID)
    set_sop_program(doc, SAMPLE_PROGRAM)
    for name, (head, lines) in section_blocks(doc).items():
        fill_section(doc, head, lines, SAMPLE[name])
    bw.set_core(doc, SAMPLE_TITLE)
    bw.save_stable(doc, path)


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else bw.default_out()
    template = out / TEMPLATE_NAME
    sample = out / SAMPLE_NAME
    build_template(template)
    build_sample(template, sample)
    for p in (template, sample):
        print(f"wrote {p.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
