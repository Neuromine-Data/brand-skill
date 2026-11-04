# SOP

Rules for writing a standard operating procedure (SOP) as a .docx in the Neuromine brand. Read https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/brand.md first, then this file. An SOP keeps the house SOP structure (title table, revision history, five numbered sections, numbered steps) in the Neuromine look. Where this file and brand.md or word.md disagree, this file wins for SOPs. "How to build an SOP" says what to do, step by step. "Rules that apply to SOPs only" says what every SOP must hold and how it must read.

## Start from the template

Download https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-sop-template.docx and fill it in place. It already holds every style, the automatic numbering, the title table, the revision history table, the header, the footer and the page setup. Each line that starts with "Guide:" says what goes there; replace it or remove it. https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-sop-sample.docx is a filled SOP for a made-up process. Never build an SOP from a blank document, from neuromine-template.docx or from an older house SOP file.

To change the template, edit `scripts/build_sop.py` and run it. Do not edit the .docx files by hand.

## Starter

```python
import copy
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

doc = Document("neuromine-sop-template.docx")

def put(paragraph, text, label=None):
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)
    if label:
        paragraph.add_run(label).bold = True
    paragraph.add_run(text)

def copy_after(paragraph):
    el = copy.deepcopy(paragraph._p)
    paragraph._p.addnext(el)
    return Paragraph(el, paragraph._parent)

for cell, text in zip(doc.tables[0].rows[1].cells, ["Data Release Sign-off", "1", "2026-11-04"]):
    put(cell.paragraphs[0], text)
for cell, text in zip(doc.tables[1].rows[1].cells, ["1", "2026-11-04", "Initial SOP."]):
    put(cell.paragraphs[0], text)
doc.sections[0].footer.paragraphs[0].runs[0].text = "NEUROMINE-SOP-NNN-V1"
purpose = next(p for p in doc.paragraphs if p.text.startswith("Guide: why"))
put(purpose, "This SOP sets out how a data release is signed off.")
second = copy_after(purpose)
put(second, "A release that skips sign-off can ship with a wrong count.")
step = next(p for p in doc.paragraphs if p.style.name == "SOP Step")
put(step, "Collect the release files")
doc.core_properties.title = "Data Release Sign-off"
doc.core_properties.author = "Neuromine data team"
doc.save("NEUROMINE-SOP-NNN-V1. Data Release Sign-off.docx")
```

## How to build an SOP

With python-docx, in this order. `scripts/build_sop.py` `build_sample` does every step and is the worked example.

1. Copy the template. Download neuromine-sop-template.docx and open it with `Document(path)`. Save under the SOP file name at the end, never over the template.
2. Fill the title table. Row 2 of `doc.tables[0]` holds three cells under the labels Title, Version Number and Effective Date. Put the title, the version as a plain number (1) and the date as YYYY-MM-DD. Change only row 2; leave the label row alone. To change a cell, clear the runs of its one paragraph and add one run with the text, as `put` in the Starter does.
3. Add the revision row. Row 2 of `doc.tables[1]` is the first row: 1, the date, "Initial SOP." For a later revision, deepcopy the last row's XML (`row._tr`), insert it after that row with `addnext`, and fill it with 1, the new date and what changed.
4. Set the SOP id. The footer's first run holds PREFIX-SOP-NNN-V1. Write the SOP id there, for example NEUROMINE-SOP-NNN-V1. Leave the tab, the word Page and the page number field.
5. Replace each section's guide text. Under each SOP Section line (Purpose, Scope, Definitions and Acronyms, Procedures, References) the template holds one Normal "Guide:" paragraph. Put your first paragraph of text in it. For more paragraphs, deepcopy it and insert each copy after the last with `addnext`.
6. Add terms. Under Definitions and Acronyms, fill the SOP Term line: a bold run with the term and a colon, then a plain run with what it means. For each further term, deepcopy that paragraph's XML and insert it after the last term, so the automatic number (3.1, 3.2) carries. If there are no terms, remove the SOP Term line and write "None." in the Normal paragraph.
7. Add steps. Under Procedures, each step is one SOP Step line (the step name) followed by four SOP Detail lines: Inputs, Steps, Outputs and Checks, each with a bold label run. To add a step, deepcopy the SOP Step paragraph and its four SOP Detail paragraphs and insert the copies after the last detail line. The automatic number (4.1, 4.2) carries with the XML. Never type a number, and never apply another style to a step line. The last step is Final validation, which holds only Checks lines.
8. Remove every guide line. Delete each paragraph or cell text that still starts with "Guide:", with `p._p.getparent().remove(p._p)` for a paragraph. Do not leave empty paragraphs in their place.
9. Set the core properties. `doc.core_properties.title` is the SOP title; `author` and `last_modified_by` are a role or "Neuromine data team", never a person's name.
10. Name the file PREFIX-SOP-NNN-VN. Title.docx, for example "NEUROMINE-SOP-NNN-V1. Data Release Sign-off.docx", and save.
11. Run `python scripts/check_sop.py "NEUROMINE-SOP-NNN-V1. Data Release Sign-off.docx"`. Fix every FAIL and run it again until each line reads PASS. Its words line prints the body word count.

## Styles

Use only these styles. Names are as Word shows them; python-docx takes the same names.

| Use | Style | Look |
| --- | --- | --- |
| Cover title "Standard Operating Procedure" | Title | Helvetica Neue bold, 20pt, #1F4FAE |
| "Revision History" label | SOP Heading | Helvetica Neue bold, 14pt, #1A2332, no number |
| Section line (Purpose to References) | SOP Section | Helvetica Neue bold, 14pt, #1A2332, numbered 1. to 5. |
| Step name under Procedures | SOP Step | Helvetica Neue bold, 11pt, #1A2332, numbered 4.1, 4.2 in #1F4FAE |
| Term under Definitions and Acronyms | SOP Term | Helvetica Neue 11pt, #2C3E50, numbered 3.1, 3.2 in #1F4FAE |
| Inputs, Steps, Outputs and Checks lines | SOP Detail | Helvetica Neue 11pt, #2C3E50, #1F4FAE round bullet |
| Body text | Normal | Helvetica Neue 11pt, #2C3E50, line 1.15 |
| Title table and revision history | Neuromine Table (table style) | header row #F4F6FB fill, bold #1A2332, #D8DCE5 rule; body #374151 with #DFE3EB rules |
| Header and footer | Header, Footer | Courier New 9pt, #4B5563 header and #6E7581 footer, #E5E7EB rule |

Font fallback is as in word.md: Helvetica Neue with Arial, Courier New for the header and footer. The only direct formatting allowed is bold on a label run (a term, or Inputs:, Steps:, Outputs:, Checks:). No other style is used: no Heading 1 to 3, no List Bullet or List Number, no Callout, no Caption.

## Page setup

US Letter, portrait, single column, 1 inch margins on all four sides, header and footer 0.5 inch from the page edge. The same header and footer run on every page. Page 1 is the cover: the Title, the title table and the revision history. Purpose starts page 2 (page break before is set on it in the template). Do not add section breaks, page break runs or empty paragraphs.

- Header: "Neuromine" at left, a tab, "Standard Operating Procedure" at right. For an Answer ALS or ALS TDI SOP, change the header's first run to that program name.
- Footer: the SOP id at left, a tab, "Page" and a `PAGE` field at right.

## Title table

One label row and one value row, three columns: Title, Version Number, Effective Date. The labels never change. Title is the SOP title in title case, the same words as in the file name. Version Number is a plain number, 1. Effective Date is YYYY-MM-DD. Every cell is filled; `check_sop.py` fails `title-table` otherwise.

## Revision history

A label row (Version Number, Effective Date, Description) and one row per revision, oldest first. The first row reads 1, the date, "Initial SOP." Each later revision adds one row; never edit or delete an earlier row. Description is one short sentence on what changed.

## Sections

Five SOP Section lines, in this order, with the names exactly as written:

1. Purpose: why the SOP exists and what goes wrong without it, in two to four sentences.
2. Scope: what it covers, what it does not, and who runs it, by role.
3. Definitions and Acronyms: one SOP Term line per term, or "None."
4. Procedures: one Normal line on when the steps run, then the steps.
5. References: one Normal line per document or tool the SOP cites, by name, or "None."

The numbers 1. to 5. come from the SOP Section style's automatic numbering. The text of the line is the name only.

## Procedure steps

Each step is a short verb and object ("Build the list of documents") on an SOP Step line, then four SOP Detail lines:

- Inputs: what the step starts from.
- Steps: what the person does, in order, in one to four sentences.
- Outputs: what the step leaves behind.
- Checks: one sentence that starts with Verify and says what is verified against what ("Verify the row count of the log against the file count on the site.").

The last step is Final validation: two to four Checks lines, each a "Verify X against Y" check run before the SOP's output is handed on. Keep a step to one stage of the process; split a stage that needs two checks.

## Rules that apply to SOPs only

These hold for SOPs and for nothing else in the brand. `check_sop.py` tests the ones marked (checked).

- The five sections are Purpose, Scope, Definitions and Acronyms, Procedures and References, in that order. Never rename, merge, reorder or drop one. (checked)
- Numbers are automatic only. Never type a section or step number ("1.", "4.1") at the start of a line, and never put a Heading style on a numbered line: Word then shows two numbers in the PDF. (checked)
- One SOP per process. A second process gets its own SOP, not a sixth section.
- Word budget: 700 to 1,000 words in the body (Purpose to the end), hard cap 1,100. Cut before you go over. Do not pad a short process to reach 700; say in your hand-back that it is short. (checked: 700 to 1,100)
- Steps take the Inputs, Steps, Outputs, Checks shape, with one "what is verified against what" check per stage. (checked: every step has a Checks line)
- References reads "None." when nothing outside the SOP is cited; it is never empty. (checked)
- Cite spreadsheet and index columns by name, never by letter. Letters move when a column is added.
- No ticket ids, tracker names, board or local paths, and no staff names. Name people by role ("the data owner", "a team member on the rota"). (checked: staff names and identifier shaped tokens)
- File name: PREFIX-SOP-NNN-VN. Title.docx, with the prefix AALS- for Answer ALS, ALSTDI- for ALS TDI or NEUROMINE- for Neuromine. The data owner assigns the number; until then write NNN, in the file name and in the footer. (checked)
- A revised SOP stays at V1, with a new revision history row, unless the data owner asks for V2.
- No cover band and no logo anywhere in an SOP. The title table is the cover.
- The PDF is exported by the data owner in Word when the SOP is cleared, not by an agent. Hand back the .docx only.
- ASCII only. Word turns straight quotes into curly quotes and two hyphens into a dash as you type; replace them before you hand back. (checked)

## From brand.md and word.md

Still apply:

- Fonts and fallback (brand.md Typography, word.md Fonts): Helvetica Neue with Arial, Courier New for the header and footer, never embedded.
- Colors (brand.md Palette): the title in #1F4FAE, section and step lines in #1A2332, body in #2C3E50.
- Header and footer look (word.md Header and footer): Courier New 9pt with the #E5E7EB rule; the text is as in Page setup above.
- Table look (word.md Table recipe, steps 1 and 2) for the title table and the revision history.
- Writing (brand.md Writing): plain words, short sentences, no em or en dash, none of the banned words.
- Dates (brand.md Dates): YYYY-MM-DD in the tables.

Do not apply:

- The cover band, the logo and the wordmark (brand.md Cover page, word.md Cover page recipe).
- The numbered Heading 1 recipe ("2. Methods") and the STYLEREF header field (word.md).
- Callouts (brand.md Callouts, word.md Callout recipe).
- Charts and figures (brand.md Charts, word.md Charts). If a step needs a picture, describe it in words.
- Table captions (word.md Table recipe, step 4) on the title table and the revision history.
- `check_docx.py`. It does not run on SOPs; `check_sop.py` does.

## Voice

Follow the Writing section in brand.md. Write for someone who runs the process for the first time: short plain sentences, one action per sentence in Steps, the present tense, and "you" or the role, never "we". Describe the process as it runs today. If something in the process is broken, say so in one line and leave the fix out of the SOP.

## Checklist before sending

- Started from neuromine-sop-template.docx; no direct font, size or color changes.
- Title table filled: title, version 1, date as YYYY-MM-DD.
- Revision history has a dated row; a revision added a row and kept the old ones.
- The five sections in order, names unchanged, numbers automatic, no typed numbers.
- Every step has Inputs, Steps, Outputs and Checks; Final validation is last.
- No "Guide:" line left and no empty paragraph where one was.
- References filled, or "None."
- Body between 700 and 1,000 words, never over 1,100.
- No ticket id, path or staff name; people by role.
- Footer SOP id and file name match: PREFIX-SOP-NNN-VN. Title.docx.
- Core properties: title set, author a role or "Neuromine data team".
- `python scripts/check_sop.py "PREFIX-SOP-NNN-V1. Title.docx"` prints PASS on every line.
