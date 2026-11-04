# Word

Rules for writing a .docx in the Neuromine brand. Read https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/brand.md first; this file maps those rules onto Word. Where the two disagree, this file wins for Word.

## Start from the template

Download https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-template.docx, open it (python-docx: `Document("neuromine-template.docx")`), delete its one empty body paragraph, then add content. The template already holds every style, the cover band with the logo, the header, the footer, the page setup, the theme colors and the font fallback. https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-sample.docx shows every recipe below. Never restyle text by hand when a style exists: apply the style.

To change the template, edit `scripts/build_word.py` and run it. Do not edit the .docx files by hand. An SOP does not use this template; see SOPs.

## Starter

```python
from docx import Document
doc = Document("neuromine-template.docx")
empty = doc.paragraphs[0]._p
empty.getparent().remove(empty)
doc.add_paragraph("Release report", style="Cover Eyebrow")
doc.add_paragraph("Release 8 summary", style="Title")
doc.add_paragraph("One line on what the document is", style="Subtitle")
doc.sections[0].first_page_footer.paragraphs[0].text = "November 2026"
h1 = doc.add_paragraph("1. Summary", style="Heading 1")
h1.paragraph_format.page_break_before = True
doc.sections[0].header.paragraphs[0]._p.xpath(".//w:fldSimple//w:t")[0].text = "1. Summary"
doc.save("neuromine-release-8-summary.docx")
```

## Styles

Use only these styles. Names are as Word shows them; python-docx takes the same names.

| Use | Style | Look |
| --- | --- | --- |
| Cover title | Title | Helvetica Neue bold, 20pt, #1F4FAE, right indent keeps it off the right third of the page |
| Cover subtitle | Subtitle | Helvetica Neue 14pt, #374151 |
| Cover eyebrow | Cover Eyebrow | Courier New bold 9pt, caps, #4B5563 |
| Cover date | Cover Date | Courier New 9pt, #4B5563 (first-page footer) |
| Section | Heading 1 | Helvetica Neue bold, 14pt, #1A2332 |
| Subsection | Heading 2 | Helvetica Neue bold, 11pt, #1A2332 |
| Minor heading | Heading 3 | Helvetica Neue bold, 11pt, #374151 |
| Body | Normal | Helvetica Neue 11pt, #2C3E50, line 1.15, 6.6pt after |
| Bullets | List Bullet | #1F4FAE round bullet (see Lists) |
| Steps | List Number | #1F4FAE numbers |
| Quotation | Quote | Helvetica Neue 11pt, #374151, indented both sides, no italics |
| Key point | Callout | #EEF4FF fill, 3px #1F4FAE left bar |
| Table or chart label | Caption | Courier New 9pt, #4B5563 |
| Tables | Neuromine Table (table style) | see Table recipe |

The sizes are the document sizes in brand.md: 20pt title, 14pt section heading, 11pt body and table cells, 9pt eyebrow, caption, header and footer. No other size is used. There are no rules under headings. Do not use Heading 4 or lower; split the section instead.

## Fonts

| Role | Font | Office-safe fallback |
| --- | --- | --- |
| Headings, title, body | Helvetica Neue | Arial |
| Eyebrow, captions, header, footer, cover date | Courier New | none needed |

The template's font table names Arial as the fallback (`altName`) for Helvetica Neue, so Word swaps it in where Helvetica Neue is missing, as on most Windows machines. Courier New ships with Windows and macOS. A render you make yourself in a sandbox shows the fallback; do not change the fonts in the styles because of it. Never embed fonts.

## Page setup

US Letter, portrait, single column, 1 inch margins on all four sides. Header and footer sit 0.5 inch from the page edge. Do not add section breaks that change any of this.

## Header and footer

Page 1 is the cover and uses its own header and footer ("different first page" is on).

- Header, page 2 onward: "Neuromine" at left, a tab, then a `STYLEREF "Heading 1"` field at right that shows the current section. Courier New 9pt, #4B5563, a 1px #E5E7EB rule below. For an Answer ALS or ALS TDI document, change the text of the header's first run to that program name.
- Footer, page 2 onward: a `PAGE` field at right, nothing at left. Courier New 9pt, #6E7581, a 1px #E5E7EB rule above.
- Cover header: the brand blue band with the logo. Cover footer: the date, style Cover Date, for example "November 2026".
- All of this comes from the Header, Footer and Cover Date styles; do not format it by hand.

Write each Heading 1 as number, full stop, space, name ("2. Methods") so the header reads well.

## Cover page recipe

The template cover has a full-width #1F4FAE band across the top quarter of the page, with the white Neuromine logo at its top left and the NEUROMINE wordmark beside it (brand.md Wordmark). The band is a picture behind the text in the first-page header. Keep it; do not draw a shape or add the logo anywhere else.

1. Cover Eyebrow: two to four words for the kind of document ("Release report"). Its space before puts the block below the band.
2. Title: the document title. No word over 13 characters, since a long word at 20pt sits badly beside the band. When the brief fixes a longer word, hyphenate it at a syllable (end the first part with a hyphen, then a line break run: `run.add_break()`). `check_docx.py` fails `title-width` on any Title word over 13 characters.
3. Subtitle: one line on what the document is.
4. Put the date in the first-page footer, `doc.sections[0].first_page_footer.paragraphs[0]` (style Cover Date). The template holds the placeholder "Month Year" there; replace it.
5. Start the first Heading 1 on page 2 with `paragraph_format.page_break_before = True`. Set it on the first Heading 1 only. Later sections follow on the same page, and the heading styles' keep with next stops a heading sitting alone at the foot of a page. Never add empty paragraphs or page break runs to push content down. `check_docx.py` fails `breaks` otherwise.

The logo is white and appears only on the brand blue band. Never put it on white, a tint or a photo. If a document needs a white cover (the band removed from the first-page header), write "Neuromine" as the eyebrow and use no logo.

## Table recipe

1. Add the table and set `table.style = "Neuromine Table"`. Row 1 is the header row: #F4F6FB fill, bold #1A2332 text, a 1px #D8DCE5 rule under it, repeats on each page. Also set `w:tblHeader` on row 1 yourself (`build_word.py` `add_table` shows how); do not rely on the style alone.
2. Body rows are #374151 text with a 1px #DFE3EB rule between rows and a #DFE3EB outer border; there are no inner vertical rules. Cell padding is built in.
3. Headers are plain words in sentence case, never raw column names. Text columns are left aligned. Number columns are right aligned in the body rows and in the header cell over them. `build_word.py` `add_table` shows how; `check_docx.py` fails `numbers-right` otherwise.
4. Put a Caption paragraph directly under the table: "Table 1. What it shows. Source: Release 8 index." In Word the caption goes under the table, not above it as brand.md says for other formats; the checker looks for it there.
5. Keep one header row. Do not merge cells or add fills to body rows.

## Callout recipe

One short paragraph in style Callout for the one thing a reader must take away from a section. Start it with a bold label run ("Key point. "), then the text in a plain run. No icon, no other bold text inside. At most one callout per Heading 1 section. For an error, a success or a warning, say it in the body text; Word has no status box style.

## Lists

List Bullet gives the #1F4FAE round bullet with the text indented 0.25 inch and 6px between items. List Number gives #1F4FAE numbers with the same spacing. Use List Number only for steps or ranked items. A list holds only the items there are. The bullet num id is 90 and the number num id is 91 (each points at the abstractNum of the same id). Keep lists one level deep. To restart numbering, add a new `w:num` and put its numId on each paragraph of the new list:

```python
num = doc.part.numbering_part.element.add_num(91)
num.add_lvlOverride(ilvl=0).add_startOverride(1)
p = doc.add_paragraph("First step", style="List Number")
p._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = num.numId
```

## Charts

The template's theme accent colors are the categorical series, so a chart made in Word picks them up in order. When you color a chart yourself, use the categorical series in this order:

1. #008300
2. #E87BA4
3. #4A3AA7
4. #1BAF7A
5. #EDA100
6. #882255

A single series is #1F4FAE. For ordered data use the sequential series and for above or below data the diverging series, both in brand.md. Axis text and tick labels #6A7683; grid lines and baseline #D9E1EA, few and light. Label every axis in sentence case with its unit. Insert charts as PNG at the text width (6.5 in), then a Caption paragraph that names the figure and the source: "Figure 1. Visits by month. Source: Release 8 index."

With matplotlib: figure 6.5 by 3.25 in, 200 dpi, transparent background, labels 9pt in #6A7683, Arial (matplotlib cannot use the Office fallback), no top or right spines, left and bottom spines #D9E1EA, series colors in the order above. Insert with `doc.add_picture(path, width=Inches(6.5))`, then add the Caption paragraph. Leaving the picture paragraph in Normal is fine. `chart_png` in `build_word.py` is a worked example.

## Voice

Voice and prose: follow the Writing section in brand.md.

## SOPs

SOPs have their own format file, https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/sop.md, with its own template and checker; this file and check_docx.py do not cover them.

## Checklist before sending

- Started from the Neuromine template; no direct font, size or color changes.
- Cover has the band, eyebrow, title, subtitle and date; the first Heading 1 starts page 2.
- `page_break_before` is on the first Heading 1 only; no empty paragraphs or page break runs push content down.
- Headings go H1, H2, H3 in order with no skipped level; each H1 reads "N. Name".
- Every table uses Neuromine Table and has a caption under it; numbers and the header cells over them are right aligned.
- At most one callout per Heading 1 section.
- Chart colors follow the order above; every chart has a caption naming its source.
- Header field set: write the first Heading 1 text into the STYLEREF field's cached run, as the Starter does. The template's updateFields setting makes Word refresh STYLEREF and PAGE on open; a viewer that trusts the cached text shows "1. Section" if you skip this.
- Cover date replaced; "Month Year" is gone.
- `python scripts/check_docx.py your-file.docx` prints PASS on every line.
- File name has no spaces, for example `neuromine-release-8-summary.docx`.
