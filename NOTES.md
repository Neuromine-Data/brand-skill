# NOTES

Choices made while building the skill, where the request left something open. Newest section last.

## Design system, 2026-10-20

- Source: the Neuromine Variant Tool tokens.css, STYLE.md and variant_reports.py, read without changing them, plus a talk deck theme. design/DECISIONS.md lists what was taken and left out.
- Fonts: Helvetica Neue with Arial fallback and Courier New mono for every output, not the web system stack of the tool. One family is plainer than one per medium.
- Status backgrounds are flat (the end color of the tool's gradient), because the brand has no gradients.
- A document size scale in pt (9, 11, 14, 20) was added, since memos need one and neither source had it.
- Logo clear space (half the logo height) and minimum size (38px on screen, 0.5in on a slide) were added; neither source wrote a rule.

## Skill shape, 2026-11-04

- The rules work as an Agent Skill. An agent fetches SKILL.md, reads references/brand.md, reads the one format file, starts from the matching template in assets/ and runs the checker.
- references/brand.md holds the brand rules in thirteen sections: Palette, Typography, Headings, Body text, Tables, Header and footer, Cover page, Callouts, Lists, Charts, Page setup, Dates, Writing.
- Charts have a Sequential and a Diverging series built from the brand blue steps and the error reds; tokens.json has them as chart_seq_1 to 6 and chart_div_1 to 6.
- scripts/prose.py: find_banned(text) returns (word, snippet) pairs for the em dash, en dash and the banned words and phrases in brand.md. The banned list in brand.md is written in inline code so the page can name the words.
- scripts/check_all.py runs every scripts/check_<ext>.py against assets/*sample*.<ext> (check_deck.py against *sample*.pptx, check_html.py against *report*.html) and shows a checker's SKIP lines under its PASS line.

## Word format, 2026-11-04

- Styles: Title 20pt bold brand blue; Subtitle 14pt Text 2; Cover Eyebrow and Cover Date, Caption, Header and Footer in Courier New 9pt; Heading 1 14pt Ink, Heading 2 11pt Ink, Heading 3 11pt Text 2, all bold Helvetica Neue; Normal 11pt Text. Heading 3 takes Text 2 to read apart from Heading 2 at the same size.
- Body font, size and color sit in docDefaults and Normal is empty, so the Neuromine Table cell color is not overridden by Normal in Word.
- The cover band is one PNG (brand blue, 8.5 by 2.75 in, the logo at 0.75 in high, 1 in from the left edge and 0.5 in from the top, the wordmark beside it) anchored behind the text at the page corner in the first-page header. On a white cover the eyebrow reads Neuromine and there is no logo.
- Header: "Neuromine" as plain text, a tab, STYLEREF Heading 1.
- Tables put the Caption under the table, as the checker expects; brand.md says above for other formats, and word.md wins for Word.
- The chart uses Arial in matplotlib (Helvetica Neue is not on Windows). Builds are byte stable on one machine; another machine may give other chart bytes if its Arial or matplotlib differs.
- SOPs are not covered here: they get their own format file, references/sop.md, in the next version.

## PowerPoint format, 2026-11-04

- Eight layouts: Title, Section, Title and Content, Two Content, Table, Chart, Picture, Closing.
- Title and Closing are brand blue with the white logo at 0.75 in from the top and left edges, 0.75 in high, the NEUROMINE wordmark beside it, the eyebrow placeholder (idx 15), a Tint 3 bar, the white 54pt headline and the Tint 3 22pt subline. The logo and the wordmark are on these two layouts only.
- Content layouts are white with the 0.18 in brand blue strip, the headline 40pt bold in brand blue and the footer: the deck name at left and the slide number at right, both Courier New 12pt. brand.md says a content slide footer is the page number only; powerpoint.md wins for PowerPoint.
- The python-pptx default theme carries gradient fills and shadows in its format scheme; the build replaces them with flat fills and empty effect lists, and check_deck.py fails on any gradient or shadow.
- The template holds one guide slide per layout, each slide's text and notes its own instruction. The sample is about building a deck with the skill, so every number in it comes from brand.md.
- python-pptx stamps its own author; set_core sets the author to the Neuromine data team and fixed dates.

## Markdown and HTML formats, 2026-11-04

- Markdown: line 1 eyebrow `NEUROMINE` in inline code, the title H1 on line 3, the italic subtitle and date on line 5, a source line under each figure.
- The Markdown figure is the Word sample chart: build_markdown_figure.py calls figure_png in build_word.py, so both are the same bytes.
- HTML: no link element at all; the stacks are "Helvetica Neue", Arial, sans-serif and "Courier New", monospace. check_html.py reads the brand hex list from design/tokens.json.
- The header mark is `<span class="mark">Neuromine</span>` in plain Courier New Caption text. The cover has a brand blue band with the white logo as a data URI and the wordmark beside it, the only place either appears.
- check_html.py builds its Windows browser path from SYSTEMDRIVE, so the file holds no absolute path.
- The md sample links to an unnumbered H3 fragment, and the HTML sample to an id like s4, since a fragment that is a hash and a number reads as a ticket id to a public-repo scan.

## Excel format, 2026-11-04

- Named styles: the seventeen sheet styles Neuromine Title, Subtitle, Eyebrow, Date, Table Title, Header, Text, Text Wrap, Integer, Decimal, Percent, Total Label, Total Integer, Total Decimal, Total Percent, Caption and Note, plus Neuromine Date Cell for a date column. check_xlsx.py requires all eighteen.
- The header row is brand blue #1F4FAE with bold white text, not the light #F4F6FB head of brand.md; excel.md wins for Excel.
- The eyebrow reads NEUROMINE, the print header Neuromine. No logo and no wordmark in a workbook.
- Fonts name Helvetica Neue with the Swiss family, so Windows Excel shows Arial; Excel has no fallback list.
- save writes each formula's computed value as its cached value and sets fullCalcOnLoad, so Protected View shows the numbers.

## Neuromine brand skill, 2026-11-04

The skill moved into its own public repo, Neuromine-Data/brand-skill, under the MIT license.

- Name: neuromine-brand-skill. The style names (Neuromine Table, Neuromine Title and the rest), the theme names, the eyebrow NEUROMINE, the header and print header Neuromine, the HTML mark and title suffix " - Neuromine" and the file author "Neuromine data team" all read Neuromine. Answer ALS and ALS TDI are named only where a rule lists the programs the brand serves.
- Palette, fonts, sizes, chart series, layouts, styles and checks did not change. The logo is the same file under a new name, design/assets/neuromine-logo-light.png; it is the white logo the Variant Tool header serves, and there is no other Neuromine logo.
- Wordmark, read from the Variant Tool header (variant_reports.py SITE_TITLE and header CSS, docs/design/tokens.css --vr-font-head): the text NEUROMINE, typed in capitals with no text-transform; weight 700; no letter-spacing set, so 0; white on brand blue; 19px beside a 54px logo (the !important header rules, which win over the 22px and 50px rules later in the file); font stack Roboto, Helvetica Neue, Arial. Here it is Helvetica Neue 700 with Arial, as for every heading, since Roboto is not on every machine.
- Wordmark size is kept as the ratio 0.35 of the logo height: 19pt beside the 0.75in logo on a slide and in the Word cover band, 18px beside the 50px HTML logo.
- The Variant Tool gap between the logo and the title comes from its page template, not its CSS, so it was not read. The gap here is the brand clear space, half the logo height (24px on a page, the nearest spacing step).
- In PowerPoint the wordmark is a text box on the Title and Closing layouts. In Word it is drawn into the cover band PNG with the Arial Bold file matplotlib finds, so the band bytes, like the chart, are stable on one machine only.
- Source lines that named one program's release now read "Source: Release 8 index." in the format files; brand.md gives one example per program.
- Not checked by eye: nobody opened the built files in Word, PowerPoint or Excel, or looked at the printed HTML (no window may open). The Word band PNG was looked at as an image. Someone should open neuromine-sample.docx and neuromine-sample.pptx and confirm the wordmark sits level with the logo and clear of the eyebrow.

## SOP format, 2026-11-04

SOPs used to start from the house SOP template as it is, in its own look, and were not checked. They now start from a Neuromine SOP template built from tokens.json that keeps the house structure, with references/sop.md and scripts/check_sop.py. The house template was read with python-docx for its structure only; it is not copied or published here.

What happened to each part of the house template:

| House template | Here | How |
| --- | --- | --- |
| Title "Standard Operating Procedure" on the cover | Kept, restyled | Title style, 20pt bold brand blue, Helvetica Neue |
| Logo picture on the cover | Dropped | No logo in an SOP; the title table is the cover |
| Title table, one row of three cells, each a bold label over a placeholder | Kept, restyled | Neuromine Table with a label row (Title, Version Number, Effective Date) over a value row, the same fields |
| "Revision History" title and table, a label row and three empty rows | Kept, restyled | SOP Heading 14pt Ink; Neuromine Table with the same three labels and one row; a revision adds a row |
| "Last Modified By" and "Date" lines over underscores | Dropped | It asked for a person's name; the revision history and the file properties hold the date and a role |
| Page break and the title repeated on page 2 | Restyled, title dropped | Purpose has page break before; the header carries "Standard Operating Procedure" on every page |
| Five sections as bold underlined Normal lines with automatic numbers 1 to 5 | Kept, restyled | SOP Section style, 14pt bold Ink, numbered by the style; "Definitions/Acronyms" reads "Definitions and Acronyms" |
| Term lines 3.1, 3.2 and step lines 4.1, 4.2 from two lists with typed prefixes in the level text (3.%2, 4.%2) | Kept, restyled | One multilevel list: sections %1., steps %1.%2, terms %1.%4, so the numbers follow the section and no level text is typed |
| "Example bullets" under each step line | Kept, restyled | SOP Detail lines with the brand blue bullet, one each for Inputs, Steps, Outputs and Checks |
| "Instruction:" lines | Restyled | "Guide:" lines, as on the PowerPoint guide slides; check_sop.py fails while one is left |
| Empty header | Restyled | "Neuromine" at left, "Standard Operating Procedure" at right, Courier New 9pt with the rule |
| Footer "Page \| n" | Kept, restyled | The SOP id at left, "Page n" at right, Courier New 9pt with the rule |
| Letter, 1 inch margins, header and footer 0.5 inch | Kept | Same values |
| Calibri, black text, embedded symbol font | Dropped | Helvetica Neue with Arial fallback, Courier New, brand colors, no embedded font |
| File author naming a person | Dropped | "Neuromine data team" with fixed dates, so the build is byte stable |

Choices:

- The SOP styles are not Heading styles. A Heading style on a numbered line shows two numbers in a PDF when a typed number is also there, and the checker fails any Heading or Title style on a numbered line.
- Word ties a list level to one paragraph style. With steps and terms both on the second level of the list, Word dropped the term numbers; terms moved to their own level (shown as 3.1, 3.2) and the PDF text was read to confirm each number shows once.
- Numbering sits in the styles, so a paragraph copied with deepcopy or added with the style keeps its number.
- The sample keeps its asset name, neuromine-sop-sample.docx. check_sop.py accepts that name and checks the name the SOP would carry (footer id, a full stop, a space and the title) against PREFIX-SOP-NNN-VN. Title.docx instead.
- check_sop.py holds the staff names of the public-repo guard as sha256 hashes of the lower case name, so the published script holds no name. The board checker fails if the hashes and its guard list differ.
- check_sop.py also fails non-ASCII text, a house rule for SOPs since the first ones were written.
- The sample is a made-up process, a monthly brand check of published documents, at about 800 words. It names no person, identifier or release count.
- Rendered once to PDF in a hidden Word to read the numbers; nobody has opened the files in Word by eye.
