# Brand

Shared brand rules for every Neuromine output: documents, decks, sheets, pages and figures. Read this before any format file. Quiet and clinical: brand blue on white, dark ink for text, thin grey rules, plenty of space, one strong element per page. No gradients, clip art, stock photos, ornamental icons or color for its own sake. Use no color, font or size that is not here. design/tokens.json holds the same values for code.

## Palette

Brand blue is the accent.

- Brand: #1F4FAE. Headlines on slides, the document title, card titles, links, the single series in a chart, the background behind the logo.
- Brand dark: #163D8C. Text set in blue on a white slide, pressed buttons.
- Brand light: #4070D9. Hover state on a web page only. Never for text.
- Tint 1: #EEF4FF. Lightest fill, for a callout or a highlighted panel.
- Tint 2: #DCE8FF. Hover fill.
- Tint 3: #C4D9FF. Selected fill, and the eyebrow, rule and subline on a brand blue slide.

Ink and text.

- Ink: #1A2332. Headings below the title.
- Text: #2C3E50. Body copy.
- Text 2: #374151. Secondary text and table cells.
- Caption: #4B5563. Captions, eyebrows, source lines, header and footer text.
- Placeholder: #6E7581. Muted text that still has to be read (placeholders, "no data" cells, page numbers).
- Muted: #9CA3AF. Disabled or decorative marks only. Never for text a reader needs.

Surfaces and rules.

- Page: #F8F9FA. Page background behind cards.
- Surface: #FFFFFF. Cards, slides, table body, documents.
- Border: #8A9097. Borders that must be seen (inputs, card edges, dividers that separate content).
- Rule: #E5E7EB. Light dividers, header and footer rules, the frame around a screenshot.
- Table border #DFE3EB, table head #F4F6FB, table head rule #D8DCE5, row hover #EEF3FB.

Status. Each has a strong color, a flat background and a text color for the box. Status colors mean status only, never decoration; the one chart use is the Diverging series under Charts.

- Error #DC3545, background #F8D7DA, text #721C24.
- Success #28A745, background #D4EDDA, text #155724.
- Info #1F4FAE (brand), background #D1ECF1, text #0C5460.
- Warning #B45309, used as text color for an inline warning.
- Required #E02424: the asterisk on a required field.

Contrast.

- Text reaches 4.5:1 against its background. Every text color above passes on white and on Page.
- A border or chart line a reader must see reaches 3:1. Use Border #8A9097 for these, not Rule #E5E7EB.
- On brand blue, use only white or Tint 3 for text.
- Chart colors are for marks (bars, points, lines), not text.

## Typography

- Headings: Helvetica Neue 600, fallback Arial, then sans-serif. On the web: "Helvetica Neue", Arial, sans-serif.
- Body: Helvetica Neue 400, same fallback. Line height 1.6 on the web, 1.15 in Office files.
- Labels and page furniture (header and footer text, page numbers, file names, commands and code): Courier New 400, fallback monospace.
- Weights: 400 body, 500 button labels, 600 headings and card titles, 700 the eyebrow and slide headlines.
- Sizes on a web page (px): 11 eyebrow, 12 chips and small notes, 13 buttons and captions, 14 body and table cells, 15 card titles, 28 the one page heading.
- Sizes on a slide (pt, 16:9 at 13.333 by 7.5 in): 12 page number, 14 eyebrow, 16 command line and chart labels, 22 title slide subline, 40 content slide headline, 54 title slide headline.
- Sizes in a document (pt): 9 eyebrow, caption and source line, 11 body and table cells, 14 section heading, 20 document title.
- Sentence case everywhere: titles, headings, buttons, table headers, axis labels. No all caps, except the small eyebrow label, uppercase with a little letter spacing (0.3px on the web, 2pt on a slide), and the NEUROMINE wordmark (see Cover page). Program names keep their own casing: Answer ALS, ALS TDI, Neuromine.

## Headings

- Title (H1 on a page): Helvetica Neue 600, Brand #1F4FAE; 20pt in a document, 28px on a page, 40pt on a content slide.
- Section heading (H2): Helvetica Neue 600, Ink #1A2332, 14pt in a document, 15px on a page.
- Sub heading (H3): Helvetica Neue 600, Ink #1A2332, at body size.
- An eyebrow may sit above the title: two to four words, uppercase, 700, Caption on white, Tint 3 on brand blue.
- No rules under headings. Levels never skip.

## Body text

11pt in a document, 14px on a page, colour Text #2C3E50, left aligned, never justified. Paragraph spacing about 0.6em. Bold marks a key term on first use only; structural labels such as callout labels are bold every time. No italics for emphasis.

## Tables

- Header row: fill #F4F6FB, text Ink #1A2332, weight 600, a 1px #D8DCE5 rule under it.
- Body rows: text Text 2 #374151, a 1px #DFE3EB rule between rows, no vertical rules, outer border #DFE3EB. On the web, 6px radius and row hover #EEF3FB.
- Headers are plain words in sentence case, never raw column names or snake_case. The same column gets the same header everywhere.
- Text left aligned, numbers right aligned with tabular figures, the header over a number column right aligned too.
- A caption line above each table: "Table N. What it shows."

## Header and footer

There is no dark logo, so on white the program name is written in text.

- Header: "Neuromine" (or the program the document is for, Answer ALS or ALS TDI) at left; the section name at right. Courier New, Caption #4B5563, 9pt. A 1px Rule #E5E7EB under the header.
- Footer: page number at right in Courier New, Placeholder #6E7581, 9pt; nothing at left. A 1px Rule #E5E7EB above the footer.
- On a content slide the footer is the page number only, bottom right, Placeholder at 12pt.

## Cover page

- A full-width brand blue #1F4FAE band across the top of the page, about one quarter of the page high, with the white logo at its top left.
- Below the band, left aligned: the eyebrow (Caption), the title (Helvetica Neue 600, Brand, 20pt or larger), the subtitle (Text 2).
- Date at bottom left in Courier New, Caption.
- A title or closing slide is brand blue all over: white logo top left at 0.75in high, eyebrow in Tint 3, a short Tint 3 bar, headline in white at 54pt, subline in Tint 3 at 22pt.

The logo is design/assets/neuromine-logo-light.png, the white Neuromine mark (web copy: https://dataportal.answerals.org/logo_light.png).

- Use it only on brand blue #1F4FAE. Never on white, a tint, a photo or any other color, where it disappears. On white, write the program name in text instead.
- Keep clear space of at least half the logo's height on every side.
- Never stretch, squash, recolor, rotate, crop or add a shadow. Scale it with its sides locked.
- Smallest size: 38px high on a screen, 0.5in on a slide.
- Once per page, top left.

Wordmark: the lettering NEUROMINE sits to the right of the logo, as in the Variant Tool header.

- Helvetica Neue 700, fallback Arial (the Variant Tool sets it in Roboto first; Roboto is not on every device). White #FFFFFF on brand blue #1F4FAE only.
- All capitals as typed, NEUROMINE, with no letter spacing (0) and no other word beside it.
- Size about 0.35 of the logo height (19 to 54 in the Variant Tool): 19pt beside the 0.75in logo on a slide or the document cover band, 18px beside the 50px logo on a page.
- Centered on the logo's height, half the logo height to its right (24px on a page, the nearest spacing step).
- Used only next to the logo: the title and closing slides, the document cover band and the HTML cover band. Never alone and never on white. The eyebrow NEUROMINE in a Markdown file or a sheet is an eyebrow, not the wordmark, and keeps the eyebrow style.

## Callouts

Tint 1 #EEF4FF fill, 3px Brand #1F4FAE left border, 8px to 12px inside padding, text Text #2C3E50, square corners in documents and slides, 8px radius on the web. The label (Note, Key point) is bold. For an error, a success or a warning use the status box instead: the status background and text color, full width. Each format file sets the callout limit.

## Lists

Round bullet in Brand #1F4FAE, text in Text, indented about 0.25in (or 20px), 6px between items. Numbered lists use the same indent with the number in Brand. A list holds only the items there are. Two levels at most.

## Charts

Use the series in this order.

- Categorical: #008300, #E87BA4, #4A3AA7, #1BAF7A, #EDA100, #882255. Keep a category on the same color across every chart in a set. More than six categories: group the small ones into "Other" or split the chart. No seventh color.
- Single series: Brand #1F4FAE for every mark.
- Sequential, light to dark: #EEF4FF, #DCE8FF, #C4D9FF, #4070D9, #1F4FAE, #163D8C.
- Diverging, blue to red, light in the middle: #163D8C, #1F4FAE, #C4D9FF, #F8D7DA, #DC3545, #721C24. Blue is the low or favourable end. This is the only chart use of a status color.
- Axis text and tick labels #6A7683. Grid lines and baseline #D9E1EA, few and light.
- No 3D, no shadows, no gradients or patterns in data marks.
- Label every axis in sentence case with its unit.
- A source line under every chart in Caption at caption size, naming the program and the release: "Source: Answer ALS Release 8 index." or "Source: ALS TDI Release 2 index."

## Page setup

- Documents: US Letter, portrait, single column, 1 inch margins on all four sides. Square corners.
- Slides: 16:9, 13.333 by 7.5 in. A content slide is white with a brand blue strip 0.18in wide down the left edge.
- Web pages: spacing in steps of 4, 6, 8, 12, 14, 16, 20 and 24px only. Radius 4 buttons, 6 inputs and tables, 8 cards and status boxes. Soft blue tinted shadows on cards only (0 2px 8px rgba(31, 79, 174, 0.08)). No shadows in documents, slides or charts.

## Dates

Full dates are day, month, year with no ordinal and no leading zero: `3 November 2026`. Where only the month is shown (a cover date), write `November 2026`. In a file name use `2026-11-03`.

## Writing

These rules apply to every word in every format.

The default voice is neutral third person. The request overrides it: when it asks for a voice, or the occasion clearly calls for one of the options below, use that.

- Plain, first person: "I looked at the Release 8 index and found that most participants have a cell line."
- Neutral, third person: "The Release 8 index shows that most participants have a cell line."
- Formal: "Neuromine has reviewed the Release 8 index, which shows that most participants have a cell line."

Plain first person for a personal note; neutral third person for a report, an analysis or anything with no named author; formal for a partner-facing or contractual document. One voice per document; never mix.

Characters and punctuation: ASCII only, no smart quotes and no symbols outside plain keyboard characters. No em dash and no en dash; use a comma, a colon, a full stop or parentheses. Hyphens only inside compound words. No exclamation marks.

Plain prose: plain words and short sentences, one idea each; active voice; concrete nouns and numbers over adjectives; say the finding first, then the evidence. On a page or slide, a headline states the point in one short sentence, not a topic ("Release 8 adds cell line fields" beats "Cell lines"). A caption says what the reader is looking at and where it came from, in one sentence. Labels are nouns, buttons are verbs.

Banned patterns, with the fix:

- Filler openers (`In today's fast-paced world`, `It is important to note`): start with the finding.
- Summary closers (`In conclusion`, `Overall,`): stop when the point is made.
- Openers and closers that talk to the reader (`Great question`, `I hope this helps`, an offer to help more): leave them out.
- Stock words (`delve`, `leverage`, `robust`, `seamless`, `utilize`, `cutting-edge`, `comprehensive`, `empower`, `crucial`, `landscape`, `journey`, `unlock`, `harness`, `tapestry`, `navigate`): use the plain word.
- Lists padded to three items: list only the items there are.
- Bullet walls when a sentence does the job: write the sentence.
- A bold label on every bullet: write the bullet as a sentence.
- Rhetorical questions: state the answer.
- `Not only X but also Y`: write "X and Y".
- Stacked hedges (`may potentially`): one hedge at most, or a number.

scripts/prose.py fails on dashes, the stock words and the quoted phrases above.

Content honesty: do not invent numbers, names, sources or quotes. Program facts come from the Lou pages and data facts from the data pages; if neither has it, say so. Counts may be repeated only when quoted from a Lou page. Mark missing data with the format's missing marker. Say what was done and what was not; if a step was skipped or a check failed, say so. No comments in code.

Length: brief by default. The output is as long as the content, never padded to a page count.
