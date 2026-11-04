# HTML

Rules for writing a self-contained .html report in the Neuromine brand. Read https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/brand.md first; this file maps those rules onto HTML and CSS. Where the two disagree, this file wins for HTML. The reader turns the HTML into PDF, so do not produce a PDF.

## Start from the template

Download https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-report.html and replace the body content; the body is a sample of every block below. Keep the stylesheet whole. Append document-specific CSS at the end of the same `<style>` block. No second `<style>` block, and no `style` attribute anywhere except `style="width: N%"` on a `.bar`.

## File rules

- One .html file. No framework, no build step, no `<script>`.
- No `<link>` element at all: no remote stylesheet, no web font, no preconnect. The fonts are the installed Helvetica Neue and Arial. No remote images. Ordinary citation hyperlinks (`<a href="https://...">`) are allowed.
- Images are inline: SVG markup or a `data:` URI. The cover logo is a `data:` URI of design/assets/neuromine-logo-light.png.
- UTF-8 text with LF line endings, ASCII only as brand.md says. Write entities only where HTML needs them: `&amp;`, `&lt;`, `&gt;`, and quotes inside attributes.
- `<html lang="en">`, `<meta charset="utf-8">`, a viewport meta, `<meta name="color-scheme" content="light">`, and a `<title>` of the form "Document title - Neuromine".

## Tokens

Brand colors and font families are defined once as custom properties on `:root`; rules use `var(...)` for those. Dimensions and other values are plain. Exceptions: literal values inside `@page` (where `var()` does not work), and inline SVG may use presentation attributes with literal brand hex (`fill="#1F4FAE"`).

| Property | Value | Brand token |
| --- | --- | --- |
| `--page` | #F8F9FA | Page |
| `--surface` | #FFFFFF | Surface |
| `--brand` | #1F4FAE | Brand |
| `--brand-dark` | #163D8C | Brand dark |
| `--brand-light` | #4070D9 | Brand light, link hover only |
| `--tint-1` | #EEF4FF | Tint 1 |
| `--tint-3` | #C4D9FF | Tint 3 |
| `--ink` | #1A2332 | Ink |
| `--text` | #2C3E50 | Text |
| `--text-2` | #374151 | Text 2 |
| `--caption` | #4B5563 | Caption |
| `--placeholder` | #6E7581 | Placeholder |
| `--border` | #8A9097 | Border |
| `--rule` | #E5E7EB | Rule |
| `--table-border`, `--table-head`, `--table-head-rule`, `--row-hover` | #DFE3EB, #F4F6FB, #D8DCE5, #EEF3FB | Tables |
| `--on-brand` | #FFFFFF | text on brand blue |
| `--series-1` to `--series-6` | #008300, #E87BA4, #4A3AA7, #1BAF7A, #EDA100, #882255 | Categorical |
| `--seq-1` to `--seq-6` | #EEF4FF, #DCE8FF, #C4D9FF, #4070D9, #1F4FAE, #163D8C | Sequential, light to dark |
| `--div-1` to `--div-6` | #163D8C, #1F4FAE, #C4D9FF, #F8D7DA, #DC3545, #721C24 | Diverging, blue to red |
| `--chart-axis`, `--chart-grid` | #6A7683, #D9E1EA | Chart axis text, grid and baseline |

Fonts:

- `--font-body`: "Helvetica Neue", Arial, sans-serif. Headings use it too, at weight 600.
- `--font-mono`: "Courier New", monospace.
- Helvetica Neue is not on Windows or in most agent sandboxes, so a render there uses Arial. Do not change the stacks because of that, and never add a web font link.

## Light only

The report is light. Set `color-scheme: light` on `:root` and add the color-scheme meta. Do not add `prefers-color-scheme` styles.

## Page and print

- Screen: body on `--page`, content in a `.page` column on `--surface`, about 50rem wide. Spacing in steps of 4, 6, 8, 12, 14, 16, 20 and 24px.
- PDF route: Chrome or Edge, File > Print > Save as PDF, background graphics on, browser headers and footers off. "Print preview" in these rules means that renderer. wkhtmltopdf and similar tools do not draw margin-box page numbers.
- `@page { size: letter portrait; margin: 1in; }`, matching brand.md page setup and the Word template.
- Page numbers come from `@page` only: `@bottom-right { content: counter(page); }` in Courier New 9pt #6E7581, none on the first page (`@page :first`). There is no screen footer.
- `print-color-adjust: exact` on body so fills print.
- In `@media print`: hide `nav` and `.no-print`, drop the `.page` padding and width, make the cover fill page 1 with `break-after: page`.
- Every h1 section starts a new page with its header block at the top: `.doc-header { break-before: page; break-after: avoid; }`.
- Headings and captions `break-after: avoid`; blockquote, callout, figure and `pre` `break-inside: avoid`.
- Tables: `thead` repeats (`display: table-header-group`) and a `tr` never splits. Long tables may break between rows. Add class `short` to a table of up to about 12 rows to keep it whole (`table.short { break-inside: avoid; }`).
- A figure taller than the printable area is split into two figures, each with its own caption.
- `.page-break` forces a break.

## Header block

- `<header class="doc-header">` right before each `h1`, with the mark at left and the section at right ("2. Counts by site", same number as the h1), a 1px `--rule` line below.
- The mark is `<span class="mark">Neuromine</span>`: the program name in text, never the logo (there is no dark logo for white). It always reads "Neuromine".
- Header text: Courier New 12px, `--caption`, sentence case.

## Cover

`<section class="cover">`: a `.cover-band` across the top in `--brand`, about a quarter of the page high, with the white logo at its top left (50px high on screen, never under 38px, clear space of at least half its height) and the wordmark beside it: `<span class="wordmark">NEUROMINE</span>` after the `img`, both in a `.lockup`, 18px, 700, white, no letter spacing, 24px to the right of the logo (brand.md Wordmark). Below the band, left aligned: the eyebrow (two to four words, 11px, 700, uppercase, 0.3px letter spacing, `--caption`), the title (`.cover-title`, 600, 28px, `--brand`), the subtitle (`--text-2`) and the date at the bottom (`.cover-date`, Courier New, `--caption`, `4 November 2026` per brand.md). The logo goes on the band only, once.

## Headings

- `h1` once per section, numbered in order from 1 ("2. Counts by site"): 28px, `--brand`.
- `h2` 15px and `h3` 14px, `--ink`. All weight 600, line height 1.25. No rules under headings.
- No `h4` or deeper. Do not skip a level. Sentence case.
- HTML numbers `h1` sections, while Markdown numbers `##` under one document H1; converting between them shifts every level by one on purpose.

## Body text

14px, line height 1.6, `--text`, left aligned, paragraphs `margin: 0 0 8px`. Links in `--brand`, `--brand-light` on hover, a 2px `--brand` outline on `a:focus-visible`. Bold marks a key term on first use only; no italics for emphasis.

## Code

- Inline `code`: `--font-mono` 13px on `--tint-1`.
- `pre` block: `--font-mono` 13px, `--tint-1` fill, 8px 12px padding, no border, `break-inside: avoid`. Long lines scroll horizontally on screen and wrap in print.
- No syntax coloring.

## Lists

- `ul`: no default bullets. Each `li::before` is a round 6px `--brand` dot, 4px from the left; text indented 20px; 6px between items.
- `ol`: numbers in `--brand` (`li::marker`), same indent.
- Nest at most one level. List only the items there are.

## Tables

- A caption line right before the table: `<div class="caption" id="t1">Table 1. Participants by site and visit</div>`, numbered in order. `<caption>` is not used; tie the line to the table with `aria-labelledby="t1"` on the `table`.
- Always `<thead>` with `<th scope="col">` cells: `--table-head` fill, `--ink` text, weight 600, a 1px `--table-head-rule` line under it, left aligned.
- Body rows: `--text-2` text, a 1px `--table-border` line between rows, no vertical lines, outer border `--table-border`, 6px radius, row hover `--row-hover` on screen. Cell padding 6px 8px.
- First column names the row and is left aligned. Number columns get `class="num"` (right aligned) on both `th` and `td`; the table uses `font-variant-numeric: tabular-nums`.
- No empty cells: `-` means missing, `n/a` means not applicable. Headers are plain words in sentence case, never raw column names. Units go in the header.

## Callouts and quotes

- Callout: `<div class="callout"><strong>Key point:</strong> ...</div>`. `--tint-1` fill, 3px `--brand` left border, 8px 12px padding, 8px radius, `--text`. Labels: Note, Key point, Warning, always bold. One callout per h1 section at most.
- Quote: `<blockquote>` with the quote in `<p>` and the source in a `<footer>` (Courier New 13px, `--caption`). `--text-2`, indented both sides, no fill and no left bar, so it never looks like a callout.

## Figures and charts

- Every chart or image is a `<figure>` with a `<figcaption class="caption">` above it: "Figure 1. ...", numbered in order, 13px `--caption`. A source line (`<p class="source">Source: Release 8 index.</p>`) closes the figure.
- Simple bar charts are pure CSS (see `.bar-chart` in the template): one `.bar` per row, width as a percent of the axis maximum. A single series is `--brand` (a plain `.bar`); several categories take classes `s1` to `s6` in brand order. The value label is a `.bar-value` span after the `.bar` inside `.bar-track`, just past the bar's end, in `--text-2`. Never text inside a bar: chart colors are for marks, not text.
- More complex charts are inline SVG (see Figure 2 in the template) in the same colors: axis and tick text in #6A7683, grid lines and baseline in #D9E1EA, few and light, Arial for SVG text.
- Give a CSS chart `role="img"` and an `aria-label` that lists the data. Inline SVG needs `role="img"`, a `<title>` and a `<desc>` that lists the data.
- Every chart needs a text or table equivalent nearby.
- No 3D, no shadows, no gradients or patterns in marks. A legend only when there is more than one series and the lines are not labeled.
- Use the categorical series in order; `--seq-*` for ordered values, `--div-*` for values either side of a midpoint (blue the low or favorable end).

## Voice

Voice and prose: follow the Writing section in brand.md.

## Checklist before sending

Each line is one `scripts/check_html.py` check unless marked human-only.

1. UTF-8, LF line endings, one final newline.
2. `lang="en"`, charset, viewport and color-scheme metas, title ends in " - Neuromine".
3. No `<script>`, no `<link>`, no remote images.
4. One `<style>` block that starts with the template stylesheet unchanged.
5. No `style` attribute except width on `.bar`.
6. No brand hex outside `:root`, `@page` and inline SVG.
7. Every mark reads "Neuromine".
8. No `h4` to `h6`; h1 numbered from 1, each right after its header block with the same number; no skipped levels.
9. No screen footer block (`.doc-footer`).
10. Every table has a "Table N." caption line before it, `aria-labelledby`, a `thead` and `scope="col"`; every figure has a "Figure N." caption.
11. No empty cells; number cells and their header carry `num`.
12. Callouts use Note, Key point or Warning, at most one per h1 section.
13. Charts and SVG have `role="img"`; SVG has `<title>` and `<desc>`.
14. Quotes have a `p` and a `footer`.
15. No bare `&`, no `<` in text.
16. print-render: PASS or SKIP; a SKIP means print by hand and check the page count. On PASS the page count looks right.
17. Human-only: print preview shows the cover alone on page 1, each section on a new page, page numbers from page 2, and no table split mid-row.
18. Human-only: one voice throughout, dates as `4 November 2026`, every chart has a text or table equivalent and a source line.
19. `python scripts/check_html.py file.html` passes.
