# Excel

Rules for writing a .xlsx in the Neuromine brand. Read [brand.md](brand.md) first; this file maps those rules onto Excel. Where the two disagree, this file wins for Excel.

## Start from the template

Open `assets/neuromine-template.xlsx` (openpyxl: `load_workbook("assets/neuromine-template.xlsx")`). It holds the named styles, the theme colors and fonts, and one formatted sheet named "Sheet1" with the title block, a table title row, a header row, print setup and the freeze pane. Rename that sheet before you fill it. `assets/neuromine-sample.xlsx` shows the helpers below in use: two data sheets, each with a table, a total row, a chart and source lines, and a Notes sheet. Its numbers are made up for the example. Apply the named styles; never format a cell by hand when a style exists.

To change the template, edit `scripts/build_excel.py` and run it. Do not edit the .xlsx files by hand.

## Helpers

`scripts/neuromine_excel.py` needs only openpyxl and does nothing at import. Add the scripts folder to the path, then import it: `sys.path.insert(0, "scripts")`, `import neuromine_excel`. It also holds the brand constants (fonts, `BRAND`, `INK`, `CHART`, `SEQUENTIAL`, `DIVERGING`, `SINGLE`, `FMT_INT`, `TITLE_ROW`, `HEADER_ROW` and the rest).

- `setup_sheet(ws, name, title, subtitle, when, landscape=False)`: renames a sheet and sets the title block, row heights, freeze pane, gridlines off, print setup, print titles and header and footer.
- `new_sheet(wb, name, title, subtitle, when, landscape=False)`: a new sheet with all of `setup_sheet` applied; returns it. `landscape` flips the orientation only.
- `set_widths(ws, widths)`: column widths from column A, in Excel character units. Set them before `add_table`.
- `add_table(ws, heads, rows, styles, first_row=HEADER_ROW, total=None, caption=None, autofilter=False, title=None)`: `title` in Neuromine Table Title in column A on the row above the header row, header row, body rows with one style name per column in `styles`, an optional total row, the caption on the next row, and an AutoFilter on header and body when `autofilter` is true. `total` is the label, or a list with the label first and one entry per column: `None` for the default (SUM for integer and decimal columns, empty otherwise), `""` for empty, or a formula. Any body or total cell can be a `(formula, value)` pair: the formula goes in the cell and `value`, computed in Python, is kept for `save` to write as the cached result. A default SUM total gets its value from the body rows. In formulas `{row}`, `{first}` and `{last}` become the current, first body and last body row. Clears leftover template header cells to the right of `heads`, and the template's table title when `title` is not given. Raises the header row to 30 when a header is longer than its column or holds a newline. Returns `(first_body_row, last_body_row, next_free_row)`.
- `add_chart(ws, kind, first, last, ncols, anchor, palette="categorical", label_col=1)`: a native chart over body rows `first` to `last`, labels in `label_col`, series in the next `ncols - 1` columns. `kind` is "column", "bar" or "line"; `palette` is "categorical", "sequential", "diverging" or "single". A categorical chart with one series uses brand blue. The legend sits at the top and does not overlay the plot area. Returns the chart.
- `chart_text(size, color, bold=False)`: the rich text block the charts use for axis and legend text.
- `finish(wb)`: sets `fullCalcOnLoad` and drops any empty sheet still named "Sheet" or "Sheet1".
- `save(wb, path)`: calls `finish`, saves with openpyxl, then writes each pair's value into its formula cell as the cached value. Use it instead of `wb.save`.

## Starter

```python
import sys; sys.path.insert(0, "scripts")
import neuromine_excel as a
from datetime import date
from openpyxl import load_workbook
wb = load_workbook("assets/neuromine-template.xlsx")
ws = wb.active
a.setup_sheet(ws, "Visits", "Visits by site", "Example counts by site and visit", date(2026, 11, 4))
a.set_widths(ws, [24, 14, 14])
first, last, nxt = a.add_table(ws, ["Site", "Screening", "Month 3"], [["Site A", 120, 108], ["Site B", 95, 91]],
    ["Neuromine Text", "Neuromine Integer", "Neuromine Integer"], total="All sites",
    caption="Table 1. Participants by site and visit", title="Participants by site and visit")
ws.cell(row=nxt, column=1, value="Source: example numbers, not release data.").style = "Neuromine Note"
a.add_chart(ws, "column", first, last, 3, f"A{nxt + 2}")
ws.cell(row=nxt + 20, column=1, value="Figure 1. Participants by site and visit").style = "Neuromine Caption"
a.save(wb, "neuromine-visits-by-site.xlsx")
```

## Fonts

| Role | Font | Fallback |
| --- | --- | --- |
| Title, table title | Helvetica Neue bold | Arial |
| Body, table cells, source lines | Helvetica Neue 11pt | Arial |
| Eyebrow, date, captions, print header and footer | Helvetica Neue 9pt | Arial |

Labels and page furniture use the body family. Courier New is not used in a workbook, which holds no code or commands.

Excel has no per-cell fallback list. The styles name Helvetica Neue with the Swiss font family, so Excel on a machine without it (most Windows machines) shows Arial or its nearest sans serif. Do not change the fonts in the template because a render you made yourself shows Arial. Never embed fonts. The theme fonts are Helvetica Neue for headings and body.

This file overrides generic xlsx guidance on fonts (Arial or Times): Neuromine workbooks use the template fonts.

## Named styles

Use only these cell styles. They show under Home > Cell Styles; openpyxl takes the same names (`cell.style = "Neuromine Text"`). The template has seventeen styles for the rows of a sheet, plus Neuromine Date Cell for a date column.

| Use | Style | Look |
| --- | --- | --- |
| Eyebrow (A1) | Neuromine Eyebrow | Helvetica Neue bold 9pt, #4B5563 |
| Sheet title (A2) | Neuromine Title | Helvetica Neue bold 20pt, #1F4FAE |
| Subtitle (A3) | Neuromine Subtitle | Helvetica Neue 14pt, #374151 |
| Date (A4) | Neuromine Date | Helvetica Neue 9pt, #4B5563, `d mmmm yyyy` |
| Table title (row above the header) | Neuromine Table Title | Helvetica Neue bold 11pt, #1A2332, left, no border |
| Table header row | Neuromine Header | Helvetica Neue bold 11pt, white #FFFFFF on #1F4FAE fill, left, wrap on |
| Text cells | Neuromine Text | Helvetica Neue 11pt, #374151, left, thin #DFE3EB rule below |
| Long text cells | Neuromine Text Wrap | as Neuromine Text, wrap on, top aligned |
| Whole numbers | Neuromine Integer | as Neuromine Text, right, `#,##0` |
| Decimals | Neuromine Decimal | as Neuromine Text, right, `#,##0.0` |
| Percentages | Neuromine Percent | as Neuromine Text, right, `0.0%` |
| Total row label | Neuromine Total Label | Helvetica Neue bold 11pt, #1A2332, left, thin #8A9097 rule above |
| Whole number total | Neuromine Total Integer | as Neuromine Total Label, right, `#,##0` |
| Decimal total | Neuromine Total Decimal | as Neuromine Total Label, right, `#,##0.0` |
| Percent total | Neuromine Total Percent | as Neuromine Total Label, right, `0.0%` |
| Table or chart label | Neuromine Caption | Helvetica Neue 9pt, #4B5563 |
| Source or note line under a caption | Neuromine Note | Helvetica Neue 9pt, #4B5563 |
| Dates in a table | Neuromine Date Cell | as Neuromine Text, right, `yyyy-mm-dd` |

Normal is Helvetica Neue 11pt #2C3E50. A text column whose entries would need a width over 40 uses Neuromine Text Wrap at width 40. Need another number format (two decimals)? Copy the nearest style in the script under a new `Neuromine` name and rebuild; do not set formats cell by cell.

## Title block

Rows 1 to 4 of every sheet, in column A, no merged cells:

1. A1, Neuromine Eyebrow: "NEUROMINE". The mark reads "Neuromine".
2. A2, Neuromine Title: the sheet title, short, sentence case.
3. A3, Neuromine Subtitle: one line on what the sheet holds, with the data source and units when the caption does not give them.
4. A4, Neuromine Date: the date as a real date value, not text.

Row 5 stays empty. Row 6 is the table title, in Neuromine Table Title in A6 (`add_table(..., title=...)`): a short name for the table, sentence case, no "Table N." number. The table header row is row 7. The title row sits outside the frozen and printed rows, so the freeze pane and print titles name the header row alone. The logo never goes in a workbook: it is white and only sits on brand blue.

## Table header row

Row 7. Neuromine Header: brand blue #1F4FAE fill, bold white #FFFFFF text, left aligned, wrap on. The template row is 22 high; `add_table` makes it 30 when a header wraps. This is the one place a workbook differs from the light table head in brand.md: a sheet has no other strong element, and white on brand blue reads in every viewer, LibreOffice and Google Sheets included. Headers are plain words in sentence case, never raw column names.

## Body rows and totals

No banding: each body cell has a thin #DFE3EB bottom border (Excel border style `thin`, openpyxl `Side(style="thin")`) and no vertical rules. Do not fill body rows. Gridlines are off on every sheet. First column left aligned (Neuromine Text); number columns right aligned with a number style. One header row, no merged cells, no blank rows or columns inside a table. Every table has a Neuromine Table Title cell in column A directly above its header row, and keeps a Neuromine Caption cell in column A on the row under the table, or under the total row: "Table 1. Participants by site and visit". The caption goes under the table, as in Word; the source line ("Source: Release 8 index.") goes on the row under the caption in Neuromine Note.

A total row puts its label in Neuromine Total Label and each total in the Total style matching the column (Neuromine Total Integer, Neuromine Total Decimal or Neuromine Total Percent), with a thin #8A9097 rule above. A percent column total is a ratio formula (total numerator over total denominator) in Neuromine Total Percent, never a SUM of percentages and never plain Neuromine Percent. A date column total stays empty.

## Number and date formats

- Whole numbers: `#,##0`. Decimals: `#,##0.0`. Percentages: `0.0%`. Store values as numbers, never as text.
- Percentages are stored as fractions: 0.173 shows as 17.3%.
- Negatives with a minus sign, not brackets or red.
- Dates in tables: `yyyy-mm-dd`. Title block date: `d mmmm yyyy`, which shows as 4 November 2026.
- Keep one format per column. Say the units in A3 or the caption.

## Formulas and data

- Calculated cells stay formulas, and every formula is written with its computed value: pass a `(formula, value)` pair to `add_table` and save with `save`. Excel opens any file from email or a download in Protected View, which does not recalculate, so a formula with no cached value shows blank there, and so does any chart that reads it. `fullCalcOnLoad` stays set, so Excel recomputes once editing is enabled. Example: `("=D{row}/B{row}", month6 / screening)`.
- A blank cell means missing; 0 means zero. Never swap them.
- Note the data source and units in A3, the caption or the source line. Counts come from a named release; never invent them.
- Counts and identifiers only: never put participant level rows in a workbook that leaves the team.
- Do not re-save the file through LibreOffice to recalculate it. It reorders the named styles so openpyxl cannot resolve them and rewrites the theme. If cached values are needed, compute them in Python and write both.

## Column widths and freeze panes

- First column 24, number columns 14 (Excel character units). Widen to fit the longest entry plus 2, at most 40; wrap header text and long text instead of going wider.
- Do not hide columns or rows.
- Freeze panes at A8 so the title block, table title and header row stay in view (`ws.freeze_panes = "A8"`). For a table wider than the screen, freeze at B8 so the first column stays too.

## Print setup

Portrait, US Letter, 1 inch margins, fit to one page wide. Row 7, the header row, repeats on each printed page. Print header: "Neuromine" left and the sheet name right, Helvetica Neue 9pt #4B5563; footer: page number right, Helvetica Neue 9pt #6E7581. The template and `new_sheet` set all of this.

## Sheets

Name sheets in one or two plain words ("Visits", "Notes"), no default "Sheet1" in a sent file. Each sheet gets the title block. For a second sheet use `new_sheet`; never `copy_worksheet`, which drops the sheet view, the header and footer and the print titles.

## Notes sheet

Prose that is not tabular data (sources, definitions, callouts, lists, quotes, a summary) goes on a sheet named Notes, as one two-column table and nothing else:

- Heads `["Label", "Note"]`, styles `["Neuromine Text", "Neuromine Text Wrap"]`, widths `[24, 60]` (the one place a text column is wider than 40), title "Notes", caption "Table N. Notes".
- One item per row, with its label in column A:
  - `Note`, `Key point` or `Warning` for a callout. At most one `Key point` per sheet.
  - `Step 1`, `Step 2`, ... for a numbered list, in order from 1.
  - A short topic word or two for a bullet item ("Source", "Blank cells").
  - `Quote` for a quote: the quote text in straight quotes, then the source, in the Note cell.
  - `Summary` for a summary paragraph.
- No merged cells, no blank rows, no text outside the table.

```python
ws = a.new_sheet(wb, "Notes", "Notes on the data", "Sources, definitions and known gaps", date(2026, 11, 4))
a.set_widths(ws, [24, 60])
a.add_table(ws, ["Label", "Note"], [["Key point", "Most participants reached month 6."],
    ["Source", "Example numbers, not release data."], ["Step 1", "Take the counts from the release index."]],
    ["Neuromine Text", "Neuromine Text Wrap"], caption="Table 2. Notes", title="Notes")
a.save(wb, "neuromine-visits-by-site.xlsx")
```

## Wide sheets

The default is one table per sheet, portrait, one page wide. These exceptions are allowed; state each in the subtitle or caption:

1. More than about eight columns: turn the page sideways with `new_sheet(..., landscape=True)`, freeze at B8.
2. A counts or listing sheet with many rows: no chart, Neuromine styles on every column, AutoFilter on the header row (`autofilter=True`), no native Excel Table, the row count in the subtitle. Never participant level rows.
3. More than one table on a sheet: each has its own title row directly above its header row (pass `first_row` as the header row; the title goes on the row above) and its own caption, two empty rows between one table's caption (or source line) and the next table's title, numbered Table 1, Table 2.
4. A dashboard sheet: tables and charts side by side in column blocks, only when each block keeps its caption and the sheet still prints one page wide with the page turned sideways.

Never hide, merge or shrink text to fit.

## Charts

Use native Excel charts, not pictures. `add_chart` makes three kinds: column, bar and line. Line charts get 2.25pt lines and no markers. The theme accents are the categorical series, so a new chart picks them up in order: #008300, #E87BA4, #4A3AA7, #1BAF7A, #EDA100, #882255. A single series is brand blue #1F4FAE.

Four palettes, from brand.md: categorical (above) by default; single for one series; sequential for ordered bands such as releases or age bands, light to dark #EEF4FF, #DCE8FF, #C4D9FF, #4070D9, #1F4FAE, #163D8C (with fewer than six series it uses the darkest ones); diverging for signed values around a zero baseline, #163D8C, #1F4FAE, #C4D9FF, #F8D7DA, #DC3545, #721C24 (the category axis crosses at 0, labels stay at the bottom). Never color unrelated categories with the sequential series. Category axis line and gridlines 9525 EMU (0.75pt) #D9E1EA, no chart border, legend at the top (`legendPos t`) with overlay off (`overlay 0`) so it never sits on the plot area or the category axis labels, axis text 9pt #6A7683, legend text 9pt #374151. No chart title inside the chart: put a Neuromine Caption cell under it ("Figure 1. Participants by site and visit") and the source line under that in Neuromine Note. Place charts below the table, starting in column A.

## Voice

Voice and prose: follow the Writing section in brand.md.

## Checklist before sending

Each line is one `check_xlsx.py` check unless marked human.

- Started from the template: all eighteen Neuromine named styles and the Neuromine theme colors are present.
- Every sheet has the title block in A1:A4, a real date in A4 and an empty row 5.
- Every cell from row 6 down uses a Neuromine style, with no direct font, fill or number format.
- Row 7 is in Neuromine Header, bold white text on brand blue #1F4FAE.
- Each body column uses one style; number and date cells hold numbers or dates, not text.
- No total row sums a percent column; a percent total is in Neuromine Total Percent.
- No merged cells, no hidden rows or columns.
- Every sheet: freeze pane at A8 or B8, gridlines off, print titles row 7, one page wide, "Neuromine" in the print header.
- Every table has a Neuromine Table Title in column A directly above its header row.
- Every table has a "Table N." caption; every chart has a "Figure N." caption below it.
- Chart series use brand chart colors; no chart titles; the legend is at the top, off the plot area.
- A Notes sheet holds one table; column A labels are Note, Key point, Warning, Summary, Quote, Step N or a one or two word topic; Step rows run in order from 1; at most one Key point.
- `fullCalcOnLoad` is set when the workbook holds formulas.
- Every formula cell carries a cached value, so Protected View shows numbers and charts.
- No sheet named Sheet or Sheet1; file name has no spaces, for example `neuromine-visits-by-site.xlsx`.
- No em dash, en dash or banned word in any text cell (the prose check).
- Human: the palette fits the data (sequential for ordered, diverging for signed), blanks mean missing, sources and units are stated, no participant level rows.
- `python scripts/check_xlsx.py file.xlsx` passes.
