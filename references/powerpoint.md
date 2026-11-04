# PowerPoint

Rules for writing a .pptx in the Neuromine brand. Read [brand.md](brand.md) first; this file maps those rules onto PowerPoint.

## Start from the template

Open `assets/neuromine-template.pptx` (python-pptx: `Presentation("assets/neuromine-template.pptx")`). It holds the slide size, the master, the eight layouts, the theme colors and fonts, the Neuromine Table style and one guide slide per layout, whose text and notes say what goes where. Drop the guide slides (see Starter) and add slides from the layouts. `assets/neuromine-sample.pptx` uses each layout and each helper. No direct font, size or color on placeholder text: the master and layouts set them. The helpers in `scripts/neuromine_pptx.py` are the allowed exceptions for tables, charts, images and callouts; do not add others.

Slide size: 16:9 widescreen, 13.333 by 7.5 in. Side margins 0.9 in on content slides, 0.75 in on the title and closing slides. Do not change them. To change the template, edit `scripts/build_powerpoint.py` and run it. Do not edit the .pptx files by hand.

## Helpers

`scripts/neuromine_pptx.py` needs only python-pptx and lxml and does nothing at import. Add the scripts folder to the path, then import it: `sys.path.insert(0, "scripts")`, `import neuromine_pptx`.

- `layout_named(prs, name)`: the layout with that name.
- `add_hf(slide, types, values)`: copies the footer (`"ftr"`), slide number (`"sldNum"`) or date (`"dt"`) placeholders onto the slide; `values` maps `"ftr"` and `"dt"` to their text.
- `set_lines(frame, lines, numbered=False)`: fills a text frame from `(level, text)` pairs; level 0 or 1. With `numbered=True` each level 0 item is numbered "1." "2." in Brand #1F4FAE; level 1 items keep the round bullet.
- `add_table(slide, placeholder, rows, numeric_cols=None)`: a Neuromine Table in the placeholder's place. `numeric_cols` lists the column indexes to right align; left out, a column is numeric when every body cell is a number.
- `add_chart(placeholder, chart_type, data, palette="categorical")`: a native chart, styled; returns the graphic frame.
- `add_image(slide, path, caption)`: on the Picture layout, the picture fitted into the box with the caption as alt text and as the caption line.
- `add_callout(slide, placeholder, text, align="top")`: the brand callout in the placeholder's place. With `align="middle"` it sits level with the middle of the box.
- `set_core(prs, title)`: sets the title, the author (the Neuromine data team) and fixed dates in the file properties.

## Starter

```python
import sys; sys.path.insert(0, "scripts")
import neuromine_pptx as d
from pptx import Presentation
prs = Presentation("assets/neuromine-template.pptx")
for sld in list(prs.slides._sldIdLst):
    prs.part.drop_rel(sld.rId); prs.slides._sldIdLst.remove(sld)
foot = {"ftr": "Release 8 summary", "dt": "November 2026"}
s = prs.slides.add_slide(d.layout_named(prs, "Title"))
s.placeholders[15].text = "Neuromine"
s.shapes.title.text = "Release 8 summary"
s.placeholders[1].text = "One line on what the deck is"
d.add_hf(s, {"dt"}, foot)
s = prs.slides.add_slide(d.layout_named(prs, "Title and Content"))
s.shapes.title.text = "Release 8 adds cell line fields"
d.set_lines(s.placeholders[1].text_frame, [(0, "First point"), (0, "Second point"), (0, "Third point")])
d.add_hf(s, {"ftr", "sldNum"}, foot)
s = prs.slides.add_slide(d.layout_named(prs, "Closing"))
s.shapes.title.text = "Thank you"
d.set_core(prs, "Release 8 summary")
prs.save("neuromine-release-8-summary.pptx")
```

## Layouts

Use only these layouts, by name (`layout_named`).

| Layout | Use | Placeholders |
| --- | --- | --- |
| Title | First slide | eyebrow (idx 15), title, subline (idx 1), date (idx 10) via `add_hf(slide, {"dt"}, values)` |
| Section | Start of each section | title "N. Name", one line of text (idx 1) |
| Title and Content | Most slides | title, content (idx 1) |
| Two Content | Side by side, for example a chart and its points | title, left (idx 1), right (idx 2) |
| Table | One table per slide | title, caption above (idx 14), table (idx 13) |
| Chart | One chart per slide | title, chart (idx 13), caption (idx 14) |
| Picture | One image per slide | title, picture (idx 13), caption (idx 14) |
| Closing | Last slide | eyebrow (idx 15), title, text (idx 1) |

Title and Closing are brand blue #1F4FAE all over, like the cover page in brand.md: the white logo top left at 0.75 in from the top and left edges and 0.75 in high with the NEUROMINE wordmark beside it (white, bold, 19pt, from brand.md), the eyebrow in Tint 3 #C4D9FF, a short Tint 3 bar, the headline in white and the subline in Tint 3. They have no footer. Every other layout is white with the 0.18 in brand blue strip down the left edge and has the footer. The logo and wordmark are on the Title and Closing layouts only; never add either to a slide.

## Fonts

| Role | Font | Fallback |
| --- | --- | --- |
| Titles | Helvetica Neue bold (theme heading font) | Arial |
| Body, tables, charts, captions | Helvetica Neue (theme body font) | Arial |
| Footer, page number, date, command lines | Courier New | monospace |

The theme names Helvetica Neue, so text set to the theme fonts picks it up. On a machine without it PowerPoint substitutes a font, most often Arial. Agent sandboxes rarely have Helvetica Neue, so a render you make yourself shows the fallback; do not change the theme fonts because of it. Never embed fonts.

## Text sizes

| Text | Size | Color |
| --- | --- | --- |
| Title and closing headline | 54pt bold | white |
| Content and section headline | 40pt bold | Brand #1F4FAE |
| Title subline | 22pt | Tint 3 #C4D9FF |
| Section text | 22pt | Text 2 #374151 |
| Body level 1 | 22pt | Text #2C3E50 |
| Body level 2, table text, chart labels, command lines | 16pt | Text 2 #374151; chart labels #6A7683 |
| Eyebrow | 14pt bold, uppercase, 2pt letter spacing | Tint 3 #C4D9FF |
| Caption, date | 14pt | Caption #4B5563 (Tint 3 on blue) |
| Footer, page number | Courier New 12pt | Caption #4B5563, Placeholder #6E7581 |

These come from the master and layouts; do not set them by hand. Only the sizes 54, 40, 22, 16, 14 and 12 are used.

## Titles

At 40pt a title takes about 35 characters per line across the 11.5 in title box and may run to two lines. Keep it to 55 characters at most and say the point, not the topic. If a title runs over, shorten it; if it cannot be shortened, split the slide. Never shrink the inherited size.

## Colors

- Content slide background: Surface #FFFFFF, with the Brand #1F4FAE strip at the left edge. Title and closing background: Brand #1F4FAE.
- Titles: Brand #1F4FAE. On blue: white, with Tint 3 #C4D9FF for the eyebrow, bar and subline.
- Body text: Text #2C3E50. Secondary text: Text 2 #374151. Captions and footer: Caption #4B5563. Page number: Placeholder #6E7581.
- Footer rule: Rule #E5E7EB, 1pt.
- Chart series in this order (theme Accent 1 to 6): #008300, #E87BA4, #4A3AA7, #1BAF7A, #EDA100, #882255. Single series, Sequential and Diverging are in brand.md.
- No shadows, gradients, glows or reflections anywhere. The theme has none; add none.

## Footer

Content layouts show a footer row 6.88 in from the top, below a 1pt Rule #E5E7EB line at 6.78 in: the deck name at left in the footer placeholder (idx 11), for example "Release 8 summary", and the slide number at right (idx 12, a slide number field). python-pptx does not copy the footer and slide number placeholders onto new slides. Call `add_hf(slide, {"ftr", "sldNum"}, values)` on each content slide. The date shows only on the Title slide.

## Bullets

- Both levels use a round Brand #1F4FAE bullet; level 1 is 22pt Text, level 2 is 16pt Text 2. 6pt after each item. Keep to these two levels (python-pptx levels 0 and 1).
- At most six bullets per slide and about ten words per bullet. No full stop at the end of a bullet.
- For steps or ranked items use `set_lines(frame, lines, numbered=True)`. At most six items. The numbers come in Brand #1F4FAE from the helper; never type the numbers into the text.
- Use the Section layout, not a bullet, for a one line statement.

## Tables

Use the Table layout, or Two Content for a table beside its points, and `add_table`. It sets Neuromine Table with the header row on and banded rows off, pads cells 4.5pt top and bottom and 8pt left and right, sets 16pt text, and aligns each header cell with its column: right over a numeric column, left otherwise. Row height is the box height divided by the row count, at least 0.5 in and at most 0.75 in. Pass `numeric_cols` when a text column could pass for numbers. Header row: Table head #F4F6FB fill, bold Ink #1A2332 text, a 1pt #D8DCE5 rule under it. Body rows: Text 2 #374151, 1pt #DFE3EB rules between rows and round the outside, no vertical rules. One header row, no merged cells, at most 7 rows and 5 columns; split a bigger table. On the Table layout the caption sits above the table: "Table 1. What it shows. Source: the release."

## Charts

Use the Chart layout and `add_chart` with one of COLUMN_CLUSTERED, BAR_CLUSTERED, LINE, PIE or DOUGHNUT (`XL_CHART_TYPE`); other types raise ValueError. `palette` is `"categorical"`, `"single"` (one series in Brand), `"sequential"` (ordered data) or `"diverging"` (above and below a midpoint), from brand.md. Column and bar get a set gap width; lines are 2.25pt with no markers; pie and doughnut color each slice. Labels 16pt in #6A7683, baseline and value gridlines 0.75pt #D9E1EA, legend at the bottom when there is more than one series, no chart title. Put the label in the caption placeholder with the source: "Figure 1. Orders by quarter. Source: Neuromine Release 8 index." One chart per slide. Never paste chart images.

## Images

Use the Picture layout and `add_image`. It scales the picture to fit the box, keeps the aspect ratio and left aligns it; never crop or stretch. PNG or JPG at least 150 dpi at the placed size. No borders, shadows, reflections or rounded frames. Every picture and chart has alt text (the caption text is fine). `add_image` sets it; for a chart, `frame._element.nvGraphicFramePr.cNvPr.set("descr", caption)`.

## Callouts

`add_callout` replaces a content placeholder with the brand.md callout: a Tint 1 #EEF4FF box with a 3pt Brand #1F4FAE bar at its left edge and the text 22pt Text #2C3E50 inside, square corners. By default it sits at the top of the box. On Two Content, when the other side holds a table or chart, pass `align="middle"`. At most one per slide, for the one thing the reader must take away. A callout never sits alone on a slide; see Fill the slide.

## Fill the slide

- A callout never sits alone. Put it on Two Content next to the points, table or chart it comes from, or make the finding the one line of text on a Section slide.
- A Title and Content slide holds at least three bullets, a table, or a callout plus bullets. `check_deck.py` fails a table on Title and Content or Two Content with fewer than 7 rows whose rows are shorter than the box height divided by the row count (at most 0.75 in), within 1pt.
- `check_deck.py` fails a Title and Content or Two Content slide whose content holds fewer than three non-empty paragraphs (a callout counts as one) and no table, chart or picture.

## From brand.md

Applies: palette, typography, chart palettes, logo rules, voice, lists, callouts, dates. Does not apply: document page setup, header and footer, cover geometry, body text density; the sections here replace them.

## Voice

Voice and prose: follow the Writing section in brand.md.

## Checklist before sending

- The file started from `assets/neuromine-template.pptx` (Neuromine Table style present) and its guide slides are gone.
- Every slide uses one of the eight layouts above; the first is Title and the last is Closing.
- Every slide but Title and Closing has the deck name in the footer and a slide number.
- No placeholder run has a direct font, size or color.
- Every table uses Neuromine Table, fits 7 rows and 5 columns, has no merged cells, and right aligns numbers and the header cells over them.
- Fill the slide: no callout alone; every Title and Content or Two Content slide holds three paragraphs, a table, a chart or a picture.
- Content placeholders keep to six bullets and two levels; numbered lists use `set_lines(..., numbered=True)`, none typed into the text.
- Every title is at most 55 characters. Every Chart and Picture slide holds its object with alt text and a caption; every Table slide holds a table and a caption.
- No shadow, gradient, glow or reflection anywhere. File name has no spaces, for example `neuromine-release-8-summary.pptx`.
- Human check: each section starts with a Section slide ("N. Name"), the logo shows only on the blue slides, and nothing sits outside the side margins or below the footer rule.
- `python scripts/check_deck.py file.pptx` passes.
