# Markdown

Dialect: GitHub Flavored Markdown. Tables and column alignment are GFM extensions. Target viewers are GitHub, VS Code preview and Obsidian. Fonts, colors and bullet shapes are the viewer's; here the brand is structure and voice.

Rules for writing a .md file (README, report, notes) in the Neuromine brand. Read https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/brand.md first. Where the two disagree, this file wins for Markdown. https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-sample.md follows every rule below.

## Title line and eyebrow

- Line 1 is the eyebrow: `NEUROMINE` in inline code, uppercase. For a document about another program, Answer ALS or ALS TDI, write that program's name in uppercase instead; the checker expects `NEUROMINE`, so say so when it fails on that line alone.
- Line 3 is the title as the only H1: `# Title`.
- Line 5 is one italic line with the subtitle and date: `*Subtitle - 3 November 2026*`. Dates follow brand.md.
- Lines 2 and 4 are blank. No front matter above line 1.
- Then the intro paragraph.

## Headings

- One H1 per file (the title). Sections are H2, subsections H3. Never use H4 or deeper; split the section instead.
- Do not skip a level (no H3 straight under H1).
- Number H2 sections in reports: `## 2. Results`. Do not number H3. READMEs and notes are not numbered. Number every H2 or none.
- Sentence case. No trailing punctuation. No bold or links inside a heading.
- One blank line before and after every heading.

## Body text

- Short paragraphs, one idea each. One blank line between paragraphs.
- Bold for a key term on first use only; table labels, figure labels and callout labels are bold every time. Italic for titles of works, the subtitle line and a quote source. No italics for emphasis, no underline, no HTML.

## Lists

- Bullets use `-` only (never `*` or `+`).
- Steps that must happen in order use `1.`, `2.`, `3.` so the raw file reads in order too.
- No blank lines between items. Nest at most one level. A nested item indents to the parent's content column: two spaces under `-`, three under `1.`.
- Items are fragments or full sentences, not a mix, within one list. List only the items there are.

## Tables

- Always a header row, then the `| --- |` separator row.
- Text columns left aligned (`---`), number columns right aligned (`---:`). Never center.
- No empty cells: `-` means missing, `n/a` means not applicable.
- First column names the row. Headers are plain words in sentence case, never raw column names. Keep units in the header, not in each cell.
- One line of bold text above the table as its label: `**Table 1. Visits by site**`.

## Figures

A Markdown deliverable carries each chart as a PNG file saved next to the .md file. Never a data URI, never an ASCII chart in a code block, never HTML.

- Make the PNG with matplotlib using the Word recipe, `chart_png` in `scripts/build_word.py`: 6.5 by 3.25 in, 200 dpi, transparent background, Arial, series in the brand chart order (Brand #1F4FAE for a single series), 9pt axis, tick and legend text in #6A7683, grid lines and axis lines in #D9E1EA, no top or right spines. Import it or copy its settings. `scripts/build_markdown_figure.py` builds the sample's figure this way.
- Name the file `<document-name>-figure-N.png` and save it in the same folder as the .md, for example `report-figure-1.png` beside `report.md`.
- One line of bold text above the image as its label, like the table label: `**Figure 1. Visits by month**`, then a blank line, then the image.
- Insert it as `![Figure 1. Visits by month](report-figure-1.png)`. The alt text repeats the label; it is never empty. The path is relative.
- Number figures in order from 1, separately from tables.
- A source line under the figure, as brand.md says: `Source: Release 8 index.`
- Put a table with the same data near the figure, so the content survives when the image does not load.

```text
**Figure 1. Visits by month**

![Figure 1. Visits by month](report-figure-1.png)

Source: Release 8 index.
```

## Callouts

Markdown has no fill, so the callout is a plain blockquote that starts with a bold label. Do not use GitHub `> [!NOTE]` alerts; they add colors and icons the brand does not have.

```text
> **Note:** Counts exclude participants who withdrew consent.
```

Labels: `Note`, `Key point`, `Warning`. One callout per H2 section at most; the intro before the first H2 counts as a section.

## Quotes

A quote is a blockquote with no bold label. After the quote, an empty `>` line, then the source in italics.

```text
> The portal made the release files easy to find.
>
> *Researcher survey, March*
```

## Code blocks

- Fenced with three backticks and always a language tag (`python`, `sql`, `bash`, `text`).
- Never indent code to make a block. Inline code for file names, paths, commands and field names.
- Keep blocks short; link to the full script instead of pasting it.

## Links

- Inline links with descriptive text: `[the build script](scripts/build_word.py)`, never "click here" or a bare URL in prose.
- Relative paths for files in the same repo. Full `https://` URLs for outside pages.
- A document that lives in no repo uses full URLs or no link.
- A `#fragment` link uses GitHub's heading slug: lowercase, spaces to hyphens, punctuation dropped (`### What the table shows` is `#what-the-table-shows`).

## Front matter

None for reports, READMEs and notes. Use YAML front matter only when a tool needs it (for example `SKILL.md`).

## Voice

Voice and prose: follow the Writing section in brand.md.

## Line length and file

- Prose: one paragraph per line, no hard wrap (diffs stay clean on GitHub).
- Code blocks and tables: keep lines under 100 characters where possible.
- UTF-8 with LF line endings and one final newline. ASCII only, as brand.md says.

## Checklist before sending

1. UTF-8, LF endings, exactly one final newline.
2. Line 1 eyebrow, line 3 the only H1, line 5 the italic subtitle and date; lines 2 and 4 blank.
3. No front matter.
4. H2 and H3 only, no skipped levels.
5. H2 numbered `N.` in order from 1, or none numbered.
6. Headings have no trailing punctuation, bold or links.
7. Bullets use `-`, never `*` or `+`.
8. Every table has a bold `**Table N.**` label, a `---` or `---:` separator and no empty cells.
9. Every chart is a PNG beside the file, `<document-name>-figure-N.png`, inserted as `![Figure N. Label](file)` under a bold `**Figure N.**` label numbered from 1, with a table of the same data nearby.
10. Every code fence has a language tag.
11. Callouts use `Note`, `Key point` or `Warning`, no `[!NOTE]` alerts, at most one per section.
12. No HTML.
13. Relative links resolve and `#fragment` links match a heading.
14. Human only: one voice throughout, short paragraphs, descriptive link text, right-aligned number columns, a source line under every figure.
15. `python scripts/check_md.py file.md` passes.
