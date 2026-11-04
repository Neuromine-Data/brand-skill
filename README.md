# brand-skill

The Neuromine brand skill tells an AI agent how to make documents, decks, sheets, pages and figures in the Neuromine brand, for Neuromine, Answer ALS and ALS TDI outputs. It holds the brand rules (palette, type, logo and wordmark, layout, voice), one rules file per format, a ready-made template for each format, a checker for each format and the scripts that build the templates. It holds no data.

## Give this to an agent

Paste this line to the agent:

Fetch https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/SKILL.md and follow it.

The brand rules every format shares are in https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/brand.md.

## Formats

- Word (.docx), memos and reports with a cover band: rules https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/word.md, template https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-template.docx, sample https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-sample.docx, checker https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/check_docx.py
- PowerPoint (.pptx), 16:9 decks in eight layouts: rules https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/powerpoint.md, template https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-template.pptx, sample https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-sample.pptx, checker https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/check_deck.py
- Excel (.xlsx), counts tables and small summaries with native charts: rules https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/excel.md, template https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-template.xlsx, sample https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-sample.xlsx, checker https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/check_xlsx.py
- Markdown (.md), READMEs, reports and notes: rules https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/markdown.md, sample to copy https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-sample.md with its figure https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-sample-figure-1.png, checker https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/check_md.py
- HTML (.html), single-file reports that print to PDF: rules https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/html.md, template and sample in one file https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-report.html, checker https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/check_html.py
- SOP (.docx), standard operating procedures in the house SOP structure: rules https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/sop.md, template https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-sop-template.docx, sample https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-sop-sample.docx, checker https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/check_sop.py

## Check a finished file

Download the checker for the format and https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/prose.py into one folder, then run the checker on the file:

    python check_docx.py visit-review.docx

Each checker prints one PASS or FAIL line per check and exits 1 on any FAIL. Fix every FAIL and run it again. check_html.py also needs https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/neuromine-report.html and https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/design/tokens.json at the same paths as in this repo, so run it from a clone. In a clone, https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/check_all.py runs every checker on the samples in assets/. https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/README.md says what each script checks.

## Rebuild the templates

The Word, PowerPoint, Excel and SOP templates and samples are built from https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/design/tokens.json and the logo, the same bytes on every run on one machine. Change the look in tokens.json or the build script, never in the built file. From a clone:

    python scripts/build_word.py
    python scripts/build_powerpoint.py
    python scripts/build_excel.py
    python scripts/build_markdown_figure.py
    python scripts/build_sop.py

Each writes to assets/, or to the folder given as its only argument. The scripts are https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/build_word.py, https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/build_powerpoint.py, https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/build_excel.py, https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/build_markdown_figure.py and https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/build_sop.py. They need Python 3.10 or later and install python-docx, python-pptx, openpyxl and matplotlib if they are missing.

## Layout

- SKILL.md: the entry page an agent reads first.
- references/: brand.md and one rules file per format.
- assets/: the templates and samples, listed in https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/assets/README.md.
- scripts/: the checkers, the runner, the build scripts and the two helper modules, neuromine_pptx.py and neuromine_excel.py.
- design/: tokens.json (every color, font and size as values for code), DECISIONS.md (where each value came from) and assets/neuromine-logo-light.png, the white logo https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/design/assets/neuromine-logo-light.png.
- NOTES.md: the choices made while building the skill.
- CHANGELOG.md: one dated line per version.

## License

MIT, see https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/LICENSE.
