# brand-skill

The Neuromine brand skill tells an AI agent how to make documents, decks, sheets, pages and figures in the Neuromine brand, for Neuromine, Answer ALS and ALS TDI outputs. It holds the brand rules, one rules file, template and checker per format, and the scripts that build the templates. It holds no data.

## Give this to an agent

Paste this line to the agent:

Fetch https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/SKILL.md and follow it.

The brand rules every format shares are in https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/brand.md.

## Formats

- Word (.docx): memos and reports with a cover band, rules https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/word.md
- PowerPoint (.pptx): 16:9 decks in eight layouts, rules https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/powerpoint.md
- Excel (.xlsx): counts tables and small summaries with native charts, rules https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/excel.md
- Markdown (.md): READMEs, reports and notes, rules https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/markdown.md
- HTML (.html): single-file reports that print to PDF, rules https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/html.md
- SOP (.docx): standard operating procedures in the house SOP structure, rules https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/references/sop.md

## Maintainers

Run the checker for the format on the finished file; https://raw.githubusercontent.com/Neuromine-Data/brand-skill/main/scripts/README.md says which checker to use and how. The templates and samples are built from design/tokens.json by the build scripts in scripts/, so change the look there and never edit a built file by hand.

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
