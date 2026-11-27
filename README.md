# textbook-docx (v2.4)

Open-source Agent Skill and Python CLI for converting Markdown chapters into polished, editable Word textbooks.

## One-line install

```bash
npx skills add SanchitKRai/textbook-docx_skill --all --global
```

## Install Python dependencies

```bash
python -m pip install -r requirements.txt
```

## Basic usage

```bash
python scripts/build_textbook.py chapter.md --output textbook.docx
```

## Full textbook from multiple chapters

```bash
python scripts/build_textbook.py 01-introduction.md 02-cell-structure.md 03-cell-cycle.md --output cell-biology.docx --title "Cell Biology" --subtitle "BSc Life Sciences" --author "Sanchit" --institution "University of Delhi" --date "2026"
```

Use `--help` to see all options. Existing outputs are protected; pass `--overwrite` only when you intend to replace one.

## What's new in v2.4

- **Native Word Mathematical Equations:** Completely replaces fragile inline transparent PNGs with native Word typography:
  - **Simple inline equations** (`$E=mc^2$`): rendered using native Word text runs with math variable italics, upright operators, and native superscript (`w:vertAlign="superscript"`) / subscript (`w:vertAlign="subscript"`). Fully inline with surrounding text; never invisible.
  - **Complex inline equations** (`$\frac{dN}{dt}=rN$`): converted to native Word Office Math Markup Language (OMML `<m:oMath>`) inside the paragraph run sequence, giving stacked fractions and crisp typography.
  - **Display equations** (`$$...$$`): rendered as centered standalone native OMML equation paragraphs (`<m:oMathPara>`).
  - **Opaque fallback:** PNG rasterization with opaque white background if native parsing fails.
- **Table of Contents (Levels 1–3):** Standardized heading mapping (`#` H1, `##` H2, `###` H3) and inserted Word dynamic TOC field (`TOC \o "1-3" \h \z \u`) spanning discrete OpenXML runs with `w:updateFields="true"` in settings.
- **Code Block Pagination:** Code blocks ≤25 lines keep lines together (`w:keepLines`), while longer blocks split cleanly across page boundaries.
- **Decoupled Reference Numbering:** Dedicated numbering instance with `startOverride="1"` guarantees bibliography starts at 1 independently of preceding lists.

## Features

- Title page and document metadata
- Word TOC field with heading levels 1–3 (auto-prompts on open; or right-click > Update Field)
- Multiple chapters, combined in supplied order
- A4 and Letter page sizes; custom running header/footer text and page numbers
- Markdown headings, emphasis, inline code, lists, simple tables, callout boxes, and fenced code
- Local images with chapter-based figure numbering
- Native Word equations (OMML and formatted text runs)
- Numbered bibliography from a terminal References section restarting at 1
- DOCX reopen validation

## How to update Table of Contents in Word

The Table of Contents is generated using Word's dynamic OpenXML field code. When opening the generated `.docx` file in Microsoft Word:
1. When prompted with *"This document contains fields that may refer to other files. Do you want to update the fields in this document?"*, click **Yes**.
2. Alternatively, right-click anywhere on the Table of Contents and select **Update Field** (or select the field and press **F9**).

## Known limitations

- TOC page numbers are calculated by Word's layout engine upon opening/updating. Standalone Python generators cannot precalculate exact pagination without running Word.
- Complex arbitrary LaTeX macro packages (TikZ, PSTricks) are not supported.
- Remote images, automatic diagram generation, automatic research/fact-checking, and complex arbitrary page-layout automation are not included.

## License

MIT. See LICENSE.
