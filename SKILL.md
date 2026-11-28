---
name: textbook-docx_skill
description: Converts one or more Markdown chapters into an editable, formatted Word textbook with a cover, populated contents preview, equations, figures, tables, lists, code blocks, callouts, and references. Use when creating or updating academic textbooks, university notes, course readers, revision guides, or structured study documents from Markdown.
license: MIT
metadata:
  author: SanchitKRai
  version: "3.0.0"
---

# textbook-docx_skill

## Runtime and dependency pre-check

Use `python3` for every command. Check for Pandoc before building:

```bash
python3 -c "import docx; print('python-docx OK')"
command -v pandoc || python3 -c "import pypandoc; print(pypandoc.get_pandoc_path())"
```

Install missing dependencies only in a virtual environment or with `--user`; do not modify the system Python:

```bash
python3 -m venv .venv
. .venv/bin/activate
python3 -m pip install -r requirements.txt
```

The `pypandoc_binary` wheel supplies a Pandoc executable when a system Pandoc is unavailable. LibreOffice (`soffice`) and `pdftotext` are required for mandatory rendered-output verification, but are system packages rather than Python dependencies.

## Supported Markdown

Markdown is parsed by Pandoc (not a custom line parser). Supported syntax includes:

| Syntax | Behavior / limit |
|---|---|
| ATX headings `#` through `######` | Converted to Word heading styles; TOC preview lists levels 1–3. |
| Soft-wrapped paragraphs | CommonMark/Pandoc paragraph joining. A blank line starts a new paragraph. |
| `**bold**`, `*italic*`, `` `code` `` | Inline formatting. |
| `[label](URL)` and bare URLs | Editable hyperlinks. Relative local targets must exist. |
| `$...$`, `$$...$$` | Pandoc TeX math converted to native Word OMML. Complex LaTeX support follows Pandoc's math reader; unsupported macros are reported by Pandoc and must not be silently accepted. |
| Fenced code blocks | Monospaced, shaded Pandoc code style; long blocks may flow across pages. |
| Nested bullet/ordered lists | Pandoc list nesting and numbering. |
| GFM pipe tables | Separator cells require at least three dashes; alignment markers are honored by Pandoc. Header rows repeat in Word where supported. |
| `> blockquotes` | Pandoc blockquotes; consecutive quote lines belong to one block. |
| `![alt](path "title")` | Local images. Image alt text is preserved; caption numbering restarts per chapter (Figure N.M). Paths with spaces should be angle-bracketed or quoted according to Markdown syntax. |
| `---` | Horizontal rule. |
| `## References` | Paragraphs, bullets, and numbered references are preserved. Reference paragraphs receive a hanging indent. |

Math accepted by Pandoc includes fractions, roots (including indexed roots), superscripts/subscripts, sums, integrals, standard functions, text commands, vector/hat/overline accents, delimiters, matrices and cases supported by Pandoc's LaTeX math reader. Arbitrary TeX packages, custom macro definitions, TikZ/PSTricks, and external TeX package features are not supported. If Pandoc reports a conversion warning/error, resolve it or preserve the source expression visibly; do not claim a successful mathematical conversion without rendered inspection.

## CLI

```bash
python3 scripts/build_textbook.py chapter1.md chapter2.md \
  --output textbook.docx --title "Book Title" --subtitle "Course / edition" \
  --author "Author" --institution "Institution" --date "2026"
```

The CLI preserves the established flags: positional input chapters; `-o/--output`; `--title`, `--subtitle`, `--author`, `--institution`, `--date`; `--page-size A4|LETTER`; `--header-text`, `--footer-text`; `--no-toc`; `--overwrite`; `--version`. Never pass `--overwrite` unless replacement is explicitly intended.

## Mandatory build verification

After generation:

1. Check the process exit code and output file.
2. Convert the DOCX to PDF using LibreOffice: `soffice --headless --convert-to pdf --outdir <verify-dir> <file.docx>`.
3. Extract text using `pdftotext <file.pdf> -` and check for missing content, raw LaTeX (`\\sum`, `\\int`, `\\frac`), raw Markdown syntax, currency corruption, duplicate/incorrect figure labels, and blank pages. Check page count with `pdfinfo` when available.
4. Run the repository test suite: `python3 -m pytest -q`.
5. Run the OOXML schema validator if available in the environment. The project validator checks repository-authored XML ordering/duplicates; record any externally sourced template errors separately.
6. Do not claim Microsoft Word-specific behavior is verified unless tested in Microsoft Word. Use the manual checklist below.

## TOC and Word manual check

A populated, static heading preview is included so the contents page is useful immediately. A genuine Word TOC field is also present for updating page numbers. In Word, click within the TOC and choose **Update Table** (or press `Ctrl+A`, then `F9`) and select **Update entire table**. Check heading levels, hyperlinks, and page numbers after update. LibreOffice rendering is not a substitute for a real Microsoft Word check.

## Agent workflow

1. Inspect the supplied Markdown and local assets. Preserve supplied scientific terminology and citations; do not invent sources.
2. Run the dependency pre-check. If needed, install into `.venv` or with `--user`, never into system Python.
3. Build with the CLI, preserving source chapter order.
4. Run all mandatory rendered-output and test checks above.
5. Return the editable DOCX, optionally the verification PDF, and state any unverified Word-specific behavior or unsupported syntax.
