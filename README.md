# textbook-docx_skill (v3.0.0)

An Agent Skill and Python CLI for turning Markdown chapters into editable, professionally formatted Word textbooks. Pandoc handles Markdown and LaTeX-to-OMML conversion; `python-docx` adds front matter, document metadata, headers/footers, TOC preview, and post-processing.

## Install the Agent Skill

```bash
npx skills add SanchitKRai/textbook-docx_skill --skill textbook-docx_skill --all --global
```

If the installed skills CLI version does not support `--skill`, install the repository and select the `textbook-docx_skill/` skill directory in the installer.

## Python setup

```bash
python3 -m venv .venv
. .venv/bin/activate
python3 -m pip install -r requirements.txt
```

Pandoc is supplied by `pypandoc_binary` if no system Pandoc exists. For rendered verification install LibreOffice Writer/Math and `poppler-utils` (`pdftotext`, `pdfinfo`).

## Build

```bash
python3 scripts/build_textbook.py examples/01_cell_structure.md examples/02_cell_division.md \
  --output cell-biology.docx --title "Cell Biology" --subtitle "BSc Life Sciences" \
  --author "Author Name" --institution "University of Delhi" --date "2026"
```

CLI compatibility is maintained: positional chapter inputs, `-o/--output`, `--title`, `--subtitle`, `--author`, `--institution`, `--date`, `--page-size A4|LETTER`, `--header-text`, `--footer-text`, `--no-toc`, `--overwrite`, and `--version`.

## Supported Markdown and math

| Feature | Example | Notes |
|---|---|---|
| Headings | `# Chapter`, `## Section`, `### Subsection` | Levels 1–3 appear in contents preview. |
| Wrapped paragraphs | soft line breaks within a paragraph | Pandoc joins them according to CommonMark. |
| Emphasis | `**bold**`, `*italic*`, `` `code` `` | Inline formatting. |
| Links | `[text](https://example.org)` | Real Word hyperlinks. Bare URLs are autolinked by Pandoc. |
| Currency | `Costs $5 and $10` | Pandoc math delimiter rule avoids treating these prices as math. |
| Equations | `$E=mc^2$`, `$$\\sum_{i=1}^n i$$` | Pandoc emits native Word OMML. |
| Advanced math | `\\int`, `\\sqrt[3]{x}`, `\\begin{pmatrix}...` | Pandoc-supported LaTeX math subset; no custom packages/macros. |
| Lists | nested bullets and numbered lists | Pandoc preserves nesting/continuity. |
| Tables | GFM pipe tables with `:---`, `---:`, `:---:` | Separator cells use 3+ dashes; alignment preserved. |
| Callouts | `> **Key idea:** ...` | Blockquote content is preserved as a coherent block. |
| Code | fenced blocks | Shaded/monospaced Word code style. |
| Figures | `![Alt text]("figures/image file.png" "Caption")` | Local images; alt text retained; Figure N.M restarts for each chapter. |
| Horizontal rule | `---` | Word horizontal rule. |
| References | `## References` followed by paragraphs or lists | Plain APA paragraphs, bullets and numbered refs supported with hanging indent. |

**Known limits:** arbitrary TeX packages (TikZ/PSTricks), custom TeX macro definitions, remote image fetching, and complex page-layout automation are outside the supported feature set. PDF renderers can differ slightly from Microsoft Word.

## Contents page

A static heading preview is populated at build time. A genuine Word TOC field is also embedded for dynamic page numbers. In Word, click the TOC and choose **Update Table** or press `Ctrl+A`, then `F9`; select **Update entire table**. Verify page numbers in Word before formal submission.

## Validation

```bash
python3 -m pytest -q
soffice --headless --convert-to pdf --outdir build examples/sample_output.docx
pdftotext build/sample_output.pdf -
pdfinfo build/sample_output.pdf
```

The CI workflow executes the stress-book render, text leakage checks, tests, and OOXML structural validation. Microsoft Word visual validation remains a manual check.

## License

MIT; see [LICENSE](LICENSE).
