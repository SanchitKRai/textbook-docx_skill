---
name: textbook-docx
description: Creates complete, professionally formatted academic textbooks and study documents as editable Microsoft Word DOCX files from one or more Markdown chapters. Use for textbooks, university notes, revision guides, reports, and whitepapers with title pages, tables of contents, figures, equations, tables, and references.
license: MIT
metadata:
  author: SanchitKRai
  version: "2.4.0"
---

# textbook-docx

Use the bundled Python CLI to turn prepared Markdown chapter files into a formatted Word textbook. Prefer the CLI over writing a custom DOCX generator.

## Requirements and setup

- Python 3.9+
- Install dependencies from the skill directory: `python -m pip install -r requirements.txt`
- The generator is `scripts/build_textbook.py`, relative to this SKILL.md.
- Ask before installing packages or changing the user's environment.

## Supported features

- Multiple Markdown chapters combined in the order supplied, with page breaks between chapters.
- Title page with title, subtitle, author, institution, and date.
- Genuine Word Table of Contents field supporting heading levels 1–3 (`TOC \o "1-3" \h \z \u`) spanning discrete OpenXML runs (`begin`, `instrText`, `separate`, placeholder text, and `end`). Word prompts to update fields upon opening via `<w:updateFields w:val="true"/>`. If not updated automatically, right-click the TOC area and choose **Update Field** (or press **F9**).
- A4 or US Letter page size; customizable running header/footer text and page numbers.
- Markdown headings (Levels 1–3: `#` Chapter, `##` Section, `###` Subsection), paragraphs, bold, italic, inline code, lists, simple pipe tables, and blockquote callouts.
- Fenced code blocks with automatic line-keeping pagination (`w:keepLines`) for short blocks (≤25 lines) to avoid splitting across pages, while long blocks split cleanly across pages.
- Local images using `![Caption](relative/path.png)`, with chapter-based figure numbering.
- **Native Word Mathematical Equations (v2.4):**
  - **Simple Inline Equations (`$E=mc^2$`):** Rendered as native Word text runs with `Cambria Math` typography, mathematical variable italics, upright operators, and native superscript (`w:vertAlign="superscript"`) / subscript (`w:vertAlign="subscript"`). Fully inline with surrounding text; never invisible.
  - **Complex Inline Equations (`$\frac{dN}{dt}=rN$`):** Rendered using native Word Office Math Markup Language (OMML `<m:oMath>`) embedded directly inside the paragraph run sequence, providing stacked vertical fractions and radical symbols.
  - **Display Equations (`$$...$$`):** Rendered as centered standalone OMML equation paragraphs (`<m:oMathPara>`).
  - **Fallback:** High-DPI opaque PNG rasterization if native rendering is unavailable, with graceful text fallback.
- Numbered bibliography from a terminal `## References` section with dedicated numbering sequence that always restarts at 1 independently of preceding lists.
- Output validation by reopening the generated DOCX.

## Workflow

1. Identify the subject, academic level, intended audience, chapter scope, and output path. Ask only for essential missing details.
2. Prepare accurate Markdown files. Keep each chapter in its own .md file when producing a multi-chapter book. Save local figure assets beside the relevant Markdown file or in a relative subfolder.
3. Add a terminal References section to a chapter when a bibliography is needed. Supply real, verified source details; never invent references.
4. Check that Python and dependencies are available. If dependencies are missing, ask permission before installing.
5. Run the CLI using the exact supported interface:

   ```bash
   python "<skill-dir>/scripts/build_textbook.py" chapter1.md chapter2.md --output textbook.docx --title "Book Title" --author "Author" --institution "Institution" --subtitle "Course or edition" --date "2026"
   ```

6. Optional flags: `--page-size A4|LETTER`, `--header-text TEXT`, `--footer-text TEXT`, `--no-toc`, and `--overwrite`. Never use `--overwrite` unless the user explicitly authorizes replacement.
7. Check the exit code and success message. The CLI verifies the saved DOCX is non-empty and can be reopened.
8. Return the DOCX and a concise contents summary. Keep the Markdown source files when the user may want future edits.

## Markdown conventions

- Use `#` for chapter titles (Heading 1), `##` for major sections (Heading 2), and `###` for subsections (Heading 3).
- Use standard Markdown bullet and numbered lists.
- Use pipe tables with a separator row.
- Use blockquotes for callout boxes, e.g. `> **Key concept:** Explanation.`
- Use local images, e.g. `![Mitochondrion](figures/mitochondrion.png)`.
- Use `$E=mc^2$` for inline math and a pair of `$$` lines for display math.
- Put references under a terminal `## References` heading as a bullet or numbered list.

## Quality and limitations

- Write at the requested academic level, define specialized terms, and preserve user-provided terminology.
- Distinguish established findings from hypotheses and uncertainty.
- Do not fabricate facts, experimental data, citations, or sources. Respect copyright.
- **TOC Updating:** Dynamic Word TOC field codes require Word's layout engine to compute page numbers. The document enables `w:updateFields="true"` so Microsoft Word prompts to calculate the TOC on open. Users can also press **F9** or right-click the TOC and select **Update Field**.
- **Equation Support:** Standard mathematical syntax (fractions, powers, indices, radicals, Greek letters, operators) is rendered natively into Word OMML and text runs without rasterization. Arbitrary LaTeX packages (TikZ, PSTricks) are not supported.
- The tool does not browse for sources, automatically create diagrams, fetch remote images, or independently fact-check content.
- Wide tables and complex page layouts may need manual refinement.
- Do not claim unsupported features or successful generation unless the CLI actually completes.
