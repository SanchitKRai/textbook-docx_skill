from __future__ import annotations
import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "build_textbook.py"
VALIDATOR = ROOT / "scripts" / "validate_ooxml.py"
EX1 = ROOT / "examples" / "01_cell_structure.md"
EX2 = ROOT / "examples" / "02_cell_division.md"

def build(tmp_path, chapters=(EX1, EX2), name="stress.docx", extra=()):
    out = tmp_path / name
    cmd = [sys.executable, str(SCRIPT), *map(str, chapters), "--output", str(out),
           "--title", "Cell Biology Regression Book", "--author", "Test Author",
           "--institution", "University Test", *extra]
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.returncode == 0, f"CLI failed: {result.stderr}\n{result.stdout}"
    assert out.is_file() and out.stat().st_size > 0
    return out

def render(docx_path: Path, out_dir: Path):
    soffice = shutil.which("soffice")
    pdftotext = shutil.which("pdftotext")
    assert soffice, "LibreOffice (soffice) is required; tests must not silently skip"
    assert pdftotext, "pdftotext is required; tests must not silently skip"
    out_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)], capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    pdf = out_dir / (docx_path.stem + ".pdf")
    assert pdf.is_file() and pdf.stat().st_size > 0
    text = subprocess.run([pdftotext, "-layout", str(pdf), "-"], capture_output=True, text=True, check=True).stdout
    pages = text.count("\f")
    assert pages > 1
    # No empty pages: every form-feed-delimited page must contain visible text.
    page_texts = [p.strip() for p in text.split("\f") if p.strip()]
    assert len(page_texts) == pages if text.endswith("\f") else len(page_texts) == pages + 1
    return pdf, text


def test_01_pandoc_math_sum_integral_functions_matrices_and_indexed_roots(tmp_path):
    docx = build(tmp_path)
    doc = Document(docx)
    xml = doc._element.xml
    assert "m:oMath" in xml
    assert "m:f" in xml or "m:rad" in xml
    assert doc._element.xpath(".//m:count")
    for mc_pr in doc._element.xpath(".//m:mcPr"):
        count = mc_pr.find(qn("m:count"))
        mc_jc = mc_pr.find(qn("m:mcJc"))
        if count is not None:
            assert mc_jc is not None and mc_pr.index(count) < mc_pr.index(mc_jc)
    assert all(not (run_properties.find(qn("m:nor")) is not None and run_properties.find(qn("m:sty")) is not None)
               for run_properties in doc._element.xpath(".//m:rPr"))
    pdf, text = render(docx, tmp_path / "render")
    for raw in (r"\sum", r"\int", r"\frac", r"\sqrt", r"\begin{pmatrix}"):
        assert raw not in text
    assert "Cell Biology Regression Book" in text


def test_02_currency_dollars_not_math(tmp_path):
    docx = build(tmp_path, chapters=(EX1,))
    pdf, text = render(docx, tmp_path / "currency")
    assert "costs $5 and the staining kit costs $10" in text
    doc = Document(docx)
    assert any("costs $5" in p.text and "$10" in p.text for p in doc.paragraphs)


def test_03_plain_paragraph_bullet_and_numbered_references_preserved(tmp_path):
    md = tmp_path / "refs.md"
    md.write_text("# References test\n\n## References\n\nSmith, J. (2020). *A title*. Publisher. https://example.org/a\n\n- Jones, K. (2021). Another title.\n\n1. Lee, M. (2022). Numbered reference.\n", encoding="utf-8")
    docx = build(tmp_path, (md,), "refs.docx")
    pdf, text = render(docx, tmp_path / "refs-render")
    for phrase in ("Smith, J. (2020)", "Jones, K. (2021)", "Lee, M. (2022)"):
        assert phrase in text
    doc = Document(docx)
    refs_idx = next(i for i,p in enumerate(doc.paragraphs) if p.text.strip() == "References")
    assert doc.paragraphs[refs_idx+1].paragraph_format.left_indent is not None
    assert doc.paragraphs[refs_idx+1].paragraph_format.first_line_indent is not None


def test_04_commonmark_soft_wrapped_lines_joined(tmp_path):
    docx = build(tmp_path, (EX1,))
    doc = Document(docx)
    matches = [p for p in doc.paragraphs if "hard-wrapped across several source lines" in p.text]
    assert len(matches) == 1
    assert "continuous paragraph" in matches[0].text


def test_05_figure_numbers_restart_per_chapter_paths_with_spaces_and_alt_text(tmp_path):
    docx = build(tmp_path)
    pdf, text = render(docx, tmp_path / "figures")
    assert "Figure 1.1: Chapter 1 cell diagram" in text
    assert "Figure 2.1: Chapter 2 dividing cell" in text
    assert "Figure 2.2: Chapter 2 cell division stages" in text
    with zipfile.ZipFile(docx) as zf:
        xml = zf.read("word/document.xml").decode("utf-8")
        media = [x for x in zf.namelist() if x.startswith("word/media/")]
    assert 'descr="Cell diagram"' in xml
    assert 'descr="Dividing cell"' in xml
    assert len(media) >= 2


def test_06_gfm_short_separators_alignment_and_repeating_header(tmp_path):
    md = tmp_path / "tables.md"
    md.write_text("# Tables\n\n| Left | Center | Right |\n|:--|:-:|-:|\n| alpha | beta | gamma |\n\nAfter table.\n", encoding="utf-8")
    docx = build(tmp_path, (md,), "tables.docx")
    doc = Document(docx)
    tables = doc.tables
    assert tables
    tbl_pr = tables[0]._tbl.tblPr
    assert tbl_pr.index(tbl_pr.find(qn("w:tblLayout"))) < tbl_pr.index(tbl_pr.find(qn("w:tblLook")))
    first_row_xml = tables[0].rows[0]._tr.xml
    assert "tblHeader" in first_row_xml
    headers = tables[0].rows[0]._tr.xpath("./w:trPr/w:tblHeader")
    assert len(headers) == 1 and headers[0].get(qn("w:val")) == "on"
    alignments = [cell.paragraphs[0].alignment for cell in tables[0].rows[0].cells]
    assert alignments[0] == WD_ALIGN_PARAGRAPH.LEFT
    assert alignments[1] == WD_ALIGN_PARAGRAPH.CENTER
    assert alignments[2] == WD_ALIGN_PARAGRAPH.RIGHT
    pdf, text = render(docx, tmp_path / "table-render")
    assert all(x in text for x in ("alpha", "beta", "gamma", "After table."))


def test_07_nested_lists_and_numbered_sequence_continues(tmp_path):
    docx = build(tmp_path)
    pdf, text = render(docx, tmp_path / "lists")
    assert "Selectively permeable boundary" in text and "Regulates transport" in text
    assert "1.   Prepare the sample." in text or "1. Prepare the sample." in text
    assert "2.   Observe using a microscope." in text or "2. Observe using a microscope." in text
    assert "3.   Record the results." in text or "3. Record the results." in text
    doc = Document(docx)
    ilvls = [int(x) for x in doc._element.xpath(".//w:numPr/w:ilvl/@w:val")]
    assert ilvls and max(ilvls) >= 1


def test_08_inline_links_math_inside_bold_headings_and_horizontal_rule(tmp_path):
    docx = build(tmp_path)
    pdf, text = render(docx, tmp_path / "inline")
    assert "OpenStax Biology" in text
    assert "https://www.ncbi.nlm.nih.gov/books/" in text
    assert "the energy relation is" in text
    assert "CHAPTERBREAKTOKEN" not in text
    with zipfile.ZipFile(docx) as zf:
        rels = zf.read("word/_rels/document.xml.rels").decode("utf-8")
        document_xml = zf.read("word/document.xml").decode("utf-8")
    math_xml = re.findall(r"<m:oMath>(.*?)</m:oMath>", document_xml, re.S)
    assert math_xml and "".join(re.findall(r"<m:t>(.*?)</m:t>", math_xml[0])) == "E=mc2"
    assert "https://openstax.org/details/books/biology-2e" in rels
    assert "https://www.ncbi.nlm.nih.gov/books/" in rels
    assert "<w:pict>" in document_xml and "o:hr=\"t\"" in document_xml, "Markdown --- must become a Word horizontal-rule drawing"


def test_09_callout_body_not_bold_blue_and_spaces_after_blocks(tmp_path):
    docx = build(tmp_path, (EX1,))
    doc = Document(docx)
    callouts = [t for t in doc.tables if t._tbl.tblPr.find(qn("w:tblCaption")) is not None
                and t._tbl.tblPr.find(qn("w:tblCaption")).get(qn("w:val")) == "Callout"]
    assert callouts, "Expected blockquote content grouped in a callout table"
    tbl_pr = callouts[0]._tbl.tblPr
    assert tbl_pr.index(tbl_pr.find(qn("w:tblLayout"))) < tbl_pr.index(tbl_pr.find(qn("w:tblLook")))
    cell = callouts[0].cell(0, 0)
    paragraphs = cell.paragraphs
    assert len(paragraphs) == 1, "Consecutive blockquote lines should remain one coherent paragraph"
    runs = [r for r in paragraphs[0].runs if r.text.strip()]
    assert runs and "Key concept:" in "".join(r.text for r in runs)
    body_runs = [r for r in runs if "plasma membrane is selectively" in r.text or "It regulates" in r.text]
    assert body_runs and all(r.bold is not True for r in body_runs)
    assert "w:shd" in cell._tc.xml and "w:tcBorders" in cell._tc.xml
    assert paragraphs[0].paragraph_format.space_after.pt >= 0


def test_10_ooxml_structure_audit(tmp_path):
    docx = build(tmp_path)
    result = subprocess.run([sys.executable, str(VALIDATOR), str(docx)], capture_output=True, text=True)
    assert result.returncode == 0, result.stderr + result.stdout
    assert "audit passed" in result.stdout.lower()


def test_11_document_metadata_is_clean(tmp_path):
    md = tmp_path / "simple.md"; md.write_text("# Simple\n\nText.\n", encoding="utf-8")
    docx = build(tmp_path, (md,), "metadata.docx")
    doc = Document(docx)
    props = doc.core_properties
    assert props.author == "Test Author"
    assert props.comments == ""
    assert props.created.year >= 2026
    assert props.modified.year >= 2026
    assert props.last_modified_by == "Test Author"
    out2 = tmp_path / "noauthor.docx"
    result = subprocess.run([sys.executable, str(SCRIPT), str(md), "-o", str(out2), "--no-toc"], capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    assert Document(out2).core_properties.author == ""


def test_12_toc_populated_heading_preview_not_heading_1(tmp_path):
    docx = build(tmp_path)
    doc = Document(docx)
    paragraphs = doc.paragraphs
    toc_idx = next(i for i,p in enumerate(paragraphs) if p.text == "Table of Contents")
    assert paragraphs[toc_idx].style.name == "Title"
    preview = "\n".join(p.text for p in paragraphs[toc_idx+1:])
    assert "Chapter 1: Cell Structure" in preview
    assert "Cost and resource example" in preview
    settings = doc.settings.element.xml
    assert "w:updateFields" in settings
    assert 'TOC \\o "1-3"' in doc._element.xml


def test_13_numbering_uses_pandoc_definitions_not_hardcoded_abstract_id():
    source = SCRIPT.read_text(encoding="utf-8")
    assert "abstract_num_id=7" not in source
    assert "abstract_num_id = 7" not in source


def test_fresh_output_version_and_page_size_cli(tmp_path):
    v = subprocess.run([sys.executable, str(SCRIPT), "--version"], capture_output=True, text=True)
    assert v.returncode == 0 and "3.0.0" in v.stdout
    md = tmp_path / "letter.md"; md.write_text("# Letter\n\nText.\n", encoding="utf-8")
    out = tmp_path / "letter.docx"
    result = subprocess.run([sys.executable, str(SCRIPT), str(md), "-o", str(out), "--page-size", "LETTER", "--no-toc"], capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    assert Document(out).sections[0].page_width.inches == 8.5


def test_14_unsupported_math_is_visible_and_warned(tmp_path):
    md = tmp_path / "unsupported.md"
    md.write_text("# Unsupported math\n\nExpression: $\\myUnknownMacro{x}$.\n", encoding="utf-8")
    out = tmp_path / "unsupported.docx"
    result = subprocess.run([sys.executable, str(SCRIPT), str(md), "-o", str(out)], capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    assert "unsupported math command" in result.stderr.lower()
    assert str(md) in result.stderr and r"\myUnknownMacro{x}" in result.stderr
    pdf, text = render(out, tmp_path / "unsupported-render")
    assert "Unconverted math source" in text
    assert r"\myUnknownMacro{x}" in text


def test_cli_help_keeps_all_legacy_flags():
    result = subprocess.run([sys.executable, str(SCRIPT), "--help"], capture_output=True, text=True)
    assert result.returncode == 0
    for flag in ("--output", "--title", "--subtitle", "--author", "--institution", "--date", "--page-size", "--header-text", "--footer-text", "--no-toc", "--overwrite", "--version"):
        assert flag in result.stdout, f"legacy CLI flag removed: {flag}"


def test_stress_render_has_no_raw_syntax_or_blank_pages(tmp_path):
    docx = build(tmp_path)
    pdf, text = render(docx, tmp_path / "strict-render")
    for token in (r"\sum", r"\int", r"\frac", r"\sqrt", "CHAPTERBREAKTOKEN", "!["):
        assert token not in text, f"raw source leaked into PDF: {token}"
    prose = text.replace("2 ** number_of_doublings", "POWER_OPERATOR_IN_CODE")
    assert "**Cost" not in prose and "**Chapter" not in prose and "__" not in prose
    for token in ("costs $5", "costs $10", "Figure 1.1:", "Figure 2.1:", "Figure 2.2:", "Alberts, B.", "Cooper, G.", "Lodish, H."):
        assert token in text, f"rendered content missing: {token}"
