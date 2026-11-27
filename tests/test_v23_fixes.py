import re
import subprocess
import sys
from pathlib import Path
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "build_textbook.py"

def test_version_cli():
    """Verify version 2.3.0 is reported by CLI."""
    result = subprocess.run(
        [sys.executable, str(SCRIPT), "--version"],
        capture_output=True, text=True
    )
    assert result.returncode == 0
    assert "textbook-docx 2.4.0" in (result.stdout + result.stderr)

def test_toc_field_structure_and_headings(tmp_path):
    """Test genuine Word TOC field structure with heading levels 1-3."""
    md_content = """# Chapter One

## Section One

### Subsection One

Paragraph content.
"""
    chapter = tmp_path / "chapter_toc.md"
    chapter.write_text(md_content, encoding="utf-8")
    output = tmp_path / "toc_test.docx"

    result = subprocess.run(
        [sys.executable, str(SCRIPT), str(chapter), "--output", str(output), "--title", "TOC Test"],
        capture_output=True, text=True
    )
    assert result.returncode == 0, result.stderr
    assert output.is_file()

    doc = docx.Document(str(output))

    # Verify TOC field structure in XML
    doc_xml = doc._element.xml
    assert 'TOC \\o "1-3" \\h \\z \\u' in doc_xml
    assert 'w:fldCharType="begin"' in doc_xml
    assert 'w:fldCharType="separate"' in doc_xml
    assert 'w:fldCharType="end"' in doc_xml
    assert 'Table of Contents (Right-click' in doc_xml

    # Verify updateFields in settings
    settings_xml = doc.settings.element.xml
    assert 'w:updateFields' in settings_xml

    # Verify Headings 1, 2, and 3 are present
    headings = [(p.style.name, p.text) for p in doc.paragraphs if p.style.name.startswith("Heading")]
    h_dict = {h[1]: h[0] for h in headings}
    assert h_dict.get("Chapter One") == "Heading 1"
    assert h_dict.get("Section One") == "Heading 2"
    assert h_dict.get("Subsection One") == "Heading 3"

def test_inline_and_display_equations_regression(tmp_path):
    """Regression test: Inline math stays inside surrounding text; display equation is centered."""
    md_content = """# Math Chapter

The mass-energy relation is $E=mc^2$ in modern physics.

- Bullet with inline math $p=mv$ momentum.

Here is an invalid expression: $\\unsupportedmacro{test}$ gracefully handled.

$$
\\frac{dN}{dt}=rN
$$
"""
    chapter = tmp_path / "math_chapter.md"
    chapter.write_text(md_content, encoding="utf-8")
    output = tmp_path / "math_test.docx"

    result = subprocess.run(
        [sys.executable, str(SCRIPT), str(chapter), "--output", str(output)],
        capture_output=True, text=True
    )
    assert result.returncode == 0, result.stderr

    doc = docx.Document(str(output))

    # 1. Verify inline equation in paragraph
    matching = [p for p in doc.paragraphs if "mass-energy relation is" in p.text]
    assert len(matching) == 1, "Inline equation must be inside its surrounding paragraph"
    p_inline = matching[0]

    assert p_inline.alignment != WD_ALIGN_PARAGRAPH.CENTER
    # Check that text before and after exist in the same paragraph
    assert "mass-energy relation is" in p_inline.text
    assert "in modern physics." in p_inline.text

    # Verify native equation formatting exists in paragraph
    has_math = any(r.font.superscript is True and "2" in r.text for r in p_inline.runs) or "m:oMath" in p_inline._p.xml
    assert has_math, "Paragraph must contain native equation formatting or superscript"

    # 2. Verify inline equation in bullet list
    bullet_matching = [p for p in doc.paragraphs if "Bullet with inline math" in p.text]
    assert len(bullet_matching) == 1
    p_bullet = bullet_matching[0]
    assert p_bullet.style.name == "List Bullet"
    assert "momentum." in p_bullet.text
    bullet_has_math = any(r.font.name == "Cambria Math" for r in p_bullet.runs) or "m:oMath" in p_bullet._p.xml
    assert bullet_has_math, "Bullet item must contain native math formatting"

    # 3. Verify unrenderable fallback in paragraph
    fb_matching = [p for p in doc.paragraphs if "Here is an invalid expression:" in p.text]
    assert len(fb_matching) == 1
    p_fb = fb_matching[0]
    assert "\\unsupportedmacro{test}" in p_fb.text
    assert "gracefully handled." in p_fb.text
    assert any("unsupportedmacro" in r.text for r in p_fb.runs)
    assert any(r.font.name == "Cambria Math" for r in p_fb.runs)

    # 4. Verify display equation is preserved as a centered paragraph
    display_paras = [p for p in doc.paragraphs if "m:oMathPara" in p._p.xml or (p.alignment == WD_ALIGN_PARAGRAPH.CENTER and len(p.runs) > 0)]
    assert len(display_paras) >= 1, "Display equation must be a centered paragraph"

def test_code_block_pagination(tmp_path):
    """Test code block pagination and formatting: shaded table cells with Consolas font.

    Short blocks keep_together=True; long blocks keep_together=False.
    Code blocks rendered inside single-cell tables with background shading and borders.
    """
    short_code = "\n".join([f"x = {i}" for i in range(5)])
    long_code = "\n".join([f"line_{i} = True" for i in range(35)])

    md_content = f"""# Code Chapter

```python
{short_code}
```

```python
{long_code}
```
"""
    chapter = tmp_path / "code_chapter.md"
    chapter.write_text(md_content, encoding="utf-8")
    output = tmp_path / "code_test.docx"

    result = subprocess.run(
        [sys.executable, str(SCRIPT), str(chapter), "--output", str(output)],
        capture_output=True, text=True
    )
    assert result.returncode == 0, result.stderr

    doc = docx.Document(str(output))

    # Code blocks are now inside table cells — find all table cells containing Consolas text
    code_cells = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if any(r.font.name == "Consolas" for r in p.runs):
                        code_cells.append((cell, p))
                        break
    assert len(code_cells) == 2, f"Expected 2 code blocks in table cells, found {len(code_cells)}"

    short_cell, short_p = code_cells[0]
    long_cell, long_p = code_cells[1]

    # Short block should keep lines together
    assert short_p.paragraph_format.keep_together is True
    # Long block should allow splitting
    assert long_p.paragraph_format.keep_together is False
    # Both should use mono line spacing
    assert short_p.paragraph_format.line_spacing == 1.0

    # Verify shading is applied (F5F5F5 background)
    short_shd = short_cell._tc.xml
    assert 'w:shd' in short_shd
    assert 'F5F5F5' in short_shd

    # Verify border is applied
    assert 'w:tcBorders' in short_shd

    # Verify font properties
    short_runs = [r for r in short_p.runs if r.font.name == "Consolas"]
    assert len(short_runs) > 0
    assert short_runs[0].font.size == docx.shared.Pt(9)

def test_reference_numbering_restart(tmp_path):
    """Test reference numbering restarts at 1 and does not inherit from previous lists."""
    md_content = """# Chapter with Lists

1. Item One
2. Item Two
3. Item Three

## References

1. Ref Alpha (2020).
2. Ref Beta (2022).
"""
    chapter = tmp_path / "refs_chapter.md"
    chapter.write_text(md_content, encoding="utf-8")
    output = tmp_path / "refs_test.docx"

    result = subprocess.run(
        [sys.executable, str(SCRIPT), str(chapter), "--output", str(output)],
        capture_output=True, text=True
    )
    assert result.returncode == 0, result.stderr

    doc = docx.Document(str(output))
    list_items = [p for p in doc.paragraphs if p.style.name == "List Number"]
    assert len(list_items) == 5

    body_num_id = list_items[0]._p.xpath(".//w:numPr/w:numId/@w:val")[0]
    ref_num_id = list_items[3]._p.xpath(".//w:numPr/w:numId/@w:val")[0]
    assert body_num_id != ref_num_id

    numbering_xml = doc.part.numbering_part.element.xml
    num_match = re.search(rf'<w:num [^>]*w:numId="{ref_num_id}"[^>]*>(.*?)</w:num>', numbering_xml, re.DOTALL)
    assert num_match is not None
    assert 'w:startOverride w:val="1"' in num_match.group(1)
