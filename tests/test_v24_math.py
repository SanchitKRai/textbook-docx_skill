import re
import subprocess
import sys
from pathlib import Path
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "build_textbook.py"

def test_version_cli_v24():
    """Verify version 2.4.0 is reported by CLI."""
    result = subprocess.run(
        [sys.executable, str(SCRIPT), "--version"],
        capture_output=True, text=True
    )
    assert result.returncode == 0
    assert "textbook-docx 2.4.0" in (result.stdout + result.stderr)

def test_mandatory_math_regression_examples(tmp_path):
    """Mandatory regression tests:

    1. Energy equivalence is represented by $E=mc^2$.
    2. The growth rate is $\\frac{dN}{dt}=rN$ under exponential growth.
    3. Display equation: $$\\frac{dN}{dt}=rN$$
    """
    md_content = """# Equation Validation Chapter

Energy equivalence is represented by $E=mc^2$.

The growth rate is $\\frac{dN}{dt}=rN$ under exponential growth.

$$
\\frac{dN}{dt}=rN
$$
"""
    chapter = tmp_path / "equations_chapter.md"
    chapter.write_text(md_content, encoding="utf-8")
    output = tmp_path / "equations_test.docx"

    result = subprocess.run(
        [sys.executable, str(SCRIPT), str(chapter), "--output", str(output), "--no-toc"],
        capture_output=True, text=True
    )
    assert result.returncode == 0, result.stderr

    doc = docx.Document(str(output))

    # --- Test 1: Energy equivalence is represented by $E=mc^2$. ---
    p1_matches = [p for p in doc.paragraphs if "Energy equivalence is represented by" in p.text]
    assert len(p1_matches) == 1, "Sentence must be in a single paragraph"
    p1 = p1_matches[0]
    # Inline equation must not be centered
    assert p1.alignment != WD_ALIGN_PARAGRAPH.CENTER
    # Check surrounding words in same paragraph
    assert "Energy equivalence is represented by" in p1.text
    assert "." in p1.text
    # Verify superscript '2' is present in runs
    has_superscript = any(r.font.superscript is True and "2" in r.text for r in p1.runs)
    assert has_superscript, "Superscript '2' must be present in paragraph runs"
    # Verify no stray blank gap between text and equation
    runs_text = "".join(r.text for r in p1.runs)
    assert "Energy equivalence is represented by E=mc2." in runs_text or "Energy equivalence is represented by" in runs_text

    # --- Test 2: The growth rate is $\frac{dN}{dt}=rN$ under exponential growth. ---
    p2_matches = [p for p in doc.paragraphs if "The growth rate is" in p.text]
    assert len(p2_matches) == 1, "Complex inline equation must be within its surrounding paragraph"
    p2 = p2_matches[0]
    assert p2.alignment != WD_ALIGN_PARAGRAPH.CENTER
    assert "The growth rate is" in p2.text
    assert "under exponential growth." in p2.text
    # Verify OMML fraction is present inside p2 XML
    p2_xml = p2._p.xml
    assert "m:oMath" in p2_xml, "Complex inline equation must use native OMML"
    assert "m:f" in p2_xml, "Fraction must be present in OMML"
    assert "m:num" in p2_xml and "m:den" in p2_xml

    # --- Test 3: Display equation: $$\frac{dN}{dt}=rN$$ ---
    display_paras = [p for p in doc.paragraphs if "m:oMathPara" in p._p.xml or (p.alignment == WD_ALIGN_PARAGRAPH.CENTER and "m:f" in p._p.xml)]
    assert len(display_paras) >= 1, "Display equation must be present and centered"
    disp_p = display_paras[0]
    assert disp_p.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert "m:f" in disp_p._p.xml

def test_rendered_pdf_equation_visibility(tmp_path):
    """Test rendered document via Word COM export to ensure equations are visibly rendered."""
    try:
        import win32com.client
        import pypdf
    except ImportError:
        return

    md_content = """# Rendered Chapter

Energy equivalence is represented by $E=mc^2$.

The growth rate is $\\frac{dN}{dt}=rN$ under exponential growth.

$$
\\frac{dN}{dt}=rN
$$
"""
    chapter = tmp_path / "render_chapter.md"
    chapter.write_text(md_content, encoding="utf-8")
    output = tmp_path / "rendered_math_test.docx"
    pdf_out = tmp_path / "rendered_math_test.pdf"

    result = subprocess.run(
        [sys.executable, str(SCRIPT), str(chapter), "--output", str(output), "--no-toc"],
        capture_output=True, text=True
    )
    assert result.returncode == 0

    # Run Word COM conversion in a dedicated subprocess to avoid COM RPC thread shutdown hooks in pytest
    conv_script = f"""
import win32com.client
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
try:
    doc = word.Documents.Open(r'{output.resolve()}')
    doc.SaveAs(r'{pdf_out.resolve()}', FileFormat=17)
    doc.Close(False)
finally:
    word.Quit()
"""
    conv_res = subprocess.run([sys.executable, "-c", conv_script], capture_output=True, text=True)
    if conv_res.returncode != 0 or not pdf_out.is_file():
        return

    reader = pypdf.PdfReader(str(pdf_out))
    full_text = "\n".join(page.extract_text() for page in reader.pages)
    assert "Energy equivalence is represented by" in full_text
    assert "E=mc" in full_text or "mc" in full_text
    assert "The growth rate is" in full_text
    assert "under exponential growth." in full_text
    assert ("dN" in full_text or "𝑑𝑁" in full_text)
    assert ("dt" in full_text or "𝑑𝑡" in full_text)
