from pathlib import Path
import subprocess
import sys
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "build_textbook.py"
SAMPLE = ROOT / "examples" / "sample_chapter.md"

def test_cli_generates_valid_docx(tmp_path):
    output = tmp_path / "sample.docx"
    result = subprocess.run(
        [sys.executable, str(SCRIPT), str(SAMPLE), "--output", str(output),
         "--title", "CLI Test Textbook", "--author", "CI Test"],
        capture_output=True, text=True,
    )
    assert result.returncode == 0, result.stderr
    assert output.is_file() and output.stat().st_size > 0
    doc = Document(str(output))
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "CLI Test Textbook" in text
    assert "Learning objectives" in text
    assert "References" in text

def test_cli_refuses_overwrite_without_flag(tmp_path):
    output = tmp_path / "existing.docx"
    output.write_bytes(b"keep")
    result = subprocess.run(
        [sys.executable, str(SCRIPT), str(SAMPLE), "--output", str(output)],
        capture_output=True, text=True,
    )
    assert result.returncode != 0
    assert output.read_bytes() == b"keep"
