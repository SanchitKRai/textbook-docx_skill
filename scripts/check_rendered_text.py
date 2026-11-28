#!/usr/bin/env python3
"""Strict PDF text/page checks shared by CI and local verification."""
import re, subprocess, sys
from pathlib import Path
if len(sys.argv) != 3:
    raise SystemExit("usage: check_rendered_text.py FILE.pdf TEXT.txt")
pdf, txt = map(Path, sys.argv[1:])
text = txt.read_text(encoding="utf-8", errors="replace")
pages = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, check=True).stdout).group(1))
page_text = [p.strip() for p in text.split("\f") if p.strip()]
assert len(page_text) == pages, f"blank/unreadable PDF page: {len(page_text)} text pages vs {pages} physical pages"
for token in (r"\sum", r"\int", r"\frac", r"\sqrt", "CHAPTERBREAKTOKEN", "!["):
    assert token not in text, f"raw Markdown/TeX leaked into PDF: {token}"
prose = text.replace("2 ** number_of_doublings", "POWER_OPERATOR_IN_CODE")
assert "**Cost" not in prose and "**Chapter" not in prose and "__" not in prose
for token in ("costs $5", "costs $10", "Figure 1.1:", "Figure 2.1:", "Figure 2.2:", "Alberts, B.", "Cooper, G.", "Lodish, H."):
    assert token in text, f"required rendered content missing: {token}"
print(f"Rendered PDF checks passed: {pages} pages, no blank pages or raw syntax leakage")
