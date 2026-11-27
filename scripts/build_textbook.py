#!/usr/bin/env python3
"""Build editable academic DOCX textbooks from Markdown using Pandoc."""
from __future__ import annotations
import argparse
import datetime as dt
import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt, RGBColor

VERSION = "3.0.0"
NAVY = "1B365D"
BLUE = "2C5282"
PALE = "EAF1F8"
GRAY = "F4F6F8"


def pandoc_path() -> str:
    exe = shutil.which("pandoc")
    if exe:
        return exe
    try:
        import pypandoc
        path = pypandoc.get_pandoc_path()
        if path and Path(path).is_file():
            return path
    except Exception:
        pass
    raise RuntimeError("Pandoc not found. Install with `python -m pip install 'pypandoc_binary>=1.13,<2'` or install system Pandoc.")


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")


def add_cell_borders(cell, color="D7DEE7", size="5") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn("w:" + edge))
        if node is None:
            node = OxmlElement("w:" + edge)
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:color"), color)


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_table_spacing(table, before=100, after=120) -> None:
    tbl_pr = table._tbl.tblPr
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "autofit")
    # Word table spacing is represented by cell paragraph spacing; add consistent
    # after-spacing to paragraphs in the table without inserting duplicate tblW.
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(3)


def add_page_field(paragraph) -> None:
    r = paragraph.add_run()
    begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve"); instr.text = " PAGE "
    separate = OxmlElement("w:fldChar"); separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t"); text.text = "1"
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, separate, text, end): r._r.append(el)


def add_toc_field(paragraph) -> None:
    r = paragraph.add_run()
    begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve"); instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    separate = OxmlElement("w:fldChar"); separate.set(qn("w:fldCharType"), "separate")
    for el in (begin, instr, separate): r._r.append(el)
    r2 = paragraph.add_run("Contents are listed above; update page numbers in Word (Ctrl+A, F9).")
    r2.italic = True; r2.font.size = Pt(9); r2.font.color.rgb = RGBColor.from_string("6B7280")
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    r3 = paragraph.add_run(); r3._r.append(end)


def add_toc_preview(doc: Document, heading_records: list[tuple[int, str]], page_break=True) -> None:
    p = doc.add_paragraph("Table of Contents", style="Title")
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.outline_level = 9
    for level, title in heading_records:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(max(0, level - 1) * 0.22)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(title)
        r.font.name = "Calibri"
        r.font.size = Pt(10 if level == 1 else 9)
        if level == 1: r.bold = True
    p = doc.add_paragraph()
    add_toc_field(p)
    if page_break:
        doc.add_page_break()


def parse_heading_records(paths: list[Path]) -> list[tuple[int, str]]:
    records = []
    for path in paths:
        md = path.read_text(encoding="utf-8")
        in_code = False
        for line in md.splitlines():
            if line.lstrip().startswith("```") or line.lstrip().startswith("~~~"):
                in_code = not in_code; continue
            if in_code: continue
            m = re.match(r"^\s*(#{1,3})\s+(.+?)\s*#*\s*$", line)
            if m and not re.match(r"References\s*$", m.group(2), re.I):
                heading = re.sub(r"\s+#+\s*$", "", m.group(2))
                # Pandoc interprets inline Markdown in headings; keep the static
                # TOC preview consistent and free of raw emphasis/code markers.
                heading = re.sub(r"\*\*(.*?)\*\*", r"\1", heading)
                heading = re.sub(r"(?<!\*)\*(.*?)\*(?!\*)", r"\1", heading)
                heading = re.sub(r"`([^`]+)`", r"\1", heading)
                records.append((len(m.group(1)), heading))
    return records


def find_asset_paths(markdown: str, source_dir: Path) -> list[Path]:
    found = []
    for m in re.finditer(r"!\[[^\]]*\]\((?:<([^>]+)>|([^\s)]+))(?:\s+['\"][^'\"]*['\"])?\)", markdown):
        raw = m.group(1) or m.group(2)
        if re.match(r"^[a-z]+://", raw, re.I):
            continue
        path = (source_dir / raw).resolve()
        if path.exists(): found.append(path)
    return found


def annotate_markdown(markdown: str, source: Path, chapter_index: int, image_counter: list[int]) -> str:
    """Add stable per-chapter figure captions/alt text; Pandoc parses the Markdown."""
    out = []
    for line in markdown.splitlines():
        m = re.match(r"^(\s*)!\[([^\]]*)\]\((?:<([^>]+)>|([^\s)]+))(?:\s+(['\"])(.*?)\5)?\)\s*$", line)
        if m:
            indent, alt, p1, p2, quote, title = m.groups()
            target = p1 or p2
            image_counter[0] += 1
            caption = title or alt or Path(target).stem
            alt_text = f"Figure {chapter_index}.{image_counter[0]}: {caption}"
            # Ensure Pandoc gets a title/caption, and keep spaces-containing paths intact.
            resolved_target = (source.parent / target).resolve()
            if not resolved_target.is_file():
                raise FileNotFoundError(f"Image referenced by {source}: {target}")
            escaped_target = "<" + str(resolved_target).replace("\\", "/") + ">"
            out.append(f'{indent}![{alt_text}]({escaped_target} "Figure {chapter_index}.{image_counter[0]}: {caption}")')
        else:
            out.append(line)
    return "\n".join(out) + "\n"


def extract_image_alt_texts(paths: list[Path]) -> list[str]:
    alts = []
    pattern = re.compile(r"^\s*!\[([^\]]*)\]\((?:<([^>]+)>|([^\s)]+))(?:\s+(['\"])(.*?)\4)?\)\s*$", re.M)
    for source in paths:
        for match in pattern.finditer(source.read_text(encoding="utf-8")):
            alts.append(match.group(1) or match.group(5) or Path(match.group(2) or match.group(3)).stem)
    return alts


def apply_image_alt_texts(doc: Document, alts: list[str]) -> None:
    nodes = list(doc._element.body.iter())
    pictures = [node for node in nodes if node.tag.endswith("}docPr")]
    for node, alt in zip(pictures, alts):
        node.set("descr", alt)


SUPPORTED_TEX_COMMANDS = {
    "frac", "dfrac", "tfrac", "sqrt", "left", "right", "middle", "text", "mathrm", "mathbf", "mathit", "mathsf", "mathtt", "mathbb", "mathcal", "mathfrak", "operatorname", "overline", "underline", "vec", "hat", "bar", "tilde", "dot", "ddot", "widehat", "widetilde", "overrightarrow", "overleftarrow", "overbrace", "underbrace", "sum", "prod", "int", "iint", "iiint", "oint", "lim", "sin", "cos", "tan", "sec", "csc", "cot", "arcsin", "arccos", "arctan", "log", "ln", "exp", "max", "min", "arg", "sup", "inf", "det", "gcd", "begin", "end", "matrix", "pmatrix", "bmatrix", "Bmatrix", "vmatrix", "Vmatrix", "cases", "aligned", "align", "split", "smallmatrix", "array", "alpha", "beta", "gamma", "delta", "epsilon", "varepsilon", "zeta", "eta", "theta", "vartheta", "iota", "kappa", "lambda", "mu", "nu", "xi", "pi", "varpi", "rho", "varrho", "sigma", "varsigma", "tau", "upsilon", "phi", "varphi", "chi", "psi", "omega", "Gamma", "Delta", "Theta", "Lambda", "Xi", "Pi", "Sigma", "Upsilon", "Phi", "Psi", "Omega", "times", "cdot", "div", "pm", "mp", "le", "leq", "ge", "geq", "ne", "neq", "approx", "equiv", "sim", "propto", "infty", "partial", "nabla", "to", "rightarrow", "leftarrow", "Rightarrow", "Leftarrow", "Leftrightarrow", "in", "notin", "subset", "subseteq", "cap", "cup", "forall", "exists", "quad", "qquad", "displaystyle", "limits", "nolimits", ",", ";", "!"
}
MATH_TOKEN = re.compile(r"\$\$(.+?)\$\$|(?<!\$)\$(?!\s)(.+?)(?<!\s)\$(?!\d)", re.S)


def preserve_unsupported_math(markdown: str, source: Path) -> str:
    """Warn with source/line and visibly preserve unknown TeX commands as styled source."""
    def replace(match):
        expr = (match.group(1) or match.group(2) or "").strip()
        commands = re.findall(r"(?<!\\)\\([A-Za-z]+)", expr)
        unknown = sorted({name for name in commands if name not in SUPPORTED_TEX_COMMANDS})
        if not unknown:
            return match.group(0)
        line = markdown[:match.start()].count("\n") + 1
        print(f"warning: {source}:{line}: unsupported math command(s) {', '.join('\\' + x for x in unknown)}; preserved source expression: {expr}", file=sys.stderr)
        if match.group(1) is not None:
            return f"\n\n**Unconverted math source:** `{expr}`\n\n"
        return f"*Unconverted math source: `{expr}`*"
    return MATH_TOKEN.sub(replace, markdown)


def normalize_short_table_separators(markdown: str) -> str:
    """Pad one/two-dash pipe-table separators to GFM's three-dash minimum."""
    lines = markdown.splitlines()
    for idx, line in enumerate(lines):
        if not line.strip().startswith("|") or idx == 0:
            continue
        previous = lines[idx - 1].strip()
        if not previous.startswith("|"):
            continue
        cells = line.strip().strip("|").split("|")
        normalized = []
        valid = True
        for cell in cells:
            token = cell.strip()
            match = re.fullmatch(r"(:)?(-+)(:)?", token)
            if not match:
                valid = False; break
            normalized.append((":" if match.group(1) else "") + "---" + (":" if match.group(3) else ""))
        if valid:
            lines[idx] = "| " + " | ".join(normalized) + " |"
    return "\n".join(lines)


def prepare_markdown(paths: list[Path], work: Path) -> tuple[list[Path], list[tuple[int, str]]]:
    prepared = []
    heading_records = []
    for idx, source in enumerate(paths, 1):
        data = preserve_unsupported_math(source.read_text(encoding="utf-8"), source)
        data = normalize_short_table_separators(data)
        refs = re.search(r"(?im)^#{1,3}\s+References\s*$", data)
        if refs:
            front, refbody = data[:refs.start()], data[refs.end():]
            # Let Pandoc parse APA paragraphs / list refs. Hanging-indent style applied post-conversion.
            data = front + "\n## References\n" + refbody
        counter = [0]
        data = annotate_markdown(data, source, idx, counter)
        target = work / f"chapter_{idx:02d}.md"
        target.write_text(data, encoding="utf-8")
        prepared.append(target)
        heading_records.extend(parse_heading_records([target]))
    return prepared, heading_records


def extract_refs_blocks(doc: Document) -> None:
    """Apply hanging indents to paragraphs under References heading."""
    in_refs = False
    for p in doc.paragraphs:
        if p.style.name.startswith("Heading") and p.text.strip().lower() == "references":
            in_refs = True; continue
        if in_refs:
            if p.style.name.startswith("Heading"):
                break
            p.paragraph_format.left_indent = Inches(0.28)
            p.paragraph_format.first_line_indent = Inches(-0.28)
            p.paragraph_format.space_after = Pt(6)


def add_cover_and_chrome(doc: Document, title, subtitle, author, institution, date, header, footer, toc_enabled, heading_records, page_size):
    # Rebuild document body while retaining section/settings relationships from Pandoc template.
    body = doc._element.body
    sect_pr = body.sectPr
    for child in list(body):
        if child is not sect_pr: body.remove(child)
    for section in doc.sections:
        section.page_width, section.page_height = (Mm(210), Mm(297)) if page_size == "A4" else (Inches(8.5), Inches(11))
        section.top_margin = section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(1)
        section.right_margin = Inches(0.85)
        section.different_first_page_header_footer = True
        hp = section.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rr = hp.add_run(header or title); rr.italic = True; rr.font.size = Pt(8); rr.font.color.rgb = RGBColor.from_string("65758B")
        fp = section.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if footer: fp.add_run(footer + " | ")
        add_page_field(fp)
    for _ in range(3): doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title); r.bold = True; r.font.name = "Calibri"; r.font.size = Pt(28); r.font.color.rgb = RGBColor.from_string(NAVY)
    for val in (subtitle, author, institution, date):
        if val:
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rr = p.add_run(val); rr.font.size = Pt(13); rr.font.color.rgb = RGBColor.from_string(BLUE)
    doc.add_page_break()
    if toc_enabled:
        add_toc_preview(doc, heading_records)


def insert_ppr_ordered(p_pr, element) -> None:
    order = {
        "pStyle": 0, "keepNext": 1, "keepLines": 2, "pageBreakBefore": 3,
        "framePr": 4, "widowControl": 5, "numPr": 6, "suppressLineNumbers": 7,
        "pBdr": 8, "shd": 9, "tabs": 10, "spacing": 21, "ind": 22,
        "jc": 26, "outlineLvl": 30, "rPr": 33,
    }
    rank = order.get(element.tag.split("}")[-1], 50)
    for idx, child in enumerate(p_pr):
        child_rank = order.get(child.tag.split("}")[-1], 50)
        if child_rank > rank:
            p_pr.insert(idx, element)
            return
    p_pr.append(element)


def add_spacing_after_tables(doc: Document) -> None:
    """Separate tables/callouts from following content without duplicating table properties."""
    body = doc._element.body
    for child in list(body):
        if child.tag != qn("w:tbl"):
            continue
        nxt = child.getnext()
        if nxt is None or nxt.tag == qn("w:sectPr"):
            continue
        if nxt.tag == qn("w:p"):
            from docx.text.paragraph import Paragraph
            paragraph = Paragraph(nxt, doc._body)
            current = paragraph.paragraph_format.space_before
            current_pt = current.pt if current is not None else 0
            if current_pt < 5:
                paragraph.paragraph_format.space_before = Pt(5)
        else:
            spacer = OxmlElement("w:p")
            ppr = OxmlElement("w:pPr")
            spacing = OxmlElement("w:spacing")
            spacing.set(qn("w:before"), "80")
            spacing.set(qn("w:after"), "80")
            ppr.append(spacing)
            spacer.append(ppr)
            child.addnext(spacer)


def insert_tcpr_ordered(tc_pr, element) -> None:
    order = {"tcW": 0, "gridSpan": 1, "hMerge": 2, "vMerge": 3, "tcBorders": 4, "shd": 5, "noWrap": 6, "tcMar": 7, "textDirection": 8, "tcFitText": 9, "vAlign": 10, "hideMark": 11, "headers": 12, "cellIns": 13, "cellDel": 14, "cellMerge": 15, "tcPrChange": 16}
    rank = order.get(element.tag.split("}")[-1], 50)
    for idx, child in enumerate(tc_pr):
        if order.get(child.tag.split("}")[-1], 50) > rank:
            tc_pr.insert(idx, element); return
    tc_pr.append(element)


def set_callout_cell(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        node = OxmlElement("w:" + edge)
        node.set(qn("w:val"), "single"); node.set(qn("w:sz"), "8")
        node.set(qn("w:color"), "C8D8E8"); borders.append(node)
    insert_tcpr_ordered(tc_pr, borders)
    shading = OxmlElement("w:shd"); shading.set(qn("w:val"), "clear"); shading.set(qn("w:fill"), PALE)
    insert_tcpr_ordered(tc_pr, shading)
    margins = OxmlElement("w:tcMar")
    for side, width in (("top", "100"), ("start", "140"), ("bottom", "100"), ("end", "140")):
        item = OxmlElement("w:" + side); item.set(qn("w:w"), width); item.set(qn("w:type"), "dxa"); margins.append(item)
    insert_tcpr_ordered(tc_pr, margins)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def convert_blockquotes_to_callout_tables(doc: Document) -> None:
    """Group consecutive Pandoc Block Text paragraphs into one shaded one-cell callout table."""
    body = doc._element.body
    children = list(body)
    groups = []
    current = []
    for child in children:
        is_block = child.tag == qn("w:p") and child.find(qn("w:pPr")) is not None and child.find(qn("w:pPr")).find(qn("w:pStyle")) is not None and child.find(qn("w:pPr")).find(qn("w:pStyle")).get(qn("w:val")) == "BlockText"
        if is_block:
            current.append(child)
        elif current:
            groups.append(current); current = []
    if current: groups.append(current)
    # Process in reverse document order so saved child indices remain stable.
    for group in reversed(groups):
        first_index = body.index(group[0])
        table = doc.add_table(rows=1, cols=1)
        tbl = table._tbl
        tbl.getparent().remove(tbl)
        body.insert(first_index, tbl)
        cell = table.cell(0, 0)
        # Apply borders directly; the minimal Pandoc-generated template does not
        # necessarily include the built-in Table Grid style.
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        default_p = cell.paragraphs[0]._p
        cell._tc.remove(default_p)
        for paragraph in group:
            body.remove(paragraph)
            cell._tc.append(paragraph)
        set_callout_cell(cell)
        for paragraph in cell.paragraphs:
            paragraph.paragraph_format.left_indent = Inches(0)
            paragraph.paragraph_format.right_indent = Inches(0)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(3)
        tbl_pr = tbl.tblPr
        caption = OxmlElement("w:tblCaption"); caption.set(qn("w:val"), "Callout")
        # tblCaption follows tblLook and is kept after standard table properties.
        tbl_pr.append(caption)


def style_document(doc: Document) -> None:
    convert_blockquotes_to_callout_tables(doc)
    styles = doc.styles
    get_style = lambda name: styles.get_by_id(name.replace(" ", ""), 1)
    normal = get_style("Normal")
    normal.font.name = "Georgia"; normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.12
    for name, size in (("Heading 1", 16), ("Heading 2", 13), ("Heading 3", 11)):
        st = get_style(name); st.font.name = "Calibri"; st.font.size = Pt(size)
        st.font.bold = True; st.font.color.rgb = RGBColor.from_string(NAVY if name == "Heading 1" else BLUE)
        st.paragraph_format.keep_with_next = True
    # Pandoc blockquotes become Block Text paragraphs; style them as restrained callouts.
    if any(s.style_id == "BlockText" for s in styles):
        st = get_style("Block Text")
        st.font.name = "Georgia"; st.font.size = Pt(10.5); st.font.bold = False
        st.paragraph_format.left_indent = Inches(0.18)
        st.paragraph_format.right_indent = Inches(0.12)
        st.paragraph_format.space_before = Pt(5); st.paragraph_format.space_after = Pt(8)
    if any(s.style_id == "SourceCode" for s in styles):
        st = get_style("Source Code")
        st.font.name = "Consolas"; st.font.size = Pt(9)
        st.paragraph_format.left_indent = Inches(0.12)
        st.paragraph_format.space_before = Pt(4); st.paragraph_format.space_after = Pt(8)
        st.paragraph_format.keep_together = True
    # Native table header repeat and consistent alignment.
    for table in doc.tables:
        caption = table._tbl.tblPr.find(qn("w:tblCaption"))
        if table.rows and (caption is None or caption.get(qn("w:val")) != "Callout"):
            set_repeat_table_header(table.rows[0])
        set_table_spacing(table)
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    # Source-code paragraphs receive explicit shading and padding.
    for p in doc.paragraphs:
        if p.style.name == "Source Code":
            p_pr = p._p.get_or_add_pPr()
            shd = OxmlElement("w:shd"); shd.set(qn("w:fill"), GRAY); shd.set(qn("w:val"), "clear"); insert_ppr_ordered(p_pr, shd)

    for p in doc.paragraphs:
        if p.style.name.lower().startswith("source code"):
            p.paragraph_format.space_before = Pt(5); p.paragraph_format.space_after = Pt(8)
    extract_refs_blocks(doc)
    add_spacing_after_tables(doc)


def set_metadata(doc: Document, title: str, author: str | None) -> None:
    props = doc.core_properties
    now = dt.datetime.now(dt.timezone.utc).replace(tzinfo=None, microsecond=0)
    props.title = title
    props.author = author or ""
    props.subject = ""
    props.comments = ""
    props.keywords = ""
    props.category = ""
    props.created = now
    props.modified = now
    props.last_modified_by = author or ""


def fix_settings_order(doc: Document) -> None:
    settings = doc.settings.element
    existing = settings.find(qn("w:updateFields"))
    if existing is not None: settings.remove(existing)
    update = OxmlElement("w:updateFields"); update.set(qn("w:val"), "true")
    # CT_Settings schema puts updateFields before compatibility/document protection tail nodes.
    insert_before = ("w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars", "w:rsids", "m:mathPr", "w:attachedSchema", "w:themeFontLang", "w:clrSchemeMapping", "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures", "w:forceUpgrade", "w: Captions")
    qnames = [qn(x) for x in insert_before if " " not in x]
    for child in list(settings):
        if child.tag in qnames:
            settings.insert(settings.index(child), update); break
    else:
        settings.append(update)


def postprocess(path: Path, title: str, author: str | None, subtitle: str | None, institution: str | None, date: str | None, header: str | None, footer: str | None, toc: bool, heading_records, page_size: str, image_alts: list[str]) -> None:
    doc = Document(path)
    # Save converted body, then inject front matter before the existing body contents.
    body = doc._element.body
    sect_pr = body.sectPr
    content = [el for el in list(body) if el is not sect_pr]
    for el in content: body.remove(el)
    # front matter is inserted using python-docx; existing content nodes are then appended back.
    add_cover_and_chrome(doc, title, subtitle, author, institution, date, header, footer, toc, heading_records, page_size)
    for el in content: body.insert(body.index(sect_pr) if sect_pr is not None else len(body), el)
    # Replace Pandoc marker paragraph with an actual page break between chapters.
    for p in list(doc.paragraphs):
        if p.text.strip() == "CHAPTERBREAKTOKEN":
            p.clear()
            p.add_run().add_break(WD_BREAK.PAGE)
    apply_image_alt_texts(doc, image_alts)
    style_document(doc)
    set_metadata(doc, title, author)
    fix_settings_order(doc)
    doc.save(path)


def validate_docx(path: Path) -> None:
    doc = Document(path)
    if not doc.paragraphs and not doc.tables:
        raise RuntimeError("Generated DOCX has no body content")
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        if bad: raise RuntimeError(f"DOCX ZIP integrity error: {bad}")
        xml = zf.read("word/document.xml").decode("utf-8", "replace")
        for forbidden in ("<w:tblW",):
            # Multiple table widths per tblPr are a known validity defect; check each table.
            pass
        for match in re.finditer(r"<w:tblPr(?:\s[^>]*)?>(.*?)</w:tblPr>", xml, re.S):
            if match.group(1).count("<w:tblW") > 1:
                raise RuntimeError("Duplicate w:tblW in table properties")
        if "<w:updateFields" not in zf.read("word/settings.xml").decode("utf-8", "replace"):
            raise RuntimeError("Missing updateFields setting")


def build(inputs, output, title=None, subtitle=None, author=None, institution=None, date=None, page_size="A4", header_text=None, footer_text=None, toc=True, overwrite=False):
    paths = [Path(x).expanduser().resolve() for x in inputs]
    out = Path(output).expanduser().resolve()
    if not paths or any(not p.is_file() or p.suffix.lower() != ".md" for p in paths):
        raise ValueError("Provide one or more existing Markdown (.md) chapter files")
    if out.exists() and not overwrite:
        raise FileExistsError(f"{out} already exists; pass --overwrite to replace")
    pandoc = pandoc_path()
    title = title or paths[0].stem.replace("_", " ").title()
    image_alts = extract_image_alt_texts(paths)
    with tempfile.TemporaryDirectory(prefix="textbook_docx_") as tmp:
        work = Path(tmp)
        prepared, headings = prepare_markdown(paths, work)
        combined = work / "combined.md"
        combined.write_text("\n\nCHAPTERBREAKTOKEN\n\n".join(p.read_text(encoding="utf-8") for p in prepared), encoding="utf-8")
        reference = work / "reference.docx"
        template = subprocess.run([pandoc, "--print-default-data-file=reference.docx"], check=True, capture_output=True)
        reference.write_bytes(template.stdout)
        args = [pandoc, str(combined), "--from=markdown+tex_math_dollars+pipe_tables+fenced_code_blocks+implicit_figures+link_attributes+autolink_bare_uris", "--to=docx", "--reference-doc", str(reference), "--standalone", "--output", str(out)]
        result = subprocess.run(args, capture_output=True, text=True)
        if result.stderr.strip():
            print(result.stderr.strip(), file=sys.stderr)
        if result.returncode:
            raise RuntimeError("Pandoc failed: " + (result.stderr.strip() or result.stdout.strip()))
    postprocess(out, title, author, subtitle, institution, date, header_text, footer_text, toc, headings, page_size, image_alts)
    validate_docx(out)
    return out


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Build editable Word textbooks from Markdown chapters using Pandoc.")
    parser.add_argument("inputs", nargs="+", type=Path)
    parser.add_argument("-o", "--output", required=True, type=Path)
    for flag in ("title", "subtitle", "author", "institution", "date", "header-text", "footer-text"):
        parser.add_argument("--" + flag)
    parser.add_argument("--page-size", choices=("A4", "LETTER"), default="A4")
    parser.add_argument("--no-toc", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--version", action="version", version=f"textbook-docx {VERSION}")
    a = parser.parse_args(argv)
    try:
        path = build(a.inputs, a.output, a.title, a.subtitle, a.author, a.institution, a.date, a.page_size, a.header_text, a.footer_text, not a.no_toc, a.overwrite)
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    print(f"Created and validated: {path}")
    if not a.no_toc: print("TOC preview is populated; update fields in Word with Ctrl+A, F9 to refresh page numbers.")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
