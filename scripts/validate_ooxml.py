#!/usr/bin/env python3
"""Repository-focused OOXML integrity/schema-order audit for generated DOCX files.

This checks package integrity, XML well-formedness, and ordering/uniqueness rules
for elements emitted or modified by this project. It is not a substitute for the
full ECMA-376 XSD validator; configure DOCX_XSD_VALIDATOR to run an external
ISO/IEC 29500 validator when one is available in the environment.
"""
from __future__ import annotations
import argparse
import os
import subprocess
import sys
import zipfile
from pathlib import Path
from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}

def local(tag: str) -> str:
    return etree.QName(tag).localname

def check_order(parent, child_order, label, errors):
    actual = [local(child.tag) for child in parent]
    ranks = {name: idx for idx, name in enumerate(child_order)}
    relevant = [(ranks[name], name) for name in actual if name in ranks]
    seq = [rank for rank, _ in relevant]
    if seq != sorted(seq):
        errors.append(f"Invalid {label} child order: {' -> '.join(name for _, name in relevant)}")

def validate(path: Path) -> list[str]:
    errors = []
    with zipfile.ZipFile(path) as zf:
        corrupt = zf.testzip()
        if corrupt:
            errors.append(f"ZIP CRC failure: {corrupt}")
        xml_files = [n for n in zf.namelist() if n.endswith(".xml") or n.endswith(".rels")]
        parsed = {}
        for name in xml_files:
            try:
                parsed[name] = etree.fromstring(zf.read(name))
            except etree.XMLSyntaxError as exc:
                errors.append(f"Malformed XML {name}: {exc}")
        doc = parsed.get("word/document.xml")
        if doc is not None:
            for tbl_pr in doc.xpath(".//w:tblPr", namespaces=NS):
                if len(tbl_pr.xpath("./w:tblW", namespaces=NS)) > 1:
                    errors.append("Duplicate w:tblW in w:tblPr")
            tc_order = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"]
            for tc_pr in doc.xpath(".//w:tcPr", namespaces=NS):
                check_order(tc_pr, tc_order, "w:tcPr", errors)
            p_order = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
            for p_pr in doc.xpath(".//w:pPr", namespaces=NS):
                check_order(p_pr, p_order, "w:pPr", errors)
        settings = parsed.get("word/settings.xml")
        if settings is not None:
            updates = settings.xpath("./w:updateFields", namespaces=NS)
            if len(updates) > 1:
                errors.append("Duplicate w:updateFields in settings.xml")
            if updates:
                for name in ("hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade"):
                    later = settings.xpath(f"./w:{name}", namespaces=NS)
                    if later and settings.index(updates[0]) > settings.index(later[0]):
                        errors.append(f"w:updateFields appears after w:{name}; settings.xml order is invalid")
                        break
        # All package relationship parts must parse and internal targets must exist.
        names = set(zf.namelist())
        for rel_name in [n for n in names if n.endswith(".rels")]:
            rel_root = parsed.get(rel_name)
            if rel_root is None: continue
            base = Path(rel_name).parent.parent if rel_name != "_rels/.rels" else Path(".")
            for rel in rel_root:
                if rel.get("TargetMode") == "External": continue
                target = rel.get("Target", "")
                if not target: continue
                resolved = (base / target).as_posix()
                while "/../" in resolved:
                    parts = resolved.split("/"); i = parts.index(".."); resolved = "/".join(parts[:i-1] + parts[i+1:])
                if resolved not in names and not target.startswith("#"):
                    errors.append(f"Broken relationship target in {rel_name}: {target} -> {resolved}")
    external = os.environ.get("DOCX_XSD_VALIDATOR")
    if external:
        proc = subprocess.run([external, str(path)], capture_output=True, text=True)
        if proc.returncode:
            errors.append("External ISO OOXML XSD validator failed: " + (proc.stderr or proc.stdout).strip())
    return errors

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    args = parser.parse_args()
    errors = validate(args.docx)
    if errors:
        for err in errors: print("FAIL:", err, file=sys.stderr)
        return 1
    print(f"OOXML package/schema-order audit passed: {args.docx}")
    if not os.environ.get("DOCX_XSD_VALIDATOR"):
        print("NOTE: Full ISO/IEC 29500 XSD validation was not run; no external validator is configured.")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
