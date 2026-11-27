#!/usr/bin/env python3
"""Textbook DOCX builder v2.4: Markdown chapters to formatted Word documents with native equation support."""
import argparse, re, sys, tempfile
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Mm, Pt, RGBColor
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import xml.sax.saxutils as saxutils

NAVY = RGBColor(27, 54, 93)
SLATE = RGBColor(44, 82, 130)
MUTED = RGBColor(74, 85, 104)

GREEK_AND_SYMBOLS = {
    r'\alpha': 'α', r'\beta': 'β', r'\gamma': 'γ', r'\delta': 'δ', r'\epsilon': 'ε',
    r'\varepsilon': 'ε', r'\zeta': 'ζ', r'\eta': 'η', r'\theta': 'θ', r'\vartheta': 'ϑ',
    r'\iota': 'ι', r'\kappa': 'κ', r'\lambda': 'λ', r'\mu': 'μ', r'\nu': 'ν',
    r'\xi': 'ξ', r'\pi': 'π', r'\varpi': 'ϖ', r'\rho': 'ρ', r'\varrho': 'ϱ',
    r'\sigma': 'σ', r'\varsigma': 'ς', r'\tau': 'τ', r'\upsilon': 'υ', r'\phi': 'φ',
    r'\varphi': 'ϕ', r'\chi': 'χ', r'\psi': 'ψ', r'\omega': 'ω',
    r'\Gamma': 'Γ', r'\Delta': 'Δ', r'\Theta': 'Θ', r'\Lambda': 'Λ', r'\Xi': 'Ξ',
    r'\Pi': 'Π', r'\Sigma': 'Σ', r'\Upsilon': 'Υ', r'\Phi': 'Φ', r'\Psi': 'Ψ', r'\Omega': 'Ω',
    r'\times': '×', r'\cdot': '·', r'\div': '÷', r'\pm': '±', r'\mp': '∓',
    r'\le': '≤', r'\leq': '≤', r'\ge': '≥', r'\geq': '≥', r'\ne': '≠', r'\neq': '≠',
    r'\approx': '≈', r'\equiv': '≡', r'\sim': '∼', r'\propto': '∝',
    r'\infty': '∞', r'\partial': '∂', r'\nabla': '∇',
    r'\to': '→', r'\rightarrow': '→', r'\leftarrow': '←', r'\Rightarrow': '⇒',
    r'\Leftarrow': '⇐', r'\Leftrightarrow': '⇔',
    r'\in': '∈', r'\notin': '∉', r'\subset': '⊂', r'\subseteq': '⊆',
    r'\cap': '∩', r'\cup': '∪', r'\forall': '∀', r'\exists': '∃',
    r'\quad': '  ', r'\qquad': '    ', r'\,': ' ', r'\;': ' ', r'\!': '',
}

def escape_xml(s):
    return saxutils.escape(s)

def parse_bracket_content(s, idx):
    """Extract content inside balanced { ... } starting at s[idx] == '{'."""
    assert s[idx] == '{'
    depth = 0
    start = idx + 1
    for i in range(idx, len(s)):
        if s[i] == '{':
            depth += 1
        elif s[i] == '}':
            depth -= 1
            if depth == 0:
                return s[start:i], i + 1
    return s[start:], len(s)

def latex_to_omml(latex_expr):
    """Convert LaTeX math expression to OMML <m:oMath> XML string."""
    s = latex_expr.strip()
    for sym, repl in GREEK_AND_SYMBOLS.items():
        s = re.sub(re.escape(sym) + r'(?![a-zA-Z])', repl, s)

    def parse_tokens(text):
        out = []
        i = 0
        n = len(text)
        while i < n:
            c = text[i]
            if c.isspace():
                i += 1
                continue

            if text[i:].startswith(r'\frac'):
                i += 5
                while i < n and text[i].isspace(): i += 1
                if i < n and text[i] == '{':
                    num, i = parse_bracket_content(text, i)
                else:
                    num = text[i] if i < n else ''
                    i += 1
                while i < n and text[i].isspace(): i += 1
                if i < n and text[i] == '{':
                    den, i = parse_bracket_content(text, i)
                else:
                    den = text[i] if i < n else ''
                    i += 1
                num_omml = "".join(parse_tokens(num))
                den_omml = "".join(parse_tokens(den))
                out.append(f'<m:f><m:num>{num_omml}</m:num><m:den>{den_omml}</m:den></m:f>')
                continue

            if text[i:].startswith(r'\sqrt'):
                i += 5
                while i < n and text[i].isspace(): i += 1
                if i < n and text[i] == '{':
                    arg, i = parse_bracket_content(text, i)
                else:
                    arg = text[i] if i < n else ''
                    i += 1
                arg_omml = "".join(parse_tokens(arg))
                out.append(f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>{arg_omml}</m:e></m:rad>')
                continue

            if c in ('^', '_'):
                is_sup = (c == '^')
                i += 1
                while i < n and text[i].isspace(): i += 1
                if i < n and text[i] == '{':
                    script_content, i = parse_bracket_content(text, i)
                else:
                    script_content = text[i] if i < n else ''
                    i += 1
                script_omml = "".join(parse_tokens(script_content))

                prev = out.pop() if out else '<m:r><m:t></m:t></m:r>'

                next_is_comp = False
                comp_content = None
                j = i
                while j < n and text[j].isspace(): j += 1
                if j < n and text[j] in ('^', '_') and text[j] != c:
                    j += 1
                    while j < n and text[j].isspace(): j += 1
                    if j < n and text[j] == '{':
                        comp_content, j = parse_bracket_content(text, j)
                    else:
                        comp_content = text[j] if j < n else ''
                        j += 1
                    i = j
                    next_is_comp = True

                if next_is_comp:
                    comp_omml = "".join(parse_tokens(comp_content))
                    sub_omml = comp_omml if is_sup else script_omml
                    sup_omml = script_omml if is_sup else comp_omml
                    out.append(f'<m:sSubSup><m:e>{prev}</m:e><m:sub>{sub_omml}</m:sub><m:sup>{sup_omml}</m:sup></m:sSubSup>')
                elif is_sup:
                    out.append(f'<m:sSup><m:e>{prev}</m:e><m:sup>{script_omml}</m:sup></m:sSup>')
                else:
                    out.append(f'<m:sSub><m:e>{prev}</m:e><m:sub>{script_omml}</m:sub></m:sSub>')
                continue

            out.append(f'<m:r><m:t>{escape_xml(c)}</m:t></m:r>')
            i += 1

        return out

    body = "".join(parse_tokens(s))
    return f'<m:oMath {nsdecls("m")}>{body}</m:oMath>'

def is_complex_math(expr):
    """Determine if math expression requires structured OMML (fractions, roots, integrals, etc.)."""
    return any(k in expr for k in [r'\frac', r'\sqrt', r'\int', r'\sum', r'\prod', r'\lim', r'\matrix', r'\begin'])

def render_simple_inline_math(p, expr):
    """Render simple inline math into paragraph using native Word runs with superscript/subscript."""
    s = expr.strip()
    for sym, repl in GREEK_AND_SYMBOLS.items():
        s = re.sub(re.escape(sym) + r'(?![a-zA-Z])', repl, s)

    tokens = re.split(r'(\^\{[^}]+\}|\^[0-9a-zA-Z]|_{[^}]+}|_[0-9a-zA-Z])', s)
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith('^'):
            val = tok[2:-1] if tok.startswith('^{') else tok[1:]
            r = p.add_run(val)
            r.font.superscript = True
            r.font.name = 'Cambria Math'
        elif tok.startswith('_'):
            val = tok[2:-1] if tok.startswith('_{') else tok[1:]
            r = p.add_run(val)
            r.font.subscript = True
            r.font.name = 'Cambria Math'
        else:
            sub_tokens = re.findall(r'[a-zA-Z]+|[0-9]+|[^a-zA-Z0-9\s]+|\s+', tok)
            for st in sub_tokens:
                r = p.add_run(st)
                r.font.name = 'Cambria Math'
                if re.fullmatch(r'[a-zA-Z]', st):
                    r.italic = True

def render_inline_math(p, expr):
    """Prioritize native Word equation support: simple expressions use text runs, complex expressions use OMML."""
    if is_complex_math(expr):
        try:
            omml_str = latex_to_omml(expr)
            omath = parse_xml(omml_str)
            p._p.append(omath)
            return
        except Exception:
            pass
    try:
        render_simple_inline_math(p, expr)
        return
    except Exception:
        pass
    # PNG fallback if native runs and OMML fail
    try:
        import matplotlib.pyplot as plt
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
            path = Path(f.name)
        f.close()
        try:
            fig, ax = plt.subplots(figsize=(0.01, 0.01), dpi=240)
            ax.axis('off')
            ax.text(0.5, 0.5, f'${expr}$', fontsize=11, ha='center', va='center')
            fig.savefig(path, transparent=False, facecolor='white', bbox_inches='tight', pad_inches=0.01)
            plt.close(fig)
            r = p.add_run()
            r.add_picture(str(path), height=Pt(11))
            return
        finally:
            path.unlink(missing_ok=True)
    except Exception:
        r = p.add_run(expr)
        r.italic = True
        r.font.name = 'Cambria Math'

def render_display_math(d, expr):
    """Render centered display equation using native OMML with centered paragraph alignment."""
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    try:
        omml_body = latex_to_omml(expr)
        omml_para = f'<m:oMathPara {nsdecls("m")}>{omml_body}</m:oMathPara>'
        p._p.append(parse_xml(omml_para))
        return
    except Exception:
        pass
    # PNG fallback for display equation
    try:
        import matplotlib.pyplot as plt
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
            path = Path(f.name)
        f.close()
        try:
            fig, ax = plt.subplots(figsize=(6, 1.0), dpi=220)
            ax.axis('off')
            ax.text(0.5, 0.5, f'${expr}$', fontsize=16, ha='center', va='center')
            fig.savefig(path, transparent=False, facecolor='white', bbox_inches='tight', pad_inches=0.08)
            plt.close(fig)
            pic = p.add_run().add_picture(str(path))
            if pic.width > Inches(5.5):
                pic.width = Inches(5.5)
            return
        finally:
            path.unlink(missing_ok=True)
    except Exception:
        r = p.add_run(expr)
        r.italic = True
        r.font.name = 'Cambria Math'
        r.font.size = Pt(14)

def field(p, code, placeholder=None):
    """Insert Word field code with separate runs and optional placeholder text."""
    for kind, text in [('begin', None), ('instr', code), ('separate', None)]:
        r = p.add_run()
        e = OxmlElement('w:fldChar' if kind != 'instr' else 'w:instrText')
        if kind == 'instr':
            e.set(qn('xml:space'), 'preserve')
            e.text = ' ' + text + ' '
        else:
            e.set(qn('w:fldCharType'), kind)
        r._r.append(e)
    if placeholder:
        r_ph = p.add_run(placeholder)
        r_ph.font.name = 'Calibri'
        r_ph.font.size = Pt(10)
        r_ph.font.italic = True
        r_ph.font.color.rgb = MUTED
    r_end = p.add_run()
    e_end = OxmlElement('w:fldChar')
    e_end.set(qn('w:fldCharType'), 'end')
    r_end._r.append(e_end)

def add_num_instance(doc, abstract_num_id=7, start=1):
    """Create a decoupled numbering sequence starting at the given start index."""
    try:
        numbering = doc.part.numbering_part.element
        existing_nums = [
            int(n.get(qn('w:numId')))
            for n in numbering.findall(qn('w:num'))
            if n.get(qn('w:numId')) is not None and n.get(qn('w:numId')).isdigit()
        ]
        new_num_id = max(existing_nums, default=0) + 1
        new_num = parse_xml(f'''
            <w:num {nsdecls("w")} w:numId="{new_num_id}">
                <w:abstractNumId w:val="{abstract_num_id}"/>
                <w:lvlOverride w:ilvl="0">
                    <w:startOverride w:val="{start}"/>
                </w:lvlOverride>
            </w:num>
        ''')
        numbering.append(new_num)
        return new_num_id
    except Exception:
        return None

def set_list_number(p, num_id):
    """Assign specific numbering instance to paragraph."""
    if num_id is not None:
        pPr = p._p.get_or_add_pPr()
        numPr = OxmlElement('w:numPr')
        ilvl = OxmlElement('w:ilvl')
        ilvl.set(qn('w:val'), '0')
        num_elem = OxmlElement('w:numId')
        num_elem.set(qn('w:val'), str(num_id))
        numPr.append(ilvl)
        numPr.append(num_elem)
        pPr.append(numPr)

def make_doc(title, size='A4', header=None, footer=None):
    d = Document()
    for s in d.sections:
        s.page_width, s.page_height = (Mm(210), Mm(297)) if size == 'A4' else (Inches(8.5), Inches(11))
        s.top_margin = s.bottom_margin = Inches(1)
        s.left_margin = Inches(1.15)
        s.right_margin = Inches(1)
        s.different_first_page_header_footer = True
        p = s.header.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(header or title)
        r.font.name = 'Calibri'
        r.font.size = Pt(9)
        r.font.italic = True
        r.font.color.rgb = MUTED
        p = s.footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if footer:
            p.add_run(footer + ' | ')
        field(p, 'PAGE')
    st = d.styles['Normal']
    st.font.name = 'Georgia'
    st.font.size = Pt(11)
    st.paragraph_format.line_spacing = 1.15
    st.paragraph_format.space_after = Pt(6)
    for n, sz, col in [('Heading 1', 16, NAVY), ('Heading 2', 13, SLATE), ('Heading 3', 11, SLATE)]:
        st = d.styles[n]
        st.font.name = 'Calibri'
        st.font.size = Pt(sz)
        st.font.bold = True
        st.font.color.rgb = col
        st.paragraph_format.keep_with_next = True
    return d

def cover(d, title, author, institution, subtitle, date):
    for _ in range(3):
        d.add_paragraph()
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.bold = True
    r.font.name = 'Calibri'
    r.font.size = Pt(30)
    r.font.color.rgb = NAVY
    for value in (subtitle, author, institution, date):
        if value:
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(value)
            r.font.size = Pt(13)
            r.font.color.rgb = SLATE
    d.add_page_break()

def inline(p, text):
    """Format inline markdown: bold, italic, inline code, and inline math."""
    tick = chr(96)
    rx = re.compile(
        r'(\*\*.+?\*\*|(?<!\*)\*[^*]+\*|' + re.escape(tick) + r'[^' + re.escape(tick) + r']+' + re.escape(tick) + r'|\$[^$]+\$)'
    )
    pos = 0
    for m in rx.finditer(text):
        p.add_run(text[pos:m.start()])
        x = m.group()
        if x.startswith('**'):
            p.add_run(x[2:-2]).bold = True
        elif x.startswith(tick):
            r = p.add_run(x[1:-1])
            r.font.name = 'Consolas'
        elif x.startswith('$') and x.endswith('$') and len(x) > 2:
            render_inline_math(p, x[1:-1])
        else:
            p.add_run(x[1:-1]).italic = True
        pos = m.end()
    p.add_run(text[pos:])

def callout(d, title, body):
    t = d.add_table(rows=1, cols=1)
    t.style = 'Light Shading Accent 1'
    c = t.cell(0, 0)
    p = c.paragraphs[0]
    r = p.add_run(title)
    r.bold = True
    r.font.color.rgb = NAVY
    inline(c.add_paragraph(), body)

def _set_cell_shading(cell, color_hex):
    """Apply background shading to a table cell via OpenXML."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color_hex)
    shading.set(qn('w:val'), 'clear')
    shading.set(qn('w:color'), 'auto')
    tcPr.append(shading)

def _set_cell_borders(cell, color_hex='D0D0D0', size='4'):
    """Apply subtle borders to all four sides of a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), size)
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color_hex)
        borders.append(el)
    tcPr.append(borders)

def add_code_block(d, block):
    """Render code block in a shaded, bordered single-cell table with monospaced font.

    Short blocks (≤25 lines) keep lines together to prevent page splitting.
    Long blocks split cleanly across pages.
    """
    # Use a single-cell table for reliable background + border in all renderers
    tbl = d.add_table(rows=1, cols=1)
    tbl.style = 'Normal Table'
    # Remove any auto-width stretching issues
    tbl.autofit = True
    cell = tbl.cell(0, 0)

    # Apply light grey background and subtle border
    _set_cell_shading(cell, 'F5F5F5')
    _set_cell_borders(cell, 'D0D0D0', '4')

    # Set cell margins for consistent indentation / padding
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    mar = OxmlElement('w:tcMar')
    for side, val in [('top', '60'), ('left', '120'), ('bottom', '60'), ('right', '120')]:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:w'), val)
        el.set(qn('w:type'), 'dxa')
        mar.append(el)
    tcPr.append(mar)

    # Clear the default empty paragraph and write code
    p = cell.paragraphs[0]
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    if len(block) <= 25:
        p.paragraph_format.keep_together = True
    else:
        p.paragraph_format.keep_together = False
    r = p.add_run('\n'.join(block))
    r.font.name = 'Consolas'
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0x2D, 0x2D, 0x2D)

    # Add spacing around the table
    tbl_elem = tbl._tbl
    tblPr = tbl_elem.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl_elem.insert(0, tblPr)
    # Set table width to 100%
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    tblPr.append(tblW)

    return tbl

def render(d, text, base, chapter, figs, refs):
    lines = text.splitlines()
    i = 0
    code = False
    block = []
    current_num_id = None

    while i < len(lines):
        s = lines[i].strip()
        if s.startswith(chr(96) * 3):
            current_num_id = None
            if code:
                add_code_block(d, block)
                block = []
                code = False
            else:
                code = True
            i += 1
            continue
        if code:
            block.append(lines[i])
            i += 1
            continue
        if not s:
            current_num_id = None
            i += 1
            continue
        if re.match(r'^#{1,2}\s+References$', s, re.I):
            current_num_id = None
            refs.extend(re.findall(r'^\s*(?:[-*+] |\d+[.)] )(.+)$', '\n'.join(lines[i + 1:]), re.M))
            break
        if s == '$$':
            current_num_id = None
            eq = []
            i += 1
            while i < len(lines) and lines[i].strip() != '$$':
                eq.append(lines[i].strip())
                i += 1
            if i < len(lines):
                i += 1
            render_display_math(d, ' '.join(eq))
            continue
        if s.startswith('$$') and s.endswith('$$') and len(s) > 4:
            current_num_id = None
            render_display_math(d, s[2:-2])
            i += 1
            continue
        h = re.match(r'^(#{1,6})\s+(.+)', s)
        if h:
            current_num_id = None
            count = len(h.group(1))
            level = min(count, 3)
            p = d.add_paragraph(h.group(2), style='Heading ' + str(level))
            if count == 1 and len(d.paragraphs) > 1:
                p.paragraph_format.page_break_before = True
            i += 1
            continue
        im = re.fullmatch(r'!\[([^]]*)\]\(([^)]+)\)', s)
        if im:
            current_num_id = None
            figs[0] += 1
            cap, src = im.groups()
            path = Path(src)
            if not path.is_absolute():
                path = base / path
            if not path.is_file():
                raise FileNotFoundError('Missing figure: ' + str(path))
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.keep_with_next = True
            p.add_run().add_picture(str(path), width=Inches(5))
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(f'Figure {chapter}.{figs[0]}: ')
            r.bold = True
            p.add_run(cap)
            i += 1
            continue
        if s.startswith('>'):
            current_num_id = None
            q = s[1:].strip()
            m = re.match(r'\*\*(.+?)\*\*:?\s*(.*)', q)
            callout(d, m.group(1), m.group(2) or m.group(1)) if m else callout(d, 'Note', q)
            i += 1
            continue
        if s.startswith('|') and i + 1 < len(lines) and re.match(r'^\s*\|?\s*:?-{3,}', lines[i + 1]):
            current_num_id = None
            rows = [s, lines[i + 1].strip()]
            i += 2
            while i < len(lines) and '|' in lines[i]:
                rows.append(lines[i].strip())
                i += 1
            rows = [[x.strip() for x in row.strip('|').split('|')] for row in rows if not re.match(r'^\s*\|?\s*:?-{3,}', row)]
            if rows:
                t = d.add_table(rows=1, cols=max(map(len, rows)))
                t.style = 'Light Shading Accent 1'
                for ri, row in enumerate(rows):
                    cells = t.rows[0].cells if ri == 0 else t.add_row().cells
                    for ci, x in enumerate(row):
                        cells[ci].text = ''
                        inline(cells[ci].paragraphs[0], x)
                        if ri == 0:
                            for r in cells[ci].paragraphs[0].runs:
                                r.bold = True
            continue
        m = re.match(r'^[-*+]\s+(.+)', s)
        if m:
            current_num_id = None
            inline(d.add_paragraph(style='List Bullet'), m.group(1))
            i += 1
            continue
        m = re.match(r'^\d+[.)]\s+(.+)', s)
        if m:
            if current_num_id is None:
                current_num_id = add_num_instance(d, abstract_num_id=7, start=1)
            p = d.add_paragraph(style='List Number')
            set_list_number(p, current_num_id)
            inline(p, m.group(1))
            i += 1
            continue
        current_num_id = None
        p = d.add_paragraph()
        inline(p, s)
        i += 1
    if block:
        add_code_block(d, block)

def build(inputs, out, title=None, author=None, institution=None, subtitle=None, date=None, size='A4', header=None, footer=None, toc=True, overwrite=False):
    files = [Path(x).expanduser().resolve() for x in inputs]
    out = Path(out).expanduser().resolve()
    if not files or any(not f.is_file() or f.suffix.lower() != '.md' for f in files):
        raise ValueError('Provide existing Markdown (.md) chapter files.')
    if out.exists() and not overwrite:
        raise FileExistsError(str(out) + ' exists; pass --overwrite to replace.')
    title = title or files[0].stem.replace('_', ' ').title()
    d = make_doc(title, size, header, footer)
    d.core_properties.title = title
    if author:
        d.core_properties.author = author
    cover(d, title, author, institution, subtitle, date)
    if toc:
        d.add_paragraph('Table of Contents', style='Heading 1')
        p = d.add_paragraph()
        field(p, 'TOC \\o "1-3" \\h \\z \\u', placeholder='Table of Contents (Right-click and select "Update Field" in Word to populate)')
        d.add_page_break()
    refs = []
    figs = [0]
    for n, f in enumerate(files, 1):
        if n > 1:
            d.add_page_break()
        render(d, f.read_text(encoding='utf-8'), f.parent, n, figs, refs)
    if refs:
        d.add_paragraph('References', style='Heading 1')
        ref_num_id = add_num_instance(d, abstract_num_id=7, start=1)
        for ref in refs:
            p = d.add_paragraph(style='List Number')
            set_list_number(p, ref_num_id)
            inline(p, ref)
    upd = OxmlElement('w:updateFields')
    upd.set(qn('w:val'), 'true')
    d.settings.element.append(upd)
    out.parent.mkdir(parents=True, exist_ok=True)
    d.save(str(out))
    if not out.is_file() or out.stat().st_size == 0:
        raise RuntimeError('DOCX output missing or empty.')
    Document(str(out))
    return out

def main(argv=None):
    p = argparse.ArgumentParser(description='Build a textbook DOCX from one or more Markdown chapters.')
    p.add_argument('inputs', nargs='+', type=Path)
    p.add_argument('-o', '--output', required=True, type=Path)
    for flag in ('title', 'subtitle', 'author', 'institution', 'date', 'header-text', 'footer-text'):
        p.add_argument('--' + flag)
    p.add_argument('--page-size', choices=('A4', 'LETTER'), default='A4')
    p.add_argument('--no-toc', action='store_true')
    p.add_argument('--overwrite', action='store_true')
    p.add_argument('--version', action='version', version='textbook-docx 2.4.0')
    a = p.parse_args(argv)
    try:
        result = build(a.inputs, a.output, a.title, a.author, a.institution, a.subtitle, a.date, a.page_size, a.header_text, a.footer_text, not a.no_toc, a.overwrite)
    except (OSError, ValueError, RuntimeError) as e:
        print('error:', e, file=sys.stderr)
        return 2
    print('Created and validated:', result)
    if not a.no_toc:
        print('If needed, update the TOC field in Word (right-click > Update Field).')
    return 0

if __name__ == '__main__':
    raise SystemExit(main())
