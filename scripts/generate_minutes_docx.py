#!/usr/bin/env python3
"""
generate_minutes_docx.py: branded meeting-minutes document generator

Turns a structured plain-text file into a Word document with a fixed house style:
  - 1-inch margins all around
  - Two-column metadata header table (primary left / light tint right)
  - Heading 1: bold 14pt primary colour with primary bottom-border divider
  - Heading 2: bold 12pt dark accent
  - Body: Calibri 11pt, 30pt spacing before/after
  - Three-column action items table (Owner | Action | Deadline)
    with a primary-colour header row and alternating white/tint data rows

Input file format:
    ##TITLE## MEETING MINUTES
    ##SUBTITLE## Organization, Topic
    ##META##
    Date: Monday, March 2, 2026
    Time: 11:00 AM to 12:00 PM
    Location: Teams/Zoom
    Participants: Name One, Name Two, Name Three
    ##SUMMARY## Executive summary text (omit line entirely if no summary)
    ##SECTION## Section Header
    Prose content...
    ##SECTION## Another Section Header
    More prose...
    ##ACTIONITEMS##
    - Action description — Owner — Deadline
    - Another action — Owner TBD
    ##BYLINE## Prepared by Jordan Lee, Executive Assistant, May 12, 2026

Usage:
    python3 generate_minutes_docx.py --input minutes_input.txt --output "Meeting Minutes.docx"
"""

import argparse
import re
import sys
from pathlib import Path
from copy import deepcopy

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches, Emu
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("python-docx not installed. Run: pip install python-docx --break-system-packages")
    sys.exit(1)


# ── Brand palette: swap these three values for your organization's colours ───
PRIMARY  = "1F3A5F"   # Heading 1, title, table header fill, dividers
PRIMARY_DARK     = "142840"   # Heading 2
TINT_LIGHT   = "EAF0F7"   # Metadata table right column, alternating action rows
GRAY_TEXT    = "5C5A56"   # Subtitle
GRAY_BORDER  = "CCCCCC"   # Action table data row borders
WHITE        = "FFFFFF"
NEAR_BLACK   = "2B2B2B"   # Body text

def hex_to_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ── XML helpers ───────────────────────────────────────────────────────────────

def set_cell_shading(cell, fill_hex):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  fill_hex)
    # Remove existing shd if present
    for existing in tcPr.findall(qn("w:shd")):
        tcPr.remove(existing)
    tcPr.append(shd)


def set_cell_borders(cell, color_hex, sz="1"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for existing in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(existing)
    tcBdr = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   "single")
        el.set(qn("w:sz"),    sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color_hex)
        tcBdr.append(el)
    tcPr.append(tcBdr)


def override_cell_bottom_border(cell, color_hex, sz="6"):
    """Override only the bottom border of a cell, leaving other borders intact.
    Used in add_metadata_table to make row dividers visible against each column's
    background: white on primary (left col), primary on tint (right col).
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBdr = tcPr.find(qn("w:tcBorders"))
    if tcBdr is None:
        tcBdr = OxmlElement("w:tcBorders")
        tcPr.append(tcBdr)
    for existing in tcBdr.findall(qn("w:bottom")):
        tcBdr.remove(existing)
    el = OxmlElement("w:bottom")
    el.set(qn("w:val"),   "single")
    el.set(qn("w:sz"),    sz)
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), color_hex)
    tcBdr.append(el)


def set_cell_width(cell, width_dxa):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for existing in tcPr.findall(qn("w:tcW")):
        tcPr.remove(existing)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"),    str(width_dxa))
    tcW.set(qn("w:type"), "dxa")
    tcPr.append(tcW)


# Table geometry. Word reads <w:tblInd> two ways depending on the document's
# compatibility mode: up to Word 2010 (mode 14, what python-docx's template
# declares) the border lands at tblInd minus the first cell's left margin;
# from Word 2013 (mode 15, every new Word document, and any file Word
# converts) the border lands at tblInd itself. The 1.5.4 fix, tblInd=108 over
# the default 108 left cell margin, was right in mode 14 only: in mode 15, or
# pasted into a new document, every table band sat 108 dxa inside the margin.
# Measured 2026-09-24 in Word at 150 dpi: margin 150 px, tables 161 px.
# Zero indent over a zero left cell margin puts the border on the margin in
# both modes; the text inset the cell margin used to give comes back as
# paragraph indent (CELL_LEFT_PAD_PT), so cell text does not move.
TABLE_INDENT_DXA = 0
CELL_LEFT_MARGIN_DXA = 0
CELL_RIGHT_MARGIN_DXA = 108
CELL_LEFT_PAD_PT = 5.4   # the 108 dxa left cell margin this replaces

# Schema order of <w:tblPr> children (CT_TblPr). Elements appended out of
# order, e.g. tblW after tblLook, are invalid OOXML that renderers tolerate
# today and are free to stop tolerating.
_TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
                "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
                "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
                "tblCellMar", "tblLook", "tblCaption", "tblDescription"]


def _tblpr_insert(tblPr, el):
    """Insert el into tblPr at its schema position."""
    rank = _TBLPR_ORDER.index(el.tag.split("}")[1])
    for child in tblPr:
        tag = child.tag.split("}")[1]
        if tag in _TBLPR_ORDER and _TBLPR_ORDER.index(tag) > rank:
            child.addprevious(el)
            return
    tblPr.append(el)


def cell_indent(pt):
    """Paragraph left indent for text in a table cell: the visual inset pt,
    plus the padding the zeroed left cell margin no longer supplies."""
    return Pt(pt + CELL_LEFT_PAD_PT)


def set_table_dimensions(tblPr, col_widths):
    """
    Set the table's overall width AND its tblGrid to match the real per-cell
    column widths (col_widths, a list of dxa ints, one per column).

    python-docx's doc.add_table() seeds tblPr with a default
    <w:tblW w:type="auto" w:w="0"/> and a tblGrid that splits the table width
    evenly across columns. Every table builder in this file was appending a
    second, correct <w:tblW> without removing that default — leaving two
    conflicting tblW elements in tblPr, which is invalid OOXML. Renderers
    disagree on which one wins: LibreOffice used the appended (correct) one,
    masking the bug in every preview; Word used the first (auto/0) one
    combined with the still-untouched, evenly-split tblGrid, which visibly
    widened the left column with blank space. Always call this instead of
    hand-rolling a tblW element.
    """
    for existing in tblPr.findall(qn("w:tblW")):
        tblPr.remove(existing)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"),    str(sum(col_widths)))
    tblW.set(qn("w:type"), "dxa")
    _tblpr_insert(tblPr, tblW)

    # Table indent and cell margins, see TABLE_INDENT_DXA. With no explicit
    # <w:tblInd> the table hung 108 dxa into the left margin (2026-09-10,
    # 1.5.4); with tblInd=108 it sat 108 dxa inside it in Word 2013+ mode
    # (2026-09-24). Both are set explicitly so neither default applies.
    for tag in ("w:tblInd", "w:tblCellMar"):
        for existing in tblPr.findall(qn(tag)):
            tblPr.remove(existing)
    tblInd = OxmlElement("w:tblInd")
    tblInd.set(qn("w:w"),    str(TABLE_INDENT_DXA))
    tblInd.set(qn("w:type"), "dxa")
    _tblpr_insert(tblPr, tblInd)
    tblCellMar = OxmlElement("w:tblCellMar")
    for side, w in (("top", 0), ("left", CELL_LEFT_MARGIN_DXA),
                    ("bottom", 0), ("right", CELL_RIGHT_MARGIN_DXA)):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"),    str(w))
        m.set(qn("w:type"), "dxa")
        tblCellMar.append(m)
    _tblpr_insert(tblPr, tblCellMar)

    tbl_el = tblPr.getparent()
    for existing in tbl_el.findall(qn("w:tblGrid")):
        tbl_el.remove(existing)
    tblGrid = OxmlElement("w:tblGrid")
    for w in col_widths:
        gridCol = OxmlElement("w:gridCol")
        gridCol.set(qn("w:w"), str(w))
        tblGrid.append(gridCol)
    # tblGrid must immediately follow tblPr as a sibling under <w:tbl>.
    tblPr.addnext(tblGrid)


def add_bottom_divider(doc):
    """Empty paragraph with a primary bottom border — the section divider line."""
    p = doc.add_paragraph()
    add_bottom_border_to_paragraph(p)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    return p


def add_bottom_border_to_paragraph(p):
    """Apply a primary bottom-border rule directly to an existing paragraph, flush
    under its last line of text — no added whitespace between text and rule."""
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    "4")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), PRIMARY)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def set_para_spacing(p, before=0, after=0):
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)


# ── Document sections ─────────────────────────────────────────────────────────

def add_title_block(doc, title_text, subtitle_text):
    """MEETING MINUTES + subtitle line. Styled as the document Title so it is neither
    a body paragraph nor an outline heading; the run formatting overrides the look."""
    p = doc.add_paragraph()
    p.style = "Title"
    clear_para_bottom_border(p)
    set_para_spacing(p, before=0, after=4)
    run = p.add_run(title_text.upper())
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = hex_to_rgb(PRIMARY)
    run.font.name = "Calibri"

    p2 = doc.add_paragraph()
    set_para_spacing(p2, before=2, after=10)
    run2 = p2.add_run(subtitle_text)
    run2.font.size = Pt(11)
    run2.font.color.rgb = hex_to_rgb(GRAY_TEXT)
    run2.font.name = "Calibri"


def mark_header_row(row):
    """Mark a table row as a repeating header row (w:tblHeader). This is what a screen
    reader and the PDF tag tree read as the header cells; without it every cell is data."""
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))


def clear_para_bottom_border(p):
    """Override a style's bottom border on one paragraph (the default Title style draws one)."""
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "nil")
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_metadata_table(doc, meta_rows):
    """
    Two-column table: primary left column (labels), light-tint right column (values).
    meta_rows: list of (label, value) tuples.
    Total width 9360 dxa. Left col 2200, right col 7160.
    """
    if not meta_rows:
        return

    LEFT_W  = 2200
    RIGHT_W = 7160

    tbl = doc.add_table(rows=len(meta_rows), cols=2)

    mark_header_row(tbl.rows[0])
    tbl.style = "Table Grid"

    # Set table width via tblPr
    tblPr = tbl._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl._tbl.insert(0, tblPr)
    set_table_dimensions(tblPr, [LEFT_W, RIGHT_W])

    # Remove default table borders
    for existing in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(existing)

    for i, (label, value) in enumerate(meta_rows):
        row = tbl.rows[i]
        left_cell  = row.cells[0]
        right_cell = row.cells[1]

        # Widths
        set_cell_width(left_cell,  LEFT_W)
        set_cell_width(right_cell, RIGHT_W)

        # Shading
        set_cell_shading(left_cell,  PRIMARY)
        set_cell_shading(right_cell, TINT_LIGHT)

        # Borders — hairline primary on all sides
        set_cell_borders(left_cell,  PRIMARY, sz="1")
        set_cell_borders(right_cell, PRIMARY, sz="1")

        # Left cell: bold white label
        lp = left_cell.paragraphs[0]
        set_para_spacing(lp, before=4, after=4)
        lp.paragraph_format.left_indent = cell_indent(6)
        lr = lp.add_run(label)
        lr.bold = True
        lr.font.color.rgb = hex_to_rgb(WHITE)
        lr.font.size = Pt(10)
        lr.font.name = "Calibri"

        # Right cell: plain text on tint
        rp = right_cell.paragraphs[0]
        set_para_spacing(rp, before=4, after=4)
        rp.paragraph_format.left_indent = cell_indent(6)
        rr = rp.add_run(value)
        rr.font.size = Pt(10)
        rr.font.name = "Calibri"
        rr.font.color.rgb = hex_to_rgb(NEAR_BLACK)

    # Row-divider border fix:
    # set_cell_borders applies one color to all four sides, making the bottom
    # border invisible in the left (primary) column. Post-process all intermediate
    # rows: white divider on the primary fill (left), primary divider on the tint (right).
    total_rows = len(meta_rows)
    for i in range(total_rows - 1):
        override_cell_bottom_border(tbl.rows[i].cells[0], WHITE,       sz="6")
        override_cell_bottom_border(tbl.rows[i].cells[1], PRIMARY, sz="6")

    # Spacing after table
    doc.add_paragraph().paragraph_format.space_after = Pt(6)


def add_heading1(doc, text):
    """Heading 1: bold 14pt primary colour, with the primary divider on the same paragraph
    (no separate empty line) so the rule sits flush under the heading text.
    Styled as a real Heading 1 so Word's outline, a screen reader's heading
    navigation and the PDF tag tree all see it; the run formatting below
    overrides the style's look (2026-09-17: every minutes file before this
    had no outline at all, caught by a visual-heading check)."""
    p = doc.add_paragraph()
    p.style = "Heading 1"
    set_para_spacing(p, before=14, after=4)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = hex_to_rgb(PRIMARY)
    run.font.name = "Calibri"
    add_bottom_border_to_paragraph(p)


def add_heading2(doc, text):
    """Heading 2: bold 12pt dark accent. A real Heading 2 style, see add_heading1."""
    p = doc.add_paragraph()
    p.style = "Heading 2"
    set_para_spacing(p, before=10, after=2)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = hex_to_rgb(PRIMARY_DARK)
    run.font.name = "Calibri"


def _is_bullet(line):
    """True if line starts with a bullet/dash marker."""
    s = line.lstrip()
    return s.startswith("- ") or s.startswith("* ") or s.startswith("• ")

def _is_nested_bullet(line):
    """True if line is indented AND starts with a bullet/dash marker."""
    return (line.startswith("  ") or line.startswith("\t")) and _is_bullet(line)

def _strip_bullet(line):
    """Remove leading bullet/dash marker from a line."""
    s = line.lstrip()
    for prefix in ("- ", "* ", "• "):
        if s.startswith(prefix):
            return s[len(prefix):]
    return s

_ITALIC_SPAN_RE = re.compile(r'_([^_]+)_')
_EMPH_SPAN_RE = re.compile(r'\*\*(.+?)\*\*|_([^_]+)_', re.S)

def _styled_run(p, text, size, bold=False, italic=False):
    """One body run in the standard font, optionally bold and/or italic."""
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.name = "Calibri"
    run.font.color.rgb = hex_to_rgb(NEAR_BLACK)
    if bold:
        run.bold = True
    if italic:
        run.italic = True
    return run

def _add_text_runs(p, text, size=11):
    """
    Add text to a paragraph as one or more runs. **wrapped** spans render bold,
    _wrapped_ spans render italic, and italics nest inside a bold span so that
    **bold with _emphasis_** renders bold-italic. Neither marker nests in
    itself. Everything else renders as a normal run in the standard body
    font/size/color. Used by both prose paragraphs and bullets so inline
    emphasis works in either.
    """
    pos = 0
    for m in _EMPH_SPAN_RE.finditer(text):
        if m.start() > pos:
            _styled_run(p, text[pos:m.start()], size)
        if m.group(1) is not None:
            inner = m.group(1)
            ipos = 0
            for im in _ITALIC_SPAN_RE.finditer(inner):
                if im.start() > ipos:
                    _styled_run(p, inner[ipos:im.start()], size, bold=True)
                _styled_run(p, im.group(1), size, bold=True, italic=True)
                ipos = im.end()
            if ipos < len(inner):
                _styled_run(p, inner[ipos:], size, bold=True)
        else:
            _styled_run(p, m.group(2), size, italic=True)
        pos = m.end()
    if pos < len(text):
        _styled_run(p, text[pos:], size)

def _add_prose_para(doc, text):
    """Single prose paragraph. **wrapped** renders bold, _wrapped_ italic."""
    p = doc.add_paragraph()
    set_para_spacing(p, before=4, after=4)
    _add_text_runs(p, text)

def _add_bullet_para(doc, text, nested=False):
    """Bullet using Word List Bullet style. **wrapped** bold, _wrapped_ italic."""
    style = "List Bullet 2" if nested else "List Bullet"
    try:
        p = doc.add_paragraph(style=style)
    except Exception:
        p = doc.add_paragraph()
    set_para_spacing(p, before=2, after=2)
    if nested:
        p.paragraph_format.left_indent = Inches(0.5)
    _add_text_runs(p, text)

def add_body_paragraph(doc, text):
    """
    Render a body block, detecting structure line by line:
    - Lines starting with '- ', '* ', or '• ' → bullet points
    - Lines indented AND starting with those markers → nested bullets
    - Everything else → prose paragraphs (consecutive non-list lines are joined)
    """
    if not text.strip():
        return

    lines = text.splitlines()
    prose_buffer = []

    def flush_prose():
        if prose_buffer:
            combined = " ".join(l.strip() for l in prose_buffer if l.strip())
            if combined:
                _add_prose_para(doc, combined)
            prose_buffer.clear()

    for line in lines:
        if not line.strip():
            # Blank line: flush prose buffer as its own paragraph
            flush_prose()
        elif _is_nested_bullet(line):
            flush_prose()
            _add_bullet_para(doc, _strip_bullet(line), nested=True)
        elif _is_bullet(line):
            flush_prose()
            _add_bullet_para(doc, _strip_bullet(line), nested=False)
        else:
            # Regular prose line — accumulate
            prose_buffer.append(line)

    flush_prose()


def add_breakdown_table(doc, caption, rows):
    """
    Stylized breakdown table matching the Action Items table: a primary header row
    (the horizontal bar) with bold white text, and alternating white/tint data
    rows. rows[0] is the header row; rows[1:] are data rows. Used for ##TABLE##
    blocks.
    """
    if not rows:
        return

    if caption:
        cp = doc.add_paragraph()
        set_para_spacing(cp, before=6, after=2)
        cr = cp.add_run(caption)
        cr.bold = True
        cr.font.size = Pt(11)
        cr.font.color.rgb = hex_to_rgb(PRIMARY_DARK)
        cr.font.name = "Calibri"

    n_data_cols = max(len(r) for r in rows)
    n_rows = len(rows)

    # Full margin-to-margin width, matching the metadata and action-items
    # tables (6.5in usable width at 1in margins = 9360 dxa). Previously this
    # was hardcoded to 5600 dxa total, leaving a gap on the right instead of
    # spanning the page.
    TABLE_FULL_WIDTH = 9360
    if n_data_cols == 2:
        # Preserve the original 3:1 label/value proportion, scaled to full width.
        data_widths = [int(TABLE_FULL_WIDTH * 3 / 4), TABLE_FULL_WIDTH - int(TABLE_FULL_WIDTH * 3 / 4)]
    else:
        base = TABLE_FULL_WIDTH // n_data_cols
        data_widths = [base] * n_data_cols
        data_widths[-1] += TABLE_FULL_WIDTH - base * n_data_cols  # remainder to last col

    tbl = doc.add_table(rows=n_rows, cols=n_data_cols)

    mark_header_row(tbl.rows[0])
    tbl.style = "Table Grid"

    tblPr = tbl._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl._tbl.insert(0, tblPr)
    set_table_dimensions(tblPr, data_widths)
    for existing in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(existing)

    for i in range(n_rows):
        row = tbl.rows[i]
        is_header = (i == 0)
        fill = PRIMARY if is_header else (WHITE if (i % 2 == 1) else TINT_LIGHT)
        cells_data = list(rows[i]) + [""] * (n_data_cols - len(rows[i]))
        for j in range(n_data_cols):
            cell = row.cells[j]
            set_cell_width(cell, data_widths[j])
            set_cell_shading(cell, fill)
            set_cell_borders(cell, PRIMARY if is_header else GRAY_BORDER, sz="1")
            p = cell.paragraphs[0]
            set_para_spacing(p, before=3, after=3)
            p.paragraph_format.left_indent = cell_indent(4)
            if n_data_cols == 2 and j == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(cells_data[j])
            run.font.size = Pt(10)
            run.font.name = "Calibri"
            if is_header:
                run.bold = True
                run.font.color.rgb = hex_to_rgb(WHITE)
            else:
                run.font.color.rgb = hex_to_rgb(NEAR_BLACK)
                if j == 0:
                    run.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(6)


def add_action_items_table(doc, items):
    """
    Three-column table: Owner | Action | Deadline
    Primary header row, alternating white/tint data rows.
    Column widths: 1979 | 5581 | 1800 dxa
    """
    OWNER_W    = 1979
    ACTION_W   = 5581
    DEADLINE_W = 1800

    # Section label above table: a real Heading 1, like every other section
    p = doc.add_paragraph()
    p.style = "Heading 1"
    set_para_spacing(p, before=14, after=2)
    run = p.add_run("Action Items")
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = hex_to_rgb(PRIMARY)
    run.font.name = "Calibri"
    add_bottom_border_to_paragraph(p)

    if not items:
        p2 = doc.add_paragraph()
        set_para_spacing(p2, before=4, after=4)
        r = p2.add_run("No specific action items were identified. Review source material for any follow-ups.")
        r.italic = True
        r.font.size = Pt(10)
        r.font.color.rgb = hex_to_rgb(GRAY_TEXT)
        r.font.name = "Calibri"
        return

    total_rows = 1 + len(items)  # header + data
    tbl = doc.add_table(rows=total_rows, cols=3)
    mark_header_row(tbl.rows[0])
    tbl.style = "Table Grid"

    col_widths = [OWNER_W, ACTION_W, DEADLINE_W]

    # Set table width
    tblPr = tbl._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl._tbl.insert(0, tblPr)
    set_table_dimensions(tblPr, col_widths)
    for existing in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(existing)

    headers = ["Owner", "Action", "Deadline"]

    # Header row — solid primary, bold white
    hrow = tbl.rows[0]
    for j, (cell, header, w) in enumerate(zip(hrow.cells, headers, col_widths)):
        set_cell_width(cell, w)
        set_cell_shading(cell, PRIMARY)
        set_cell_borders(cell, PRIMARY, sz="1")
        p = cell.paragraphs[0]
        set_para_spacing(p, before=3, after=3)
        p.paragraph_format.left_indent = cell_indent(4)
        run = p.add_run(header)
        run.bold = True
        run.font.color.rgb = hex_to_rgb(WHITE)
        run.font.size = Pt(10)
        run.font.name = "Calibri"

    # Data rows — alternating white / light tint
    for i, item in enumerate(items):
        # Parse "Action description — Owner — Deadline"
        parts = [x.strip() for x in item.split(" — ")]
        if len(parts) >= 3:
            action, owner, deadline = parts[0], parts[1], parts[2]
        elif len(parts) == 2:
            action, owner, deadline = parts[0], parts[1], "TBD"
        else:
            action, owner, deadline = parts[0], "Owner TBD", "TBD"

        fill = WHITE if (i % 2 == 0) else TINT_LIGHT
        row = tbl.rows[i + 1]
        values = [owner, action, deadline]

        for j, (cell, val, w) in enumerate(zip(row.cells, values, col_widths)):
            set_cell_width(cell, w)
            set_cell_shading(cell, fill)
            set_cell_borders(cell, GRAY_BORDER, sz="1")
            p = cell.paragraphs[0]
            set_para_spacing(p, before=3, after=3)
            p.paragraph_format.left_indent = cell_indent(4)
            run = p.add_run(val)
            run.font.size = Pt(10)
            run.font.name = "Calibri"
            run.font.color.rgb = hex_to_rgb(NEAR_BLACK)
            # Bold the owner cell
            if j == 0:
                run.bold = True


def add_byline(doc, text):
    """Right-aligned italic small text, preceded by a thin primary rule."""
    add_bottom_divider(doc)
    p = doc.add_paragraph()
    set_para_spacing(p, before=4, after=0)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(text)
    run.italic = True
    run.font.size = Pt(9)
    run.font.color.rgb = hex_to_rgb(GRAY_TEXT)
    run.font.name = "Calibri"


# ── Parser ────────────────────────────────────────────────────────────────────

def parse_content(text):
    """
    Returns a list of (marker, content) tuples.
    Markers: TITLE, SUBTITLE, META_ROW, SUMMARY_BODY, SECTION_HEADER,
             SECTION_BODY, ACTION_ITEM, BYLINE
    """
    blocks = []
    current_mode = None
    body_lines = []
    table_rows = []
    table_caption = [None]  # boxed so the nested flush can rebind without nonlocal

    def flush_body(mode):
        if body_lines:
            body_text = "\n".join(body_lines).strip()
            if body_text:
                if mode == "SUMMARY":
                    blocks.append(("SUMMARY_BODY", body_text))
                elif mode == "SECTION":
                    blocks.append(("SECTION_BODY", body_text))
        body_lines.clear()

    def flush_table():
        if table_rows:
            blocks.append(("TABLE", (table_caption[0], [list(r) for r in table_rows])))
            table_rows.clear()
        table_caption[0] = None

    for raw_line in text.splitlines():
        line = raw_line.strip()

        if line.startswith("##TITLE##"):
            flush_body(current_mode)
            flush_table()
            current_mode = None
            blocks.append(("TITLE", line[len("##TITLE##"):].strip()))

        elif line.startswith("##SUBTITLE##"):
            flush_body(current_mode)
            flush_table()
            current_mode = None
            blocks.append(("SUBTITLE", line[len("##SUBTITLE##"):].strip()))

        elif line.startswith("##META##"):
            flush_body(current_mode)
            flush_table()
            current_mode = "META"

        elif line.startswith("##SUMMARY##"):
            flush_body(current_mode)
            flush_table()
            current_mode = "SUMMARY"
            inline = line[len("##SUMMARY##"):].strip()
            if inline:
                body_lines.append(inline)

        elif line.startswith("##SECTION##"):
            flush_body(current_mode)
            flush_table()
            current_mode = "SECTION"
            blocks.append(("SECTION_HEADER", line[len("##SECTION##"):].strip()))

        elif line.startswith("##TABLE##"):
            flush_body(current_mode)
            flush_table()
            current_mode = "TABLE"
            cap = line[len("##TABLE##"):].strip()
            table_caption[0] = cap if cap else None

        elif line.startswith("##ACTIONITEMS##"):
            flush_body(current_mode)
            flush_table()
            current_mode = "ACTIONS"

        elif line.startswith("##BYLINE##"):
            flush_body(current_mode)
            flush_table()
            current_mode = None
            blocks.append(("BYLINE", line[len("##BYLINE##"):].strip()))

        else:
            # Continuation content
            if current_mode == "META":
                if ":" in line:
                    key, _, val = line.partition(":")
                    blocks.append(("META_ROW", (key.strip(), val.strip())))
            elif current_mode == "ACTIONS":
                stripped = line.lstrip("-•").strip()
                if stripped:
                    blocks.append(("ACTION_ITEM", stripped))
            elif current_mode == "TABLE":
                if "|" in line:
                    table_rows.append([c.strip() for c in line.split("|")])
            elif current_mode in ("SUMMARY", "SECTION"):
                body_lines.append(raw_line)

    flush_body(current_mode)
    flush_table()
    return blocks


# ── Main assembler ────────────────────────────────────────────────────────────

def build_document(blocks, output_path):
    doc = Document()

    # 1-inch margins all around 
    for section in doc.sections:
        section.top_margin    = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin   = Inches(1.0)
        section.right_margin  = Inches(1.0)

    # Default style
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    # Collect meta rows and action items to render as tables
    meta_rows    = []
    action_items = []

    title_text    = "MEETING MINUTES"
    subtitle_text = ""
    byline_text   = ""

    # First pass: collect metadata, title, subtitle, byline
    for marker, content in blocks:
        if marker == "TITLE":
            title_text = content
        elif marker == "SUBTITLE":
            subtitle_text = content
        elif marker == "META_ROW":
            meta_rows.append(content)
        elif marker == "ACTION_ITEM":
            action_items.append(content)
        elif marker == "BYLINE":
            byline_text = content

    # Render title block
    # the generic banner names no meeting; the subtitle line does
    generic = title_text.strip().upper() == "MEETING MINUTES"
    doc.core_properties.title = subtitle_text if generic and subtitle_text else title_text
    add_title_block(doc, title_text, subtitle_text)

    # Render metadata table immediately after title
    if meta_rows:
        add_metadata_table(doc, meta_rows)

    # Second pass: render body content in order
    pending_section_header = None
    for marker, content in blocks:
        if marker in ("TITLE", "SUBTITLE", "META_ROW", "ACTION_ITEM", "BYLINE"):
            continue  # already handled

        elif marker == "SUMMARY_BODY":
            add_heading1(doc, "Executive Summary")
            add_body_paragraph(doc, content)

        elif marker == "SECTION_HEADER":
            pending_section_header = content

        elif marker == "SECTION_BODY":
            if pending_section_header:
                # Detect subheading: if header contains a period-number prefix (e.g. "1.")
                # it's a Heading 1, otherwise check for indented/sub style
                add_heading1(doc, pending_section_header)
                pending_section_header = None
            add_body_paragraph(doc, content)

        elif marker == "TABLE":
            if pending_section_header:
                add_heading1(doc, pending_section_header)
                pending_section_header = None
            caption, rows = content
            add_breakdown_table(doc, caption, rows)

    # Flush any section header without body
    if pending_section_header:
        add_heading1(doc, pending_section_header)

    # Action items table
    add_action_items_table(doc, action_items)

    # Byline
    if byline_text:
        add_byline(doc, byline_text)

    doc.save(output_path)
    print(f"Saved: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Generate branded meeting minutes .docx")
    parser.add_argument("--input",  required=True, help="Path to structured minutes text file")
    parser.add_argument("--output", required=True, help="Output .docx path")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: input file not found: {input_path}")
        sys.exit(1)

    text = input_path.read_text(encoding="utf-8")
    blocks = parse_content(text)

    if not blocks:
        print("Error: no parseable content found in input file.")
        sys.exit(1)

    build_document(blocks, args.output)


if __name__ == "__main__":
    main()
