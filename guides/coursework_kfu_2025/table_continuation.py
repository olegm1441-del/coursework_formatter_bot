"""
Phase 3, Rule 1 — Table continuation (hybrid: LRPB + geometry fallback).

Primary signal — w:lastRenderedPageBreak (LRPB):
  Word writes these elements into the XML every time it saves after rendering.
  They mark the exact position of page breaks in the last Word render.
  For table rows: if a row contains LRPB, that row straddles a page break
  → we split just before that row.
  For body paragraphs: LRPB resets our cumulative-height tracker, preventing
  error accumulation across pages.

Geometry fallback (for tables with no LRPB in their rows):
  Estimated from font/spacing constants and cell text length.
  Used only when LRPB data is absent or insufficient.
"""

from __future__ import annotations

import copy
from copy import deepcopy  # used in _merge_tables and _split_table
import logging
import math
import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .docx_utils import xml_has_image, is_source_or_note_line, FormattingReport

logger = logging.getLogger(__name__)

# ── Unit helpers ─────────────────────────────────────────────────────────────

EMU_PER_PT  = 12700   # 1 pt  = 12 700 EMU  (python-docx stores lengths in EMU)
TWIP_PER_PT = 20      # 1 pt  = 20 twips    (w:trHeight val is in twips)

def _emu_pt(v: int) -> float: return v / EMU_PER_PT
def _twip_pt(v: int) -> float: return v / TWIP_PER_PT


# ── w:lastRenderedPageBreak helpers ──────────────────────────────────────────

_LRPB_TAG = qn("w:lastRenderedPageBreak")


def _para_has_lrpb(p_elem) -> bool:
    """True if this paragraph contains w:lastRenderedPageBreak."""
    return p_elem.find(".//" + _LRPB_TAG) is not None


def _row_lrpb_index(rows) -> int:
    """
    Return the index of the first row (after the header, index > 0) that
    contains w:lastRenderedPageBreak, or -1 if none found.
    A positive result means the page break is WITHIN that row — we should
    split just BEFORE it (split_after = index - 1).
    """
    for i, row in enumerate(rows):
        if i == 0:
            continue   # header row — never split before it
        if row._tr.find(".//" + _LRPB_TAG) is not None:
            return i
    return -1


def _all_row_lrpb_indices(rows) -> list[int]:
    """
    Return ALL row indices (after the header, index > 0) that contain
    w:lastRenderedPageBreak.  Used for tables spanning 3+ pages.

    Each index means: the page break is WITHIN that row → split just
    BEFORE it (split_after = index - 1).

    Returns an empty list if no LRPB is found after the header row.
    """
    result: list[int] = []
    for i, row in enumerate(rows):
        if i == 0:
            continue   # header row — never split before it
        if row._tr.find(".//" + _LRPB_TAG) is not None:
            result.append(i)
    return result

# ── Page geometry ─────────────────────────────────────────────────────────────

# Safety margin subtracted from body height so we don't overfill a page.
# Accounts for rounding + minor rendering differences between LO and Word.
_PAGE_BUFFER_PT = 36

# Minimum column width (pt) for column-width optimisation.
# Columns narrower than this are "phantom" or overly squeezed.
_MIN_COL_PT = 48.0   # ≈ 1.27 cm


def _body_height_pt(doc: Document) -> float:
    s = doc.sections[0]
    return _emu_pt(s.page_height - s.top_margin - s.bottom_margin) - _PAGE_BUFFER_PT


def _body_width_pt(doc: Document) -> float:
    s = doc.sections[0]
    return _emu_pt(s.page_width - s.left_margin - s.right_margin)


# ── Height estimators ─────────────────────────────────────────────────────────

# KFU body: Times New Roman 14 pt, 1.5 line spacing → ~21 pt/line
_BODY_LINE_PT  = 14 * 1.5
# Table cells: Times New Roman 12 pt, 1.0 line spacing → ~12 pt/line
_TABLE_LINE_PT = 12 * 1.0
# Empirical chars-per-line for 14 pt TNR in a 17 cm body column.
# Lowered from 68 → 62 to avoid underestimating multi-line paragraphs
# (shorter effective measure due to first-line indent + word-wrap).
_BODY_CHARS_PER_LINE = 62

# Approx pt per char for 12 pt TNR (used to derive chars-per-column)
_PT_PER_CHAR_TABLE = 6.0

# Top+bottom cell padding in pt (default Word cell margins ≈ 2.25 pt each side)
_CELL_PADDING_PT = 4.5


def _estimate_para_height(p) -> float:
    """Estimated rendered height of a body paragraph in points."""
    text = (p.text or "").strip()
    n_lines = max(1, math.ceil(len(text) / _BODY_CHARS_PER_LINE)) if text else 1

    line_h = _BODY_LINE_PT
    try:
        pf = p.paragraph_format
        ls = pf.line_spacing
        if ls is not None:
            # python-docx may return:
            #  • a Length subclass (Emu, Twips, …) with .pt for exact/atLeast rules
            #  • a plain float multiplier (e.g. 1.5) for auto rule
            #  • a raw int in 240ths-of-a-line when rule is unset (older python-docx)
            # Detection order: .pt first (handles all Length objects correctly),
            # then float multiplier, then 240ths fallback.
            # NOTE: WD_LINE_SPACING.EXACTLY == 4 (not 1), so checking int(rule)==1
            #       was wrong — we now rely on type detection instead.
            if hasattr(ls, "pt"):
                # Length object: .pt converts to points regardless of sub-type
                line_h = float(ls.pt)
            elif isinstance(ls, float):
                # Pure Python float → line spacing multiplier (e.g. 1.5)
                line_h = 14 * ls
            elif isinstance(ls, int):
                ls_i = int(ls)
                if ls_i > 10:
                    # Raw 240ths-of-a-line value (240=single, 360=1.5×, 480=double)
                    line_h = 14 * (ls_i / 240)
                else:
                    # Small integer treated as a multiplier (rare)
                    line_h = 14 * ls_i
    except Exception:
        pass

    sb = sa = 0.0
    try:
        if p.paragraph_format.space_before:
            sb = p.paragraph_format.space_before.pt
        if p.paragraph_format.space_after:
            sa = p.paragraph_format.space_after.pt
    except Exception:
        pass

    return n_lines * line_h + sb + sa


def _tbl_col_widths_pt(tbl_elem) -> list[float]:
    """
    Read actual column widths (in pt) from w:tblGrid / w:gridCol w:w (twips).
    Returns an empty list if not present.
    """
    tblGrid = tbl_elem.find(qn("w:tblGrid"))
    if tblGrid is None:
        return []
    widths = []
    for gc in tblGrid.findall(qn("w:gridCol")):
        w_val = gc.get(qn("w:w"))
        if w_val and w_val.isdigit():
            widths.append(_twip_pt(int(w_val)))
    return widths


def _cell_margins_pt(cell_elem) -> float:
    """
    Return total vertical cell margin (top + bottom) in pt from w:tcPr/w:tcMar.
    Falls back to _CELL_PADDING_PT if not specified.
    """
    tcPr = cell_elem.find(qn("w:tcPr"))
    if tcPr is None:
        return _CELL_PADDING_PT
    tcMar = tcPr.find(qn("w:tcMar"))
    if tcMar is None:
        return _CELL_PADDING_PT
    total = 0.0
    found = False
    for side in ("w:top", "w:bottom"):
        el = tcMar.find(qn(side))
        if el is not None:
            w_type = el.get(qn("w:type"), "dxa")
            val = el.get(qn("w:w"), "0")
            if val.lstrip("-").isdigit():
                if w_type == "dxa":
                    total += _twip_pt(int(val))
                elif w_type == "nil":
                    pass   # zero
            found = True
    return total if found else _CELL_PADDING_PT


def _para_font_size_pt(p_elem) -> float:
    """
    Read font size (pt) from the paragraph's rPr or its first run's rPr.
    Checks paragraph-level rPr first (w:pPr/w:rPr), then first w:r/w:rPr.
    Falls back to _TABLE_LINE_PT.
    """
    # Paragraph-level run properties (pPr > rPr)
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is not None:
        rPr = pPr.find(qn("w:rPr"))
        if rPr is not None:
            sz = rPr.find(qn("w:sz"))
            if sz is not None:
                val = sz.get(qn("w:val"))
                if val and val.isdigit():
                    return int(val) / 2

    # First run's rPr
    for r in p_elem.findall(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is not None:
            sz = rPr.find(qn("w:sz"))
            if sz is not None:
                val = sz.get(qn("w:val"))
                if val and val.isdigit():
                    return int(val) / 2

    return _TABLE_LINE_PT   # default: 12 pt


def _para_line_height_pt(p_elem, font_pt: float) -> float:
    """
    Resolve actual single-line rendered height (pt) for a paragraph,
    reading w:spacing w:line + w:lineRule from the paragraph's pPr.
    """
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return font_pt
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        return font_pt

    line_val = spacing.get(qn("w:line"))
    line_rule = spacing.get(qn("w:lineRule"), "auto")

    if line_val and line_val.lstrip("-").isdigit():
        lv = int(line_val)
        if line_rule == "exact":
            # Exact: value is in twips
            return _twip_pt(lv)
        elif line_rule == "atLeast":
            # At-least: value in twips, but could be taller
            return max(font_pt, _twip_pt(lv))
        else:
            # "auto" (default): value is in 240ths of a line
            # 240 = single spacing; 360 = 1.5x
            return font_pt * (lv / 240.0)

    return font_pt


def _para_spacing_pt(p_elem) -> tuple[float, float]:
    """Return (space_before_pt, space_after_pt) for a paragraph."""
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return 0.0, 0.0
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        return 0.0, 0.0
    sb = sa = 0.0
    before = spacing.get(qn("w:before"))
    after  = spacing.get(qn("w:after"))
    if before and before.lstrip("-").isdigit():
        sb = _twip_pt(int(before))
    if after and after.lstrip("-").isdigit():
        sa = _twip_pt(int(after))
    return sb, sa


def _estimate_cell_height(cell, col_w_pt: float) -> float:
    """
    Estimate total height of a table cell in points.

    Accounts for:
    - All paragraphs in the cell (not just concatenated text)
    - Per-paragraph font size, line spacing, space_before, space_after
    - Proportional TNR character width for line-wrap estimation
    - Cell top+bottom margins from w:tcMar
    """
    p_elems = cell._element.findall(qn("w:p"))
    if not p_elems:
        return _TABLE_LINE_PT + _CELL_PADDING_PT

    total_h = 0.0
    for p_elem in p_elems:
        font_pt = _para_font_size_pt(p_elem)
        line_h  = _para_line_height_pt(p_elem, font_pt)
        # If no explicit line spacing is set in the paragraph XML, the cell
        # inherits the document's Normal style (typically 1.5× in KFU docs).
        # Apply 1.5× as a conservative default to avoid underestimating row height.
        if abs(line_h - font_pt) < 0.5:   # "line_h == font_pt" means unset (single)
            line_h = font_pt * 1.5
        sb, sa  = _para_spacing_pt(p_elem)

        # TNR avg char width ≈ 0.50 × font size (conservative — Cyrillic glyphs are wider than Latin)
        pt_per_char  = font_pt * 0.50
        chars_per_line = max(4, int(col_w_pt / pt_per_char))

        # Gather text from all runs (preserves multi-run paragraphs)
        text = "".join(
            (r.find(qn("w:t")).text or "")
            for r in p_elem.findall(qn("w:r"))
            if r.find(qn("w:t")) is not None
        ).strip()

        n_lines = max(1, math.ceil(len(text) / chars_per_line)) if text else 1
        total_h += n_lines * line_h + sb + sa

    # Cell top+bottom margins
    cell_margin = _cell_margins_pt(cell._element)
    return total_h + cell_margin


def _estimate_row_height(row, body_width_pt: float, col_widths_pt: list[float] | None = None) -> float:
    """
    Estimated rendered height of a table row in points.

    Priority:
    1. Explicit w:trHeight (hRule=exact) → use as-is
    2. Explicit w:trHeight (hRule=atLeast) → use as minimum
    3. Estimate from cell content via _estimate_cell_height
    """
    tr = row._tr
    trPr = tr.find(qn("w:trPr"))
    explicit_min = 0.0
    if trPr is not None:
        trH = trPr.find(qn("w:trHeight"))
        if trH is not None:
            val = trH.get(qn("w:val"))
            h_rule = trH.get(qn("w:hRule"), "atLeast")
            if val and val.lstrip("-").isdigit():
                h = _twip_pt(int(val))
                if h > 2:
                    if h_rule == "exact":
                        return h   # exact → trust it completely
                    else:
                        explicit_min = h   # atLeast → use as lower bound

    cells = row.cells
    if not cells:
        return max(explicit_min, _TABLE_LINE_PT + _CELL_PADDING_PT)

    num_cols = len(cells)

    # Per-cell column width: actual XML widths preferred
    if col_widths_pt and len(col_widths_pt) >= num_cols:
        col_ws = col_widths_pt
    else:
        equal_w = max(20.0, body_width_pt / num_cols)
        col_ws = [equal_w] * num_cols

    max_h = 0.0
    seen: set[int] = set()
    col_idx = 0
    for cell in cells:
        cid = id(cell._element)
        if cid in seen:
            col_idx += 1
            continue
        seen.add(cid)

        col_w_pt = col_ws[col_idx] if col_idx < len(col_ws) else max(20.0, body_width_pt / num_cols)
        cell_h = _estimate_cell_height(cell, col_w_pt)
        max_h = max(max_h, cell_h)
        col_idx += 1

    return max(explicit_min, max_h)


# ── Body element iteration ────────────────────────────────────────────────────

def _iter_body(doc: Document):
    """
    Yield (kind, xml_element, py_object) for each direct child of <w:body>.
    kind ∈ {"paragraph", "table"}
    """
    body = doc.element.body
    para_map  = {p._element: p for p in doc.paragraphs}
    table_map = {t._element: t for t in doc.tables}

    for child in body:
        local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if local == "p" and child in para_map:
            yield "paragraph", child, para_map[child]
        elif local == "tbl" and child in table_map:
            yield "table", child, table_map[child]


# ── Table number extraction ───────────────────────────────────────────────────

_TBL_NUM_RE = re.compile(
    r"(?:таблица|table)\s+(\d+(?:\.\d+){0,2})",
    re.IGNORECASE,
)


def _extract_table_num(text: str) -> str | None:
    m = _TBL_NUM_RE.search(text.strip())
    return m.group(1) if m else None


# ── "Продолжение таблицы X" paragraph ────────────────────────────────────────

_FORMATTER_RSID = "00CF0001"   # stamp on formatter-inserted paragraphs; never set by Word


def _make_continuation_para(table_num: str) -> OxmlElement:
    """
    Build <w:p> for "Продолжение таблицы X.Y.Z":
      right-aligned, Times New Roman 14 pt, not bold, no indent, keep_with_next.
    """
    p = OxmlElement("w:p")
    # Stamp a unique rsidR so _is_formatter_continuation can identify this
    # paragraph even after Phase 1 reformats its content.
    p.set(qn("w:rsidR"), _FORMATTER_RSID)

    pPr = OxmlElement("w:pPr")

    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "right")
    pPr.append(jc)

    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:right"), "0")
    ind.set(qn("w:firstLine"), "0")
    pPr.append(ind)

    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "360")       # 1.5× line spacing (360 / 240 = 1.5)
    spacing.set(qn("w:lineRule"), "auto")
    pPr.append(spacing)

    # keep_with_next so "Продолжение" doesn't hang alone without the table
    keep = OxmlElement("w:keepNext")
    pPr.append(keep)

    p.append(pPr)

    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    rFonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(attr), "Times New Roman")
    rPr.append(rFonts)

    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), "28")   # 14 pt = 28 half-points
        rPr.append(el)

    for tag in ("w:b", "w:bCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), "0")    # explicitly not bold
        rPr.append(el)

    r.append(rPr)

    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = f"Продолжение таблицы {table_num}"
    r.append(t)
    p.append(r)

    return p


# ── Table split ───────────────────────────────────────────────────────────────

_MIN_DATA_ROWS = 1   # at least this many data rows must remain on the first page

# In-memory split hints for tables that were merged from a student-provided
# continuation pair during the current formatting run.
# key: id(tbl_xml_element), value: split_after_row
_MERGED_SPLIT_HINTS: dict[int, int] = {}


def _split_table(tbl_elem, split_after_row: int, table_num: str) -> bool:
    """
    Split tbl_elem after row index split_after_row.

    After the call the document body looks like:
        tbl1  (rows 0 … split_after_row)
        <w:p> "Продолжение таблицы X.Y.Z"
        tbl2  (header_copy + rows split_after_row+1 … end)

    Returns True on success, False if split is not meaningful.
    """
    rows = tbl_elem.findall(qn("w:tr"))
    total = len(rows)

    # Clamp split point: keep ≥ MIN_DATA_ROWS data rows on first page,
    # and ≥ 1 row after the split.
    split_after_row = max(_MIN_DATA_ROWS, min(split_after_row, total - 2))

    if split_after_row >= total - 1:
        logger.debug("_split_table: nothing to split (split_after=%d, total=%d)", split_after_row, total)
        return False

    # Deep-copy entire table → becomes the second (continuation) table
    tbl2 = copy.deepcopy(tbl_elem)

    # ── Strip stale LRPB markers from tbl2 ────────────────────────────────
    # The rows moved to tbl2 carry w:lastRenderedPageBreak elements that
    # reflect the ORIGINAL document layout.  After the split, those page
    # boundaries are no longer valid.  Removing them prevents a second
    # formatter pass from treating them as fresh split signals and
    # re-splitting the continuation table.
    for lrpb_elem in list(tbl2.findall(".//" + _LRPB_TAG)):
        parent = lrpb_elem.getparent()
        if parent is not None:
            parent.remove(lrpb_elem)

    # ── Trim tbl1: remove rows after the split ─────────────────────────────
    tbl1_rows = tbl_elem.findall(qn("w:tr"))
    for row in tbl1_rows[split_after_row + 1:]:
        tbl_elem.remove(row)

    # ── Trim tbl2: remove rows up to and including split_after_row ─────────
    tbl2_rows = tbl2.findall(qn("w:tr"))
    header_copy = tbl2_rows[0]          # deep-copy of row 0 (already copied)
    for row in tbl2_rows[: split_after_row + 1]:
        tbl2.remove(row)

    # ── Prepend header row to tbl2 ─────────────────────────────────────────
    first_data_rows = tbl2.findall(qn("w:tr"))
    if first_data_rows:
        tbl2.insert(list(tbl2).index(first_data_rows[0]), header_copy)
    else:
        tbl2.append(header_copy)

    # Mark the prepended header as a repeating table header (w:tblHeader)
    trPr = header_copy.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        header_copy.insert(0, trPr)
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))

    # ── Insert "Продолжение" paragraph + tbl2 after tbl1 ──────────────────
    cont_para = _make_continuation_para(table_num)
    tbl_elem.addnext(tbl2)         # body: tbl1 → tbl2
    tbl_elem.addnext(cont_para)    # body: tbl1 → cont_para → tbl2  ✓

    logger.info(
        "table_continuation: split '%s' after row %d/%d",
        table_num, split_after_row, total - 1,
    )
    return True


# ── Table merging (pre-existing student splits) ───────────────────────────────

_CONT_RE = re.compile(r"продолжени", re.IGNORECASE)
_TBL_WORD_RE = re.compile(r"таблиц", re.IGNORECASE)


def _is_student_continuation(text: str) -> bool:
    """
    True if paragraph text looks like a student-written standalone
    'Продолжение таблицы X.Y.Z' header.

    Guard: text must be short (≤30 chars) — long paragraphs are prose
    that merely happen to contain those words mid-sentence.
    30 chars covers table numbers up to e.g. "100.10.10" depth.
    """
    if len(text) > 30:
        return False
    return bool(_CONT_RE.search(text) and _TBL_WORD_RE.search(text))


def _is_formatter_continuation(para) -> bool:
    """
    True if this paragraph is a formatter-inserted 'Продолжение таблицы' paragraph.

    The formatter stamps w:rsidR="_FORMATTER_RSID" on every continuation paragraph
    it creates.  Phase 1 (safe_formatter) never reads or modifies rsidR attributes,
    so the stamp survives all formatting passes and is the only reliable discriminator.
    """
    try:
        return para._element.get(qn("w:rsidR")) == _FORMATTER_RSID
    except Exception:
        return False


def _rows_match(row1, row2) -> bool:
    """Compare cell texts of two table rows (True if identical)."""
    cells1 = [c.text.strip() for c in row1.cells]
    cells2 = [c.text.strip() for c in row2.cells]
    return cells1 == cells2


def _merge_tables(tbl1_elem, tbl1_obj, tbl2_elem, tbl2_obj) -> None:
    """
    Append all rows from tbl2 into tbl1.
    If the first row of tbl2 is identical to the first row of tbl1 (duplicate
    header), skip it.
    """
    rows1 = tbl1_obj.rows
    rows2 = tbl2_obj.rows

    if not rows2:
        return

    start_row = 0
    if rows1 and rows2 and _rows_match(rows1[0], rows2[0]):
        start_row = 1   # skip duplicate header

    for row in rows2[start_row:]:
        tbl1_elem.append(deepcopy(row._tr))

    logger.debug("_merge_tables: appended %d row(s) from tbl2", len(rows2) - start_row)


def _tbl_has_lrpb(tbl_obj) -> bool:
    """True if any row (after row 0) in the table contains w:lastRenderedPageBreak."""
    return _row_lrpb_index(tbl_obj.rows) > 0


def _optimize_table_col_widths(tbl_xml, body_width_pt: float) -> bool:
    """
    Ensure no column is narrower than _MIN_COL_PT and total width ≤ body_width_pt.

    Algorithm:
      1. Scale all columns down proportionally if total > body_width_pt.
      2. Identify undersized columns; redistribute deficit from wider donor columns.

    Updates both w:tblGrid/w:gridCol and each w:tc/w:tcPr/w:tcW (honouring
    w:gridSpan for merged cells).

    Returns True if any width was changed.
    """
    grid = tbl_xml.find(qn("w:tblGrid"))
    if grid is None:
        return False
    gridcols = grid.findall(qn("w:gridCol"))
    if not gridcols:
        return False

    widths = [int(c.get(qn("w:w"), 0)) / TWIP_PER_PT for c in gridcols]
    n = len(widths)
    total = sum(widths)
    if total < 1:
        return False

    changed = False

    # Step 1: scale down if total exceeds body width
    if total > body_width_pt + 0.5:
        scale = body_width_pt / total
        widths = [w * scale for w in widths]
        total = sum(widths)
        changed = True

    # Step 2: redistribute to fix undersized columns (up to n iterations)
    for _ in range(n):
        undersized = [(i, _MIN_COL_PT - widths[i]) for i in range(n)
                      if widths[i] < _MIN_COL_PT - 0.5]
        if not undersized:
            break
        donors = [i for i in range(n) if widths[i] > _MIN_COL_PT + 0.5]
        if not donors:
            break
        total_deficit = sum(d for _, d in undersized)
        total_donor_excess = sum(widths[i] - _MIN_COL_PT for i in donors)
        take_frac = min(1.0, total_donor_excess / total_deficit)

        for i, deficit in undersized:
            widths[i] += deficit * take_frac
        actual_taken = total_deficit * take_frac
        for i in donors:
            donor_excess = widths[i] - _MIN_COL_PT
            widths[i] -= actual_taken * (donor_excess / total_donor_excess)
        changed = True

    if not changed:
        return False

    # Round to integer twips, keep total consistent
    twip_widths = [max(1, round(w * TWIP_PER_PT)) for w in widths]

    # Apply to grid
    for col_el, tw in zip(gridcols, twip_widths):
        col_el.set(qn("w:w"), str(tw))

    # Apply to each row's cells (respecting gridSpan)
    for tr in tbl_xml.findall(qn("w:tr")):
        col_idx = 0
        for tc in tr.findall(qn("w:tc")):
            if col_idx >= n:
                break
            tcPr = tc.find(qn("w:tcPr"))
            gridSpan_el = tcPr.find(qn("w:gridSpan")) if tcPr is not None else None
            span = int(gridSpan_el.get(qn("w:val"), 1)) if gridSpan_el is not None else 1
            span = max(1, min(span, n - col_idx))

            cell_tw = sum(twip_widths[col_idx: col_idx + span])

            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.append(tcW)
            tcW.set(qn("w:w"), str(cell_tw))
            tcW.set(qn("w:type"), "dxa")

            col_idx += span

    return True


def apply_table_merging(doc: Document) -> int:
    """
    Phase 3 pre-pass — detect tables that the student manually split with a
    "Продолжение таблицы X.Y.Z" paragraph and merge them back into a single table.

    After merging, apply_table_continuation will re-split any table that genuinely
    overflows the page and insert a correctly formatted continuation paragraph.

    Returns the number of merges performed.
    """
    _MERGED_SPLIT_HINTS.clear()
    body_elems = list(_iter_body(doc))
    n = len(body_elems)

    # Collect merge jobs:
    # (cont_para_indices, tbl2_index, tbl1_index)
    merge_jobs: list[tuple[list[int], int, int]] = []

    # Track the index of the most recent table
    last_tbl_idx: int | None = None

    i = 0
    while i < n:
        kind, xml_elem, py_obj = body_elems[i]

        if kind == "table":
            # ── Detect silent splits: consecutive tables with no body text ──
            # Students sometimes split a table by just inserting a page break
            # mid-table, creating two (or more) adjacent tables with no
            # "Продолжение" paragraph between them.  Merge if:
            #   • the immediately preceding body element is also a table, AND
            #   • all elements between them are empty paragraphs, AND
            #   • both tables have the same number of grid columns.
            if last_tbl_idx is not None:
                inter = body_elems[last_tbl_idx + 1 : i]
                has_content = any(
                    k == "paragraph" and (po.text or "").strip()
                    for k, _, po in inter
                )
                if not has_content:
                    _, tbl1_xml, tbl1_obj = body_elems[last_tbl_idx]
                    tbl2_obj = py_obj
                    tbl1_grid = tbl1_xml.find(qn("w:tblGrid"))
                    tbl2_grid = xml_elem.find(qn("w:tblGrid"))
                    tbl1_cols = len(tbl1_grid.findall(qn("w:gridCol"))) if tbl1_grid is not None else 0
                    tbl2_cols = len(tbl2_grid.findall(qn("w:gridCol"))) if tbl2_grid is not None else 0
                    if tbl1_cols > 0 and tbl1_cols == tbl2_cols:
                        # Only merge if at least one part has LRPB — guarantees
                        # we can re-split at the correct position.  Without LRPB
                        # we have no reliable split point and the student's break
                        # (even if imperfect) is better than none.
                        if not (_tbl_has_lrpb(tbl1_obj) or _tbl_has_lrpb(tbl2_obj)):
                            last_tbl_idx = i
                            i += 1
                            continue
                        empty_para_idxs = [
                            idx
                            for idx, (k, _, _) in enumerate(body_elems[last_tbl_idx + 1 : i], start=last_tbl_idx + 1)
                            if k == "paragraph"
                        ]
                        merge_jobs.append((empty_para_idxs, i, last_tbl_idx))
                        last_tbl_idx = i
                        i += 1
                        continue

            last_tbl_idx = i

        elif kind == "paragraph":
            text = (py_obj.text or "").strip()
            if (_is_student_continuation(text)
                    and not _is_formatter_continuation(py_obj)
                    and last_tbl_idx is not None):
                # Collect the continuation paragraph + any trailing empty paras
                cont_indices = [i]
                j = i + 1
                while j < n:
                    k2, _, po2 = body_elems[j]
                    if k2 != "paragraph":
                        break
                    t2 = (po2.text or "").strip()
                    if t2:
                        break   # non-empty non-table element — stop
                    cont_indices.append(j)
                    j += 1

                # j must now point to the second table part
                if j < n and body_elems[j][0] == "table":
                    # Only merge if LRPB is present — without it we cannot
                    # re-split correctly and student's break is better than none
                    _, tbl1_xml_c, tbl1_obj_c = body_elems[last_tbl_idx]
                    _, _tbl2_xml_c, tbl2_obj_c = body_elems[j]
                    if not (_tbl_has_lrpb(tbl1_obj_c) or _tbl_has_lrpb(tbl2_obj_c)):
                        i = j + 1
                        last_tbl_idx = j
                        continue
                    merge_jobs.append((cont_indices, j, last_tbl_idx))
                    # Skip past everything we just catalogued
                    i = j + 1
                    last_tbl_idx = i - 1  # tbl2 is now the "last table"
                    continue

        i += 1

    if not merge_jobs:
        logger.info("table_merging: no pre-existing splits found")
        return 0

    # Apply in reverse order so indices remain valid
    merged = 0
    for cont_indices, tbl2_idx, tbl1_idx in reversed(merge_jobs):
        try:
            _, tbl1_xml, tbl1_obj = body_elems[tbl1_idx]
            _, tbl2_xml, tbl2_obj = body_elems[tbl2_idx]
            # Preserve the student's original split point (header + N rows in
            # the first part => split_after_row = len(rows_part1) - 1).
            # We will prioritise this hint over stale LRPB markers after merge.
            split_hint = max(_MIN_DATA_ROWS, len(tbl1_obj.rows) - 1)
            _merge_tables(tbl1_xml, tbl1_obj, tbl2_xml, tbl2_obj)
            _MERGED_SPLIT_HINTS[id(tbl1_xml)] = split_hint

            # Remove continuation paragraphs
            for ci in cont_indices:
                xe = body_elems[ci][1]
                parent = xe.getparent()
                if parent is not None:
                    parent.remove(xe)

            # Remove the second table element
            parent = tbl2_xml.getparent()
            if parent is not None:
                parent.remove(tbl2_xml)

            merged += 1
            logger.info(
                "table_merging: merged tbl@%d ← tbl@%d (removed %d cont para(s))",
                tbl1_idx, tbl2_idx, len(cont_indices),
            )
        except Exception:
            logger.exception(
                "table_merging: failed to merge tbl@%d ← tbl@%d", tbl1_idx, tbl2_idx
            )

    logger.info("table_merging: %d merge(s) applied", merged)
    return merged


# ── Main entry point ──────────────────────────────────────────────────────────

_MIN_ROWS_PAGE1_WARN = 4   # warn in trouble-report if fewer rows land on page 1


def apply_table_continuation(
    doc: Document,
    report: FormattingReport | None = None,
) -> int:
    """
    Walk document body elements, detect tables that overflow a page, split them.

    For tables spanning multiple pages (multiple LRPB rows), multiple splits
    are produced — one per LRPB row, applied from last to first so that row
    indices remain valid across sequential splits of the same element.

    Args:
        doc:    The loaded python-docx Document (modified in-place).
        report: Optional FormattingReport — warnings appended for tables that
                could not be split or produced very short first-page sections.

    Returns the number of splits performed.
    """
    body_h = _body_height_pt(doc)
    body_w = _body_width_pt(doc)
    logger.info("table_continuation: body_height=%.1f pt  body_width=%.1f pt", body_h, body_w)

    body_elems = list(_iter_body(doc))   # snapshot — safe to mutate doc after this

    last_tbl_num: str | None = None   # table number from the most recent caption
    # Each entry: (tbl_xml, split_after_row, table_num)
    # For multi-LRPB tables, multiple entries share the same tbl_xml, ordered
    # from highest split_after_row to lowest so they apply correctly in sequence.
    splits: list[tuple] = []

    for kind, xml_elem, py_obj in body_elems:

        if kind == "paragraph":
            text = (py_obj.text or "").strip()
            num = _extract_table_num(text)
            if num:
                last_tbl_num = num

        elif kind == "table":
            rows = py_obj.rows
            if len(rows) < 2:
                continue

            table_num = last_tbl_num or "?"

            # ── Priority 1: student split hint (from apply_table_merging) ──
            split_hint = _MERGED_SPLIT_HINTS.get(id(xml_elem))
            if split_hint is not None:
                splits.append((xml_elem, split_hint, table_num))
                logger.info(
                    "table_continuation: merged-hint split '%s' after row %d",
                    table_num, split_hint,
                )
                continue

            # ── Priority 2: w:lastRenderedPageBreak ────────────────────────
            # Collect ALL LRPB rows so tables spanning 3+ pages get multiple
            # splits.  Sort descending so each split correctly trims the tail
            # of tbl_xml without invalidating the earlier split positions.
            lrpb_rows = _all_row_lrpb_indices(rows)
            if lrpb_rows:
                for lrpb_row in sorted(lrpb_rows, reverse=True):
                    split_after = max(_MIN_DATA_ROWS, lrpb_row - 1)
                    splits.append((xml_elem, split_after, table_num))
                logger.info(
                    "table_continuation: LRPB split '%s' — %d page-break(s) at rows %s",
                    table_num, len(lrpb_rows), lrpb_rows,
                )
                continue

            # ── No signal: table cannot be auto-split ─────────────────────
            if report is not None:
                report.warn(
                    f"Таблица {table_num}: нет данных о разрыве страниц — "
                    "проверьте перенос вручную"
                )
            logger.info(
                "table_continuation: no split signal for table '%s' — skipped",
                table_num,
            )

    # Apply splits in the order they were appended.
    # For a single table with multiple LRPB positions the entries are already
    # ordered from highest to lowest split_after_row (descending sort above),
    # so each successive _split_table call operates on the shortened tbl_xml.
    n_splits = 0
    for tbl_xml, split_after_row, tbl_num in splits:
        # Warn if the first page part will be very short (< _MIN_ROWS_PAGE1_WARN rows)
        if report is not None and split_after_row < _MIN_ROWS_PAGE1_WARN - 1:
            report.warn(
                f"При переносе таблицы {tbl_num} осталось мало строк "
                f"({split_after_row + 1} стр. 1)"
            )
        if _split_table(tbl_xml, split_after_row, tbl_num):
            n_splits += 1

    logger.info("table_continuation: %d split(s) applied", n_splits)

    # Column-width optimisation: fix phantom/undersized columns and tables
    # wider than the body.  Applied to ALL tables after splitting so that
    # both original and continuation tables are corrected.
    n_col_fixed = 0
    body_elems_post = list(_iter_body(doc))
    for kind, tbl_xml, _ in body_elems_post:
        if kind != "table":
            continue
        if _optimize_table_col_widths(tbl_xml, body_w):
            n_col_fixed += 1
    if n_col_fixed:
        logger.info("table_continuation: col-width optimised %d table(s)", n_col_fixed)

    return n_splits


# Only trust a w:lastRenderedPageBreak calibration signal when we have already
# accumulated at least this fraction of the page.  LRPB markers that fire at
# very low cumulative heights are stale artefacts from the ORIGINAL layout that
# no longer reflect page boundaries in the MODIFIED document (e.g. a paragraph
# that was the first on a page in the source but is now mid-page after a table
# split was inserted above it).
_LRPB_TRUST_RATIO = 0.25   # 25 % of body height ≈ ~178 pt for a KFU page


def _lrpb_calibrate(xml_elem, current_h: float, body_h: float) -> float:
    """
    Return the new current_h after applying an optional LRPB calibration.

    Resets to 0.0 only when:
      1. The paragraph contains a w:lastRenderedPageBreak, AND
      2. current_h >= body_h * _LRPB_TRUST_RATIO
         (enough content has been seen that the LRPB is likely genuine).
    """
    if _para_has_lrpb(xml_elem) and current_h >= body_h * _LRPB_TRUST_RATIO:
        return 0.0
    return current_h


# ── Helpers for geometry-based page-break rules ───────────────────────────────

_TABLE_CAP_RE_GEOM = re.compile(
    r"^\s*(таблица|table)\s+\d+(?:\.\d+){0,2}",
    re.IGNORECASE,
)
_FIGURE_CAP_RE_GEOM = re.compile(
    r"^\s*(рис\.|рисунок|figure|fig\.)\s*\d+",
    re.IGNORECASE,
)


def _para_has_image(p_elem) -> bool:
    """True if the paragraph XML element contains an inline drawing or picture."""
    return xml_has_image(p_elem)


def _set_page_break_before(para_elem) -> None:
    """Add w:pageBreakBefore to a paragraph's pPr (idempotent)."""
    pPr = para_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        para_elem.insert(0, pPr)
    if pPr.find(qn("w:pageBreakBefore")) is None:
        pb = OxmlElement("w:pageBreakBefore")
        pPr.append(pb)


# ── Rule 4: no empty first line of page ──────────────────────────────────────

def _apply_rule4_pass(doc: Document) -> int:
    """
    Single pass of Rule 4 — remove empty paragraphs at the very top of a page.

    Conservative: only removes paragraphs with no text AND no meaningful
    spacing (space_before ≤ 2 pt). This avoids deleting intentional
    visual separators.

    Returns the number of paragraphs removed in this pass.
    """
    body_h = _body_height_pt(doc)
    body_w = _body_width_pt(doc)

    body_elems = list(_iter_body(doc))
    current_h = 0.0
    to_remove: list = []
    prev_nonempty_kind: str | None = None

    for kind, xml_elem, py_obj in body_elems:
        if kind == "paragraph":
            # LRPB calibration — only trust when enough page content was seen
            current_h = _lrpb_calibrate(xml_elem, current_h, body_h)

            h = _estimate_para_height(py_obj)

            page_overflow = (current_h + h > body_h)
            if page_overflow:
                current_h = 0.0   # new page starts

            text = (py_obj.text or "").strip()
            # A paragraph with an image but no text must never be treated as
            # "empty" — removing it would delete the figure from the document.
            is_empty = not text and not xml_has_image(xml_elem)

            if page_overflow and is_empty:
                # Preserve intentional blank lines that must remain after
                # headings and table note blocks ("Источник:" / "Примечание:").
                if prev_nonempty_kind in {"heading", "source_or_note"}:
                    current_h += h
                    continue

                # Check it's not a meaningful spacer (large space_before)
                try:
                    sb = py_obj.paragraph_format.space_before
                    if sb and sb.pt > 2:
                        current_h += h
                        continue
                except Exception:
                    pass
                to_remove.append(xml_elem)
                # current_h stays 0 — next element is still first on page
            else:
                current_h += h
                if not is_empty:
                    if _looks_like_heading(text):
                        prev_nonempty_kind = "heading"
                    elif is_source_or_note_line(text):
                        prev_nonempty_kind = "source_or_note"
                    else:
                        prev_nonempty_kind = "text"

        elif kind == "table":
            prev_nonempty_kind = "table"
            rows = py_obj.rows
            col_widths = _tbl_col_widths_pt(xml_elem)
            for rh in (_estimate_row_height(r, body_w, col_widths) for r in rows):
                current_h += rh
                if current_h > body_h:
                    current_h = rh

    removed = 0
    for elem in reversed(to_remove):
        parent = elem.getparent()
        if parent is not None:
            parent.remove(elem)
            removed += 1

    return removed


def apply_rule4_empty_first_lines(doc: Document) -> int:
    """
    Rule 4 — Remove empty paragraphs that land at the very top of a page.

    Runs iteratively until convergence: each removal can shift subsequent
    page boundaries, potentially exposing new violations that the first
    pass missed (stale LRPB calibration + cascading removals).

    Returns total number of paragraphs removed across all passes.
    """
    total = 0
    for _ in range(5):   # cap at 5 iterations to prevent infinite loops
        n = _apply_rule4_pass(doc)
        total += n
        if n == 0:
            break
    logger.info("rule4: removed %d empty first-line paragraph(s) total", total)
    return total


# ── Rule 3: no orphan table caption at page bottom ────────────────────────────

def apply_rule3_table_orphan(doc: Document) -> int:
    """
    Rule 3 (geometry) — Prevent table caption from hanging alone at page bottom.

    If a table_caption paragraph (optionally followed by a short title line)
    fits on the current page but the table's first data row does not,
    set w:pageBreakBefore on the caption so the caption and table land
    together on the next page.

    Returns the number of captions given a pageBreakBefore.
    """
    body_h = _body_height_pt(doc)
    body_w = _body_width_pt(doc)
    body_elems = list(_iter_body(doc))
    n = len(body_elems)
    current_h = 0.0
    count = 0

    i = 0
    while i < n:
        kind, xml_elem, py_obj = body_elems[i]

        if kind == "paragraph":
            current_h = _lrpb_calibrate(xml_elem, current_h, body_h)

            text = (py_obj.text or "").strip()
            h = _estimate_para_height(py_obj)

            if not _TABLE_CAP_RE_GEOM.match(text):
                if current_h + h > body_h:
                    current_h = h
                else:
                    current_h += h
                i += 1
                continue

            # Found a table caption — collect caption + possible title lines
            cap_start_h = current_h
            cap_items: list[tuple] = [(xml_elem, h)]   # (xml_elem, height)

            j = i + 1
            while j < n:
                k2, xe2, po2 = body_elems[j]
                if k2 != "paragraph":
                    break
                t2 = (po2.text or "").strip()
                # Stop at: empty para, very long line (body text), another caption
                if not t2 or len(t2) > 200 or _TABLE_CAP_RE_GEOM.match(t2):
                    break
                cap_items.append((xe2, _estimate_para_height(po2)))
                j += 1

            cap_total_h = sum(h2 for _, h2 in cap_items)

            # j should point to the table element
            if j < n and body_elems[j][0] == "table":
                tbl_py = body_elems[j][2]
                tbl_xml = body_elems[j][1]
                rows = tbl_py.rows
                if rows:
                    col_widths = _tbl_col_widths_pt(tbl_xml)
                    first_row_h = _estimate_row_height(rows[0], body_w, col_widths)

                    caption_fits     = (cap_start_h + cap_total_h <= body_h)
                    first_row_orphan = (cap_start_h + cap_total_h + first_row_h > body_h)
                    fits_fresh       = (cap_total_h + first_row_h <= body_h)

                    if caption_fits and first_row_orphan and fits_fresh:
                        _set_page_break_before(cap_items[0][0])
                        count += 1
                        logger.info(
                            "rule3: pageBreakBefore on table caption [%s]",
                            text[:50],
                        )
                        current_h = cap_total_h
                        i = j      # resume from the table element
                        continue

            # No action — advance geometry past caption + title
            current_h = cap_start_h + cap_total_h
            if current_h > body_h:
                current_h = cap_items[-1][1]
            i = j
            continue

        elif kind == "table":
            rows = py_obj.rows
            col_widths = _tbl_col_widths_pt(xml_elem)
            for row in rows:
                rh = _estimate_row_height(row, body_w, col_widths)
                if current_h + rh > body_h:
                    current_h = rh
                else:
                    current_h += rh

        i += 1

    logger.info("rule3: %d table caption(s) given pageBreakBefore", count)
    return count


# ── Rule 6: figure must stay with its caption ─────────────────────────────────

def apply_rule6_figure_orphan(doc: Document) -> int:
    """
    Rule 6 (geometry) — Prevent figure caption from being stranded at the
    top of the next page while the figure itself is on the current page.

    If an image paragraph fits on the current page but the immediately
    following figure_caption does not, set w:pageBreakBefore on the image
    so both the image and caption land on the next page together.

    Returns the number of images given a pageBreakBefore.
    """
    body_h = _body_height_pt(doc)
    body_w = _body_width_pt(doc)
    body_elems = list(_iter_body(doc))
    n = len(body_elems)
    current_h = 0.0
    count = 0

    i = 0
    while i < n:
        kind, xml_elem, py_obj = body_elems[i]

        if kind == "paragraph":
            current_h = _lrpb_calibrate(xml_elem, current_h, body_h)

            h = _estimate_para_height(py_obj)

            if not _para_has_image(xml_elem):
                if current_h + h > body_h:
                    current_h = h
                else:
                    current_h += h
                i += 1
                continue

            # Image paragraph — check if the next paragraph is a figure caption
            if i + 1 < n:
                nk, nxe, npo = body_elems[i + 1]
                if nk == "paragraph":
                    next_text = (npo.text or "").strip()
                    if _FIGURE_CAP_RE_GEOM.match(next_text):
                        caption_h   = _estimate_para_height(npo)
                        img_fits    = (current_h + h <= body_h)
                        cap_orphan  = (current_h + h + caption_h > body_h)
                        fits_fresh  = (h + caption_h <= body_h)

                        if img_fits and cap_orphan and fits_fresh:
                            _set_page_break_before(xml_elem)
                            count += 1
                            logger.info(
                                "rule6: pageBreakBefore on image before [%s]",
                                next_text[:50],
                            )
                            # Both now start fresh on next page
                            current_h = h + caption_h
                            i += 2
                            continue

            # Normal advance
            if current_h + h > body_h:
                current_h = h
            else:
                current_h += h

        elif kind == "table":
            rows = py_obj.rows
            col_widths = _tbl_col_widths_pt(xml_elem)
            for row in rows:
                rh = _estimate_row_height(row, body_w, col_widths)
                if current_h + rh > body_h:
                    current_h = rh
                else:
                    current_h += rh

        i += 1

    logger.info("rule6: %d figure(s) given pageBreakBefore", count)
    return count


# ── Rule 2: no trailing empty lines at page bottom before a heading ───────────

_HEADING_RE = re.compile(
    r"^\s*\d+(?:\.\d+)*\.?\s",   # matches "1. …" / "1.1. …" / "1.1.1. …"
)


def _looks_like_heading(text: str) -> bool:
    return bool(_HEADING_RE.match(text))


def apply_rule2_trailing_empties(doc: Document) -> int:
    """
    Rule 2 — Remove empty paragraphs that sit at the very bottom of a page
    when the next non-empty element is a heading (heading1 / heading2).

    These ghost lines appear because the geometry estimator places them
    mid-page, but Word's real line-breaking pushes them to page bottom,
    so Rule 4 (which only catches first-on-page empties) never fires.

    Strategy:
      Walk body elements in order.  Collect runs of consecutive empty
      paragraphs.  When the run is followed by a heading-like paragraph
      AND the geometry says the run straddles or is near the page
      boundary (within _BOTTOM_TOLERANCE_PT), mark the empties for removal.

    Conservative: requires the very next non-empty paragraph to be a heading
    so we don't accidentally eat intentional visual separators between sections.

    Returns the number of paragraphs removed.
    """
    body_h = _body_height_pt(doc)
    body_w = _body_width_pt(doc)
    _BOTTOM_TOLERANCE_PT = _BODY_LINE_PT * 3   # empty lines within last ~3 lines

    body_elems = list(_iter_body(doc))
    n = len(body_elems)

    current_h = 0.0
    to_remove: list = []

    i = 0
    while i < n:
        kind, xml_elem, py_obj = body_elems[i]

        if kind == "paragraph":
            text = (py_obj.text or "").strip()
            h = _estimate_para_height(py_obj)

            if current_h + h > body_h:
                current_h = h   # new page

            if not text:
                # Start of a potential empty-paragraph run
                run_start = i
                run_elems = [(xml_elem, h)]
                run_h_start = current_h   # height at start of run

                j = i + 1
                while j < n:
                    k2, xe2, po2 = body_elems[j]
                    if k2 != "paragraph":
                        break
                    t2 = (po2.text or "").strip()
                    if t2:
                        break
                    run_elems.append((xe2, _estimate_para_height(po2)))
                    j += 1

                # j now points to the first non-empty element after the run
                next_is_heading = False
                if j < n:
                    k_next, _, po_next = body_elems[j]
                    if k_next == "paragraph":
                        t_next = (po_next.text or "").strip()
                        next_is_heading = _looks_like_heading(t_next)

                if next_is_heading:
                    run_total_h = sum(rh for _, rh in run_elems)
                    heading_h = _estimate_para_height(po_next)

                    # Only remove if the heading lands on the SAME page.
                    # If the empty run already pushes past body_h → heading is on
                    # the next page → the empties are harmless bottom-of-page padding,
                    # leave them alone (user confirmed this is acceptable).
                    heading_on_next_page = (
                        run_h_start + run_total_h + heading_h > body_h
                    )
                    if not heading_on_next_page:
                        for xe, _ in run_elems:
                            to_remove.append(xe)
                        current_h = run_h_start   # pretend the run wasn't there
                        i = j
                        continue

                # Otherwise just advance normally through the run
                for _, rh in run_elems:
                    current_h += rh
                    if current_h > body_h:
                        current_h = rh
                i = j
                continue

            else:
                current_h += h

        elif kind == "table":
            rows = py_obj.rows
            for rh in (_estimate_row_height(r, body_w) for r in rows):
                current_h += rh
                if current_h > body_h:
                    current_h = rh

        i += 1

    removed = 0
    for elem in reversed(to_remove):
        parent = elem.getparent()
        if parent is not None:
            parent.remove(elem)
            removed += 1

    logger.info("rule2: removed %d trailing empty paragraph(s) before headings", removed)
    return removed
