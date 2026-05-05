"""
Phase 3 — Table formatting utilities.

### What works now (geometry-based, no LibreOffice required)
  - apply_table_merging:      stub — returns 0 (see FUTURE note below)
  - apply_table_continuation: stub — returns 0 (see FUTURE note below)
  - _optimize_table_col_widths: active — fixes oversized/phantom columns
  - apply_rule3_table_orphan: active — prevents table caption orphaned at page bottom
  - apply_rule4_empty_first_lines: active — removes empty paragraphs at page top
  - apply_rule6_figure_orphan: active — keeps image with its caption

### FUTURE: Table splitting via LibreOffice (Rule 1)
#
# The table-continuation system (merge pre-split tables → re-split at real page
# breaks → insert "Продолжение таблицы X.Y.Z" headers) requires knowing EXACTLY
# where page breaks fall after formatting.  Pure geometry estimation (without a
# rendering engine) is too unreliable for production use:
#
#   Problem A — w:lastRenderedPageBreak (LRPB) is stale.
#     Word writes LRPB markers when it saves.  After Phase 1 reformatting
#     (fonts, margins, spacing all change) the LRPBs reflect the OLD layout,
#     not the new one.  Fresh KFU-formatted documents have NO LRPB at all,
#     producing 9-12 spurious "check manually" warnings per document.
#
#   Problem B — Geometry estimator is approximate.
#     Font metrics, line-wrap, cell merges, images in cells, and Word's own
#     internal kerning all introduce errors that compound over many rows.
#     A 2% per-row error on a 50-row table → entire page off.
#
# Recommended future approach — LibreOffice headless PDF-info:
#   1. Run `soffice --headless --convert-to pdf <formatted.docx>` (separate
#      Railway service or sidecar, NOT inline — adds ~400 MB + 8-15 s startup).
#   2. Parse the PDF page-stream to find exact row → page mapping.
#   3. Split at real page breaks, insert "Продолжение таблицы X.Y.Z" headers.
#
# Required helper functions (written, now commented-out):
#   _FORMATTER_RSID         — unique rsidR stamp for formatter-inserted paragraphs
#   _make_continuation_para — builds <w:p> "Продолжение таблицы X.Y.Z"
#   _split_table            — splits tbl_xml after row N, inserts continuation para
#   _is_formatter_continuation — detects formatter-stamped continuation paras
#   _rows_match / _merge_tables — merges two table parts (undo student splits)
#   apply_table_merging     — pre-pass: detect & merge student-split table pairs
#   apply_table_continuation — main pass: split at real page breaks
#
# To re-enable: restore those functions from git history (commit before this one),
# replace the stubs below, and integrate with a LibreOffice rendering step.
"""

from __future__ import annotations

import logging
import math
import os
import re
import shutil
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


from .docx_utils import xml_has_image, is_source_or_note_line, FormattingReport
from .layout_render import LibreOfficeNotFoundError, render_docx_to_pdf
from .pdf_layout_analyzer import PdfLine, analyze_pdf_lines
from .table_split_prototype import apply_numbered_split_to_document

logger = logging.getLogger(__name__)

# ── Unit helpers ─────────────────────────────────────────────────────────────

EMU_PER_PT  = 12700   # 1 pt  = 12 700 EMU  (python-docx stores lengths in EMU)
TWIP_PER_PT = 20      # 1 pt  = 20 twips    (w:trHeight val is in twips)

def _emu_pt(v: int) -> float: return v / EMU_PER_PT
def _twip_pt(v: int) -> float: return v / TWIP_PER_PT


# ── w:lastRenderedPageBreak helpers ──────────────────────────────────────────

_LRPB_TAG = qn("w:lastRenderedPageBreak")


def _para_has_lrpb(p_elem) -> bool:
    """True if this paragraph contains w:lastRenderedPageBreak.
    Used by _lrpb_calibrate (Rule 4 geometry estimator).
    """
    return p_elem.find(".//" + _LRPB_TAG) is not None

# ── Page geometry ─────────────────────────────────────────────────────────────

# Safety margin subtracted from body height so we don't overfill a page.
# Accounts for rounding + minor rendering differences between LO and Word.
_PAGE_BUFFER_PT = 36

# Minimum column width (pt) for column-width optimisation.
# Columns narrower than this are "phantom" (invisible/accidental).
# Using 20 pt (variant C): only truly phantom columns are redistributed;
# legitimate narrow columns (e.g. 30 pt numbering column) are left as-is.
_MIN_COL_PT = 20.0   # ≈ 0.7 cm — only phantom columns


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
_CONT_NUM_RE = re.compile(
    r"продолжение\s+таблицы\s+(\d+(?:\.\d+){0,2})",
    re.IGNORECASE,
)


def _extract_table_num(text: str) -> str | None:
    m = _TBL_NUM_RE.search(text.strip())
    return m.group(1) if m else None


# ── Table merging / continuation detection helpers ────────────────────────────
# (splitting/merging logic is stubbed — see module docstring for FUTURE plan)

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


def _norm_text(text: str) -> str:
    return " ".join((text or "").split())


def _is_any_continuation_marker(text: str) -> bool:
    t = _norm_text(text)
    return bool(t and t.lower().startswith("продолжение таблицы"))


def _table_col_count(tbl_xml) -> int:
    grid = tbl_xml.find(qn("w:tblGrid"))
    if grid is not None:
        cols = grid.findall(qn("w:gridCol"))
        if cols:
            return len(cols)

    first_row = tbl_xml.find(qn("w:tr"))
    if first_row is None:
        return 0
    count = 0
    for tc in first_row.findall(qn("w:tc")):
        tcPr = tc.find(qn("w:tcPr"))
        gs = tcPr.find(qn("w:gridSpan")) if tcPr is not None else None
        span = int(gs.get(qn("w:val"), 1)) if gs is not None else 1
        count += max(1, span)
    return count


def _row_cell_texts(tr_xml) -> list[str]:
    vals: list[str] = []
    for tc in tr_xml.findall(qn("w:tc")):
        txt = "".join(
            (t.text or "")
            for t in tc.findall(".//" + qn("w:t"))
            if t.text
        )
        vals.append(_norm_text(txt))
    return vals


def _rows_match(row1_xml, row2_xml) -> bool:
    return _row_cell_texts(row1_xml) == _row_cell_texts(row2_xml)


@dataclass(frozen=True)
class RowSignature:
    row_idx: int
    key: str
    fragments: tuple[str, ...]


@dataclass(frozen=True)
class TableSignature:
    table_idx: int
    tbl_xml: object
    rows: tuple[RowSignature, ...]


@dataclass(frozen=True)
class RenderedSplitCandidate:
    table_idx: int
    tbl_xml: object
    split_after: int


@dataclass(frozen=True)
class RenderedWholeTableMoveCandidate:
    table_idx: int
    tbl_xml: object
    caption_para_xml: object


_START_HAS_COMPLETE_DATA_ROW = "has_complete_data_row"
_START_NO_COMPLETE_DATA_ROW = "no_complete_data_row"
_START_AMBIGUOUS = "ambiguous"


def _tbl_has_at_least_two_rows(tbl_xml) -> bool:
    return len(tbl_xml.findall(qn("w:tr"))) >= 2


def _is_vmerge_continue(tc_xml) -> bool:
    tcPr = tc_xml.find(qn("w:tcPr"))
    if tcPr is None:
        return False
    vm = tcPr.find(qn("w:vMerge"))
    if vm is None:
        return False
    val = vm.get(qn("w:val"))
    # w:vMerge with no val is "continue" by spec.
    return val is None or val == "continue"


def _is_split_boundary_safe(rows_xml: list, split_after: int) -> bool:
    """
    split_after is index of the last row in part 1.
    Boundary is between rows[split_after] and rows[split_after+1].
    """
    if split_after < 0 or split_after + 1 >= len(rows_xml):
        return False
    next_row = rows_xml[split_after + 1]
    for tc in next_row.findall(qn("w:tc")):
        if _is_vmerge_continue(tc):
            return False
    return True


def _find_safe_split_after(rows_xml: list, candidate_after: int) -> int | None:
    """
    Move split boundary upward until it is safe and leaves at least
    header + 1 data row in part 1.
    """
    s = candidate_after
    while s >= 1:
        if _is_split_boundary_safe(rows_xml, s):
            return s
        s -= 1
    return None


def _find_caption_number_before_table(doc: Document, tbl_xml) -> str | None:
    """
    Strict source of truth: caption paragraph before the table.
    Supports:
      - "Таблица X.X"
      - "Таблица X.X.X"
    and two-paragraph format (caption line + title line).
    """
    body = doc.element.body
    children = list(body)
    try:
        idx = children.index(tbl_xml)
    except ValueError:
        return None

    # Build a fast map of paragraph XML -> paragraph text
    para_text = {p._element: _norm_text(p.text) for p in doc.paragraphs}

    j = idx - 1
    nonempty_seen = 0
    while j >= 0 and nonempty_seen < 4:
        node = children[j]
        if node.tag == qn("w:p"):
            txt = para_text.get(node, "")
            if txt:
                nonempty_seen += 1
                m = _TBL_NUM_RE.match(txt)
                if m:
                    return m.group(1)
        elif node.tag == qn("w:tbl"):
            break
        j -= 1
    return None


def _find_caption_paragraph_before_table(doc: Document, tbl_xml):
    """
    Return the strict table caption paragraph XML and number before tbl_xml.
    The caption paragraph, not the title paragraph, is the only safe anchor for
    whole-table moves.
    """
    body = doc.element.body
    children = list(body)
    try:
        idx = children.index(tbl_xml)
    except ValueError:
        return None

    para_text = {p._element: _norm_text(p.text) for p in doc.paragraphs}

    j = idx - 1
    nonempty_seen = 0
    while j >= 0 and nonempty_seen < 4:
        node = children[j]
        if node.tag == qn("w:p"):
            txt = para_text.get(node, "")
            if txt:
                nonempty_seen += 1
                m = _TBL_NUM_RE.match(txt)
                if m:
                    return node, m.group(1)
        elif node.tag == qn("w:tbl"):
            break
        j -= 1
    return None


def _norm_match_text(text: str) -> str:
    return _norm_text(text).lower()


def _row_signature(tr_xml, row_idx: int) -> RowSignature | None:
    fragments = tuple(
        frag
        for frag in (_norm_match_text(t) for t in _row_cell_texts(tr_xml))
        if frag
    )
    if not fragments:
        return None
    return RowSignature(row_idx=row_idx, key=" || ".join(fragments), fragments=fragments)


def _collect_table_signatures(doc: Document) -> list[TableSignature]:
    out: list[TableSignature] = []
    for table_idx, table in enumerate(doc.tables):
        rows: list[RowSignature] = []
        for row_idx, tr in enumerate(table._tbl.findall(qn("w:tr"))):
            sig = _row_signature(tr, row_idx)
            if sig is not None:
                rows.append(sig)
        out.append(TableSignature(table_idx=table_idx, tbl_xml=table._tbl, rows=tuple(rows)))
    return out


def _valid_manual_continuation_table_ids(doc: Document) -> set[int]:
    """
    Return table XML ids that are already part of a valid manual continuation.
    Valid manual chains must be preserved exactly.
    """
    body = doc.element.body
    children = list(body)
    skip: set[int] = set()

    i = 1
    while i < len(children) - 1:
        prev_node = children[i - 1]
        node = children[i]
        next_node = children[i + 1]

        if prev_node.tag != qn("w:tbl") or node.tag != qn("w:p") or next_node.tag != qn("w:tbl"):
            i += 1
            continue

        p_obj = next((p for p in doc.paragraphs if p._element is node), None)
        marker_text = _norm_text(p_obj.text if p_obj is not None else "")
        if not _is_any_continuation_marker(marker_text):
            i += 1
            continue

        if _is_valid_manual_continuation_chain(doc, prev_node, node, next_node):
            skip.add(id(prev_node))
            skip.add(id(next_node))

        i += 1

    return skip


def _paragraph_has_keep_next(p_xml) -> bool:
    pPr = p_xml.find(qn("w:pPr"))
    if pPr is None:
        return False
    keep = pPr.find(qn("w:keepNext"))
    if keep is None:
        return False
    return keep.get(qn("w:val")) not in {"0", "false", "False"}


def _paragraph_is_right_aligned(p_xml) -> bool:
    pPr = p_xml.find(qn("w:pPr"))
    if pPr is None:
        return False
    jc = pPr.find(qn("w:jc"))
    return bool(jc is not None and jc.get(qn("w:val")) == "right")


def _manual_marker_matches_caption(doc: Document, tbl_xml, marker_text: str) -> bool:
    caption_num = _find_caption_number_before_table(doc, tbl_xml)
    marker_match = _CONT_NUM_RE.search(marker_text)
    marker_num = marker_match.group(1) if marker_match else None
    if caption_num is None and marker_num is None:
        return True
    return caption_num == marker_num


def _is_valid_manual_continuation_chain(doc: Document, tbl1, marker_p, tbl2) -> bool:
    marker_text = ""
    for text_node in marker_p.findall(".//" + qn("w:t")):
        marker_text += text_node.text or ""
    marker_text = _norm_text(marker_text)
    if not _is_any_continuation_marker(marker_text):
        return False
    if not _manual_marker_matches_caption(doc, tbl1, marker_text):
        return False
    if not _paragraph_is_right_aligned(marker_p):
        return False
    if not _paragraph_has_keep_next(marker_p):
        return False

    compatible = _table_col_count(tbl1) == _table_col_count(tbl2) and _table_col_count(tbl1) > 0
    rows1 = tbl1.findall(qn("w:tr"))
    rows2 = tbl2.findall(qn("w:tr"))
    headers_match = bool(rows1 and rows2 and _rows_match(rows1[0], rows2[0]))
    return compatible and headers_match and _tbl_has_at_least_two_rows(tbl2)


def _row_matches_line(sig: RowSignature, line_text: str) -> bool:
    pos = 0
    for fragment in sig.fragments:
        found = line_text.find(fragment, pos)
        if found < 0:
            return False
        pos = found + len(fragment)
    return True


def _match_row_pages(table_sig: TableSignature, pdf_lines: list[PdfLine]) -> dict[int, int] | None:
    data_rows = [sig for sig in table_sig.rows if sig.row_idx > 0]
    if len(data_rows) < 2:
        return None

    keys = [sig.key for sig in data_rows]
    if len(keys) != len(set(keys)):
        return None

    line_texts = [(_norm_match_text(line.text), line.page_num) for line in pdf_lines]
    row_pages: dict[int, int] = {}
    last_match_idx = -1

    for sig in data_rows:
        matches = [
            (idx, page_num)
            for idx, (line_text, page_num) in enumerate(line_texts)
            if idx > last_match_idx and _row_matches_line(sig, line_text)
        ]
        if len(matches) != 1:
            return None
        last_match_idx, page_num = matches[0]
        row_pages[sig.row_idx] = page_num

    return row_pages


_TOKEN_RE = re.compile(r"[0-9A-Za-zА-Яа-яЁё]+")


def _distinctive_tokens(text: str) -> set[str]:
    tokens = {
        token.lower()
        for token in _TOKEN_RE.findall(_norm_text(text))
        if len(token) >= 4 and not token.isdigit()
    }
    return tokens


def _row_distinctive_tokens(sig: RowSignature) -> set[str]:
    out: set[str] = set()
    for fragment in sig.fragments:
        out.update(_distinctive_tokens(fragment))
    return out


def _unique_data_row_tokens(data_rows: list[RowSignature]) -> dict[int, set[str]]:
    all_tokens: dict[str, int] = {}
    row_tokens: dict[int, set[str]] = {}
    for sig in data_rows:
        tokens = _row_distinctive_tokens(sig)
        row_tokens[sig.row_idx] = tokens
        for token in tokens:
            all_tokens[token] = all_tokens.get(token, 0) + 1
    return {
        row_idx: {token for token in tokens if all_tokens[token] == 1}
        for row_idx, tokens in row_tokens.items()
    }


def _line_matches_caption_number(line_text: str, num: str) -> bool:
    m = _TBL_NUM_RE.match(_norm_text(line_text))
    return bool(m and m.group(1) == num)


_DISABLED_PAGE_BREAK_VALUES = {"0", "false", "False", "off"}


def _is_active_page_break_before(page_break_elem) -> bool:
    if page_break_elem is None:
        return False
    return page_break_elem.get(qn("w:val")) not in _DISABLED_PAGE_BREAK_VALUES


def _find_page_break_before(pPr):
    if pPr is None:
        return None
    return pPr.find(qn("w:pageBreakBefore"))


def _pdf_caption_match_count(caption_num: str, pdf_lines: list[PdfLine]) -> int:
    return sum(1 for line in pdf_lines if _line_matches_caption_number(line.text, caption_num))


def _row_has_any_token_in_text(sig: RowSignature, text: str) -> bool:
    tokens = _row_distinctive_tokens(sig)
    if not tokens:
        return False
    text_tokens = _distinctive_tokens(text)
    return bool(tokens & text_tokens)


def _tokens_in_text(tokens: set[str], text: str) -> bool:
    if not tokens:
        return False
    return bool(tokens & _distinctive_tokens(text))


def _has_complete_data_row_in_page_window(
    data_rows: list[RowSignature],
    unique_tokens: dict[int, set[str]],
    page_texts: list[str],
) -> bool:
    max_window = 4
    for sig in data_rows:
        tokens = unique_tokens[sig.row_idx]
        if len(tokens) < 2:
            continue
        for start in range(len(page_texts)):
            for end in range(start + 1, min(len(page_texts), start + max_window) + 1):
                window_text = " ".join(page_texts[start:end])
                if tokens <= _distinctive_tokens(window_text):
                    return True
    return False


def _header_line_indexes(header: RowSignature, page_texts: list[str]) -> list[int]:
    header_search_limit = min(len(page_texts), 12)
    return [
        idx
        for idx, line_text in enumerate(page_texts[:header_search_limit])
        if _row_matches_line(header, line_text) or _row_has_any_token_in_text(header, line_text)
    ]


def _first_data_row_spills_to_next_page(
    first_row: RowSignature,
    first_row_tokens: set[str],
    header: RowSignature,
    caption_idx: int,
    start_page: int,
    pdf_lines: list[PdfLine],
    data_page_texts: list[str],
) -> bool:
    if len(first_row_tokens) < 2:
        return False

    start_page_joined = " ".join(data_page_texts)
    start_page_distinctive = _distinctive_tokens(start_page_joined)
    start_tokens = first_row_tokens & start_page_distinctive
    if len(start_tokens) < 2:
        return False
    next_only_tokens = first_row_tokens - start_tokens
    if not next_only_tokens:
        return False

    next_page = start_page + 1
    next_page_lines = [
        line
        for idx, line in enumerate(pdf_lines)
        if idx > caption_idx and line.page_num == next_page
    ]
    if not next_page_lines:
        return False

    next_page_texts = [_norm_match_text(line.text) for line in next_page_lines]
    next_header_indexes = _header_line_indexes(header, next_page_texts)
    if not next_header_indexes:
        return False

    next_page_data_texts = next_page_texts[(max(next_header_indexes) + 1):]
    if not next_page_data_texts:
        return False

    # Conservative continuation evidence: the next page should expose a very
    # short residue of the same first row right after the repeated header, not
    # a long prose line that happens to reuse one token.
    for start in range(min(len(next_page_data_texts), 4)):
        for end in range(start + 1, min(len(next_page_data_texts), start + 2) + 1):
            window_tokens = _distinctive_tokens(" ".join(next_page_data_texts[start:end]))
            if len(window_tokens) <= 3 and (window_tokens & next_only_tokens):
                return True

    return False


def _classify_start_page_usability(
    table_sig: TableSignature,
    caption_num: str,
    pdf_lines: list[PdfLine],
) -> str:
    """
    Conservative Patch 2.1 detector.

    It does not reconstruct the table. It only answers whether the rendered
    caption page contains one clearly complete data row. Ambiguous evidence is
    intentionally treated as no-op by the caller.
    """
    caption_matches = [
        (idx, line)
        for idx, line in enumerate(pdf_lines)
        if _line_matches_caption_number(line.text, caption_num)
    ]
    if len(caption_matches) != 1:
        return _START_AMBIGUOUS

    caption_idx, caption_line = caption_matches[0]
    start_page = caption_line.page_num
    same_page_lines = [
        line
        for idx, line in enumerate(pdf_lines)
        if idx > caption_idx and line.page_num == start_page
    ]
    if not same_page_lines:
        return _START_AMBIGUOUS

    header = next((sig for sig in table_sig.rows if sig.row_idx == 0), None)
    data_rows = [sig for sig in table_sig.rows if sig.row_idx > 0]
    if header is None or not data_rows:
        return _START_AMBIGUOUS

    data_keys = [sig.key for sig in data_rows]
    if len(data_keys) != len(set(data_keys)):
        return _START_AMBIGUOUS
    unique_tokens = _unique_data_row_tokens(data_rows)
    if any(not unique_tokens.get(sig.row_idx) for sig in data_rows):
        return _START_AMBIGUOUS

    same_page_texts = [_norm_match_text(line.text) for line in same_page_lines]
    same_page_joined = " ".join(same_page_texts)
    header_line_indexes = _header_line_indexes(header, same_page_texts)
    if not header_line_indexes and not _row_has_any_token_in_text(header, same_page_joined):
        return _START_AMBIGUOUS

    data_page_texts = same_page_texts[(max(header_line_indexes) + 1):] if header_line_indexes else same_page_texts
    data_page_joined = " ".join(data_page_texts)

    for sig in data_rows:
        if any(_row_matches_line(sig, line_text) for line_text in data_page_texts):
            return _START_HAS_COMPLETE_DATA_ROW
    if _has_complete_data_row_in_page_window(data_rows, unique_tokens, data_page_texts):
        return _START_HAS_COMPLETE_DATA_ROW

    first_row = data_rows[0]
    if _first_data_row_spills_to_next_page(
        first_row=first_row,
        first_row_tokens=unique_tokens[first_row.row_idx],
        header=header,
        caption_idx=caption_idx,
        start_page=start_page,
        pdf_lines=pdf_lines,
        data_page_texts=data_page_texts,
    ):
        return _START_NO_COMPLETE_DATA_ROW

    rows_with_start_page_tokens = [
        sig for sig in data_rows if _tokens_in_text(unique_tokens[sig.row_idx], data_page_joined)
    ]
    if not rows_with_start_page_tokens:
        return _START_NO_COMPLETE_DATA_ROW

    later_text = " ".join(
        _norm_match_text(line.text)
        for idx, line in enumerate(pdf_lines)
        if idx > caption_idx and line.page_num > start_page
    )
    split_like_rows = [
        sig for sig in rows_with_start_page_tokens
        if _tokens_in_text(unique_tokens[sig.row_idx], later_text)
    ]
    if len(split_like_rows) == 1:
        return _START_NO_COMPLETE_DATA_ROW

    return _START_AMBIGUOUS


def _find_rendered_whole_table_move_candidate(
    doc: Document,
    pdf_lines: list[PdfLine],
    diagnostics: dict[str, bool] | None = None,
) -> RenderedWholeTableMoveCandidate | None:
    manual_skip = _valid_manual_continuation_table_ids(doc)
    inspected = 0

    for table_sig in _collect_table_signatures(doc):
        inspected += 1
        if id(table_sig.tbl_xml) in manual_skip:
            logger.info(
                "rendered_whole_table_candidate table_idx=%s skip=valid_manual_continuation",
                table_sig.table_idx,
            )
            continue

        caption = _find_caption_paragraph_before_table(doc, table_sig.tbl_xml)
        if caption is None:
            logger.info(
                "rendered_whole_table_candidate table_idx=%s skip=caption_missing",
                table_sig.table_idx,
            )
            continue
        caption_para_xml, caption_num = caption
        pdf_caption_matches = _pdf_caption_match_count(caption_num, pdf_lines)
        caption_pPr = caption_para_xml.find(qn("w:pPr"))
        if _is_active_page_break_before(_find_page_break_before(caption_pPr)):
            logger.info(
                "rendered_whole_table_candidate table_idx=%s caption=%s pdf_caption_matches=%s strict_caption_found=%s skip=existing_active_page_break",
                table_sig.table_idx,
                caption_num,
                pdf_caption_matches,
                pdf_caption_matches == 1,
            )
            continue

        usability = _classify_start_page_usability(table_sig, caption_num, pdf_lines)
        logger.info(
            "rendered_whole_table_candidate table_idx=%s caption=%s pdf_caption_matches=%s strict_caption_found=%s start_page_usability=%s",
            table_sig.table_idx,
            caption_num,
            pdf_caption_matches,
            pdf_caption_matches == 1,
            usability,
        )
        if usability == _START_NO_COMPLETE_DATA_ROW:
            logger.info(
                "rendered_whole_table_candidate_selected table_idx=%s caption=%s reason=%s",
                table_sig.table_idx,
                caption_num,
                usability,
            )
            return RenderedWholeTableMoveCandidate(
                table_idx=table_sig.table_idx,
                tbl_xml=table_sig.tbl_xml,
                caption_para_xml=caption_para_xml,
            )

        if usability == _START_AMBIGUOUS and diagnostics is not None:
            diagnostics["ambiguous"] = True
        logger.info(
            "rendered_whole_table_candidate table_idx=%s caption=%s skip=%s",
            table_sig.table_idx,
            caption_num,
            "ambiguous" if usability == _START_AMBIGUOUS else "has_complete_data_row",
        )

    logger.info("rendered_whole_table_no_candidate inspected=%s", inspected)
    return None


def _ensure_page_break_before(para_elem) -> bool:
    pPr = para_elem.find(qn("w:pPr"))
    page_break = _find_page_break_before(pPr)
    if _is_active_page_break_before(page_break):
        return False
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        para_elem.insert(0, pPr)
    if page_break is None:
        pPr.append(OxmlElement("w:pageBreakBefore"))
    else:
        page_break.attrib.pop(qn("w:val"), None)
    return True


def _find_rendered_split_candidate(
    doc: Document,
    pdf_lines: list[PdfLine],
    diagnostics: dict[str, bool] | None = None,
) -> RenderedSplitCandidate | None:
    manual_skip = _valid_manual_continuation_table_ids(doc)
    inspected = 0

    for table_sig in _collect_table_signatures(doc):
        inspected += 1
        if id(table_sig.tbl_xml) in manual_skip:
            logger.info(
                "rendered_split_candidate table_idx=%s skip=valid_manual_continuation",
                table_sig.table_idx,
            )
            continue

        rows_xml = table_sig.tbl_xml.findall(qn("w:tr"))
        if len(rows_xml) < 3:
            logger.info(
                "rendered_split_candidate table_idx=%s rows=%s skip=too_few_rows",
                table_sig.table_idx,
                len(rows_xml),
            )
            continue

        row_pages = _match_row_pages(table_sig, pdf_lines)
        if row_pages is None:
            if diagnostics is not None:
                diagnostics["ambiguous"] = True
            logger.info(
                "rendered_split_candidate table_idx=%s rows=%s skip=row_mapping_ambiguous",
                table_sig.table_idx,
                len(rows_xml),
            )
            continue

        page_boundary_found = False
        for row_idx in sorted(row_pages):
            next_idx = row_idx + 1
            if next_idx not in row_pages:
                continue
            if row_pages[row_idx] < row_pages[next_idx]:
                page_boundary_found = True
                safe_after = _find_safe_split_after(rows_xml, row_idx)
                if safe_after is None or safe_after < 1:
                    if diagnostics is not None:
                        diagnostics["ambiguous"] = True
                    logger.info(
                        "rendered_split_candidate table_idx=%s row_idx=%s skip=merged_boundary_conflict",
                        table_sig.table_idx,
                        row_idx,
                    )
                    return None
                if len(rows_xml) - (safe_after + 1) < 1:
                    logger.info(
                        "rendered_split_candidate table_idx=%s row_idx=%s safe_after=%s skip=no_continuation_data_row",
                        table_sig.table_idx,
                        row_idx,
                        safe_after,
                    )
                    return None
                logger.info(
                    "rendered_split_candidate_selected table_idx=%s row_idx=%s split_after=%s",
                    table_sig.table_idx,
                    row_idx,
                    safe_after,
                )
                return RenderedSplitCandidate(
                    table_idx=table_sig.table_idx,
                    tbl_xml=table_sig.tbl_xml,
                    split_after=safe_after,
                )
        if not page_boundary_found:
            logger.info(
                "rendered_split_candidate table_idx=%s rows=%s skip=no_page_boundary",
                table_sig.table_idx,
                len(rows_xml),
            )

    logger.info("rendered_split_no_candidate inspected=%s", inspected)
    return None


def _build_continuation_para(text: str):
    """
    Create:
      - right align
      - Times New Roman 14 pt
      - no first-line indent
      - keepWithNext=True
    """
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    p.append(pPr)

    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "right")
    pPr.append(jc)

    ind = OxmlElement("w:ind")
    ind.set(qn("w:firstLine"), "0")
    ind.set(qn("w:left"), "0")
    pPr.append(ind)

    keep_next = OxmlElement("w:keepNext")
    pPr.append(keep_next)

    r = OxmlElement("w:r")
    p.append(r)
    rPr = OxmlElement("w:rPr")
    r.append(rPr)

    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "Times New Roman")
    fonts.set(qn("w:hAnsi"), "Times New Roman")
    fonts.set(qn("w:cs"), "Times New Roman")
    rPr.append(fonts)

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "28")  # 14 pt
    rPr.append(sz)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), "28")
    rPr.append(szCs)

    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    return p


def _split_table_at(doc: Document, tbl_xml, split_after: int, continuation_text: str) -> bool:
    rows = tbl_xml.findall(qn("w:tr"))
    if len(rows) < 3:  # header + at least 2 data rows to split
        return False
    if split_after < 1 or split_after >= len(rows) - 1:
        return False

    header_row = deepcopy(rows[0])
    tail_rows = [deepcopy(r) for r in rows[split_after + 1:]]
    if not tail_rows:
        return False

    # part2 must have at least header + 1 data row
    if len(tail_rows) < 1:
        return False

    tbl2 = deepcopy(tbl_xml)
    for tr in list(tbl2.findall(qn("w:tr"))):
        tbl2.remove(tr)
    tbl2.append(header_row)
    for tr in tail_rows:
        tbl2.append(tr)

    # mark repeated header row
    trPr = header_row.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        header_row.insert(0, trPr)
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))

    # trim part1
    for tr in rows[split_after + 1:]:
        tbl_xml.remove(tr)

    body = doc.element.body
    marker = _build_continuation_para(continuation_text)
    tbl_xml.addnext(marker)
    marker.addnext(tbl2)
    return True



_NUMERIC_CELL_RE = re.compile(r"^[\d\s\+\-−–,.%]+$")
_PT_PER_CHAR_NUMERIC = 6.0   # approx pt/char for 12pt TNR digits
_CELL_H_PADDING = 8.0        # left+right cell padding (pt) added to content width


def _compute_col_minimums(tbl_xml, n_cols: int) -> list[float]:
    """
    Compute per-column minimum widths (pt) in a single pass over all rows.

    For cells containing only numbers/symbols (no letters), the minimum is
    set to the width needed to render the longest value on one line:
        min_w = len(text) × _PT_PER_CHAR_NUMERIC + _CELL_H_PADDING

    For all other cells (header or text), the minimum falls back to _MIN_COL_PT.
    This protects numeric columns (e.g. "9 503 005") from being scaled so narrow
    that values wrap to multiple lines.

    Only single-column cells (gridSpan = 1) are considered.
    """
    minimums = [_MIN_COL_PT] * n_cols

    for tr in tbl_xml.findall(qn("w:tr")):
        col_idx = 0
        for tc in tr.findall(qn("w:tc")):
            if col_idx >= n_cols:
                break
            tcPr = tc.find(qn("w:tcPr"))
            gs = tcPr.find(qn("w:gridSpan")) if tcPr is not None else None
            span = int(gs.get(qn("w:val"), 1)) if gs is not None else 1
            span = max(1, min(span, n_cols - col_idx))

            if span == 1:
                for p_el in tc.findall(".//" + qn("w:p")):
                    cell_text = "".join(
                        (r.find(qn("w:t")).text or "")
                        for r in p_el.findall(qn("w:r"))
                        if r.find(qn("w:t")) is not None
                    ).strip()
                    if cell_text and _NUMERIC_CELL_RE.match(cell_text):
                        content_w = len(cell_text) * _PT_PER_CHAR_NUMERIC + _CELL_H_PADDING
                        if content_w > minimums[col_idx]:
                            minimums[col_idx] = content_w

            col_idx += span

    return minimums


def _optimize_table_col_widths(tbl_xml, body_width_pt: float) -> bool:
    """
    Ensure no column is narrower than its content minimum and total width ≤ body_width_pt.

    Algorithm:
      1. Scale all columns down proportionally if total > body_width_pt.
      2. Identify undersized columns (based on content-aware per-column minimums);
         redistribute deficit from wider donor columns.

    The per-column minimums are content-aware: numeric-only cells (digits, spaces,
    punctuation) set a minimum wide enough to display their content on one line.
    This prevents number columns from being scaled too narrow when proportionally
    shrinking a wide table.

    Updates both w:tblGrid/w:gridCol, w:tblPr/w:tblW, and each w:tc/w:tcPr/w:tcW
    (honouring w:gridSpan for merged cells).

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

    # Content-aware per-column minimums (protects numeric columns from over-shrinking)
    col_mins = _compute_col_minimums(tbl_xml, n)

    changed = False

    # Step 1: scale down if total exceeds body width
    if total > body_width_pt + 0.5:
        scale = body_width_pt / total
        widths = [w * scale for w in widths]
        total = sum(widths)
        changed = True

    # Step 2: redistribute to fix undersized columns (up to n iterations).
    # Uses per-column minimums: numeric columns have higher minimums to keep
    # values on one line; other columns use the global _MIN_COL_PT floor.
    for _ in range(n):
        undersized = [(i, col_mins[i] - widths[i]) for i in range(n)
                      if widths[i] < col_mins[i] - 0.5]
        if not undersized:
            break
        donors = [i for i in range(n) if widths[i] > col_mins[i] + 0.5]
        if not donors:
            break
        total_deficit = sum(d for _, d in undersized)
        total_donor_excess = sum(widths[i] - col_mins[i] for i in donors)
        take_frac = min(1.0, total_donor_excess / total_deficit)

        for i, deficit in undersized:
            widths[i] += deficit * take_frac
        actual_taken = total_deficit * take_frac
        for i in donors:
            donor_excess = widths[i] - col_mins[i]
            widths[i] -= actual_taken * (donor_excess / total_donor_excess)
        changed = True

    if not changed:
        return False

    # Round to integer twips, keep total consistent
    twip_widths = [max(1, round(w * TWIP_PER_PT)) for w in widths]

    # Apply to grid
    for col_el, tw in zip(gridcols, twip_widths):
        col_el.set(qn("w:w"), str(tw))

    # Update w:tblPr/w:tblW to the new column total.
    # Without this, Word uses the original (too-wide) tblW as master table width
    # and ignores the corrected gridCol / tcW values.
    tblPr = tbl_xml.find(qn("w:tblPr"))
    if tblPr is not None:
        tblW = tblPr.find(qn("w:tblW"))
        if tblW is None:
            tblW = OxmlElement("w:tblW")
            tblPr.append(tblW)
        tblW.set(qn("w:w"), str(sum(twip_widths)))
        tblW.set(qn("w:type"), "dxa")

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
    Phase 3 pre-pass — STUB (table splitting/merging disabled).

    Previously: detected student-split table pairs (table + "Продолжение
    таблицы X" paragraph + continuation table) and merged them back into one
    table so apply_table_continuation could re-split at the real page boundary.

    Disabled because reliable page-break detection requires a rendering engine
    (LibreOffice / Word) — pure geometry estimation was too unreliable.
    See module docstring for the FUTURE implementation plan.

    Returns 0 (no changes made).
    """
    body = doc.element.body
    children = list(body)
    merges = 0

    i = 1
    while i < len(children) - 1:
        prev_node = children[i - 1]
        node = children[i]
        next_node = children[i + 1]

        if prev_node.tag != qn("w:tbl") or node.tag != qn("w:p") or next_node.tag != qn("w:tbl"):
            i += 1
            continue

        p_obj = next((p for p in doc.paragraphs if p._element is node), None)
        marker_text = _norm_text(p_obj.text if p_obj is not None else "")
        if not _is_any_continuation_marker(marker_text):
            i += 1
            continue

        tbl1 = prev_node
        tbl2 = next_node

        rows1 = tbl1.findall(qn("w:tr"))
        rows2 = tbl2.findall(qn("w:tr"))
        headers_match = bool(rows1 and rows2 and _rows_match(rows1[0], rows2[0]))
        keep_manual_split = _is_valid_manual_continuation_chain(doc, tbl1, node, tbl2)

        if keep_manual_split:
            i += 1
            continue

        # Rebuild invalid split: merge tbl2 into tbl1, skipping duplicate header if present.
        start_idx = 1 if headers_match else 0
        for tr in rows2[start_idx:]:
            tbl1.append(deepcopy(tr))

        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
        parent2 = tbl2.getparent()
        if parent2 is not None:
            parent2.remove(tbl2)
        merges += 1

        # refresh snapshot after mutations
        children = list(body)
        i = max(1, i - 1)

    return merges


# ── Main entry point ──────────────────────────────────────────────────────────

def apply_table_continuation(
    doc: Document,
    report: FormattingReport | None = None,
) -> int:
    """
    Phase 3 Rule 1 — STUB (table page-break splitting disabled).

    Still active: column-width optimisation (_optimize_table_col_widths) runs
    for ALL tables — fixes oversized / phantom-narrow columns regardless of
    whether splitting is enabled.

    The splitting part is disabled because reliable page-break detection
    requires a rendering engine.  See module docstring for the FUTURE plan.

    Returns the number of tables whose widths were normalised.
    Does not split tables or insert continuation markers.
    """
    # ── Column-width optimisation (always active) ──────────────────────────
    body_w = _body_width_pt(doc)
    n_col_fixed = 0
    for kind, tbl_xml, _ in _iter_body(doc):
        if kind != "table":
            continue
        if _optimize_table_col_widths(tbl_xml, body_w):
            n_col_fixed += 1
    if n_col_fixed:
        logger.info("table_continuation: col-width optimised %d table(s)", n_col_fixed)

    return n_col_fixed


def _warn_rendered_split_unavailable(
    report: FormattingReport | None,
    reason: str,
) -> None:
    logger.info("rendered table continuation skipped: %s", reason)
    if report is not None:
        report.warn("Автоперенос таблиц по PDF временно недоступен")


@dataclass(frozen=True)
class _MarkerSplitDecision:
    eligible: bool
    split_before_row: int | None
    skip_reason: str | None


def _marker_split_enabled() -> bool:
    return os.getenv("KPFU_ENABLE_MARKER_SPLIT", "").strip().lower() in {
        "1", "true", "yes", "on",
    }


def _marker_split_apply_enabled() -> bool:
    return os.getenv("KPFU_APPLY_MARKER_SPLIT", "").strip().lower() in {
        "1", "true", "yes", "on",
    }


def _classify_marker_duplicate_rows(
    diagnostic,
    *,
    header_rows: int = 1,
) -> tuple[str, list[int]]:
    if not diagnostic.duplicate_rows:
        return "none", []

    data_duplicate_rows = sorted(
        row_index
        for row_index in diagnostic.duplicate_rows
        if row_index >= header_rows
    )
    if data_duplicate_rows:
        return "data_rows", data_duplicate_rows

    return "header_only", sorted(diagnostic.duplicate_rows)


def _non_header_rows_are_clean(
    diagnostic,
    *,
    header_rows: int = 1,
) -> bool:
    if any(row_index >= header_rows for row_index in diagnostic.missing_rows):
        return False
    if any(row_index >= header_rows for row_index in diagnostic.duplicate_rows):
        return False

    row_pages = {
        row_index: page_num
        for row_index, page_num in sorted(diagnostic.row_pages.items())
        if row_index >= header_rows
    }
    expected_rows = list(range(header_rows, diagnostic.rows_count))
    return list(row_pages) == expected_rows


def _is_header_only_duplicate_safe(
    diagnostic,
    *,
    header_rows: int = 1,
) -> bool:
    duplicate_classification, _ = _classify_marker_duplicate_rows(
        diagnostic,
        header_rows=header_rows,
    )
    return (
        duplicate_classification == "header_only"
        and _non_header_rows_are_clean(diagnostic, header_rows=header_rows)
    )


def _evaluate_marker_split_diagnostic(
    diagnostic,
    *,
    header_rows: int = 1,
) -> _MarkerSplitDecision:
    if diagnostic.error_message:
        return _MarkerSplitDecision(False, None, "mapping_error")
    if len(diagnostic.pages_detected) != 2:
        return _MarkerSplitDecision(False, None, "not_2_pages")

    duplicate_classification, _ = _classify_marker_duplicate_rows(
        diagnostic,
        header_rows=header_rows,
    )
    if duplicate_classification == "data_rows":
        return _MarkerSplitDecision(False, None, "duplicate_rows")
    if diagnostic.missing_rows not in ([], [0]):
        return _MarkerSplitDecision(False, None, "missing_rows_outside_header")
    if (
        duplicate_classification == "header_only"
        and not _is_header_only_duplicate_safe(diagnostic, header_rows=header_rows)
    ):
        return _MarkerSplitDecision(False, None, "duplicate_rows")

    row_pages = {
        row_index: page_num
        for row_index, page_num in sorted(diagnostic.row_pages.items())
        if row_index >= header_rows
    }
    expected_rows = list(range(header_rows, diagnostic.rows_count))
    if list(row_pages) != expected_rows:
        return _MarkerSplitDecision(False, None, "no_boundary")

    first_page = None
    second_page = None
    split_before_row = None
    expected_first_page, expected_second_page = diagnostic.pages_detected

    for row_index, page_num in row_pages.items():
        if page_num not in diagnostic.pages_detected:
            return _MarkerSplitDecision(False, None, "non_monotonic_pages")

        if first_page is None:
            if page_num != expected_first_page:
                return _MarkerSplitDecision(False, None, "non_monotonic_pages")
            first_page = page_num
            continue

        if second_page is None:
            if page_num == first_page:
                continue
            if page_num != expected_second_page:
                return _MarkerSplitDecision(False, None, "non_monotonic_pages")
            second_page = page_num
            split_before_row = row_index
            continue

        if page_num != second_page:
            return _MarkerSplitDecision(False, None, "non_monotonic_pages")

    if split_before_row is None:
        return _MarkerSplitDecision(False, None, "no_boundary")
    return _MarkerSplitDecision(True, split_before_row, None)


def _map_marker_split_apply_error(exc: Exception) -> str:
    text = str(exc).lower()
    if "standard table caption" in text:
        return "ordinary_without_standard_caption"
    if "tblgrid" in text or "grid" in text:
        return "unsupported_grid"
    if "complex merged header" in text:
        return "complex_merged_header"
    if "malformed" in text:
        return "malformed_numbered_row"
    if "header_rows" in text:
        return "unsupported_header_rows"
    return "mutation_error"


def _apply_marker_split_candidate(
    docx_path: Path,
    diagnostic,
    decision: _MarkerSplitDecision,
):
    doc = Document(str(docx_path))
    manual_skip = _valid_manual_continuation_table_ids(doc)
    if diagnostic.table_index in manual_skip:
        return None, "valid_manual_continuation"
    if not diagnostic.appendix_table and not diagnostic.has_standard_table_caption:
        return None, "ordinary_without_standard_caption"

    try:
        result = apply_numbered_split_to_document(
            doc,
            diagnostic.table_index,
            decision.split_before_row,
            header_rows=1,
            numbered_header=True,
            appendix_table=diagnostic.appendix_table,
            continuation_paragraph_builder=_build_continuation_para,
        )
    except Exception as exc:
        return None, _map_marker_split_apply_error(exc)

    if result.source_note_after_second is False:
        return None, "source_note_ordering_failed"

    doc.save(str(docx_path))
    return result, None


def _run_marker_split_detection_pass(docx_path: Path, *, apply_split: bool = False) -> int:
    from . import table_markers

    def _format_row_pages(row_pages: dict[int, int]) -> str:
        if not row_pages:
            return "-"
        return ",".join(
            f"{row_index}:{page_num}"
            for row_index, page_num in sorted(row_pages.items())
        )

    def _format_duplicate_rows(duplicate_rows: dict[int, list[int]]) -> str:
        if not duplicate_rows:
            return "-"
        return ",".join(
            f"{row_index}:{'/'.join(str(page) for page in pages)}"
            for row_index, pages in sorted(duplicate_rows.items())
        )

    def _format_page_spans(page_spans) -> str:
        if not page_spans:
            return "-"
        return ",".join(
            f"{span.start_row}-{span.end_row}:{span.page_num}"
            for span in page_spans
        )

    eligible_count = 0
    try:
        diagnostics = table_markers.diagnose_all_tables(docx_path, keep_temp=False)
    except Exception as exc:
        logger.info("marker_split_skipped reason=mapping_error error=%s", exc)
        return 0

    for diagnostic in diagnostics:
        logger.info(
            "marker_split_candidate table_index=%s rows=%s pages=%s row_pages=%s page_spans=%s missing_rows=%s duplicate_rows=%s",
            diagnostic.table_index,
            diagnostic.rows_count,
            diagnostic.pages_detected,
            _format_row_pages(diagnostic.row_pages),
            _format_page_spans(diagnostic.page_spans),
            diagnostic.missing_rows,
            _format_duplicate_rows(diagnostic.duplicate_rows),
        )
        duplicate_classification, duplicate_rows = _classify_marker_duplicate_rows(
            diagnostic,
            header_rows=1,
        )
        if duplicate_classification != "none":
            logger.info(
                "marker_split_duplicate_rows_classified table_index=%s classification=%s rows=%s missing_rows=%s duplicate_rows=%s page_spans=%s",
                diagnostic.table_index,
                duplicate_classification,
                duplicate_rows,
                diagnostic.missing_rows,
                _format_duplicate_rows(diagnostic.duplicate_rows),
                _format_page_spans(diagnostic.page_spans),
            )
        decision = _evaluate_marker_split_diagnostic(diagnostic, header_rows=1)
        if decision.eligible:
            if duplicate_classification == "header_only":
                logger.info(
                    "marker_split_header_duplicate_allowed table_index=%s missing_rows=%s duplicate_rows=%s page_spans=%s",
                    diagnostic.table_index,
                    diagnostic.missing_rows,
                    _format_duplicate_rows(diagnostic.duplicate_rows),
                    _format_page_spans(diagnostic.page_spans),
                )
            logger.info(
                "marker_split_boundary table_index=%s split_before_row=%s",
                diagnostic.table_index,
                decision.split_before_row,
            )
            logger.info(
                "marker_split_decision=ELIGIBLE table_index=%s",
                diagnostic.table_index,
            )
            eligible_count += 1
            if apply_split:
                result, skip_reason = _apply_marker_split_candidate(
                    docx_path,
                    diagnostic,
                    decision,
                )
                if result is not None:
                    logger.info(
                        "marker_split_applied table_index=%s split_before_row=%s first_rows=%s second_rows=%s appendix=%s continuation=%s",
                        diagnostic.table_index,
                        decision.split_before_row,
                        result.first_table_rows_count,
                        result.second_table_rows_count,
                        diagnostic.appendix_table,
                        result.continuation_paragraph_inserted,
                    )
                    return 1
                logger.info(
                    "marker_split_skipped table_index=%s reason=%s missing_rows=%s duplicate_rows=%s page_spans=%s",
                    diagnostic.table_index,
                    skip_reason,
                    diagnostic.missing_rows,
                    _format_duplicate_rows(diagnostic.duplicate_rows),
                    _format_page_spans(diagnostic.page_spans),
                )
            continue

        logger.info(
            "marker_split_skipped table_index=%s reason=%s missing_rows=%s duplicate_rows=%s page_spans=%s",
            diagnostic.table_index,
            decision.skip_reason,
            diagnostic.missing_rows,
            _format_duplicate_rows(diagnostic.duplicate_rows),
            _format_page_spans(diagnostic.page_spans),
        )

    return eligible_count


def apply_rendered_table_continuation(
    docx_path: Path,
    report: FormattingReport | None = None,
    max_passes: int = 1,
) -> int:
    """
    Phase 3 rendered table continuation entry point.

    Patch 1 only wires LibreOffice/PDF availability checks and disables the
    previous heuristic splitter. Actual rendered row-to-page splitting is added
    in a later patch.
    """
    docx_path = Path(docx_path)
    doc = Document(str(docx_path))
    if not doc.tables:
        logger.info("rendered_table_continuation_enter tables=0 pdf_lines=0 max_passes=%s", max_passes)
        logger.info("rendered_final_decision action=rendered_no_action reason=no_tables")
        return 0

    logger.info(
        "rendered_table_continuation_start tables=%s max_passes=%s",
        len(doc.tables),
        max_passes,
    )

    if _marker_split_enabled():
        apply_marker_split = _marker_split_apply_enabled()
        marker_result = _run_marker_split_detection_pass(
            docx_path,
            apply_split=apply_marker_split,
        )
        if apply_marker_split and marker_result:
            logger.info("rendered_final_decision action=marker_split_applied")
            return marker_result

    pdf_path: Path | None = None
    try:
        pdf_path = render_docx_to_pdf(docx_path)
        pdf_lines = analyze_pdf_lines(pdf_path)
    except LibreOfficeNotFoundError as exc:
        _warn_rendered_split_unavailable(report, str(exc))
        logger.info("rendered_final_decision action=rendered_no_action reason=libreoffice_unavailable")
        return 0
    except Exception as exc:
        _warn_rendered_split_unavailable(report, str(exc))
        logger.info("rendered_final_decision action=rendered_no_action reason=render_or_pdf_analysis_failed")
        return 0
    finally:
        if pdf_path is not None:
            shutil.rmtree(pdf_path.parent, ignore_errors=True)

    logger.info(
        "rendered_table_continuation_enter tables=%s pdf_lines=%s max_passes=%s",
        len(doc.tables),
        len(pdf_lines),
        max_passes,
    )

    diagnostics: dict[str, bool] = {"ambiguous": False}
    move_candidate = _find_rendered_whole_table_move_candidate(doc, pdf_lines, diagnostics)
    if move_candidate is not None:
        if not _ensure_page_break_before(move_candidate.caption_para_xml):
            logger.info(
                "rendered_final_decision action=rendered_no_action reason=whole_table_candidate_already_has_page_break table_idx=%s",
                move_candidate.table_idx,
            )
            return 0
        doc.save(str(docx_path))
        logger.info(
            "rendered_final_decision action=rendered_whole_table_move table_idx=%s",
            move_candidate.table_idx,
        )
        return 1

    candidate = _find_rendered_split_candidate(doc, pdf_lines, diagnostics)
    if candidate is None:
        if diagnostics["ambiguous"]:
            logger.info("rendered_final_decision action=rendered_skip_ambiguous reason=no_safe_rendered_candidate")
        else:
            logger.info("rendered_final_decision action=rendered_no_action reason=no_rendered_candidate")
        return 0

    num = _find_caption_number_before_table(doc, candidate.tbl_xml)
    continuation_text = f"Продолжение таблицы {num}" if num else "Продолжение таблицы"
    if not _split_table_at(doc, candidate.tbl_xml, candidate.split_after, continuation_text):
        logger.info(
            "rendered_final_decision action=rendered_no_action reason=split_mutation_failed table_idx=%s split_after=%s",
            candidate.table_idx,
            candidate.split_after,
        )
        return 0

    doc.save(str(docx_path))
    logger.info(
        "rendered_final_decision action=rendered_split table_idx=%s split_after=%s",
        candidate.table_idx,
        candidate.split_after,
    )
    return 1


# ── Remove empty paragraphs between image and figure caption ─────────────────

def remove_empty_before_figure_captions(doc: Document) -> int:
    """
    Remove empty paragraphs that appear immediately between an image paragraph
    and a figure caption ("Рис. X.Y.Z — ...").

    Students often insert a blank line between the figure and its caption.
    This leaves a visual gap in the formatted output.  We remove such blanks
    only when the paragraph immediately before the empty run contains a drawing.

    Returns the number of paragraphs removed.
    """
    paragraphs = doc.paragraphs
    n = len(paragraphs)
    to_remove: list = []

    i = 0
    while i < n:
        text = (paragraphs[i].text or "").strip()
        if _FIGURE_CAP_RE_GEOM.match(text):
            # Collect preceding empty paragraphs
            j = i - 1
            empty_elems: list = []
            while j >= 0:
                prev_text = (paragraphs[j].text or "").strip()
                if not prev_text and not _para_has_image(paragraphs[j]._element):
                    empty_elems.append(paragraphs[j]._element)
                    j -= 1
                else:
                    break
            # Only remove if the paragraph immediately before the run is an image
            if empty_elems and j >= 0 and _para_has_image(paragraphs[j]._element):
                to_remove.extend(empty_elems)
        i += 1

    removed = 0
    for elem in to_remove:
        parent = elem.getparent()
        if parent is not None:
            parent.remove(elem)
            removed += 1

    if removed:
        logger.info("remove_empty_before_figure_captions: removed %d gap paragraph(s)", removed)
    return removed


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


def _get_image_height_pt(p_elem) -> float | None:
    """
    Return the rendered height (pt) of the first drawing in a paragraph by
    reading the wp:extent cy attribute (in EMU).

    Word stores drawing dimensions in EMU (English Metric Units):
        1 pt = 12 700 EMU  (EMU_PER_PT constant)

    Returns None if no wp:extent element is found.
    """
    for drawing in p_elem.findall(".//" + qn("w:drawing")):
        for container_tag in (qn("wp:inline"), qn("wp:anchor")):
            container = drawing.find(container_tag)
            if container is not None:
                extent = container.find(qn("wp:extent"))
                if extent is not None:
                    cy = extent.get("cy")
                    if cy and cy.lstrip("-").isdigit():
                        return int(cy) / EMU_PER_PT
    return None


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

            # Image paragraph — use actual rendered height from wp:extent cy if
            # available; fall back to the generic paragraph height estimate.
            # The generic estimate returns ~21 pt (1 empty line) for image-only
            # paragraphs, massively underestimating real figure heights.
            img_h = _get_image_height_pt(xml_elem) or h

            # Image paragraph — check if the next paragraph is a figure caption.
            # Skip past any empty paragraphs between image and caption first.
            j = i + 1
            while j < n and body_elems[j][0] == "paragraph":
                nk, nxe, npo = body_elems[j]
                if (npo.text or "").strip():
                    break
                j += 1

            if j < n:
                nk, nxe, npo = body_elems[j]
                if nk == "paragraph":
                    next_text = (npo.text or "").strip()
                    if _FIGURE_CAP_RE_GEOM.match(next_text):
                        caption_h   = _estimate_para_height(npo)
                        img_fits    = (current_h + img_h <= body_h)
                        cap_orphan  = (current_h + img_h + caption_h > body_h)
                        fits_fresh  = (img_h + caption_h <= body_h)

                        if img_fits and cap_orphan and fits_fresh:
                            _set_page_break_before(xml_elem)
                            count += 1
                            logger.info(
                                "rule6: pageBreakBefore on image before [%s]",
                                next_text[:50],
                            )
                            # Both now start fresh on next page
                            current_h = img_h + caption_h
                            i = j + 1
                            continue

            # Normal advance (use img_h for accurate geometry tracking)
            if current_h + img_h > body_h:
                current_h = img_h
            else:
                current_h += img_h

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
