"""
Phase 3, Rule 1 — Table continuation (geometry-based, no LibreOffice).

For every table that overflows its page:
  1. Split it just before the row that would cross the page boundary.
  2. Insert "Продолжение таблицы X.Y.Z" (right-aligned, 14 pt, not bold).
  3. Prepend a copy of the header row to the continuation table.

Geometry is estimated from font/spacing constants and cell text length.
Accuracy: ±1–2 rows vs real Word rendering — acceptable for the typical
student coursework document where tables rarely exceed 2 pages.
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

logger = logging.getLogger(__name__)

# ── Unit helpers ─────────────────────────────────────────────────────────────

EMU_PER_PT  = 12700   # 1 pt  = 12 700 EMU  (python-docx stores lengths in EMU)
TWIP_PER_PT = 20      # 1 pt  = 20 twips    (w:trHeight val is in twips)

def _emu_pt(v: int) -> float: return v / EMU_PER_PT
def _twip_pt(v: int) -> float: return v / TWIP_PER_PT

# ── Page geometry ─────────────────────────────────────────────────────────────

# Safety margin subtracted from body height so we don't overfill a page.
# Accounts for rounding + minor rendering differences between LO and Word.
_PAGE_BUFFER_PT = 18


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
        ls_rule = pf.line_spacing_rule   # None | WD_LINE_SPACING (0=auto,1=exact,2=atLeast,…)
        if ls is not None:
            # ls_rule == 1 means "Exactly" — ls is stored in EMU, not as a multiplier
            if ls_rule is not None and int(ls_rule) == 1:
                # Exact rule: ls is an Emu object → .pt gives the fixed line height
                if hasattr(ls, "pt"):
                    line_h = ls.pt
                else:
                    # Fallback: raw int = EMU
                    try:
                        line_h = int(ls) / EMU_PER_PT
                    except (TypeError, ValueError):
                        pass
            elif isinstance(ls, (int, float)):
                # Multiplier (e.g. 1.5)
                line_h = 14 * float(ls)
            elif hasattr(ls, "pt"):
                # Fixed spacing stored as Emu (atLeast / exact via old API)
                line_h = ls.pt
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


def _cell_font_size_pt(cell) -> float:
    """
    Read font size (pt) from the first run in the first paragraph of a cell.
    Falls back to _TABLE_LINE_PT (12 pt) if not found.
    """
    for p in cell._element.findall(qn("w:p")):
        for r in p.findall(qn("w:r")):
            rPr = r.find(qn("w:rPr"))
            if rPr is not None:
                sz = rPr.find(qn("w:sz"))
                if sz is not None:
                    val = sz.get(qn("w:val"))
                    if val and val.isdigit():
                        return int(val) / 2  # half-points → points
    return _TABLE_LINE_PT


def _estimate_row_height(row, body_width_pt: float, col_widths_pt: list[float] | None = None) -> float:
    """Estimated rendered height of a table row in points."""
    # 1. Explicit height from XML (twips) — most reliable
    tr = row._tr
    trPr = tr.find(qn("w:trPr"))
    if trPr is not None:
        trH = trPr.find(qn("w:trHeight"))
        if trH is not None:
            val = trH.get(qn("w:val"))
            if val and val.isdigit():
                h = _twip_pt(int(val))
                if h > 2:
                    return h

    # 2. Estimate from cell content
    cells = row.cells
    if not cells:
        return _TABLE_LINE_PT + _CELL_PADDING_PT

    num_cols = len(cells)

    # Build per-cell column width: prefer actual w:tblGrid widths, else equal split
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

        # Use actual font size for this cell
        font_pt = _cell_font_size_pt(cell)
        # pt-per-char: TNR proportional font, empirically ~0.42 × font_pt per avg char
        pt_per_char = font_pt * 0.42
        chars_per_line = max(4, int(col_w_pt / pt_per_char))
        line_h = font_pt  # single-spaced table cell

        text = (cell.text or "").strip()
        n_lines = max(1, math.ceil(len(text) / chars_per_line)) if text else 1
        cell_h = n_lines * line_h
        max_h = max(max_h, cell_h)

        col_idx += 1

    return max_h + _CELL_PADDING_PT


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

def _make_continuation_para(table_num: str) -> OxmlElement:
    """
    Build <w:p> for "Продолжение таблицы X.Y.Z":
      right-aligned, Times New Roman 14 pt, not bold, no indent, keep_with_next.
    """
    p = OxmlElement("w:p")

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

    Guard: text must be short (≤100 chars) — long paragraphs are prose
    that merely happen to contain those words mid-sentence.
    """
    if len(text) > 100:
        return False
    return bool(_CONT_RE.search(text) and _TBL_WORD_RE.search(text))


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


def apply_table_merging(doc: Document) -> int:
    """
    Phase 3 pre-pass — detect tables that the student manually split with a
    "Продолжение таблицы X.Y.Z" paragraph and merge them back into a single table.

    After merging, apply_table_continuation will re-split any table that genuinely
    overflows the page and insert a correctly formatted continuation paragraph.

    Returns the number of merges performed.
    """
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
            last_tbl_idx = i

        elif kind == "paragraph":
            text = (py_obj.text or "").strip()
            if _is_student_continuation(text) and last_tbl_idx is not None:
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
            _merge_tables(tbl1_xml, tbl1_obj, tbl2_xml, tbl2_obj)

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

def apply_table_continuation(doc: Document) -> int:
    """
    Walk document body elements, detect tables that overflow a page, split them.
    Returns the number of splits performed.
    """
    body_h = _body_height_pt(doc)
    body_w = _body_width_pt(doc)
    logger.info("table_continuation: body_height=%.1f pt  body_width=%.1f pt", body_h, body_w)

    body_elems = list(_iter_body(doc))   # snapshot — safe to mutate doc after this

    current_h = 0.0          # running height on the current page
    last_tbl_num: str | None = None   # table number from the most recent caption
    splits: list[tuple] = []  # (tbl_xml, split_after_row, table_num)

    for kind, xml_elem, py_obj in body_elems:

        if kind == "paragraph":
            text = (py_obj.text or "").strip()
            num = _extract_table_num(text)
            if num:
                last_tbl_num = num

            h = _estimate_para_height(py_obj)
            current_h += h
            if current_h > body_h:
                current_h = h   # paragraph starts a fresh page

        elif kind == "table":
            rows = py_obj.rows

            if len(rows) < 2:
                # Single-row tables never need splitting
                current_h += sum(_estimate_row_height(r, body_w) for r in rows)
                continue

            table_num = last_tbl_num or "?"
            col_widths = _tbl_col_widths_pt(xml_elem)
            row_hs = [_estimate_row_height(r, body_w, col_widths) for r in rows]
            split_after = -1

            for row_idx, rh in enumerate(row_hs):
                current_h += rh
                if current_h > body_h:
                    if row_idx == 0:
                        # Header row alone overflows — just start a new page
                        current_h = rh
                        continue

                    split_after = row_idx - 1   # last row that fit
                    splits.append((xml_elem, split_after, table_num))

                    # Continue geometry tracking: new page starts with
                    # header row copy + the current (overflowing) row
                    current_h = row_hs[0] + rh
                    # For very long tables spanning 3+ pages we'd need to
                    # continue the inner loop here — skipped in this version
                    # because it is rare in student coursework.
                    break

    # Apply splits (each operates on an independent tbl XML element, so
    # processing in forward order is safe — no index shifting between tables)
    n_splits = 0
    for tbl_xml, split_after_row, tbl_num in splits:
        if _split_table(tbl_xml, split_after_row, tbl_num):
            n_splits += 1

    logger.info("table_continuation: %d split(s) applied", n_splits)
    return n_splits


# ── Rule 4: no empty first line of page ──────────────────────────────────────

def apply_rule4_empty_first_lines(doc: Document) -> int:
    """
    Rule 4 — Remove empty paragraphs that land at the very top of a page.

    Strategy: same geometry tracker as apply_table_continuation.
    When the running height crosses the page boundary and the next body
    element is an empty paragraph, mark it for deletion.

    Conservative: only removes paragraphs with no text AND no meaningful
    spacing (space_before ≤ 2 pt). This avoids deleting intentional
    visual separators.

    Returns the number of paragraphs removed.
    """
    body_h = _body_height_pt(doc)
    body_w = _body_width_pt(doc)

    body_elems = list(_iter_body(doc))
    current_h = 0.0
    to_remove: list = []   # xml elements to delete

    for kind, xml_elem, py_obj in body_elems:
        if kind == "paragraph":
            h = _estimate_para_height(py_obj)

            page_overflow = (current_h + h > body_h)
            if page_overflow:
                current_h = 0.0   # new page starts

            text = (py_obj.text or "").strip()
            is_empty = not text

            if page_overflow and is_empty:
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

        elif kind == "table":
            rows = py_obj.rows
            col_widths = _tbl_col_widths_pt(xml_elem)
            for rh in (_estimate_row_height(r, body_w, col_widths) for r in rows):
                current_h += rh
                if current_h > body_h:
                    current_h = rh

    # Remove from bottom to top so parent indices stay valid
    removed = 0
    for elem in reversed(to_remove):
        parent = elem.getparent()
        if parent is not None:
            parent.remove(elem)
            removed += 1

    logger.info("rule4: removed %d empty first-line paragraph(s)", removed)
    return removed


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
