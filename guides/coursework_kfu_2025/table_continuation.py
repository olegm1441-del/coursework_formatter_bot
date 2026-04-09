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


def _estimate_row_height(row, body_width_pt: float) -> float:
    """Estimated rendered height of a table row in points."""
    # 1. Explicit height from XML (twips)
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
        return _TABLE_LINE_PT

    num_cols = len(cells)
    col_w_pt = max(20.0, body_width_pt / num_cols)
    chars_per_col = max(8, int(col_w_pt / _PT_PER_CHAR_TABLE))

    max_h = _TABLE_LINE_PT
    seen: set[int] = set()
    for cell in cells:
        cid = id(cell._element)
        if cid in seen:
            continue
        seen.add(cid)
        text = (cell.text or "").strip()
        n = max(1, math.ceil(len(text) / chars_per_col)) if text else 1
        max_h = max(max_h, n * _TABLE_LINE_PT)

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
    last_para_kind: str | None = None  # kind of the most recent paragraph
    splits: list[tuple] = []  # (tbl_xml, split_after_row, table_num)

    for kind, xml_elem, py_obj in body_elems:

        if kind == "paragraph":
            text = (py_obj.text or "").strip()
            num = _extract_table_num(text)
            if num:
                last_tbl_num = num

            # Detect existing "Продолжение таблицы" paragraph written by the student
            is_continuation = bool(_TBL_NUM_RE.search(text) and "продолжение" in text.lower())
            last_para_kind = "table_continuation" if is_continuation else ("empty" if not text else "other")

            h = _estimate_para_height(py_obj)
            current_h += h
            if current_h > body_h:
                current_h = h   # paragraph starts a fresh page

        elif kind == "table":
            rows = py_obj.rows

            # If the immediately preceding paragraph is an existing
            # "Продолжение таблицы" written by the student, this table is
            # already a manually split continuation — skip it entirely so
            # we don't double-split or insert a second continuation header.
            if last_para_kind == "table_continuation":
                logger.debug(
                    "table_continuation: skipping table (preceded by existing continuation para)"
                )
                last_para_kind = None
                current_h += sum(_estimate_row_height(r, body_w) for r in rows)
                continue

            last_para_kind = None

            if len(rows) < 2:
                # Single-row tables never need splitting
                current_h += sum(_estimate_row_height(r, body_w) for r in rows)
                continue

            table_num = last_tbl_num or "?"
            row_hs = [_estimate_row_height(r, body_w) for r in rows]
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
            for rh in (_estimate_row_height(r, body_w) for r in rows):
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
                    # Check: does the run land near the page bottom?
                    # "near" = current_h after the run would be close to body_h
                    run_total_h = sum(rh for _, rh in run_elems)
                    end_h = run_h_start + run_total_h
                    near_bottom = end_h >= (body_h - _BOTTOM_TOLERANCE_PT)
                    if near_bottom:
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
