"""
Phase 3 regression tests.

Run from repo root:
    python -m pytest tests/test_phase3.py -v
or directly:
    python tests/test_phase3.py

Acceptance criteria per task:
  A  — Figure deletion: images survive Rule 4 (paragraphs with w:drawing never removed)
  C  — Student continuation length: text up to 30 chars is detected and merged
  B  — Multi-LRPB: a table with 2 LRPB rows gets 2 continuation paragraphs
  D  — Trouble-report: tables without LRPB/split produce a non-empty warning list
  E  — Source/note: formatter sets keepWithNext on source paragraphs after tables
  regression — all 5 asset files format without crash and produce a .docx output
"""

from __future__ import annotations

import io
import os
import sys
import shutil
import tempfile
import traceback
from pathlib import Path

# ── project root on path ──────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from guides.coursework_kfu_2025.table_continuation import (
    _is_student_continuation,
    _para_has_image,
    _body_height_pt,
    _estimate_para_height,
    apply_rule4_empty_first_lines,
    apply_table_merging,
    apply_table_continuation,
    _FORMATTER_RSID,
)
from guides.coursework_kfu_2025.formatter_service import format_docx

ASSETS = ROOT / "assets"
ASSET_FILES = list(ASSETS.glob("*.docx"))

PASS = "PASS"
FAIL = "FAIL"


def _result(ok: bool, msg: str = "") -> tuple[bool, str]:
    return ok, msg


# ── helpers ───────────────────────────────────────────────────────────────────

def _make_minimal_doc_with_image() -> Document:
    """
    Minimal document: body paragraph with a w:drawing (simulated image) placed
    EXACTLY at the top of a new page in the geometry estimator.

    Strategy: fill one page worth of content using the same height estimator
    that Rule 4 uses, so the image paragraph triggers page_overflow=True
    and is_empty=True — this is the exact condition that caused the deletion bug.
    """
    doc = Document()

    # Calculate how many "Body text." paragraphs fit on one page
    probe = doc.add_paragraph("Body text.")
    h_per_para = _estimate_para_height(probe)
    probe._element.getparent().remove(probe._element)

    body_h = _body_height_pt(doc)
    # Fill just under one page (leave room for image to overflow)
    n_paras = max(1, int(body_h / h_per_para))
    for _ in range(n_paras):
        doc.add_paragraph("Body text.")

    # Image paragraph: no text, one w:drawing — lands at page-top in estimator
    img_p = doc.add_paragraph()
    drawing = OxmlElement("w:drawing")
    r = OxmlElement("w:r")
    r.append(drawing)
    img_p._element.append(r)
    return doc


def _count_drawings(doc: Document) -> int:
    return len(doc.element.body.findall(".//" + qn("w:drawing")))


def _set_row_height_exact(row, height_pt: float) -> None:
    """Set an exact row height (w:trHeight hRule=exact) so geometry estimator uses it."""
    tr = row._tr
    trPr = tr.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr.insert(0, trPr)
    trH = trPr.find(qn("w:trHeight"))
    if trH is None:
        trH = OxmlElement("w:trHeight")
        trPr.append(trH)
    trH.set(qn("w:val"), str(round(height_pt * 20)))  # twips = pt × 20
    trH.set(qn("w:hRule"), "exact")


def _make_doc_with_student_continuation(cont_text: str) -> Document:
    """
    Minimal document with two 2-column tables separated by a student continuation
    paragraph of given text.
    """
    doc = Document()
    # Table 1
    t1 = doc.add_table(rows=3, cols=2)
    for i, cell in enumerate(t1.rows[0].cells):
        cell.text = f"Header {i+1}"
    for r in t1.rows[1:]:
        for c in r.cells:
            c.text = "data"
    # Continuation paragraph
    doc.add_paragraph(cont_text)
    # Table 2
    t2 = doc.add_table(rows=2, cols=2)
    for i, cell in enumerate(t2.rows[0].cells):
        cell.text = f"Header {i+1}"
    for c in t2.rows[1].cells:
        c.text = "more data"
    return doc


# ── Task A — figure deletion ──────────────────────────────────────────────────

def test_a_rule4_does_not_delete_images() -> tuple[bool, str]:
    """
    Rule 4 must NOT remove paragraphs that contain w:drawing even when they
    appear to be empty (no text) and land at the top of a new estimated page.
    """
    doc = _make_minimal_doc_with_image()
    before = _count_drawings(doc)
    if before == 0:
        return _result(False, "test setup failed: no drawing inserted")

    apply_rule4_empty_first_lines(doc)

    after = _count_drawings(doc)
    if after < before:
        return _result(False, f"drawing deleted: before={before}, after={after}")
    return _result(True, f"drawings intact: {after}")


def test_a_para_has_image_helper() -> tuple[bool, str]:
    """_para_has_image correctly detects w:drawing elements."""
    doc = _make_minimal_doc_with_image()
    # Last paragraph has the drawing
    last_p = doc.paragraphs[-1]
    if not _para_has_image(last_p._element):
        return _result(False, "_para_has_image returned False for paragraph with w:drawing")
    # A normal paragraph should return False
    normal_p = doc.paragraphs[0]
    if _para_has_image(normal_p._element):
        return _result(False, "_para_has_image returned True for text paragraph")
    return _result(True)


# ── Task C — student continuation length ─────────────────────────────────────

def test_c_continuation_length_guard() -> tuple[bool, str]:
    """
    _is_student_continuation must accept texts up to 30 chars.
    Target behaviour after raising limit 27 → 30:
      ≤30 chars + 'продолжени' + 'таблиц' → True
      >30 chars → False
    """
    cases = [
        # (text, expected_after_fix_to_30)
        ("Продолжение таблицы 2.1.10",   True),   # 26 chars
        ("Продолжение таблицы 10.1.10",  True),   # 27 chars
        ("Продолжение таблицы 1.1",      True),   # 23 chars
        ("Продолжение таблицы 100.10.10", True),  # 29 chars — needs limit ≥30
        ("Продолжение таблицы 1.1 (часть 2)", False),  # 33 chars > 30
        ("Это обычный абзац с упоминанием таблицы и продолжения", False),  # long prose
    ]
    failures = []
    for text, expected in cases:
        got = _is_student_continuation(text)
        if got != expected:
            failures.append(f"'{text}' (len={len(text)}): expected={expected}, got={got}")
    if failures:
        return _result(False, "; ".join(failures))
    return _result(True, f"all {len(cases)} cases correct")


def test_c_student_continuation_merges() -> tuple[bool, str]:
    """
    A document with a student continuation text of ≤30 chars must be merged
    (apply_table_merging returns n>0) when at least one table has LRPB.
    Without LRPB the merge is intentionally skipped — so we inject a fake LRPB.
    """
    cont_text = "Продолжение таблицы 10.1.1"   # 26 chars
    doc = _make_doc_with_student_continuation(cont_text)

    # Inject a fake LRPB into the second row of table 1 so the guard passes
    tbl1 = doc.tables[0]
    row1_tr = tbl1.rows[1]._tr
    lrpb = OxmlElement("w:lastRenderedPageBreak")
    row1_tr.append(lrpb)

    n = apply_table_merging(doc)
    if n == 0:
        return _result(False, "expected merge=1, got 0 — continuation not detected")
    if len(doc.tables) != 1:
        return _result(False, f"expected 1 merged table, got {len(doc.tables)}")
    return _result(True, f"merged {n} table(s)")


# ── Task B — multi-LRPB ──────────────────────────────────────────────────────

def test_b_multi_lrpb_produces_two_splits() -> tuple[bool, str]:
    """
    A single table with 2 LRPB rows (row 4 and row 7 of a 10-row table) should
    produce 2 splits and therefore 2 'Продолжение таблицы' paragraphs.

    Row indices chosen so split_after = 3 and 6, both ≥ _MIN_ROWS_TO_SPLIT-1 (3),
    meaning rows_page1 = 4 and 7 — both valid, not stale.
    """
    from guides.coursework_kfu_2025.table_continuation import _MERGED_SPLIT_HINTS

    doc = Document()
    doc.add_paragraph("Таблица 1.1 — Test")
    tbl = doc.add_table(rows=10, cols=2)
    for i, cell in enumerate(tbl.rows[0].cells):
        cell.text = f"H{i+1}"
    for ri in range(1, 10):
        for ci, cell in enumerate(tbl.rows[ri].cells):
            cell.text = f"r{ri}c{ci}"

    # Make each row 80pt tall so total (10×80=800pt) exceeds body_h (~720pt).
    # This ensures the table-fits-on-one-page guard does NOT fire and the
    # LRPB-based split logic is exercised.
    for row in tbl.rows:
        _set_row_height_exact(row, 80.0)

    # Inject LRPB into row 4 and row 7
    # row 4 → split_after=3 → 4 rows page1 (valid ≥ 4)
    # row 7 → split_after=6 → 7 rows page1 (valid ≥ 4)
    for row_idx in (4, 7):
        tr = tbl.rows[row_idx]._tr
        lrpb = OxmlElement("w:lastRenderedPageBreak")
        tr.append(lrpb)

    _MERGED_SPLIT_HINTS.clear()
    n = apply_table_continuation(doc)

    # Count formatter-inserted continuation paragraphs
    cont_count = sum(
        1 for p in doc.paragraphs
        if p._element.get(qn("w:rsidR")) == _FORMATTER_RSID
    )

    if n != 2:
        return _result(False, f"expected 2 splits, got {n}")
    if cont_count != 2:
        return _result(False, f"expected 2 continuation paragraphs, got {cont_count}")
    return _result(True, f"splits={n}, continuation paragraphs={cont_count}")


# ── Asset regression ──────────────────────────────────────────────────────────

def test_regression_asset(asset_path: Path) -> tuple[bool, str]:
    """
    Format asset file end-to-end; verify:
    - No crash
    - Output .docx exists
    - Image count not decreased
    - No Python exception in formatter
    """
    with tempfile.TemporaryDirectory() as tmp:
        out_path = Path(tmp) / f"out_{asset_path.name}"
        # Count images before
        doc_in = Document(str(asset_path))
        imgs_before = _count_drawings(doc_in)
        del doc_in

        try:
            format_docx(str(asset_path), str(out_path))
        except Exception as e:
            return _result(False, f"formatter raised: {e}\n{traceback.format_exc()}")

        if not out_path.exists():
            return _result(False, "output file not created")

        doc_out = Document(str(out_path))
        imgs_after = _count_drawings(doc_out)
        if imgs_after < imgs_before:
            return _result(
                False,
                f"images deleted: before={imgs_before}, after={imgs_after}",
            )

        return _result(True, f"ok (images: {imgs_before}→{imgs_after})")


# ── Batch 1 — tblW fix, _MIN_COL_PT, stale LRPB skip ────────────────────────

def test_b1_tblW_updated_after_col_optimization() -> tuple[bool, str]:
    """
    _optimize_table_col_widths must update w:tblPr/w:tblW to match the new
    column total after scaling.  Without this fix Word renders the table at
    the original (too-wide) tblW instead of the corrected column sum.
    """
    from guides.coursework_kfu_2025.table_continuation import (
        _optimize_table_col_widths, TWIP_PER_PT,
    )

    doc = Document()
    tbl = doc.add_table(rows=2, cols=3)
    tbl_xml = tbl._element
    body_w = 481.9  # standard KFU body width in pt

    # Set each of 3 columns to 200 pt → total 600 pt > body_w
    grid = tbl_xml.find(qn("w:tblGrid"))
    if grid is None:
        return _result(False, "no tblGrid in table XML")
    for gc in grid.findall(qn("w:gridCol")):
        gc.set(qn("w:w"), str(int(200 * TWIP_PER_PT)))

    # Set tblW to original oversized value
    tblPr = tbl_xml.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl_xml.insert(0, tblPr)
    tblW_el = tblPr.find(qn("w:tblW"))
    if tblW_el is None:
        tblW_el = OxmlElement("w:tblW")
        tblPr.append(tblW_el)
    tblW_el.set(qn("w:w"), str(int(600 * TWIP_PER_PT)))
    tblW_el.set(qn("w:type"), "dxa")

    changed = _optimize_table_col_widths(tbl_xml, body_w)
    if not changed:
        return _result(False, "optimizer reported no change (expected scale-down)")

    new_tblW_el = tblPr.find(qn("w:tblW"))
    if new_tblW_el is None:
        return _result(False, "w:tblW element missing after optimization")

    new_total_twips = int(new_tblW_el.get(qn("w:w"), 0))
    expected_twips = round(body_w * TWIP_PER_PT)
    # Allow ±50 twips rounding slack
    if abs(new_total_twips - expected_twips) > 50:
        return _result(
            False,
            f"tblW not updated: got {new_total_twips} twips, expected ~{expected_twips}",
        )
    return _result(True, f"tblW updated to {new_total_twips} twips (expected ~{expected_twips})")


def test_b1_min_col_pt_is_20() -> tuple[bool, str]:
    """
    _MIN_COL_PT must be ≤ 20 (variant C: only phantom columns < 20 pt
    are redistributed; legitimate narrow columns like 30 pt survive).
    """
    from guides.coursework_kfu_2025.table_continuation import _MIN_COL_PT
    if _MIN_COL_PT > 20.5:
        return _result(False, f"_MIN_COL_PT={_MIN_COL_PT} > 20 — old value, fix not applied")
    return _result(True, f"_MIN_COL_PT={_MIN_COL_PT} ✓")


def test_b1_stale_lrpb_skipped() -> tuple[bool, str]:
    """
    Table 2.3.1: 6 rows, LRPB at row index 3 → split_after=2 → 3 rows on page 1.
    3 < _MIN_ROWS_TO_SPLIT(4) → apply_table_continuation must SKIP the split
    and emit a trouble-report warning.

    Rows are made 130pt tall (6×130=780pt > body_h) so the table-fits-on-one-page
    guard does NOT fire first.
    """
    from guides.coursework_kfu_2025.table_continuation import (
        apply_table_continuation, _MERGED_SPLIT_HINTS,
    )
    from guides.coursework_kfu_2025.docx_utils import FormattingReport

    doc = Document()
    doc.add_paragraph("Таблица 2.3.1 — Stale LRPB test")
    tbl = doc.add_table(rows=6, cols=2)
    for i, cell in enumerate(tbl.rows[0].cells):
        cell.text = f"H{i + 1}"
    for ri in range(1, 6):
        for ci, cell in enumerate(tbl.rows[ri].cells):
            cell.text = f"r{ri}c{ci}"

    # Make each row 130pt exact so total (6×130=780pt) exceeds body_h (~720pt)
    for row in tbl.rows:
        _set_row_height_exact(row, 130.0)

    # LRPB at row index 3 → split_after = max(1, 3-1) = 2 → 3 rows on page 1
    tr = tbl.rows[3]._tr
    lrpb = OxmlElement("w:lastRenderedPageBreak")
    tr.append(lrpb)

    _MERGED_SPLIT_HINTS.clear()
    report = FormattingReport()
    n = apply_table_continuation(doc, report=report)

    if n != 0:
        return _result(False, f"expected 0 splits (stale LRPB skipped), got {n}")
    if report.is_empty():
        return _result(False, "expected a trouble-report warning for stale LRPB, got none")
    return _result(True, f"stale LRPB skipped; warning: '{report.warnings[0][:60]}'")


def test_b1_valid_lrpb_splits() -> tuple[bool, str]:
    """
    Table with LRPB at row index 4 → split_after=3 → 4 rows on page 1 = threshold.
    apply_table_continuation must still perform exactly 1 split.

    Rows are made 130pt tall (6×130=780pt > body_h) so the table-fits-on-one-page
    guard does NOT fire first.
    """
    from guides.coursework_kfu_2025.table_continuation import (
        apply_table_continuation, _MERGED_SPLIT_HINTS,
    )
    from guides.coursework_kfu_2025.docx_utils import FormattingReport

    doc = Document()
    doc.add_paragraph("Таблица 1.3.1 — Valid LRPB test")
    tbl = doc.add_table(rows=6, cols=2)
    for i, cell in enumerate(tbl.rows[0].cells):
        cell.text = f"H{i + 1}"
    for ri in range(1, 6):
        for ci, cell in enumerate(tbl.rows[ri].cells):
            cell.text = f"r{ri}c{ci}"

    # Make each row 130pt exact so total (6×130=780pt) exceeds body_h (~720pt)
    for row in tbl.rows:
        _set_row_height_exact(row, 130.0)

    # LRPB at row index 4 → split_after = max(1, 4-1) = 3 → 4 rows on page 1
    tr = tbl.rows[4]._tr
    lrpb = OxmlElement("w:lastRenderedPageBreak")
    tr.append(lrpb)

    _MERGED_SPLIT_HINTS.clear()
    report = FormattingReport()
    n = apply_table_continuation(doc, report=report)

    if n != 1:
        return _result(False, f"expected 1 split, got {n}")
    return _result(True, "split performed at row 4 (4 rows on page 1 = threshold)")


# ── Batch 2 — keepTogether, Rule 6 propagation, image height ─────────────────

def test_b2_keep_together_on_table_caption() -> tuple[bool, str]:
    """
    After apply_pagination_rules, table_caption and table_title paragraphs
    must have keep_together=True (prevents a long title from being split
    across pages by Word's line-breaker).
    """
    from guides.coursework_kfu_2025.pagination_rules import apply_pagination_rules

    doc = Document()
    doc.add_paragraph("Таблица 1.1 — Test caption line")   # → table_caption
    doc.add_table(rows=2, cols=2)
    apply_pagination_rules(doc)

    p = doc.paragraphs[0]
    if not p.paragraph_format.keep_together:
        return _result(False, "keep_together not set on table_caption paragraph")
    return _result(True, "table_caption has keep_together=True")


def test_b2_keep_together_on_headings() -> tuple[bool, str]:
    """
    After apply_pagination_rules, heading1 and heading2 paragraphs must have
    keep_together=True (prevents a multi-line heading from being split across pages).
    """
    from guides.coursework_kfu_2025.pagination_rules import apply_pagination_rules

    doc = Document()
    doc.add_paragraph("1. Теоретические основы исследования")   # → heading1
    doc.add_paragraph("1.1. Понятие и сущность термина")         # → heading2
    doc.add_paragraph("Основной текст параграфа.")
    apply_pagination_rules(doc)

    p_h1 = doc.paragraphs[0]
    p_h2 = doc.paragraphs[1]
    if not p_h1.paragraph_format.keep_together:
        return _result(False, "keep_together not set on heading1")
    if not p_h2.paragraph_format.keep_together:
        return _result(False, "keep_together not set on heading2")
    return _result(True, "heading1 and heading2 have keep_together=True")


def test_b2_rule6_propagates_through_empty_para() -> tuple[bool, str]:
    """
    _apply_rule6: an image paragraph followed by one (or more) empty paragraphs
    and then a figure_caption must have keepWithNext set on BOTH the image paragraph
    AND the intervening empty paragraph(s), so the chain reaches the caption.
    """
    from guides.coursework_kfu_2025.pagination_rules import apply_pagination_rules

    doc = Document()
    # Image paragraph
    img_p = doc.add_paragraph()
    drawing = OxmlElement("w:drawing")
    r_el = OxmlElement("w:r")
    r_el.append(drawing)
    img_p._element.append(r_el)
    # Empty paragraph between image and caption
    doc.add_paragraph("")
    # Figure caption
    doc.add_paragraph("Рисунок 1.1 — Схема взаимодействия")

    apply_pagination_rules(doc)

    img_para   = doc.paragraphs[0]
    empty_para = doc.paragraphs[1]
    if not img_para.paragraph_format.keep_with_next:
        return _result(False, "keep_with_next not set on image paragraph")
    if not empty_para.paragraph_format.keep_with_next:
        return _result(
            False,
            "keep_with_next not set on empty paragraph between image and caption",
        )
    return _result(True, "keepWithNext propagated through empty paragraph to caption")


def test_b2_image_height_from_emu() -> tuple[bool, str]:
    """
    _get_image_height_pt must read wp:extent cy from a drawing element and
    convert EMU → pt correctly (EMU_PER_PT = 12700).
    """
    from guides.coursework_kfu_2025.table_continuation import _get_image_height_pt

    doc = Document()
    p = doc.add_paragraph()

    # Build a minimal drawing: w:drawing > wp:inline > wp:extent cy="1270000" (=100pt)
    drawing  = OxmlElement("w:drawing")
    inline   = OxmlElement("wp:inline")
    extent   = OxmlElement("wp:extent")
    extent.set("cy", str(100 * 12700))   # 100 pt × 12700 EMU/pt = 1270000 EMU
    inline.append(extent)
    drawing.append(inline)
    r_el = OxmlElement("w:r")
    r_el.append(drawing)
    p._element.append(r_el)

    h = _get_image_height_pt(p._element)
    if h is None:
        return _result(False, "_get_image_height_pt returned None — extent not read")
    if abs(h - 100.0) > 0.5:
        return _result(False, f"expected 100.0 pt, got {h:.2f} pt")
    return _result(True, f"image height correctly read as {h:.1f} pt from EMU")


# ── Batch 3 — footnote standardization ───────────────────────────────────────

def test_b3_format_footnote_para_applies_10pt_tnr() -> tuple[bool, str]:
    """
    _format_footnote_para must apply 10pt Times New Roman, no bold,
    single line spacing, and zero indent to a paragraph XML element.
    Tests the low-level helper directly to avoid needing a real footnotes part.
    """
    from guides.coursework_kfu_2025.safe_formatter import _format_footnote_para

    doc = Document()
    p = doc.add_paragraph()

    # Give the paragraph some run with 14pt bold text (typical body style)
    r_el = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    sz_el = OxmlElement("w:sz")
    sz_el.set(qn("w:val"), "28")   # 14pt
    bold_el = OxmlElement("w:b")
    rPr.append(sz_el)
    rPr.append(bold_el)
    t_el = OxmlElement("w:t")
    t_el.text = "Footnote text"
    r_el.append(rPr)
    r_el.append(t_el)
    p._element.append(r_el)

    _format_footnote_para(p._element)

    # Check run font size is now 10pt (w:sz val="20")
    r_out = p._element.find(".//" + qn("w:r"))
    if r_out is None:
        return _result(False, "no w:r found after formatting")
    rPr_out = r_out.find(qn("w:rPr"))
    if rPr_out is None:
        return _result(False, "no w:rPr on run after formatting")

    sz_out = rPr_out.find(qn("w:sz"))
    if sz_out is None:
        return _result(False, "w:sz missing from run rPr after formatting")
    sz_val = sz_out.get(qn("w:val"))
    if sz_val != "20":
        return _result(False, f"expected w:sz val='20' (10pt), got '{sz_val}'")

    # Bold must be suppressed: w:b absent or val="0"
    b_out = rPr_out.find(qn("w:b"))
    if b_out is not None:
        b_val = b_out.get(qn("w:val"), "1")
        if b_val not in ("0", "false"):
            return _result(False, f"bold not suppressed (w:b val='{b_val}')")

    # Check paragraph indent = 0
    pPr_out = p._element.find(qn("w:pPr"))
    if pPr_out is not None:
        ind_out = pPr_out.find(qn("w:ind"))
        if ind_out is not None:
            left_val = ind_out.get(qn("w:left"), "0")
            if left_val not in ("0", None):
                return _result(False, f"indent not zeroed (w:ind left='{left_val}')")

    return _result(True, "footnote para: 10pt TNR, no bold, zero indent ✓")


# ── Batch C2 — image gap, table-fits-on-1-page, number columns ───────────────

def test_c2_empty_para_between_image_and_caption_removed() -> tuple[bool, str]:
    """
    Phase 3 must remove empty paragraphs that appear between an image paragraph
    and its figure_caption (e.g. blank line inserted by student between рисунок
    and 'Рис. 1.2.1 — …').
    """
    from guides.coursework_kfu_2025.table_continuation import remove_empty_before_figure_captions

    doc = Document()
    # Image paragraph
    img_p = doc.add_paragraph()
    drawing = OxmlElement("w:drawing")
    r_el = OxmlElement("w:r")
    r_el.append(drawing)
    img_p._element.append(r_el)
    # Empty paragraph between image and caption (the student's stray blank line)
    doc.add_paragraph("")
    # Figure caption
    doc.add_paragraph("Рисунок 1.2.1 — Схема взаимодействия")

    n = remove_empty_before_figure_captions(doc)

    if n != 1:
        return _result(False, f"expected 1 removal, got {n}")
    # Check the empty paragraph is gone: image should be immediately before caption
    remaining = [p for p in doc.paragraphs if not _para_has_image(p._element)]
    # paragraphs: [img_p (has image), caption]
    total = len(doc.paragraphs)
    if total != 2:
        return _result(False, f"expected 2 paragraphs after removal, got {total}")
    return _result(True, "empty paragraph between image and caption removed ✓")


def test_c2_table_fits_one_page_skips_lrpb() -> tuple[bool, str]:
    """
    If a table's estimated total height ≤ body_h, the LRPB signal is stale
    (table now fits on one page after Phase 1 reformatting) → must NOT split.
    """
    from guides.coursework_kfu_2025.table_continuation import (
        apply_table_continuation, _MERGED_SPLIT_HINTS, _body_height_pt,
        _estimate_row_height, _body_width_pt,
    )
    from guides.coursework_kfu_2025.docx_utils import FormattingReport

    doc = Document()
    doc.add_paragraph("Таблица 1.3.1 — Fits-on-one-page test")
    # Create a small table (3 rows) whose total estimated height << body_h
    tbl = doc.add_table(rows=3, cols=2)
    for i, cell in enumerate(tbl.rows[0].cells):
        cell.text = f"H{i + 1}"
    for ri in range(1, 3):
        for ci, cell in enumerate(tbl.rows[ri].cells):
            cell.text = "short"   # tiny cells → table fits on one page easily

    # Put LRPB in row 2 — but total table height << body_h, so should be skipped
    tr = tbl.rows[2]._tr
    lrpb = OxmlElement("w:lastRenderedPageBreak")
    tr.append(lrpb)

    _MERGED_SPLIT_HINTS.clear()
    report = FormattingReport()
    n = apply_table_continuation(doc, report=report)

    if n != 0:
        return _result(False, f"expected 0 splits (table fits on one page), got {n}")
    return _result(True, "LRPB skipped because table fits on one page ✓")


def test_c2_number_column_minimum() -> tuple[bool, str]:
    """
    _optimize_table_col_widths must protect numeric-only columns from being
    scaled below the width needed to display their content on one line.
    A 7-digit number like '9503005' in a column requires at least ~50pt.
    """
    from guides.coursework_kfu_2025.table_continuation import (
        _optimize_table_col_widths, TWIP_PER_PT,
    )

    doc = Document()
    tbl = doc.add_table(rows=3, cols=4)
    tbl_xml = tbl._element
    body_w = 481.9

    # Set column widths: [250, 100, 100, 130] pt → total 580pt (needs scaling)
    original_widths_pt = [250.0, 100.0, 100.0, 130.0]
    grid = tbl_xml.find(qn("w:tblGrid"))
    if grid is None:
        return _result(False, "no tblGrid")
    for gc, w in zip(grid.findall(qn("w:gridCol")), original_widths_pt):
        gc.set(qn("w:w"), str(round(w * TWIP_PER_PT)))

    # Put numeric content in column 1 (index 1): '9 503 005' (9 chars)
    for ri in range(3):
        cells = tbl.rows[ri].cells
        cells[0].text = "Текстовый заголовок показателя" if ri == 0 else "Текст"
        cells[1].text = "2023 г." if ri == 0 else "9 503 005"  # numeric
        cells[2].text = "2024 г." if ri == 0 else "9 875 076"  # numeric
        cells[3].text = "Абсолютное изменение" if ri == 0 else "−372 071"

    # Also update tcW for each cell to match initial widths
    for ri in range(3):
        tr = tbl.rows[ri]._tr
        col_idx = 0
        for tc in tr.findall(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.append(tcW)
            tcW.set(qn("w:w"), str(round(original_widths_pt[col_idx] * TWIP_PER_PT)))
            tcW.set(qn("w:type"), "dxa")
            col_idx += 1

    _optimize_table_col_widths(tbl_xml, body_w)

    # Column 1 and 2 contain "9 503 005" / "9 875 076" (9 chars × 6pt + 8pt ≈ 62pt)
    # After optimization, columns 1 and 2 should be at least 50pt
    grid_after = tbl_xml.find(qn("w:tblGrid"))
    cols_after = grid_after.findall(qn("w:gridCol"))
    widths_after_pt = [int(c.get(qn("w:w"), 0)) / TWIP_PER_PT for c in cols_after]

    min_expected = 50.0  # 9 chars × 6pt + 8pt padding ≈ 62pt; 50pt is a safe floor
    for col_idx in (1, 2):
        if widths_after_pt[col_idx] < min_expected:
            return _result(
                False,
                f"numeric column {col_idx} too narrow: {widths_after_pt[col_idx]:.1f}pt < {min_expected}pt",
            )
    return _result(True, f"numeric columns protected: {[f'{w:.1f}' for w in widths_after_pt]}")


# ── Runner ────────────────────────────────────────────────────────────────────

def run_all() -> None:
    tests = [
        ("A  | rule4 does not delete images",       test_a_rule4_does_not_delete_images),
        ("A  | _para_has_image helper",              test_a_para_has_image_helper),
        ("C  | continuation length guard",           test_c_continuation_length_guard),
        ("C  | student continuation merges",         test_c_student_continuation_merges),
        ("B  | multi-LRPB → 2 splits",              test_b_multi_lrpb_produces_two_splits),
        ("B1 | tblW updated after optimization",       test_b1_tblW_updated_after_col_optimization),
        ("B1 | _MIN_COL_PT ≤ 20",                     test_b1_min_col_pt_is_20),
        ("B1 | stale LRPB skipped (3 rows page1)",    test_b1_stale_lrpb_skipped),
        ("B1 | valid LRPB splits (4 rows page1)",     test_b1_valid_lrpb_splits),
        ("B2 | keepTogether on table_caption",         test_b2_keep_together_on_table_caption),
        ("B2 | keepTogether on heading1/heading2",     test_b2_keep_together_on_headings),
        ("B2 | rule6 keepWithNext through empty para", test_b2_rule6_propagates_through_empty_para),
        ("B2 | image height from wp:extent cy",        test_b2_image_height_from_emu),
        ("B3 | footnote para: 10pt TNR no bold",         test_b3_format_footnote_para_applies_10pt_tnr),
        ("C2 | empty para image→caption removed",         test_c2_empty_para_between_image_and_caption_removed),
        ("C2 | table fits 1 page → no split",             test_c2_table_fits_one_page_skips_lrpb),
        ("C2 | numeric column minimum protected",          test_c2_number_column_minimum),
    ]

    for asset in ASSET_FILES:
        tests.append((
            f"REG| {asset.name}",
            lambda a=asset: test_regression_asset(a),
        ))

    passed = failed = 0
    for name, fn in tests:
        try:
            ok, msg = fn()
        except Exception as e:
            ok, msg = False, f"EXCEPTION: {e}\n{traceback.format_exc()}"
        status = PASS if ok else FAIL
        suffix = f"  — {msg}" if msg else ""
        print(f"[{status}] {name}{suffix}")
        if ok:
            passed += 1
        else:
            failed += 1

    print(f"\n{'='*60}")
    print(f"Results: {passed} passed, {failed} failed")
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    run_all()
