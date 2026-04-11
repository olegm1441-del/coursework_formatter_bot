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
    A single table with 2 LRPB rows (row 2 and row 4) should produce 2 splits
    and therefore 2 'Продолжение таблицы' paragraphs in the output.
    """
    from guides.coursework_kfu_2025.table_continuation import _MERGED_SPLIT_HINTS

    doc = Document()
    doc.add_paragraph("Таблица 1.1 — Test")
    tbl = doc.add_table(rows=6, cols=2)
    for i, cell in enumerate(tbl.rows[0].cells):
        cell.text = f"H{i+1}"
    for ri in range(1, 6):
        for ci, cell in enumerate(tbl.rows[ri].cells):
            cell.text = f"r{ri}c{ci}"

    # Inject LRPB into row 2 and row 4
    for row_idx in (2, 4):
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


# ── Runner ────────────────────────────────────────────────────────────────────

def run_all() -> None:
    tests = [
        ("A  | rule4 does not delete images",       test_a_rule4_does_not_delete_images),
        ("A  | _para_has_image helper",              test_a_para_has_image_helper),
        ("C  | continuation length guard",           test_c_continuation_length_guard),
        ("C  | student continuation merges",         test_c_student_continuation_merges),
        ("B  | multi-LRPB → 2 splits",              test_b_multi_lrpb_produces_two_splits),
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
