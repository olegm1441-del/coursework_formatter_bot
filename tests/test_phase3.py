"""
Phase 3 regression tests.

Run from repo root:
    python -m pytest tests/test_phase3.py -v
or directly:
    python tests/test_phase3.py

Acceptance criteria per task:
  A  — Figure deletion: images survive Rule 4 (paragraphs with w:drawing never removed)
  C  — Student continuation length: _is_student_continuation detects ≤30 char texts
  B1 — tblW fix: _optimize_table_col_widths updates w:tblW after scaling
  B2 — keepTogether, Rule 6 propagation, image height from wp:extent
  B3 — Footnote standardisation
  C2 — Empty para between image and caption removed; numeric column minimums
  regression — all 5 asset files format without crash and produce a .docx output

  NOTE: Tests for LRPB-based table splitting (B, B1-stale/valid, C2-fits-1-page,
  C-student-merges) were removed when apply_table_merging / apply_table_continuation
  were stubbed out.  See module docstring in table_continuation.py for the future
  LibreOffice-based plan.
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


def test_c_caption_number_extraction_strict() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_continuation import _extract_table_num
    cases = [
        ("Таблица 2.3", "2.3"),
        ("Таблица 2.3.4", "2.3.4"),
        ("Продолжение таблицы 2.3", None),
        ("Таблица абв", None),
    ]
    bad = []
    for text, expected in cases:
        got = _extract_table_num(text)
        if got != expected:
            bad.append(f"{text!r}: expected={expected!r}, got={got!r}")
    return _result(not bad, "; ".join(bad) if bad else "strict caption extraction OK")


def test_c_apply_table_merging_rebuilds_invalid_split() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_continuation import apply_table_merging

    doc = Document()
    t1 = doc.add_table(rows=3, cols=2)
    t1.rows[0].cells[0].text = "H1"
    t1.rows[0].cells[1].text = "H2"
    t1.rows[1].cells[0].text = "a"
    t1.rows[1].cells[1].text = "b"
    t1.rows[2].cells[0].text = "c"
    t1.rows[2].cells[1].text = "d"

    doc.add_paragraph("Продолжение таблицы 1.1")

    # invalid continuation: header row does NOT match source header
    t2 = doc.add_table(rows=2, cols=2)
    t2.rows[0].cells[0].text = "X"
    t2.rows[0].cells[1].text = "Y"
    t2.rows[1].cells[0].text = "e"
    t2.rows[1].cells[1].text = "f"

    n = apply_table_merging(doc)
    if n != 1:
        return _result(False, f"expected 1 merge, got {n}")
    if len(doc.tables) != 1:
        return _result(False, f"expected 1 table after merge, got {len(doc.tables)}")
    texts = [p.text for p in doc.paragraphs]
    if any("Продолжение таблицы" in (t or "") for t in texts):
        return _result(False, "continuation marker paragraph was not removed for invalid split")
    return _result(True, "invalid manual split was rebuilt")


def test_c_apply_table_merging_keeps_valid_manual_split() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_continuation import apply_table_merging

    doc = Document()
    doc.add_paragraph("Таблица 1.1")
    t1 = doc.add_table(rows=3, cols=2)
    t1.rows[0].cells[0].text = "H1"
    t1.rows[0].cells[1].text = "H2"
    t1.rows[1].cells[0].text = "a"
    t1.rows[1].cells[1].text = "b"
    t1.rows[2].cells[0].text = "c"
    t1.rows[2].cells[1].text = "d"

    marker = doc.add_paragraph("Продолжение таблицы 1.1")
    marker.alignment = 2
    marker.paragraph_format.keep_with_next = True

    t2 = doc.add_table(rows=2, cols=2)
    # valid continuation header equals source header
    t2.rows[0].cells[0].text = "H1"
    t2.rows[0].cells[1].text = "H2"
    t2.rows[1].cells[0].text = "e"
    t2.rows[1].cells[1].text = "f"

    n = apply_table_merging(doc)
    if n != 0:
        return _result(False, f"expected 0 merges for valid manual split, got {n}")
    if len(doc.tables) != 2:
        return _result(False, f"expected 2 tables preserved, got {len(doc.tables)}")
    return _result(True, "valid manual split preserved")


def test_c_apply_table_merging_rebuilds_marker_without_keep_next() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_continuation import apply_table_merging

    doc = Document()
    doc.add_paragraph("Таблица 1.3.1")
    t1 = doc.add_table(rows=2, cols=2)
    t1.rows[0].cells[0].text = "H1"
    t1.rows[0].cells[1].text = "H2"
    t1.rows[1].cells[0].text = "a"
    t1.rows[1].cells[1].text = "b"

    marker = doc.add_paragraph("Продолжение таблицы 1.3.1")
    marker.alignment = 2  # right, but not tightly coupled to the continuation table

    t2 = doc.add_table(rows=2, cols=2)
    t2.rows[0].cells[0].text = "H1"
    t2.rows[0].cells[1].text = "H2"
    t2.rows[1].cells[0].text = "c"
    t2.rows[1].cells[1].text = "d"

    n = apply_table_merging(doc)
    if n != 1:
        return _result(False, f"expected malformed manual chain to be rebuilt, got {n}")
    if len(doc.tables) != 1:
        return _result(False, f"expected 1 table after rebuild, got {len(doc.tables)}")
    if any("Продолжение таблицы" in (p.text or "") for p in doc.paragraphs):
        return _result(False, "malformed continuation marker was preserved")
    if len(doc.tables[0].rows) != 3:
        return _result(False, f"expected duplicate header skipped after merge, rows={len(doc.tables[0].rows)}")
    return _result(True, "manual chain without keepWithNext rebuilt")


def test_c_apply_table_continuation_does_not_heuristic_split() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_continuation import apply_table_continuation

    doc = Document()
    for _ in range(22):
        doc.add_paragraph("Текст абзаца для заполнения страницы.")

    doc.add_paragraph("Таблица 2.3")
    doc.add_paragraph("Название таблицы")

    tbl = doc.add_table(rows=8, cols=2)
    tbl.rows[0].cells[0].text = "Колонка A"
    tbl.rows[0].cells[1].text = "Колонка B"
    for i in range(1, 8):
        tbl.rows[i].cells[0].text = f"a{i}"
        tbl.rows[i].cells[1].text = f"b{i}"

    before_rows = len(doc.tables[0].rows)
    n = apply_table_continuation(doc)
    markers = [p for p in doc.paragraphs if "Продолжение таблицы" in (p.text or "")]

    if n != 0:
        return _result(False, f"expected no width changes in split fixture, got {n}")
    if len(doc.tables) != 1:
        return _result(False, f"heuristic split created extra table(s): {len(doc.tables)}")
    if len(doc.tables[0].rows) != before_rows:
        return _result(False, f"row count changed: {before_rows} -> {len(doc.tables[0].rows)}")
    if markers:
        return _result(False, f"heuristic continuation marker inserted: {[p.text for p in markers]!r}")
    return _result(True, "heuristic split disabled")


def test_c_apply_table_continuation_width_normalization_only() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_continuation import apply_table_continuation

    doc = Document()
    tbl = doc.add_table(rows=2, cols=2)
    tbl.rows[0].cells[0].text = "A"
    tbl.rows[0].cells[1].text = "B"
    tbl.rows[1].cells[0].text = "1"
    tbl.rows[1].cells[1].text = "2"
    grid = tbl._tbl.find(qn("w:tblGrid"))
    if grid is None:
        return _result(False, "test setup failed: no tblGrid")
    for gc in grid.findall(qn("w:gridCol")):
        gc.set(qn("w:w"), "12000")

    before_tables = len(doc.tables)
    before_rows = len(doc.tables[0].rows)
    n = apply_table_continuation(doc)
    markers = [p for p in doc.paragraphs if "Продолжение таблицы" in (p.text or "")]

    if n != 1:
        return _result(False, f"expected one width-normalised table, got {n}")
    if len(doc.tables) != before_tables:
        return _result(False, f"table count changed: {before_tables} -> {len(doc.tables)}")
    if len(doc.tables[0].rows) != before_rows:
        return _result(False, f"row count changed: {before_rows} -> {len(doc.tables[0].rows)}")
    if markers:
        return _result(False, f"unexpected continuation marker: {[p.text for p in markers]!r}")
    return _result(True, "width normalisation remained active without splitting")


def test_c_apply_table_continuation_no_split_double_run_idempotent() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_continuation import apply_table_continuation

    doc = Document()
    for _ in range(22):
        doc.add_paragraph("Текст абзаца для заполнения страницы.")

    doc.add_paragraph("Таблица 3.1")
    tbl = doc.add_table(rows=8, cols=2)
    tbl.rows[0].cells[0].text = "H1"
    tbl.rows[0].cells[1].text = "H2"
    for i in range(1, 8):
        tbl.rows[i].cells[0].text = f"a{i}"
        tbl.rows[i].cells[1].text = f"b{i}"

    first = apply_table_continuation(doc)
    marker_count_1 = sum(1 for p in doc.paragraphs if "Продолжение таблицы" in (p.text or ""))
    table_count_1 = len(doc.tables)
    table_rows_1 = [len(t.rows) for t in doc.tables]

    second = apply_table_continuation(doc)
    marker_count_2 = sum(1 for p in doc.paragraphs if "Продолжение таблицы" in (p.text or ""))
    table_count_2 = len(doc.tables)
    table_rows_2 = [len(t.rows) for t in doc.tables]

    if first != 0 or second != 0:
        return _result(False, f"expected no heuristic changes, got first={first}, second={second}")
    if marker_count_2 != marker_count_1:
        return _result(False, f"marker count changed: {marker_count_1} -> {marker_count_2}")
    if table_count_2 != table_count_1:
        return _result(False, f"table count changed: {table_count_1} -> {table_count_2}")
    if table_rows_2 != table_rows_1:
        return _result(False, f"table structure changed: {table_rows_1!r} -> {table_rows_2!r}")
    return _result(True, "double run did not add markers or split structure")


def test_c_apply_rendered_table_continuation_warns_when_lo_unavailable() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.docx_utils import FormattingReport

    doc = Document()
    doc.add_paragraph("Таблица 1.1")
    doc.add_table(rows=3, cols=2)

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "in.docx"
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        report = FormattingReport()
        try:
            def raise_lo(_path):
                raise tc.LibreOfficeNotFoundError("missing LO")

            tc.render_docx_to_pdf = raise_lo
            n = tc.apply_rendered_table_continuation(path, report=report)
        finally:
            tc.render_docx_to_pdf = old_render

        reread = Document(str(path))

    if n != 0:
        return _result(False, f"expected 0 rendered splits, got {n}")
    if not report.warnings:
        return _result(False, "expected rendered split warning")
    if len(reread.tables) != 1:
        return _result(False, f"DOCX mutated unexpectedly, tables={len(reread.tables)}")
    return _result(True, "LO unavailable path warns and does not mutate")


def test_c_apply_rendered_table_continuation_warns_when_pdf_analysis_fails() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.docx_utils import FormattingReport

    doc = Document()
    doc.add_paragraph("Таблица 1.1")
    doc.add_table(rows=3, cols=2)

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "in.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "in.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        report = FormattingReport()
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path

            def raise_analysis(_path):
                raise RuntimeError("pdf parse failed")

            tc.analyze_pdf_lines = raise_analysis
            n = tc.apply_rendered_table_continuation(path, report=report)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))

    if n != 0:
        return _result(False, f"expected 0 rendered splits, got {n}")
    if not report.warnings:
        return _result(False, "expected PDF analysis warning")
    if len(reread.tables) != 1:
        return _result(False, f"DOCX mutated unexpectedly, tables={len(reread.tables)}")
    return _result(True, "PDF analysis failure warns and does not mutate")


def test_c_rendered_split_single_boundary_success() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    doc = Document()
    doc.add_paragraph("Таблица 2.3")
    tbl = doc.add_table(rows=4, cols=2)
    tbl.rows[0].cells[0].text = "H1"
    tbl.rows[0].cells[1].text = "H2"
    tbl.rows[1].cells[0].text = "alpha one"
    tbl.rows[1].cells[1].text = "beta one"
    tbl.rows[2].cells[0].text = "gamma two"
    tbl.rows[2].cells[1].text = "delta two"
    tbl.rows[3].cells[0].text = "epsilon three"
    tbl.rows[3].cells[1].text = "zeta three"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "in.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "in.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: [
                PdfLine("H1 H2", 1, 100.0, 112.0),
                PdfLine("alpha one beta one", 1, 120.0, 132.0),
                PdfLine("gamma two delta two", 2, 80.0, 92.0),
                PdfLine("epsilon three zeta three", 2, 100.0, 112.0),
            ]
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))

    if n != 1:
        return _result(False, f"expected one rendered split, got {n}")
    if len(reread.tables) != 2:
        return _result(False, f"expected 2 tables after split, got {len(reread.tables)}")
    markers = [p.text for p in reread.paragraphs if "Продолжение таблицы" in (p.text or "")]
    if markers != ["Продолжение таблицы 2.3"]:
        return _result(False, f"unexpected markers: {markers!r}")
    if [c.text for c in reread.tables[0].rows[0].cells] != [c.text for c in reread.tables[1].rows[0].cells]:
        return _result(False, "continuation table header was not repeated")
    return _result(True, "rendered single-boundary split succeeded")


def test_c_rendered_split_preserves_valid_manual_split() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    doc = Document()
    t1 = doc.add_table(rows=2, cols=2)
    t1.rows[0].cells[0].text = "H1"
    t1.rows[0].cells[1].text = "H2"
    t1.rows[1].cells[0].text = "alpha"
    t1.rows[1].cells[1].text = "beta"
    marker = doc.add_paragraph("Продолжение таблицы 1.1")
    marker.alignment = 2  # right; must be preserved exactly
    marker.paragraph_format.keep_with_next = True
    t2 = doc.add_table(rows=2, cols=2)
    t2.rows[0].cells[0].text = "H1"
    t2.rows[0].cells[1].text = "H2"
    t2.rows[1].cells[0].text = "gamma"
    t2.rows[1].cells[1].text = "delta"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "manual.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "manual.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)
        before_xml = marker._element.xml

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: [
                PdfLine("H1 H2", 1, 100.0, 112.0),
                PdfLine("alpha beta", 1, 120.0, 132.0),
                PdfLine("gamma delta", 2, 80.0, 92.0),
            ]
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))
        markers = [p for p in reread.paragraphs if "Продолжение таблицы" in (p.text or "")]

    if n != 0:
        return _result(False, f"valid manual split should be preserved, got split count {n}")
    if len(reread.tables) != 2:
        return _result(False, f"manual split table count changed: {len(reread.tables)}")
    if len(markers) != 1 or markers[0]._element.xml != before_xml:
        return _result(False, "manual continuation marker XML changed")
    return _result(True, "valid manual split preserved exactly")


def test_c_rendered_split_skips_ambiguous_repeated_rows() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    doc = Document()
    doc.add_paragraph("Таблица 2.4")
    tbl = doc.add_table(rows=4, cols=2)
    tbl.rows[0].cells[0].text = "H1"
    tbl.rows[0].cells[1].text = "H2"
    for idx in range(1, 4):
        tbl.rows[idx].cells[0].text = "same"
        tbl.rows[idx].cells[1].text = "row"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "ambiguous.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "ambiguous.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: [
                PdfLine("same row", 1, 120.0, 132.0),
                PdfLine("same row", 2, 80.0, 92.0),
                PdfLine("same row", 2, 100.0, 112.0),
            ]
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))

    if n != 0:
        return _result(False, f"ambiguous repeated rows should skip, got {n}")
    if len(reread.tables) != 1:
        return _result(False, f"ambiguous split mutated table count: {len(reread.tables)}")
    return _result(True, "ambiguous repeated rows skipped")


def test_c_rendered_split_skips_merged_boundary_conflict() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    doc = Document()
    doc.add_paragraph("Таблица 2.5")
    tbl = doc.add_table(rows=3, cols=2)
    tbl.rows[0].cells[0].text = "H1"
    tbl.rows[0].cells[1].text = "H2"
    tbl.rows[1].cells[0].text = "merge start"
    tbl.rows[1].cells[1].text = "alpha"
    tbl.rows[2].cells[0].text = "merge continue"
    tbl.rows[2].cells[1].text = "beta"
    tbl.cell(1, 0).merge(tbl.cell(2, 0))

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "merged.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "merged.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: [
                PdfLine("H1 H2", 1, 100.0, 112.0),
                PdfLine("merge start alpha", 1, 120.0, 132.0),
                PdfLine("merge continue beta", 2, 80.0, 92.0),
            ]
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))

    if n != 0:
        return _result(False, f"merged boundary conflict should skip, got {n}")
    if len(reread.tables) != 1:
        return _result(False, f"merged conflict mutated table count: {len(reread.tables)}")
    return _result(True, "merged boundary conflict skipped")


def test_c_rendered_split_marker_is_right_aligned() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    doc = Document()
    doc.add_paragraph("Таблица 2.6")
    tbl = doc.add_table(rows=3, cols=2)
    tbl.rows[0].cells[0].text = "H1"
    tbl.rows[0].cells[1].text = "H2"
    tbl.rows[1].cells[0].text = "alpha"
    tbl.rows[1].cells[1].text = "beta"
    tbl.rows[2].cells[0].text = "gamma"
    tbl.rows[2].cells[1].text = "delta"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "format.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "format.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: [
                PdfLine("H1 H2", 1, 100.0, 112.0),
                PdfLine("alpha beta", 1, 120.0, 132.0),
                PdfLine("gamma delta", 2, 80.0, 92.0),
            ]
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))
        markers = [p for p in reread.paragraphs if "Продолжение таблицы" in (p.text or "")]

    if n != 1 or len(markers) != 1:
        return _result(False, f"expected one generated marker, n={n}, markers={len(markers)}")
    pPr = markers[0]._element.find(qn("w:pPr"))
    jc = pPr.find(qn("w:jc")) if pPr is not None else None
    ind = pPr.find(qn("w:ind")) if pPr is not None else None
    keep = pPr.find(qn("w:keepNext")) if pPr is not None else None
    sz = markers[0]._element.find(".//" + qn("w:sz"))
    if jc is None or jc.get(qn("w:val")) != "right":
        return _result(False, "marker is not right-aligned")
    if ind is None or ind.get(qn("w:firstLine")) != "0":
        return _result(False, "marker first-line indent is not zero")
    if keep is None:
        return _result(False, "marker keepWithNext missing")
    if sz is None or sz.get(qn("w:val")) != "28":
        return _result(False, "marker font size is not 14pt")
    return _result(True, "generated marker formatting is correct")


def test_c_rendered_split_caption_number_and_fallback() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    def run_case(caption: str, expected_marker: str) -> tuple[bool, str]:
        doc = Document()
        doc.add_paragraph(caption)
        tbl = doc.add_table(rows=3, cols=2)
        tbl.rows[0].cells[0].text = "H1"
        tbl.rows[0].cells[1].text = "H2"
        tbl.rows[1].cells[0].text = "alpha"
        tbl.rows[1].cells[1].text = "beta"
        tbl.rows[2].cells[0].text = "gamma"
        tbl.rows[2].cells[1].text = "delta"

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "caption.docx"
            pdf_dir = Path(tmp) / "pdf"
            pdf_dir.mkdir()
            pdf_path = pdf_dir / "caption.pdf"
            pdf_path.write_bytes(b"%PDF-1.4\n")
            doc.save(path)

            old_render = tc.render_docx_to_pdf
            old_analyze = tc.analyze_pdf_lines
            try:
                tc.render_docx_to_pdf = lambda _path: pdf_path
                tc.analyze_pdf_lines = lambda _path: [
                    PdfLine("H1 H2", 1, 100.0, 112.0),
                    PdfLine("alpha beta", 1, 120.0, 132.0),
                    PdfLine("gamma delta", 2, 80.0, 92.0),
                ]
                n = tc.apply_rendered_table_continuation(path)
            finally:
                tc.render_docx_to_pdf = old_render
                tc.analyze_pdf_lines = old_analyze

            reread = Document(str(path))
            markers = [p.text for p in reread.paragraphs if "Продолжение таблицы" in (p.text or "")]

        if n != 1:
            return _result(False, f"{caption!r}: expected split, got {n}")
        if markers != [expected_marker]:
            return _result(False, f"{caption!r}: expected {expected_marker!r}, got {markers!r}")
        return _result(True, "")

    ok, msg = run_case("Таблица 2.3.4", "Продолжение таблицы 2.3.4")
    if not ok:
        return _result(False, msg)
    ok, msg = run_case("Таблица абв", "Продолжение таблицы")
    if not ok:
        return _result(False, msg)
    return _result(True, "strict caption number and fallback markers correct")


def test_c_rendered_start_page_moves_whole_table_without_complete_data_row() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    doc = Document()
    caption = doc.add_paragraph("Таблица 2.2.3")
    doc.add_paragraph("Показатели эффективности")
    tbl = doc.add_table(rows=3, cols=2)
    tbl.rows[0].cells[0].text = "Показатель"
    tbl.rows[0].cells[1].text = "Экономия"
    tbl.rows[1].cells[0].text = "Почтовые расходы"
    tbl.rows[1].cells[1].text = "переход на электронный документооборот"
    tbl.rows[2].cells[0].text = "Архивное хранение"
    tbl.rows[2].cells[1].text = "высокая экономия архива"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "move.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "move.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: [
                PdfLine("Таблица 2.2.3", 1, 686.5, 700.0),
                PdfLine("Показатели эффективности", 1, 710.8, 724.0),
                PdfLine("Показатель Экономия", 1, 741.8, 755.0),
                PdfLine("Почтовые расходы переход", 1, 763.2, 776.0),
                PdfLine("на электронный документооборот", 2, 86.8, 99.0),
                PdfLine("Архивное хранение высокая экономия архива", 2, 108.0, 121.0),
            ]
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))
        reread_caption = next(p for p in reread.paragraphs if p.text == caption.text)

    pPr = reread_caption._element.find(qn("w:pPr"))
    page_break = pPr.find(qn("w:pageBreakBefore")) if pPr is not None else None
    markers = [p.text for p in reread.paragraphs if "Продолжение таблицы" in (p.text or "")]

    if n != 1:
        return _result(False, f"expected whole-table move, got {n}")
    if page_break is None:
        return _result(False, "caption did not receive pageBreakBefore")
    if len(reread.tables) != 1:
        return _result(False, f"whole-table move should not split, got {len(reread.tables)} tables")
    if markers:
        return _result(False, f"whole-table move inserted continuation marker: {markers!r}")
    return _result(True, "whole-table move applied to caption")


def test_c_rendered_start_page_skips_ambiguous_usability() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    doc = Document()
    doc.add_paragraph("Таблица 2.2.4")
    tbl = doc.add_table(rows=3, cols=2)
    tbl.rows[0].cells[0].text = "Год"
    tbl.rows[0].cells[1].text = "Значение"
    tbl.rows[1].cells[0].text = "2023"
    tbl.rows[1].cells[1].text = "10"
    tbl.rows[2].cells[0].text = "2024"
    tbl.rows[2].cells[1].text = "10"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "ambiguous_move.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "ambiguous_move.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: [
                PdfLine("Таблица 2.2.4", 1, 690.0, 702.0),
                PdfLine("Год Значение", 1, 735.0, 748.0),
                PdfLine("2023", 1, 763.0, 776.0),
                PdfLine("10", 2, 86.0, 98.0),
                PdfLine("2024 10", 2, 108.0, 120.0),
            ]
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))
        reread_caption = next(p for p in reread.paragraphs if p.text == "Таблица 2.2.4")

    pPr = reread_caption._element.find(qn("w:pPr"))
    page_break = pPr.find(qn("w:pageBreakBefore")) if pPr is not None else None

    if n != 0:
        return _result(False, f"ambiguous start-page evidence should skip, got {n}")
    if page_break is not None:
        return _result(False, "ambiguous start-page evidence added pageBreakBefore")
    if len(reread.tables) != 1:
        return _result(False, f"ambiguous start-page evidence changed tables: {len(reread.tables)}")
    return _result(True, "ambiguous start-page usability skipped")


def test_c_rendered_start_page_keeps_table_with_clear_complete_data_row() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    doc = Document()
    doc.add_paragraph("Таблица 2.2.5")
    tbl = doc.add_table(rows=3, cols=2)
    tbl.rows[0].cells[0].text = "Показатель"
    tbl.rows[0].cells[1].text = "Эффект"
    tbl.rows[1].cells[0].text = "Почтовые расходы"
    tbl.rows[1].cells[1].text = "экономия бюджета"
    tbl.rows[2].cells[0].text = "Архивное хранение"
    tbl.rows[2].cells[1].text = "снижение затрат"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "clear_row.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "clear_row.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: [
                PdfLine("Таблица 2.2.5", 1, 400.0, 412.0),
                PdfLine("Показатель Эффект", 1, 430.0, 442.0),
                PdfLine("Почтовые расходы экономия бюджета", 1, 455.0, 467.0),
                PdfLine("Архивное хранение снижение затрат", 1, 480.0, 492.0),
            ]
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))
        reread_caption = next(p for p in reread.paragraphs if p.text == "Таблица 2.2.5")

    pPr = reread_caption._element.find(qn("w:pPr"))
    page_break = pPr.find(qn("w:pageBreakBefore")) if pPr is not None else None

    if n != 0:
        return _result(False, f"clear complete data row should not move table, got {n}")
    if page_break is not None:
        return _result(False, "clear complete data row still added pageBreakBefore")
    if len(reread.tables) != 1:
        return _result(False, f"clear complete data row changed tables: {len(reread.tables)}")
    return _result(True, "clear complete data row prevents whole-table move")


def test_c_vmerge_guard_rejects_boundary_inside_merge_zone() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_continuation import _is_split_boundary_safe

    doc = Document()
    tbl = doc.add_table(rows=4, cols=2)
    for r_idx, row in enumerate(tbl.rows):
        row.cells[0].text = f"A{r_idx}"
        row.cells[1].text = f"B{r_idx}"

    merged = tbl.cell(1, 0).merge(tbl.cell(2, 0))
    merged.text = "merged"

    rows_xml = tbl._tbl.findall(qn("w:tr"))
    if _is_split_boundary_safe(rows_xml, 1):
        return _result(False, "boundary before vMerge continuation row was considered safe")
    if not _is_split_boundary_safe(rows_xml, 2):
        return _result(False, "boundary after vMerge continuation row was considered unsafe")
    return _result(True, "vMerge guard rejects split inside merge zone")



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


def test_b2_table_source_note_normalised_and_chained() -> tuple[bool, str]:
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from guides.coursework_kfu_2025.pagination_rules import apply_pagination_rules

    doc = Document()
    tbl = doc.add_table(rows=2, cols=2)
    tbl.rows[0].cells[0].text = "H1"
    tbl.rows[0].cells[1].text = "H2"
    tbl.rows[1].cells[0].text = "a"
    tbl.rows[1].cells[1].text = "b"
    source = doc.add_paragraph("Источник: составлено автором.")
    source.alignment = WD_ALIGN_PARAGRAPH.CENTER
    note = doc.add_paragraph("Примечание: расчет ориентировочный.")
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER

    apply_pagination_rules(doc)

    last_cell_p = tbl.rows[-1].cells[-1].paragraphs[-1]
    if not last_cell_p.paragraph_format.keep_with_next:
        return _result(False, "table tail is not chained to source/note")
    if not source.paragraph_format.keep_with_next:
        return _result(False, "source is not chained to following note")
    if source.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
        return _result(False, f"source alignment not normalised: {source.alignment}")
    if note.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
        return _result(False, f"note alignment not normalised: {note.alignment}")
    if source.paragraph_format.first_line_indent is None:
        return _result(False, "source first-line indent was not restored")
    if note.paragraph_format.first_line_indent is None:
        return _result(False, "note first-line indent was not restored")
    return _result(True, "table source/note normalised and chained")


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


def test_yo_normalisation_midword_uppercase() -> tuple[bool, str]:
    """
    Words starting with uppercase but containing lowercase ё mid-word
    (e.g. "Лётчик") must have the ё replaced with е.
    Capital Ё at the start of a word must be preserved.
    """
    from guides.coursework_kfu_2025.safe_formatter import normalize_yo_in_text

    cases = [
        # (input, expected)
        ("лётчик",       "летчик"),
        ("ёж",           "еж"),
        ("Ёж",           "Ёж"),        # capital Ё: preserved
        ("Лётчик",       "Летчик"),    # starts with uppercase Л, ё is lowercase → replace
        ("ЛЁТЧИК",       "ЛЁТЧИК"),   # Ё uppercase → preserved
        ("неёмкий",      "неемкий"),
        ("Чернышёв",     "Чернышев"),
    ]
    failures = []
    for inp, expected in cases:
        got = normalize_yo_in_text(inp)
        if got != expected:
            failures.append(f"normalize_yo_in_text({inp!r}) = {got!r}, expected {expected!r}")
    if failures:
        return _result(False, "\n".join(failures))
    return _result(True, f"all {len(cases)} ё-normalisation cases correct")


def test_t_indent_body_paragraph_left_zero() -> tuple[bool, str]:
    """
    After formatting, regular body paragraphs must have:
    - left_indent = 0 (or None, not a hanging indent)
    - first_line_indent = 709 twips (≈1.25 cm)
    No hanging indent (w:hanging must not be present).
    """
    from guides.coursework_kfu_2025.safe_formatter import process_document

    doc = Document()
    # process_document requires a paragraph with text "введение" to find body start
    doc.add_paragraph("введение")
    # Simulate a paragraph that originally had a List style with hanging indent
    p = doc.add_paragraph("Это обычный абзац с текстом.")
    # Manually inject a hanging indent (simulating "List Paragraph" style effect)
    pPr = p._element.get_or_add_pPr()
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "709")
    ind.set(qn("w:hanging"), "360")
    pPr.append(ind)

    with tempfile.TemporaryDirectory() as tmp:
        inp = os.path.join(tmp, "in.docx")
        out = os.path.join(tmp, "out.docx")
        doc.save(inp)
        process_document(inp, out)
        result_doc = Document(out)

    body_paras = [p for p in result_doc.paragraphs if "обычный абзац" in (p.text or "")]
    if not body_paras:
        return _result(False, "body paragraph not found in output")

    bp = body_paras[0]
    pPr_out = bp._element.find(qn("w:pPr"))
    ind_out = pPr_out.find(qn("w:ind")) if pPr_out is not None else None

    # Check no hanging
    if ind_out is not None and ind_out.get(qn("w:hanging")):
        return _result(False, f"w:hanging still present: {ind_out.get(qn('w:hanging'))}")

    # Check left=0 (either absent or "0")
    left_val = ind_out.get(qn("w:left")) if ind_out is not None else None
    if left_val and left_val != "0":
        return _result(False, f"w:left={left_val!r} (expected 0 or absent)")

    # Check firstLine≈709
    fl_val = ind_out.get(qn("w:firstLine")) if ind_out is not None else None
    if fl_val is None or abs(int(fl_val) - 709) > 30:
        return _result(False, f"w:firstLine={fl_val!r} (expected ≈709)")

    return _result(True, f"body paragraph indent: left=0, firstLine={fl_val} ✓")


# ── Task 2 — Глава N without title ────────────────────────────────────────────

def test_t2_chapter_heading_without_title() -> tuple[bool, str]:
    """
    "Глава 1" (no title) must be classified as heading1.
    "Глава 1. Название" (with title) must still work.
    """
    from guides.coursework_kfu_2025.classifier import parse_heading1

    cases = [
        ("Глава 1",                    True),
        ("глава 2",                    True),
        ("ГЛАВА 3",                    True),
        ("Глава 1.",                   True),
        ("Глава 1. Теоретические основы", True),
        ("Глава 10. Заключение",       True),
        ("Глава",                      False),  # no number
        ("1. Теоретические основы",    True),   # normalized heading — must still work
        ("Введение",                   True),   # exact match — must still work
    ]
    failures = []
    for text, expected in cases:
        result = parse_heading1(text)
        got = result is not None
        if got != expected:
            failures.append(f"parse_heading1({text!r}) → {result}, expected match={expected}")
    if failures:
        return _result(False, "\n".join(failures))
    return _result(True, f"all {len(cases)} chapter heading cases correct")


def _add_fake_word_numbering(paragraph, ilvl_value: str = "0") -> None:
    pPr = paragraph._element.get_or_add_pPr()
    numPr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), ilvl_value)
    num_id = OxmlElement("w:numId")
    num_id.set(qn("w:val"), "1")
    numPr.append(ilvl)
    numPr.append(num_id)
    pPr.append(numPr)


def _style_name(paragraph) -> str:
    try:
        return (paragraph.style.name or "").strip().lower()
    except Exception:
        return ""


def _find_paragraph_starting_with(document: Document, prefix: str):
    for paragraph in document.paragraphs:
        if " ".join(paragraph.text.split()).startswith(prefix):
            return paragraph
    return None


def test_t2_manual_heading2_still_promoted() -> tuple[bool, str]:
    """Explicit manual heading syntax '1.1. ...' remains Heading 2."""
    from guides.coursework_kfu_2025.safe_formatter import process_document
    import tempfile, os

    doc = Document()
    doc.add_paragraph("ВВЕДЕНИЕ")
    doc.add_paragraph("1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ")
    doc.add_paragraph("1.1. Понятие конкурентоспособности организации")
    doc.add_paragraph("Обычный текст подраздела.")

    with tempfile.TemporaryDirectory() as tmp:
        inp = os.path.join(tmp, "in.docx")
        out = os.path.join(tmp, "out.docx")
        doc.save(inp)
        process_document(inp, out)
        result = Document(out)

    heading = _find_paragraph_starting_with(result, "1.1. Понятие конкурентоспособности")
    if heading is None:
        return _result(False, "manual heading2 text missing after formatting")

    if _style_name(heading) not in {"heading 2", "заголовок 2"}:
        return _result(False, f"manual heading2 style is {_style_name(heading)!r}")

    return _result(True, "manual heading2 remains Heading 2")


def test_t2_word_autonumbered_heading2_with_style_still_promoted() -> tuple[bool, str]:
    """
    A real Word-autonumbered Heading 2 may have numPr but no visible '1.1.'
    in paragraph.text. Heading style is enough structural evidence to promote it.
    """
    from guides.coursework_kfu_2025.safe_formatter import (
        auto_detect_heading2,
        clean_spaces,
        is_likely_numbered_heading2_candidate,
        is_probable_body_list_item,
        normalize_heading2_numbering,
        paragraph_has_numbering,
    )

    doc = Document()
    heading = doc.add_paragraph("Понятие конкурентоспособности организации")
    heading.style = "Heading 2"
    _add_fake_word_numbering(heading, ilvl_value="1")

    if is_probable_body_list_item(heading):
        return _result(False, "Word-autonumbered Heading 2 was classified as body/list")

    if not auto_detect_heading2(heading, current_chapter_num=1, next_paragraph_num=1):
        return _result(False, "Word-autonumbered Heading 2 was not auto-detected")

    if not is_likely_numbered_heading2_candidate(heading, 1, 1):
        return _result(False, "Word-autonumbered Heading 2 was not a heading2 candidate")

    normalized = normalize_heading2_numbering(heading, 1, 1)
    expected = "1.1. Понятие конкурентоспособности организации"
    if normalized != expected or clean_spaces(heading.text) != expected:
        return _result(False, f"unexpected Heading 2 normalization: {normalized!r}, text={heading.text!r}")

    if paragraph_has_numbering(heading):
        return _result(False, "Heading 2 Word numbering was not converted to plain text")

    if _style_name(heading) not in {"heading 2", "заголовок 2"}:
        return _result(False, f"autonumbered heading2 style is {_style_name(heading)!r}")

    return _result(True, "Word-autonumbered Heading 2 remains supported")


def test_t2_word_autonumbered_heading1_with_style_still_promoted() -> tuple[bool, str]:
    """
    A real Word-autonumbered Heading 1 may have numPr but no visible '1.'
    in paragraph.text. Heading style/outline must keep it on the heading path.
    """
    from guides.coursework_kfu_2025.safe_formatter import (
        auto_detect_numbered_heading1,
        paragraph_has_numbering,
        process_document,
    )
    import tempfile, os

    direct_doc = Document()
    direct_heading = direct_doc.add_paragraph("ТЕОРЕТИЧЕСКИЕ ОСНОВЫ КОНКУРЕНТОСПОСОБНОСТИ")
    direct_heading.style = "Heading 1"
    _add_fake_word_numbering(direct_heading)
    following_h2 = direct_doc.add_paragraph("Понятие конкурентоспособности организации")
    following_h2.style = "Heading 2"
    _add_fake_word_numbering(following_h2, ilvl_value="1")

    if not auto_detect_numbered_heading1(direct_heading, current_chapter_num=None, next_paragraph=following_h2):
        return _result(False, "Word-autonumbered Heading 1 was not auto-detected")

    doc = Document()
    doc.add_paragraph("ВВЕДЕНИЕ")
    heading1 = doc.add_paragraph("ТЕОРЕТИЧЕСКИЕ ОСНОВЫ КОНКУРЕНТОСПОСОБНОСТИ")
    heading1.style = "Heading 1"
    _add_fake_word_numbering(heading1)
    heading2 = doc.add_paragraph("Понятие конкурентоспособности организации")
    heading2.style = "Heading 2"
    _add_fake_word_numbering(heading2, ilvl_value="1")
    doc.add_paragraph("Обычный текст подраздела.")

    with tempfile.TemporaryDirectory() as tmp:
        inp = os.path.join(tmp, "in.docx")
        out = os.path.join(tmp, "out.docx")
        doc.save(inp)
        process_document(inp, out)
        result = Document(out)

    formatted_h1 = _find_paragraph_starting_with(
        result,
        "1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ КОНКУРЕНТОСПОСОБНОСТИ",
    )
    if formatted_h1 is None:
        return _result(False, "Word-autonumbered Heading 1 did not get plain-text chapter number")

    if paragraph_has_numbering(formatted_h1):
        return _result(False, "Heading 1 Word numbering remained after formatting")

    if _style_name(formatted_h1) not in {"heading 1", "заголовок 1"}:
        return _result(False, f"autonumbered heading1 style is {_style_name(formatted_h1)!r}")

    formatted_h2 = _find_paragraph_starting_with(
        result,
        "1.1. Понятие конкурентоспособности организации",
    )
    if formatted_h2 is None:
        return _result(False, "following autonumbered Heading 2 was not normalized under Heading 1")

    return _result(True, "Word-autonumbered Heading 1 remains supported")


def test_t2_word_numbered_body_items_not_promoted_to_headings() -> tuple[bool, str]:
    """
    Word-numbered body list items are not heading evidence by themselves.
    This protects real coursework lists such as "Правление и Совет директоров"
    from becoming artificial "3.1." / "8.1." Heading 2 lines.
    """
    from guides.coursework_kfu_2025.safe_formatter import (
        auto_detect_heading2,
        auto_detect_numbered_heading1,
        is_likely_numbered_heading2_candidate,
        is_probable_body_list_item,
        normalize_heading2_numbering,
    )

    doc = Document()
    previous = doc.add_paragraph("Организационная структура включает несколько элементов.")
    item = doc.add_paragraph("Правление и Совет директоров")
    _add_fake_word_numbering(item)

    if not is_probable_body_list_item(item, prev_paragraph=previous, prev_kind="body_text"):
        return _result(False, "Word-numbered body item was not classified as body_list_item")

    if auto_detect_heading2(item, current_chapter_num=3, next_paragraph_num=1, prev_kind="body_text"):
        return _result(False, "Word-numbered body item auto-detected as heading2")

    if is_likely_numbered_heading2_candidate(item, 3, 1, prev_kind="body_text"):
        return _result(False, "Word-numbered body item considered likely heading2 candidate")

    if auto_detect_numbered_heading1(item, current_chapter_num=3):
        return _result(False, "Word-numbered body item auto-detected as heading1")

    before = item.text
    normalized = normalize_heading2_numbering(item, 3, 1)
    if normalized is not None or item.text != before:
        return _result(False, f"body item was renumbered: normalized={normalized!r}, text={item.text!r}")

    return _result(True, "Word-numbered body items stay body/list items")


def test_t2_numbered_sentence_not_promoted_to_heading1() -> tuple[bool, str]:
    """
    A numbered sentence-like body paragraph must not be uppercased as Heading 1.
    Real Heading 1 syntax without sentence boundary remains allowed.
    """
    from guides.coursework_kfu_2025.classifier import parse_heading1
    from guides.coursework_kfu_2025.safe_formatter import is_heading1_promotion_safe

    doc = Document()
    body_sentence = doc.add_paragraph("1. Маркетинговый подход. Данный подход")
    parsed = parse_heading1(body_sentence.text)
    if not parsed:
        return _result(False, "test setup failed: parse_heading1 did not parse numbered sentence")
    if is_heading1_promotion_safe(body_sentence, parsed):
        return _result(False, "sentence-like numbered body paragraph considered safe heading1")

    real_heading = doc.add_paragraph("1. ТЕОРЕТИЧЕСКИЕ АСПЕКТЫ КОНКУРЕНТОСПОСОБНОСТИ")
    parsed_real = parse_heading1(real_heading.text)
    if not parsed_real or not is_heading1_promotion_safe(real_heading, parsed_real):
        return _result(False, "real explicit heading1 was rejected")

    return _result(True, "numbered sentence rejected; real heading accepted")


def test_t2_chapter_colon_heading_repaired_without_colon_artifact() -> tuple[bool, str]:
    """'Глава 2: Название' becomes '2. НАЗВАНИЕ', never '2.: НАЗВАНИЕ'."""
    from guides.coursework_kfu_2025.safe_formatter import smart_repair_heading1

    doc = Document()
    paragraph = doc.add_paragraph("Глава 2: Практические аспекты критериев")

    if not smart_repair_heading1(paragraph, paragraph.text):
        return _result(False, "smart_repair_heading1 did not repair chapter heading")

    expected = "2. ПРАКТИЧЕСКИЕ АСПЕКТЫ КРИТЕРИЕВ"
    if paragraph.text != expected:
        return _result(False, f"unexpected repaired heading: {paragraph.text!r}")

    return _result(True, "chapter heading colon artifact removed")


def test_t2_real_coursework_17_heading_regression() -> tuple[bool, str]:
    """
    Real regression: body/list paragraphs in coursework 17 must not become
    artificial headings such as "3.1. Правление..." or ALL CAPS list items.
    """
    from guides.coursework_kfu_2025.safe_formatter import is_empty_paragraph

    fixture = Path(
        "/Users/mac/Desktop/курсовые/"
        "курсова 17. Критерии и показатели конкурентоспособности организации.docx"
    )
    if not fixture.exists():
        return _result(True, f"fixture not present, skipped: {fixture}")

    with tempfile.TemporaryDirectory() as tmp:
        out_path = Path(tmp) / "coursework_17_formatted.docx"
        try:
            format_docx(str(fixture), str(out_path))
        except Exception as e:
            return _result(False, f"formatter raised on real fixture: {e}\n{traceback.format_exc()}")

        doc = Document(str(out_path))
        paragraphs = doc.paragraphs
        texts = [" ".join(p.text.split()) for p in paragraphs if " ".join(p.text.split())]

    forbidden = [
        "1. МАРКЕТИНГОВЫЙ ПОДХОД. ДАННЫЙ ПОДХОД",
        "1.1. Доля рынка продукции предприятия",
        "3.1. Правление и Совет директоров",
        "3.2. Интеграция с международными научными центрами",
        "8.1. Повышение экспортного потенциала",
        "2.:",
    ]
    found_forbidden = [
        marker
        for marker in forbidden
        if any(text.startswith(marker) or marker in text for text in texts)
    ]
    if found_forbidden:
        return _result(False, f"false heading markers found: {found_forbidden}")

    required = [
        "ВВЕДЕНИЕ",
        "1. ТЕОРЕТИЧЕСКИЕ АСПЕКТЫ",
        "1.1. Понятие",
        "2. ПРАКТИЧЕСКИЕ АСПЕКТЫ",
        "2.1. Общая характеристика",
        "ЗАКЛЮЧЕНИЕ",
        "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    ]
    missing = [
        marker
        for marker in required
        if not any(marker.lower() in text.lower() for text in texts)
    ]
    if missing:
        return _result(False, f"real headings missing after formatting: {missing}")

    for idx, paragraph in enumerate(paragraphs):
        if " ".join(paragraph.text.split()).startswith("1.3. Методы оценки конкурентоспособности"):
            if idx < 1 or not is_empty_paragraph(paragraphs[idx - 1]):
                return _result(False, "real fixture: missing blank before 1.3 heading")
            if idx >= 2 and is_empty_paragraph(paragraphs[idx - 2]):
                return _result(False, "real fixture: double blank before 1.3 heading")
            break
    else:
        return _result(False, "real fixture: 1.3 heading missing")

    return _result(True, "real coursework 17 heading regression is clean")


def test_t3_reference_subheading_centred() -> tuple[bool, str]:
    """
    After formatting, reference section headers must be CENTER aligned, bold,
    preceded by exactly one empty paragraph.
    Source entries must use regular body-style indentation:
    left=0, firstLine≈709 twips, no hanging indent.
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from guides.coursework_kfu_2025.safe_formatter import process_document
    import tempfile, os

    doc = Document()
    doc.add_paragraph("Введение")
    doc.add_paragraph("")
    doc.add_paragraph("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    doc.add_paragraph("Официальные материалы")
    doc.add_paragraph("1. Некий закон.")
    doc.add_paragraph("Интернет-ресурсы")
    doc.add_paragraph("2. Некий сайт.")

    with tempfile.TemporaryDirectory() as tmp:
        inp = os.path.join(tmp, "in.docx")
        out = os.path.join(tmp, "out.docx")
        doc.save(inp)
        process_document(inp, out)
        result_doc = Document(out)

    paras = list(result_doc.paragraphs)
    sh_idx = next((i for i, p in enumerate(paras) if "официальные" in (p.text or "").lower()), None)
    if sh_idx is None:
        return _result(False, "subheading paragraph not found in output")

    sh = paras[sh_idx]
    if sh.alignment != WD_ALIGN_PARAGRAPH.CENTER:
        return _result(False, f"subheading not centred: alignment={sh.alignment}")

    pPr_sh = sh._element.find(qn("w:pPr"))
    ind_sh = pPr_sh.find(qn("w:ind")) if pPr_sh is not None else None
    if ind_sh is not None:
        fli = ind_sh.get(qn("w:firstLine"))
        left = ind_sh.get(qn("w:left"))
        hang = ind_sh.get(qn("w:hanging"))
        if hang and int(hang) > 100:
            return _result(False, f"subheading has hanging indent: {hang}")
        if fli and int(fli) > 100:
            return _result(False, f"subheading has first-line indent: {fli}")
        if left and int(left) > 100:
            return _result(False, f"subheading has left indent: {left}")

    bold_ok = any(r.bold for r in sh.runs if r.text.strip())
    if not bold_ok:
        return _result(False, "subheading runs are not bold")

    if sh_idx == 0 or (paras[sh_idx - 1].text or "").strip():
        return _result(False, "no empty paragraph before reference subheading")

    # Check source entry body-style indent
    source_paras = [p for p in paras if "некий закон" in (p.text or "").lower()]
    if source_paras:
        sp = source_paras[0]
        pPr_sp = sp._element.find(qn("w:pPr"))
        ind_sp = pPr_sp.find(qn("w:ind")) if pPr_sp is not None else None
        if ind_sp is None:
            return _result(False, "source entry has no w:ind")
        left_v = ind_sp.get(qn("w:left"))
        first_line_v = ind_sp.get(qn("w:firstLine"))
        hang_v = ind_sp.get(qn("w:hanging"))
        if left_v not in {None, "0"}:
            return _result(False, f"source entry left={left_v!r} (expected 0)")
        if not first_line_v or abs(int(first_line_v) - 709) > 60:
            return _result(False, f"source entry firstLine={first_line_v!r} (expected ≈709)")
        if hang_v is not None:
            return _result(False, f"source entry hanging={hang_v!r} (expected absent)")

    return _result(True, "reference subheading: centred, bold, blank before; source body indent OK")


def test_t4_citation_brackets_split() -> tuple[bool, str]:
    """
    Multi-source citation brackets split; single-source with page range get hyphen→en-dash.
    p. notation is supported. Single page [5, с. 12] unchanged.
    """
    from guides.coursework_kfu_2025.safe_formatter import _split_citation_brackets_in_text

    cases = [
        # Multi-source split
        ("[21, с. 30–45, 22, с. 21–33, 5, с. 3–8, 10]",
         "[21, с. 30–45], [22, с. 21–33], [5, с. 3–8], [10]"),
        ("[12; 13; 5]",      "[12], [13], [5]"),
        ("[21, 22]",         "[21], [22]"),
        # Single source — unchanged (but hyphen normalized)
        ("[21, с. 30–45]",   "[21, с. 30–45]"),
        ("[10]",             "[10]"),
        # Hyphen → en-dash in single source range
        ("[5, с. 12-15]",    "[5, с. 12–15]"),
        ("[5, с. 12–15]",    "[5, с. 12–15]"),
        # Single page (no range)
        ("[5, с. 12]",       "[5, с. 12]"),
        # p. notation → с. in output
        ("[5, p. 12-15]",    "[5, с. 12–15]"),
        ("[5, p. 12]",       "[5, с. 12]"),
        # Mixed in sentence
        ("по данным [21, 22], а также [5, с. 3–8, 10]",
         "по данным [21], [22], а также [5, с. 3–8], [10]"),
    ]
    failures = []
    for inp, expected in cases:
        got = _split_citation_brackets_in_text(inp)
        if got != expected:
            failures.append(f"Input:    {inp!r}\nExpected: {expected!r}\nGot:      {got!r}")
    if failures:
        return _result(False, "\n\n".join(failures))
    return _result(True, f"all {len(cases)} citation cases correct")


def test_t5_list_formatting() -> tuple[bool, str]:
    """
    Numeric list items (1)/1.) after a colon-ending paragraph become а)/б)/в).
    Level-1 items get left=906 hanging=198. Level-2 items get left=963 hanging=198.
    """
    from guides.coursework_kfu_2025.safe_formatter import _normalize_plain_list_paragraphs
    from docx.oxml.ns import qn

    doc = Document()
    intro = doc.add_paragraph("Выделяют следующие виды:")
    p1 = doc.add_paragraph("1) первый вид")
    p2 = doc.add_paragraph("2) второй вид")
    p3 = doc.add_paragraph("3) третий вид")

    _normalize_plain_list_paragraphs([intro, p1, p2, p3])

    if not p1.text.startswith("а)"):
        return _result(False, f"p1 not converted: {p1.text!r}")
    if not p2.text.startswith("б)"):
        return _result(False, f"p2 not converted: {p2.text!r}")
    if not p3.text.startswith("в)"):
        return _result(False, f"p3 not converted: {p3.text!r}")

    pPr = p1._element.find(qn("w:pPr"))
    ind = pPr.find(qn("w:ind")) if pPr is not None else None
    if ind is None:
        return _result(False, "no w:ind on level-1 item")
    left = ind.get(qn("w:left"))
    hang = ind.get(qn("w:hanging"))
    if left != "906" or hang != "198":
        return _result(False, f"wrong indent: left={left}, hanging={hang} (expected 906/198)")

    return _result(True, "list items converted and indented correctly ✓")


def test_figure_caption_spacing_and_blank_font() -> tuple[bool, str]:
    """
    Figure captions require exactly one blank before the caption, but no blank
    between the caption and its Источник line. Formatter-created blanks use
    body font size.
    """
    from guides.coursework_kfu_2025.safe_formatter import (
        ensure_single_blank_before_figure_captions,
        remove_empty_between_figure_caption_and_source,
    )

    doc = Document()
    doc.add_paragraph("Текст перед рисунком.")
    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph("Рис. 1.2.1. Схема процесса")
    doc.add_paragraph("")
    doc.add_paragraph("Источник: составлено автором.")

    ensure_single_blank_before_figure_captions(doc, 0)
    remove_empty_between_figure_caption_and_source(doc, 0)

    texts = [p.text for p in doc.paragraphs]
    expected = [
        "Текст перед рисунком.",
        "",
        "Рис. 1.2.1. Схема процесса",
        "Источник: составлено автором.",
    ]
    if texts != expected:
        return _result(False, f"unexpected paragraph layout: {texts!r}")

    blank = doc.paragraphs[1]
    run = blank.runs[0] if blank.runs else None
    if run is None:
        return _result(False, "blank paragraph has no run")

    sz = run._element.get_or_add_rPr().find(qn("w:sz"))
    if sz is None or sz.get(qn("w:val")) != "28":
        val = sz.get(qn("w:val")) if sz is not None else None
        return _result(False, f"blank font size is {val}, expected 28 half-points")

    return _result(True, "figure spacing and blank font are correct")


def test_heading2_late_spacing_before_13() -> tuple[bool, str]:
    """Late/final Heading 2 formatting still leaves one blank before 1.3."""
    from guides.coursework_kfu_2025.safe_formatter import is_empty_paragraph, process_document
    import tempfile, os

    doc = Document()
    doc.add_paragraph("ВВЕДЕНИЕ")
    doc.add_paragraph("1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ")
    doc.add_paragraph("1.2. Критерии конкурентоспособности организации")
    doc.add_paragraph("Текст подраздела 1.2.")
    doc.add_paragraph("Эти критерии позволят перейти к разделу 1.3.")
    doc.add_paragraph("1.3. Методы оценки конкурентоспособности организации")
    doc.add_paragraph("В процессе оценки конкурентоспособности применяются методы.")

    with tempfile.TemporaryDirectory() as tmp:
        inp = os.path.join(tmp, "in.docx")
        out = os.path.join(tmp, "out.docx")
        doc.save(inp)
        process_document(inp, out)
        result = Document(out)

    paragraphs = result.paragraphs
    target_idx = None
    for idx, paragraph in enumerate(paragraphs):
        if " ".join(paragraph.text.split()).startswith("1.3. Методы оценки"):
            target_idx = idx
            break

    if target_idx is None:
        return _result(False, "1.3 heading not found after formatting")

    if target_idx < 1 or not is_empty_paragraph(paragraphs[target_idx - 1]):
        return _result(False, "missing blank before 1.3 heading")

    if target_idx >= 2 and is_empty_paragraph(paragraphs[target_idx - 2]):
        return _result(False, "double blank before 1.3 heading")

    if target_idx + 1 >= len(paragraphs) or not is_empty_paragraph(paragraphs[target_idx + 1]):
        return _result(False, "missing blank after 1.3 heading")

    if target_idx + 2 < len(paragraphs) and is_empty_paragraph(paragraphs[target_idx + 2]):
        return _result(False, "double blank after 1.3 heading")

    return _result(True, "1.3 heading has exactly one blank before and after")


def test_blank_before_figure_block() -> tuple[bool, str]:
    """
    A drawing paragraph that follows body text must have exactly one blank before it.
    The caption/source spacing rules remain untouched.
    """
    from guides.coursework_kfu_2025.safe_formatter import (
        ensure_single_blank_before_figure_blocks,
        is_empty_paragraph,
        remove_empty_between_figure_caption_and_source,
        paragraph_has_drawing,
    )

    doc = Document()
    doc.add_paragraph("Текст перед рисунком.")
    drawing_p = doc.add_paragraph()
    drawing = OxmlElement("w:drawing")
    run = OxmlElement("w:r")
    run.append(drawing)
    drawing_p._element.append(run)
    doc.add_paragraph("Рис. 1.1.1. Схема процесса")
    doc.add_paragraph("")
    doc.add_paragraph("Источник: составлено автором.")

    ensure_single_blank_before_figure_blocks(doc, 0)
    remove_empty_between_figure_caption_and_source(doc, 0)

    drawing_idx = None
    for idx, paragraph in enumerate(doc.paragraphs):
        if paragraph_has_drawing(paragraph):
            drawing_idx = idx
            break

    if drawing_idx is None:
        return _result(False, "drawing paragraph not found")

    if drawing_idx < 1 or not is_empty_paragraph(doc.paragraphs[drawing_idx - 1]):
        return _result(False, "missing blank before drawing paragraph")

    if drawing_idx >= 2 and is_empty_paragraph(doc.paragraphs[drawing_idx - 2]):
        return _result(False, "double blank before drawing paragraph")

    texts = [p.text for p in doc.paragraphs]
    expected = [
        "Текст перед рисунком.",
        "",
        "",
        "Рис. 1.1.1. Схема процесса",
        "Источник: составлено автором.",
    ]
    if texts != expected:
        return _result(False, f"unexpected figure block layout: {texts!r}")

    blank = doc.paragraphs[drawing_idx - 1]
    run = blank.runs[0] if blank.runs else None
    if run is None:
        return _result(False, "blank before drawing has no run")

    sz = run._element.get_or_add_rPr().find(qn("w:sz"))
    if sz is None or sz.get(qn("w:val")) != "28":
        val = sz.get(qn("w:val")) if sz is not None else None
        return _result(False, f"blank before drawing font size is {val}, expected 28 half-points")

    return _result(True, "drawing paragraph has one TNR 14 blank before it")


# ── Runner ────────────────────────────────────────────────────────────────────

def run_all() -> None:
    tests = [
        ("A  | rule4 does not delete images",          test_a_rule4_does_not_delete_images),
        ("A  | _para_has_image helper",                test_a_para_has_image_helper),
        ("C  | continuation length guard",             test_c_continuation_length_guard),
        ("C  | strict caption-number extraction",      test_c_caption_number_extraction_strict),
        ("C  | merge invalid manual split",            test_c_apply_table_merging_rebuilds_invalid_split),
        ("C  | keep valid manual split",               test_c_apply_table_merging_keeps_valid_manual_split),
        ("C  | rebuild loose manual marker",           test_c_apply_table_merging_rebuilds_marker_without_keep_next),
        ("C  | heuristic split disabled",              test_c_apply_table_continuation_does_not_heuristic_split),
        ("C  | width normalisation only",              test_c_apply_table_continuation_width_normalization_only),
        ("C  | no-split double-run idempotency",       test_c_apply_table_continuation_no_split_double_run_idempotent),
        ("C  | rendered split LO fallback",            test_c_apply_rendered_table_continuation_warns_when_lo_unavailable),
        ("C  | rendered split PDF fallback",           test_c_apply_rendered_table_continuation_warns_when_pdf_analysis_fails),
        ("C  | rendered single-boundary split",        test_c_rendered_split_single_boundary_success),
        ("C  | rendered preserves manual split",       test_c_rendered_split_preserves_valid_manual_split),
        ("C  | rendered ambiguity skip",               test_c_rendered_split_skips_ambiguous_repeated_rows),
        ("C  | rendered merged-boundary skip",         test_c_rendered_split_skips_merged_boundary_conflict),
        ("C  | rendered marker formatting",            test_c_rendered_split_marker_is_right_aligned),
        ("C  | rendered caption number/fallback",      test_c_rendered_split_caption_number_and_fallback),
        ("C  | rendered whole-table move",             test_c_rendered_start_page_moves_whole_table_without_complete_data_row),
        ("C  | rendered start-page ambiguity skip",    test_c_rendered_start_page_skips_ambiguous_usability),
        ("C  | rendered start-page complete row",      test_c_rendered_start_page_keeps_table_with_clear_complete_data_row),
        ("C  | vMerge guard",                          test_c_vmerge_guard_rejects_boundary_inside_merge_zone),
        ("B1 | tblW updated after optimization",       test_b1_tblW_updated_after_col_optimization),
        ("B1 | _MIN_COL_PT ≤ 20",                     test_b1_min_col_pt_is_20),
        ("B2 | keepTogether on table_caption",         test_b2_keep_together_on_table_caption),
        ("B2 | keepTogether on heading1/heading2",     test_b2_keep_together_on_headings),
        ("B2 | rule6 keepWithNext through empty para", test_b2_rule6_propagates_through_empty_para),
        ("B2 | table source/note chained",             test_b2_table_source_note_normalised_and_chained),
        ("B2 | image height from wp:extent cy",        test_b2_image_height_from_emu),
        ("B3 | footnote para: 10pt TNR no bold",       test_b3_format_footnote_para_applies_10pt_tnr),
        ("C2 | empty para image→caption removed",      test_c2_empty_para_between_image_and_caption_removed),
        ("C2 | numeric column minimum protected",      test_c2_number_column_minimum),
        ("T1 | ё→е normalisation (midword uppercase fix)", test_yo_normalisation_midword_uppercase),
        ("T_indent | body paragraph left=0 firstLine=709", test_t_indent_body_paragraph_left_zero),
        ("T2 | 'Глава N' without title → heading1", test_t2_chapter_heading_without_title),
        ("T2 | manual heading2 still works", test_t2_manual_heading2_still_promoted),
        ("T2 | Word-autonumbered heading2 still works", test_t2_word_autonumbered_heading2_with_style_still_promoted),
        ("T2 | Word-autonumbered heading1 still works", test_t2_word_autonumbered_heading1_with_style_still_promoted),
        ("T2 | Word-numbered body items stay body/list", test_t2_word_numbered_body_items_not_promoted_to_headings),
        ("T2 | numbered sentence not promoted to heading1", test_t2_numbered_sentence_not_promoted_to_heading1),
        ("T2 | chapter colon heading repaired", test_t2_chapter_colon_heading_repaired_without_colon_artifact),
        ("T2 | real coursework 17 heading regression", test_t2_real_coursework_17_heading_regression),
        ("T3 | reference subheading centred + source indent", test_t3_reference_subheading_centred),
        ("T4 | citation brackets split + p. notation + hyphen→en-dash", test_t4_citation_brackets_split),
        ("T5 | list а)/б)/в) formatting", test_t5_list_formatting),
        ("T6 | figure caption spacing + blank font", test_figure_caption_spacing_and_blank_font),
        ("T6 | heading2 late spacing before 1.3", test_heading2_late_spacing_before_13),
        ("T6 | blank before figure block", test_blank_before_figure_block),
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
