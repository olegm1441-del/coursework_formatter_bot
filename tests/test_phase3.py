"""
Fast Phase 3/product-rule regression tests.

Run from repo root:
    python -m pytest tests/test_phase3.py -v
or directly:
    python tests/test_phase3.py

The default runner stays cheap: synthetic DOCX/XML checks and isolated unit tests
only. Real asset formatting smoke checks are opt-in via:
    KPFU_RUN_LONG_PHASE3_TESTS=1 python tests/test_phase3.py

Product-rule coverage:
  A  — Figure deletion: images survive Rule 4 (paragraphs with w:drawing never removed)
  C  — Student continuation length: _is_student_continuation detects ≤30 char texts
  B1 — tblW fix: _optimize_table_col_widths updates w:tblW after scaling
  B2 — keepTogether, Rule 6 propagation, image height from wp:extent
  B3 — Footnote standardisation
  C2 — Empty para between image and caption removed; numeric column minimums
  T2 — Heading paragraphs/styles must not use Word autonumbering; manual
       heading text numbering remains literal text.
  M1/S1 — Marker/prototype table split rules for ordinary and appendix tables.

  NOTE: Tests for LRPB-based table splitting (B, B1-stale/valid, C2-fits-1-page,
  C-student-merges) were removed when apply_table_merging / apply_table_continuation
  were stubbed out.  See module docstring in table_continuation.py for the future
  LibreOffice-based plan.
"""

from __future__ import annotations

import io
import logging
import os
import re
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


def _section_break_positions(doc: Document) -> list[tuple[int, str]]:
    positions: list[tuple[int, str]] = []
    for idx, paragraph in enumerate(doc.paragraphs):
        pPr = paragraph._element.pPr
        if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
            positions.append((idx, paragraph.text.strip()))
    return positions




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


def test_a_rule4_preserves_front_matter_section_breaks() -> tuple[bool, str]:
    """
    Product rule: title, contents, and introduction are separated by structural
    section breaks. Rule 4 may delete visual blank paragraphs, but never a
    paragraph carrying w:sectPr.
    """
    from guides.coursework_kfu_2025.safe_formatter import process_document

    doc = Document()
    probe = doc.add_paragraph("Титульная строка")
    h_per_para = _estimate_para_height(probe)
    probe._element.getparent().remove(probe._element)

    body_h = _body_height_pt(doc)
    title_lines = max(1, int(body_h / h_per_para))
    for i in range(title_lines):
        doc.add_paragraph(f"Титульная строка {i + 1}")
    doc.add_paragraph("")
    doc.add_paragraph("СОДЕРЖАНИЕ")
    doc.add_paragraph("ВВЕДЕНИЕ 3")
    doc.add_paragraph("1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ 4")
    doc.add_paragraph("")
    doc.add_paragraph("ВВЕДЕНИЕ")
    doc.add_paragraph("Текст введения.")

    with tempfile.TemporaryDirectory() as tmp:
        inp = Path(tmp) / "front_matter_in.docx"
        out = Path(tmp) / "front_matter_out.docx"
        doc.save(inp)
        process_document(inp, out)

        formatted = Document(str(out))
        before_positions = _section_break_positions(formatted)
        if len(before_positions) < 2:
            return _result(False, f"front matter section breaks missing after Phase 1: {before_positions!r}")

        apply_rule4_empty_first_lines(formatted)
        after_positions = _section_break_positions(formatted)

    if len(after_positions) != len(before_positions):
        return _result(
            False,
            f"Rule 4 removed structural section break(s): before={before_positions!r} after={after_positions!r}",
        )
    if after_positions != before_positions:
        return _result(
            False,
            f"Rule 4 moved structural section break(s): before={before_positions!r} after={after_positions!r}",
        )
    return _result(True, "Rule 4 preserved front matter section breaks")


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
    """
    Product rule: generated ordinary-table continuation markers are
    right-aligned and keep the existing continuation marker formatting.
    """
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
    page_break = pPr.find(qn("w:pageBreakBefore")) if pPr is not None else None
    jc = pPr.find(qn("w:jc")) if pPr is not None else None
    ind = pPr.find(qn("w:ind")) if pPr is not None else None
    keep = pPr.find(qn("w:keepNext")) if pPr is not None else None
    sz = markers[0]._element.find(".//" + qn("w:sz"))
    if page_break is None:
        return _result(False, "marker pageBreakBefore missing")
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


def test_c_rendered_start_page_first_row_spill_moves_whole_table() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    doc = Document()
    caption = doc.add_paragraph("Таблица 2.3.1")
    doc.add_paragraph("Структура прямой экономии ТТС при переходе на ЭДО")
    tbl = doc.add_table(rows=3, cols=3)
    tbl.rows[0].cells[0].text = "Статья"
    tbl.rows[0].cells[1].text = "Значение"
    tbl.rows[0].cells[2].text = "Комментарий"
    tbl.rows[1].cells[0].text = "Почтовые расходы"
    tbl.rows[1].cells[1].text = "31–33"
    tbl.rows[1].cells[2].text = "отказ от бумажных отправлений"
    tbl.rows[2].cells[0].text = "Печать"
    tbl.rows[2].cells[1].text = "4–5"
    tbl.rows[2].cells[2].text = "сокращение печати"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "first_row_spill.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "first_row_spill.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: [
                PdfLine("Таблица 2.3.1", 1, 683.2, 695.0),
                PdfLine("Структура прямой экономии ТТС при переходе на ЭДО", 1, 707.4, 719.0),
                PdfLine("Статья Значение Комментарий", 1, 731.5, 743.0),
                PdfLine("Почтовые расходы 31–33 отказ от бумажных", 1, 759.6, 771.0),
                PdfLine("Статья Значение Комментарий", 2, 58.8, 70.0),
                PdfLine("отправлений", 2, 86.8, 98.0),
                PdfLine("Печать 4–5 сокращение печати", 2, 108.2, 120.0),
            ]
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))
        reread_caption = next(p for p in reread.paragraphs if p.text == caption.text)

    pPr = reread_caption._element.find(qn("w:pPr"))
    page_break = pPr.find(qn("w:pageBreakBefore")) if pPr is not None else None

    if n != 1:
        return _result(False, f"first-row spill should trigger whole-table move, got {n}")
    if page_break is None:
        return _result(False, "caption did not receive pageBreakBefore after first-row spill")
    if len(reread.tables) != 1:
        return _result(False, f"whole-table move should not split tables, got {len(reread.tables)}")
    return _result(True, "first-row spill triggered whole-table move")


def test_c_rendered_start_page_skips_existing_page_break_candidate() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    doc = Document()
    first_caption = doc.add_paragraph("Таблица 1.2.2")
    first_caption.paragraph_format.page_break_before = True
    first_tbl = doc.add_table(rows=2, cols=2)
    first_tbl.rows[0].cells[0].text = "Показатель"
    first_tbl.rows[0].cells[1].text = "Эффект"
    first_tbl.rows[1].cells[0].text = "Первый показатель"
    first_tbl.rows[1].cells[1].text = "переход на электронный обмен"

    second_caption = doc.add_paragraph("Таблица 2.3.3")
    second_tbl = doc.add_table(rows=2, cols=2)
    second_tbl.rows[0].cells[0].text = "Год"
    second_tbl.rows[0].cells[1].text = "Комментарий"
    second_tbl.rows[1].cells[0].text = "Первый год"
    second_tbl.rows[1].cells[1].text = "обучение сотрудников"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "skip_existing_break.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "skip_existing_break.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: [
                PdfLine("Таблица 1.2.2", 1, 680.0, 692.0),
                PdfLine("Показатель Эффект", 1, 705.0, 717.0),
                PdfLine("Первый показатель", 2, 80.0, 92.0),
                PdfLine("переход на электронный обмен", 2, 100.0, 112.0),
                PdfLine("Таблица 2.3.3", 3, 680.0, 692.0),
                PdfLine("Год Комментарий", 3, 705.0, 717.0),
                PdfLine("Первый год", 4, 80.0, 92.0),
                PdfLine("обучение сотрудников", 4, 100.0, 112.0),
            ]
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))
        first = next(p for p in reread.paragraphs if p.text == first_caption.text)
        second = next(p for p in reread.paragraphs if p.text == second_caption.text)

    first_pPr = first._element.find(qn("w:pPr"))
    second_pPr = second._element.find(qn("w:pPr"))
    first_pb = first_pPr.find(qn("w:pageBreakBefore")) if first_pPr is not None else None
    second_pb = second_pPr.find(qn("w:pageBreakBefore")) if second_pPr is not None else None

    if n != 1:
        return _result(False, f"expected one later whole-table move, got {n}")
    if first_pb is None:
        return _result(False, "existing pageBreakBefore was lost from first caption")
    if second_pb is None:
        return _result(False, "later candidate did not receive pageBreakBefore")
    if len(reread.tables) != 2:
        return _result(False, f"whole-table move should not split tables, got {len(reread.tables)}")
    return _result(True, "existing page-break candidate skipped and later candidate moved")


def test_c_rendered_start_page_upgrades_disabled_page_break() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    doc = Document()
    caption = doc.add_paragraph("Таблица 2.4.1")
    pPr = caption._element.get_or_add_pPr()
    disabled_break = OxmlElement("w:pageBreakBefore")
    disabled_break.set(qn("w:val"), "0")
    pPr.append(disabled_break)

    tbl = doc.add_table(rows=2, cols=2)
    tbl.rows[0].cells[0].text = "Показатель"
    tbl.rows[0].cells[1].text = "Комментарий"
    tbl.rows[1].cells[0].text = "Первый показатель"
    tbl.rows[1].cells[1].text = "обучение сотрудников"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "disabled_page_break.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "disabled_page_break.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: [
                PdfLine("Таблица 2.4.1", 1, 680.0, 692.0),
                PdfLine("Показатель Комментарий", 1, 705.0, 717.0),
                PdfLine("Первый показатель", 2, 80.0, 92.0),
                PdfLine("обучение сотрудников", 2, 100.0, 112.0),
            ]
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))
        reread_caption = next(p for p in reread.paragraphs if p.text == "Таблица 2.4.1")

    reread_pPr = reread_caption._element.find(qn("w:pPr"))
    page_break = reread_pPr.find(qn("w:pageBreakBefore")) if reread_pPr is not None else None
    page_break_val = page_break.get(qn("w:val")) if page_break is not None else None

    if n != 1:
        return _result(False, f"disabled pageBreakBefore should not block move, got {n}")
    if page_break is None:
        return _result(False, "disabled pageBreakBefore was not upgraded")
    if page_break_val in {"0", "false", "False", "off"}:
        return _result(False, f"pageBreakBefore still disabled: {page_break_val!r}")
    if len(reread.tables) != 1:
        return _result(False, f"whole-table move should not split tables, got {len(reread.tables)}")
    return _result(True, "disabled pageBreakBefore upgraded to active")


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


def test_c_rendered_start_page_first_row_spill_needs_strong_next_page_evidence() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    doc = Document()
    doc.add_paragraph("Таблица 2.3.6")
    doc.add_paragraph("Промежуточный эффект")
    tbl = doc.add_table(rows=3, cols=3)
    tbl.rows[0].cells[0].text = "Статья"
    tbl.rows[0].cells[1].text = "Значение"
    tbl.rows[0].cells[2].text = "Комментарий"
    tbl.rows[1].cells[0].text = "Почтовые расходы"
    tbl.rows[1].cells[1].text = "31–33"
    tbl.rows[1].cells[2].text = "отказ от бумажных отправлений"
    tbl.rows[2].cells[0].text = "Печать"
    tbl.rows[2].cells[1].text = "4–5"
    tbl.rows[2].cells[2].text = "сокращение печати"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "weak_spill.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "weak_spill.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: [
                PdfLine("Таблица 2.3.6", 1, 683.2, 695.0),
                PdfLine("Промежуточный эффект", 1, 707.4, 719.0),
                PdfLine("Статья Значение Комментарий", 1, 731.5, 743.0),
                PdfLine("Почтовые расходы 31–33 отказ от бумажных", 1, 759.6, 771.0),
                PdfLine("Печать 4–5", 1, 776.0, 788.0),
                PdfLine("отправлений", 2, 86.8, 98.0),
                PdfLine("сокращение печати", 2, 108.2, 120.0),
            ]
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))
        reread_caption = next(p for p in reread.paragraphs if p.text == "Таблица 2.3.6")

    pPr = reread_caption._element.find(qn("w:pPr"))
    page_break = pPr.find(qn("w:pageBreakBefore")) if pPr is not None else None

    if n != 0:
        return _result(False, f"weak next-page evidence should not trigger move, got {n}")
    if page_break is not None:
        return _result(False, "weak next-page evidence still added pageBreakBefore")
    return _result(True, "weak next-page evidence does not trigger spill detection")


def test_c_rendered_start_page_first_row_spill_ignores_later_prose_token_reuse() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    doc = Document()
    doc.add_paragraph("Таблица 2.3.7")
    doc.add_paragraph("Промежуточный эффект")
    tbl = doc.add_table(rows=3, cols=3)
    tbl.rows[0].cells[0].text = "Статья"
    tbl.rows[0].cells[1].text = "Значение"
    tbl.rows[0].cells[2].text = "Комментарий"
    tbl.rows[1].cells[0].text = "Почтовые расходы"
    tbl.rows[1].cells[1].text = "31–33"
    tbl.rows[1].cells[2].text = "отказ от бумажных отправлений"
    tbl.rows[2].cells[0].text = "Печать"
    tbl.rows[2].cells[1].text = "4–5"
    tbl.rows[2].cells[2].text = "сокращение печати"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "prose_reuse_spill.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "prose_reuse_spill.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: [
                PdfLine("Таблица 2.3.7", 1, 683.2, 695.0),
                PdfLine("Промежуточный эффект", 1, 707.4, 719.0),
                PdfLine("Статья Значение Комментарий", 1, 731.5, 743.0),
                PdfLine("Почтовые расходы 31–33 отказ от бумажных", 1, 759.6, 771.0),
                PdfLine("Печать 4–5", 1, 776.0, 788.0),
                PdfLine("Статья Значение Комментарий", 2, 58.8, 70.0),
                PdfLine("В тексте обсуждаются отправлений документов и риски обмена", 2, 86.8, 98.0),
                PdfLine("сокращение печати", 2, 108.2, 120.0),
            ]
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze

        reread = Document(str(path))
        reread_caption = next(p for p in reread.paragraphs if p.text == "Таблица 2.3.7")

    pPr = reread_caption._element.find(qn("w:pPr"))
    page_break = pPr.find(qn("w:pageBreakBefore")) if pPr is not None else None

    if n != 0:
        return _result(False, f"later prose token reuse should not trigger move, got {n}")
    if page_break is not None:
        return _result(False, "later prose token reuse still added pageBreakBefore")
    return _result(True, "later prose token reuse does not trigger spill detection")


def test_c_rendered_decision_logging_for_ambiguous_skip() -> tuple[bool, str]:
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

    log_stream = io.StringIO()
    handler = logging.StreamHandler(log_stream)
    handler.setFormatter(logging.Formatter("%(message)s"))
    old_level = tc.logger.level
    tc.logger.addHandler(handler)
    tc.logger.setLevel(logging.INFO)

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "ambiguous_logging.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "ambiguous_logging.pdf"
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
            tc.logger.removeHandler(handler)
            tc.logger.setLevel(old_level)

    logs = log_stream.getvalue()
    expected_fragments = [
        "rendered_table_continuation_enter tables=1 pdf_lines=5",
        "rendered_whole_table_candidate table_idx=0 caption=2.2.4",
        "pdf_caption_matches=1 strict_caption_found=True start_page_usability=ambiguous",
        "rendered_split_candidate table_idx=0 rows=3 skip=row_mapping_ambiguous",
        "rendered_final_decision action=rendered_skip_ambiguous",
    ]
    missing = [fragment for fragment in expected_fragments if fragment not in logs]

    if n != 0:
        return _result(False, f"ambiguous logging scenario should not mutate, got {n}")
    if missing:
        return _result(False, f"missing log fragments: {missing!r}; logs={logs!r}")
    return _result(True, "ambiguous rendered decision logs are emitted")


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


def _add_fake_style_numbering(document: Document, style_name: str, ilvl_value: str = "0") -> None:
    style = document.styles[style_name]
    style_element = style.element
    pPr = style_element.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        style_element.append(pPr)

    existing = pPr.find(qn("w:numPr"))
    if existing is not None:
        pPr.remove(existing)

    numPr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), ilvl_value)
    num_id = OxmlElement("w:numId")
    num_id.set(qn("w:val"), "42")
    numPr.append(ilvl)
    numPr.append(num_id)
    pPr.append(numPr)


def _style_has_numbering(document: Document, style_name: str) -> bool:
    style = document.styles[style_name]
    pPr = style.element.find(qn("w:pPr"))
    if pPr is None:
        return False
    return pPr.find(qn("w:numPr")) is not None


def _paragraph_has_direct_numbering(paragraph) -> bool:
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        return False
    return pPr.find(qn("w:numPr")) is not None


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


def test_t2_heading_style_numbering_is_removed() -> tuple[bool, str]:
    """
    Product rule: headings must not use Word autonumbering.
    Heading styles may carry w:numPr, which renders extra numbering even when
    heading paragraphs have no direct numPr. Manual numbering in heading text
    must remain as literal text.
    """
    from guides.coursework_kfu_2025.safe_formatter import process_document

    doc = Document()
    _add_fake_style_numbering(doc, "Heading 1", ilvl_value="0")
    _add_fake_style_numbering(doc, "Heading 2", ilvl_value="1")
    _add_fake_style_numbering(doc, "Heading 3", ilvl_value="2")

    h1_exact = doc.add_paragraph("ВВЕДЕНИЕ")
    h1_exact.style = "Heading 1"
    _add_fake_word_numbering(h1_exact)
    h1_chapter = doc.add_paragraph("1. Теоретические основы")
    h1_chapter.style = "Heading 1"
    _add_fake_word_numbering(h1_chapter)
    h2 = doc.add_paragraph("1.1. Понятие конкурентоспособности организации")
    h2.style = "Heading 2"
    _add_fake_word_numbering(h2, ilvl_value="1")
    doc.add_paragraph("Обычный текст подраздела.")
    doc.add_paragraph("Ненумерованный основной текст не должен получить numPr.")

    with tempfile.TemporaryDirectory() as tmp:
        inp = Path(tmp) / "in.docx"
        out = Path(tmp) / "out.docx"
        doc.save(inp)
        process_document(inp, out)
        formatted = Document(str(out))

    numbered_styles = [
        style_name
        for style_name in ("Heading 1", "Heading 2", "Heading 3")
        if _style_has_numbering(formatted, style_name)
    ]
    if numbered_styles:
        return _result(False, f"heading style numbering remained: {numbered_styles!r}")

    heading_texts = {
        "ВВЕДЕНИЕ",
        "1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ",
        "1.1. Понятие конкурентоспособности организации",
    }
    found = []
    for paragraph in formatted.paragraphs:
        text = " ".join((paragraph.text or "").split())
        if text in heading_texts:
            found.append(text)
            if _paragraph_has_direct_numbering(paragraph):
                return _result(False, f"direct heading numbering remained on {text!r}")
        elif text and _paragraph_has_direct_numbering(paragraph):
            return _result(False, f"numbering was added outside tables/headings: {text!r}")

    missing = sorted(heading_texts - set(found))
    if missing:
        return _result(False, f"manual heading text missing after formatting: {missing!r}")

    return _result(True, "heading style numbering removed while manual heading text stayed")


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


def test_marker_instrumentation_keeps_source_unchanged() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_markers import instrument_table_rows_copy

    doc = Document()
    tbl = doc.add_table(rows=3, cols=2)
    tbl.rows[0].cells[0].text = "Header"
    tbl.rows[1].cells[0].text = "Row one"
    tbl.rows[2].cells[0].text = "Row two"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "source.docx"
        workdir = Path(tmp) / "work"
        doc.save(src)
        before = src.read_bytes()

        instrumentation = instrument_table_rows_copy(src, 0, workdir=workdir, marker_font_size_pt=1)

        after = src.read_bytes()
        if before != after:
            return _result(False, "source docx changed after instrumentation")

        source_doc = Document(str(src))
        instrumented_doc = Document(str(instrumentation.instrumented_docx_path))
        source_text = " ".join(p.text for p in source_doc.paragraphs)
        if "KPFU_TMARK_" in source_text:
            return _result(False, "marker leaked into source document")

        marker_hits = sum(
            text.count("KPFU_TMARK_")
            for table in instrumented_doc.tables
            for row in table.rows
            for cell in row.cells
            for text in [cell.text]
        )
        if marker_hits != 3:
            return _result(False, f"expected 3 row markers in instrumented copy, got {marker_hits}")

    return _result(True, "instrumentation only changes temp copy")


def test_marker_instrumentation_only_targets_selected_table() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_markers import instrument_table_rows_copy

    doc = Document()
    first = doc.add_table(rows=2, cols=1)
    first.rows[0].cells[0].text = "First header"
    first.rows[1].cells[0].text = "First body"
    second = doc.add_table(rows=3, cols=1)
    second.rows[0].cells[0].text = "Second header"
    second.rows[1].cells[0].text = "Second row"
    second.rows[2].cells[0].text = "Third row"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "multi.docx"
        doc.save(src)
        instrumentation = instrument_table_rows_copy(src, 1, workdir=Path(tmp) / "work", marker_font_size_pt=1)
        instrumented = Document(str(instrumentation.instrumented_docx_path))

    first_text = " ".join(cell.text for row in instrumented.tables[0].rows for cell in row.cells)
    second_text = " ".join(cell.text for row in instrumented.tables[1].rows for cell in row.cells)
    if "KPFU_TMARK_" in first_text:
        return _result(False, "marker inserted into non-target table")
    if second_text.count("KPFU_TMARK_") != 3:
        return _result(False, f"expected markers only in target table rows, got text={second_text!r}")

    return _result(True, "only selected table was instrumented")


def test_marker_extract_handles_inline_text_and_missing_rows() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_markers import extract_row_pages_from_pdf_lines
    from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine

    result = extract_row_pages_from_pdf_lines(
        [
            PdfLine("prefixKPFU_TMARK_ABC123_T00_R000suffix", 27, 10.0, 20.0),
            PdfLine("bodyKPFU_TMARK_ABC123_T00_R001tail", 27, 30.0, 40.0),
        ],
        marker_salt="ABC123",
        table_index=0,
        total_rows=3,
    )

    if result.row_pages != {0: 27, 1: 27}:
        return _result(False, f"unexpected row_pages: {result.row_pages!r}")
    if result.found_rows != [0, 1]:
        return _result(False, f"unexpected found_rows: {result.found_rows!r}")
    if result.missing_rows != [2]:
        return _result(False, f"unexpected missing_rows: {result.missing_rows!r}")
    if result.duplicate_rows:
        return _result(False, f"unexpected duplicate_rows: {result.duplicate_rows!r}")

    return _result(True, "inline marker parsing and missing-row diagnostics work")


def test_marker_map_rows_to_pages_keep_temp_debug_paths() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_markers as tm

    doc = Document()
    tbl = doc.add_table(rows=3, cols=1)
    tbl.rows[0].cells[0].text = "Header"
    tbl.rows[1].cells[0].text = "Alpha"
    tbl.rows[2].cells[0].text = "Beta"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "source.docx"
        doc.save(src)

        seen_docx: dict[str, Path] = {}
        old_render = tm.render_docx_to_pdf
        old_analyze = tm.analyze_pdf_lines
        try:
            def fake_render(docx_path):
                seen_docx["path"] = Path(docx_path)
                pdf_dir = Path(tmp) / "pdf_keep"
                pdf_dir.mkdir(exist_ok=True)
                pdf_path = pdf_dir / "instrumented.pdf"
                pdf_path.write_bytes(b"%PDF-1.4\n")
                return pdf_path

            def fake_analyze(_pdf_path):
                inst_doc = Document(str(seen_docx["path"]))
                row_markers = []
                for row in inst_doc.tables[0].rows:
                    text = " ".join(cell.text for cell in row.cells)
                    match = re.search(r"KPFU_TMARK_[A-F0-9]{6}_T00_R\d{3}", text)
                    if not match:
                        raise AssertionError(f"marker not found in row text: {text!r}")
                    row_markers.append(match.group(0))
                return [
                    tm.PdfLine(f"left{row_markers[0]}right", 27, 10.0, 20.0),
                    tm.PdfLine(f"left{row_markers[1]}right", 27, 30.0, 40.0),
                    tm.PdfLine(f"left{row_markers[2]}right", 28, 50.0, 60.0),
                ]

            tm.render_docx_to_pdf = fake_render
            tm.analyze_pdf_lines = fake_analyze
            result = tm.map_table_rows_to_pages(src, 0, keep_temp=True)
        finally:
            tm.render_docx_to_pdf = old_render
            tm.analyze_pdf_lines = old_analyze

        if result.row_pages != {0: 27, 1: 27, 2: 28}:
            return _result(False, f"unexpected row_pages: {result.row_pages!r}")
        if result.instrumented_docx_path is None or not result.instrumented_docx_path.exists():
            return _result(False, "instrumented_docx_path was not preserved in keep_temp mode")
        if result.pdf_path is None or not result.pdf_path.exists():
            return _result(False, "pdf_path was not preserved in keep_temp mode")
        if result.marker_font_size_pt != 1:
            return _result(False, f"expected 1pt success path, got {result.marker_font_size_pt}")

    return _result(True, "keep_temp preserves instrumented DOCX/PDF and returns exact mapping")


def test_marker_map_rows_to_pages_falls_back_to_2pt_and_returns_debug_info() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_markers as tm

    doc = Document()
    tbl = doc.add_table(rows=3, cols=1)
    tbl.rows[0].cells[0].text = "Header"
    tbl.rows[1].cells[0].text = "Alpha"
    tbl.rows[2].cells[0].text = "Beta"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "source.docx"
        doc.save(src)

        seen_docx: dict[str, Path] = {}
        old_render = tm.render_docx_to_pdf
        old_analyze = tm.analyze_pdf_lines
        try:
            def fake_render(docx_path):
                seen_docx["path"] = Path(docx_path)
                pdf_dir = Path(tmp) / f"pdf_{Path(docx_path).stem}"
                pdf_dir.mkdir(exist_ok=True)
                pdf_path = pdf_dir / "instrumented.pdf"
                pdf_path.write_bytes(b"%PDF-1.4\n")
                return pdf_path

            def fake_analyze(_pdf_path):
                inst_doc = Document(str(seen_docx["path"]))
                full_text = " ".join(cell.text for row in inst_doc.tables[0].rows for cell in row.cells)
                markers = re.findall(r"KPFU_TMARK_[A-F0-9]{6}_T00_R\d{3}", full_text)
                if len(markers) != 3:
                    raise AssertionError(f"expected 3 markers, got {markers!r}")
                if "_2pt" not in seen_docx["path"].name:
                    return [tm.PdfLine(f"x{markers[0]}y", 27, 10.0, 20.0)]
                return [
                    tm.PdfLine(f"x{markers[0]}y", 27, 10.0, 20.0),
                    tm.PdfLine(f"x{markers[1]}y", 28, 30.0, 40.0),
                    tm.PdfLine(f"x{markers[1]}y", 29, 50.0, 60.0),
                ]

            tm.render_docx_to_pdf = fake_render
            tm.analyze_pdf_lines = fake_analyze
            result = tm.map_table_rows_to_pages(src, 0, keep_temp=False)
        finally:
            tm.render_docx_to_pdf = old_render
            tm.analyze_pdf_lines = old_analyze

        if result.marker_font_size_pt != 2:
            return _result(False, f"expected 2pt fallback, got {result.marker_font_size_pt}")
        if result.row_pages != {0: 27}:
            return _result(False, f"unexpected partial row_pages: {result.row_pages!r}")
        if result.missing_rows != [2]:
            return _result(False, f"unexpected missing_rows after fallback: {result.missing_rows!r}")
        if result.duplicate_rows != {1: [28, 29]}:
            return _result(False, f"unexpected duplicate_rows after fallback: {result.duplicate_rows!r}")
        if result.instrumented_docx_path is None or result.pdf_path is None:
            return _result(False, "debug paths should be preserved for incomplete diagnostics")
        if not result.instrumented_docx_path.exists() or not result.pdf_path.exists():
            return _result(False, "preserved debug paths do not exist")

    return _result(True, "1pt fallback to 2pt preserves diagnostics and debug artifacts")


def test_marker_instrumentation_rejects_invalid_table_index() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_markers import instrument_table_rows_copy

    doc = Document()
    doc.add_table(rows=1, cols=1)

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "source.docx"
        doc.save(src)
        try:
            instrument_table_rows_copy(src, 3, workdir=Path(tmp) / "work")
        except ValueError:
            return _result(True, "invalid table index rejected")
        except Exception as exc:
            return _result(False, f"unexpected exception type: {exc}")

    return _result(False, "expected ValueError for invalid table index")


def test_marker_page_span_summary() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_markers import summarize_row_page_spans

    spans = summarize_row_page_spans({
        0: 12,
        1: 12,
        2: 13,
        3: 13,
        5: 14,
    })
    triples = [(s.start_row, s.end_row, s.page_num) for s in spans]
    expected = [(0, 1, 12), (2, 3, 13), (5, 5, 14)]
    if triples != expected:
        return _result(False, f"unexpected page spans: {triples!r}")
    return _result(True, "row page spans are grouped correctly")


def test_marker_diagnose_all_tables_summary() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_markers as tm

    doc = Document()
    first = doc.add_table(rows=3, cols=1)
    first.rows[0].cells[0].text = "H1"
    first.rows[1].cells[0].text = "A"
    first.rows[2].cells[0].text = "B"
    second = doc.add_table(rows=2, cols=1)
    second.rows[0].cells[0].text = "H2"
    second.rows[1].cells[0].text = "C"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "diag.docx"
        doc.save(src)

        old_map = tm.map_table_rows_to_pages
        try:
            def fake_map(_docx_path, table_index, keep_temp=False):
                if table_index == 0:
                    return tm.TableMarkerResult(
                        row_pages={0: 12, 1: 12, 2: 13},
                        found_rows=[0, 1, 2],
                        missing_rows=[],
                        duplicate_rows={},
                        marker_font_size_pt=1,
                    )
                return tm.TableMarkerResult(
                    row_pages={0: 15},
                    found_rows=[0],
                    missing_rows=[1],
                    duplicate_rows={},
                    instrumented_docx_path=Path(tmp) / "inst.docx" if keep_temp else None,
                    pdf_path=Path(tmp) / "inst.pdf" if keep_temp else None,
                    marker_font_size_pt=2,
                )

            tm.map_table_rows_to_pages = fake_map
            diagnostics = tm.diagnose_all_tables(src, keep_temp=True)
        finally:
            tm.map_table_rows_to_pages = old_map

    if len(diagnostics) != 2:
        return _result(False, f"expected 2 diagnostics, got {len(diagnostics)}")
    if diagnostics[0].candidate_for_split is not True:
        return _result(False, "multi-page fully-mapped table should be candidate_for_split=yes")
    if [(s.start_row, s.end_row, s.page_num) for s in diagnostics[0].page_spans] != [(0, 1, 12), (2, 2, 13)]:
        return _result(False, f"unexpected first table spans: {diagnostics[0].page_spans!r}")
    if diagnostics[0].appendix_table is not False:
        return _result(False, "first table should not be marked as appendix table")
    if diagnostics[0].caption_detected is not False:
        return _result(False, "first table should not report caption without preceding paragraph")
    if diagnostics[1].candidate_for_split is not False:
        return _result(False, "table with missing rows should not be candidate_for_split")
    if diagnostics[1].missing_rows != [1]:
        return _result(False, f"unexpected missing rows: {diagnostics[1].missing_rows!r}")
    if diagnostics[1].marker_font_size_pt != 2:
        return _result(False, f"unexpected fallback font size: {diagnostics[1].marker_font_size_pt}")

    return _result(True, "document-level diagnostics summarize all tables")


def test_marker_diagnose_table_handles_mapping_error() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_markers as tm

    doc = Document()
    tbl = doc.add_table(rows=2, cols=1)
    tbl.rows[0].cells[0].text = "H"
    tbl.rows[1].cells[0].text = "A"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "error.docx"
        doc.save(src)

        old_map = tm.map_table_rows_to_pages
        try:
            def raise_map(_docx_path, _table_index, keep_temp=False):
                raise RuntimeError("render failed")

            tm.map_table_rows_to_pages = raise_map
            diagnostic = tm.diagnose_table(src, 0, keep_temp=False)
        finally:
            tm.map_table_rows_to_pages = old_map

    if diagnostic.error_message != "render failed":
        return _result(False, f"unexpected error_message: {diagnostic.error_message!r}")
    if diagnostic.candidate_for_split:
        return _result(False, "error diagnostic must not be candidate_for_split")
    if diagnostic.row_pages != {} or diagnostic.pages_detected != []:
        return _result(False, "error diagnostic should not report row/page mapping")
    return _result(True, "diagnose_table degrades to diagnostic error instead of crashing")


def test_marker_appendix_and_caption_metadata() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_markers as tm

    doc = Document()
    doc.add_paragraph("Таблица 1.1")
    doc.add_paragraph("Двухстрочный заголовок обычной таблицы")
    first = doc.add_table(rows=2, cols=1)
    first.rows[0].cells[0].text = "H1"
    first.rows[1].cells[0].text = "A"
    doc.add_paragraph("Приложение А")
    doc.add_paragraph("Длинная таблица по приложению")
    second = doc.add_table(rows=2, cols=1)
    second.rows[0].cells[0].text = "H2"
    second.rows[1].cells[0].text = "B"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "appendix.docx"
        doc.save(src)

        old_map = tm.map_table_rows_to_pages
        try:
            def fake_map(_docx_path, table_index, keep_temp=False):
                return tm.TableMarkerResult(
                    row_pages={0: 10, 1: 10},
                    found_rows=[0, 1],
                    missing_rows=[],
                    duplicate_rows={},
                    marker_font_size_pt=1,
                )

            tm.map_table_rows_to_pages = fake_map
            diagnostics = tm.diagnose_all_tables(src, keep_temp=False)
        finally:
            tm.map_table_rows_to_pages = old_map

    if diagnostics[0].caption_detected is not True or diagnostics[0].has_standard_table_caption is not True:
        return _result(False, "standard split table caption was not detected for first table")
    if diagnostics[0].appendix_table is not False:
        return _result(False, "first table should not be appendix table")
    if diagnostics[0].preceding_paragraph_text != "Двухстрочный заголовок обычной таблицы":
        return _result(False, f"immediate title context was not preserved: {diagnostics[0].preceding_paragraph_text!r}")
    if diagnostics[1].appendix_table is not True:
        return _result(False, "second table should be marked as appendix table")
    if diagnostics[1].caption_detected is not True:
        return _result(False, "appendix table title/caption should be detected")
    if diagnostics[1].has_standard_table_caption is not False:
        return _result(False, "appendix table title should not be treated as standard table caption")
    if diagnostics[1].preceding_paragraph_text != "Длинная таблица по приложению":
        return _result(False, f"unexpected preceding paragraph text: {diagnostics[1].preceding_paragraph_text!r}")

    return _result(True, "appendix and caption metadata are detected")


def test_marker_runtime_dry_run_clean_two_page_table_is_eligible() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.table_markers import TableMarkerDiagnostic, TablePageSpan

    diagnostic = TableMarkerDiagnostic(
        table_index=10,
        rows_count=21,
        pages_detected=[53, 54],
        row_pages={**{0: 53}, **{row: 53 for row in range(1, 18)}, 18: 54, 19: 54, 20: 54},
        found_rows=list(range(21)),
        missing_rows=[],
        duplicate_rows={},
        candidate_for_split=False,
        page_spans=[TablePageSpan(0, 17, 53), TablePageSpan(18, 20, 54)],
        appendix_table=True,
        caption_detected=True,
        has_standard_table_caption=False,
        preceding_paragraph_text="Расчет трудозатрат",
    )

    decision = tc._evaluate_marker_split_diagnostic(diagnostic, header_rows=1)
    if decision.eligible is not True:
        return _result(False, f"expected eligible decision, got {decision!r}")
    if decision.split_before_row != 18:
        return _result(False, f"expected split_before_row=18, got {decision.split_before_row!r}")
    return _result(True, "clean two-page marker mapping is marked eligible")


def test_marker_runtime_dry_run_skips_duplicate_rows() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.table_markers import TableMarkerDiagnostic

    diagnostic = TableMarkerDiagnostic(
        table_index=0,
        rows_count=4,
        pages_detected=[12, 13],
        row_pages={0: 12, 1: 12, 3: 13},
        found_rows=[0, 1, 3],
        missing_rows=[],
        duplicate_rows={2: [12, 13]},
        candidate_for_split=False,
        page_spans=[],
        appendix_table=False,
        caption_detected=True,
        has_standard_table_caption=True,
    )

    decision = tc._evaluate_marker_split_diagnostic(diagnostic, header_rows=1)
    if decision.eligible:
        return _result(False, "duplicate rows should skip dry-run eligibility")
    if decision.skip_reason != "duplicate_rows":
        return _result(False, f"unexpected skip_reason: {decision.skip_reason!r}")
    return _result(True, "duplicate rows are skipped")


def test_marker_runtime_dry_run_skips_missing_rows_outside_header() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.table_markers import TableMarkerDiagnostic

    diagnostic = TableMarkerDiagnostic(
        table_index=0,
        rows_count=5,
        pages_detected=[20, 21],
        row_pages={0: 20, 1: 20, 3: 21, 4: 21},
        found_rows=[0, 1, 3, 4],
        missing_rows=[2],
        duplicate_rows={},
        candidate_for_split=False,
        page_spans=[],
        appendix_table=False,
        caption_detected=True,
        has_standard_table_caption=True,
    )

    decision = tc._evaluate_marker_split_diagnostic(diagnostic, header_rows=1)
    if decision.eligible:
        return _result(False, "missing body rows should skip dry-run eligibility")
    if decision.skip_reason != "missing_rows_outside_header":
        return _result(False, f"unexpected skip_reason: {decision.skip_reason!r}")
    return _result(True, "missing rows outside header are skipped")


def test_marker_runtime_dry_run_skips_three_page_tables() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    from guides.coursework_kfu_2025.table_markers import TableMarkerDiagnostic

    diagnostic = TableMarkerDiagnostic(
        table_index=0,
        rows_count=6,
        pages_detected=[30, 31, 32],
        row_pages={0: 30, 1: 30, 2: 31, 3: 31, 4: 32, 5: 32},
        found_rows=[0, 1, 2, 3, 4, 5],
        missing_rows=[],
        duplicate_rows={},
        candidate_for_split=False,
        page_spans=[],
        appendix_table=False,
        caption_detected=True,
        has_standard_table_caption=True,
    )

    decision = tc._evaluate_marker_split_diagnostic(diagnostic, header_rows=1)
    if decision.eligible:
        return _result(False, "3-page table should not be eligible in v1")
    if decision.skip_reason != "not_2_pages":
        return _result(False, f"unexpected skip_reason: {decision.skip_reason!r}")
    return _result(True, "3-page tables are skipped")


def test_marker_runtime_dry_run_logs_eligible_candidate() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    import guides.coursework_kfu_2025.table_markers as tm

    log_stream = io.StringIO()
    handler = logging.StreamHandler(log_stream)
    handler.setFormatter(logging.Formatter("%(message)s"))
    old_level = tc.logger.level
    tc.logger.addHandler(handler)
    tc.logger.setLevel(logging.INFO)

    diagnostic = tm.TableMarkerDiagnostic(
        table_index=10,
        rows_count=21,
        pages_detected=[53, 54],
        row_pages={**{0: 53}, **{row: 53 for row in range(1, 18)}, 18: 54, 19: 54, 20: 54},
        found_rows=list(range(21)),
        missing_rows=[],
        duplicate_rows={},
        candidate_for_split=False,
        page_spans=[tm.TablePageSpan(0, 17, 53), tm.TablePageSpan(18, 20, 54)],
        appendix_table=True,
        caption_detected=True,
        has_standard_table_caption=False,
        preceding_paragraph_text="Расчет трудозатрат",
    )

    old_diagnose_all = tm.diagnose_all_tables
    try:
        tm.diagnose_all_tables = lambda _path, keep_temp=False: [diagnostic]
        count = tc._run_marker_split_detection_pass(Path("/tmp/fake.docx"))
    finally:
        tm.diagnose_all_tables = old_diagnose_all
        tc.logger.removeHandler(handler)
        tc.logger.setLevel(old_level)

    logs = log_stream.getvalue()
    expected_fragments = [
        "marker_split_candidate table_index=10 rows=21 pages=[53, 54]",
        "marker_split_boundary table_index=10 split_before_row=18",
        "marker_split_decision=ELIGIBLE table_index=10",
    ]
    missing = [fragment for fragment in expected_fragments if fragment not in logs]
    if count != 1:
        return _result(False, f"expected one eligible candidate, got {count}")
    if missing:
        return _result(False, f"missing log fragments: {missing!r}; logs={logs!r}")
    return _result(True, "eligible marker candidate logs are emitted")


def test_marker_runtime_dry_run_feature_flag_off_skips_detection_hook() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc

    doc = Document()
    doc.add_paragraph("Таблица 1.1")
    tbl = doc.add_table(rows=2, cols=1)
    tbl.rows[0].cells[0].text = "H"
    tbl.rows[1].cells[0].text = "A"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "flag_off.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "flag_off.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_flag = os.environ.pop("KPFU_ENABLE_MARKER_SPLIT", None)
        old_hook = tc._run_marker_split_detection_pass
        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            def fail_hook(_docx_path):
                raise AssertionError("marker dry-run hook should not be called when flag is off")

            tc._run_marker_split_detection_pass = fail_hook
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: []
            tc.apply_rendered_table_continuation(path)
        finally:
            tc._run_marker_split_detection_pass = old_hook
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze
            if old_flag is not None:
                os.environ["KPFU_ENABLE_MARKER_SPLIT"] = old_flag

    return _result(True, "feature flag off keeps marker dry-run hook disabled")


def test_marker_runtime_dry_run_only_does_not_mutate_document() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    import guides.coursework_kfu_2025.table_markers as tm

    doc = Document()
    doc.add_paragraph("Таблица 1.1.1")
    tbl = doc.add_table(rows=5, cols=3)
    tbl.rows[0].cells[0].text = "A"
    tbl.rows[0].cells[1].text = "B"
    tbl.rows[0].cells[2].text = "C"
    for i in range(1, 5):
        for j in range(3):
            tbl.rows[i].cells[j].text = f"r{i}c{j}"

    diagnostic = tm.TableMarkerDiagnostic(
        table_index=0,
        rows_count=5,
        pages_detected=[12, 13],
        row_pages={0: 12, 1: 12, 2: 12, 3: 13, 4: 13},
        found_rows=[0, 1, 2, 3, 4],
        missing_rows=[],
        duplicate_rows={},
        candidate_for_split=False,
        page_spans=[tm.TablePageSpan(0, 2, 12), tm.TablePageSpan(3, 4, 13)],
        appendix_table=False,
        caption_detected=True,
        has_standard_table_caption=True,
        preceding_paragraph_text="Таблица 1.1.1",
    )

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "dry_run_only.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "dry_run_only.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)
        before = path.read_bytes()

        old_enable = os.environ.get("KPFU_ENABLE_MARKER_SPLIT")
        old_apply = os.environ.pop("KPFU_APPLY_MARKER_SPLIT", None)
        os.environ["KPFU_ENABLE_MARKER_SPLIT"] = "1"

        old_diagnose_all = tm.diagnose_all_tables
        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tm.diagnose_all_tables = lambda _path, keep_temp=False: [diagnostic]
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: []
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tm.diagnose_all_tables = old_diagnose_all
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze
            if old_enable is None:
                os.environ.pop("KPFU_ENABLE_MARKER_SPLIT", None)
            else:
                os.environ["KPFU_ENABLE_MARKER_SPLIT"] = old_enable
            if old_apply is not None:
                os.environ["KPFU_APPLY_MARKER_SPLIT"] = old_apply

        after = path.read_bytes()

    if n != 0:
        return _result(False, f"dry-run only should not mutate, got {n}")
    if before != after:
        return _result(False, "dry-run only changed document bytes")
    return _result(True, "dry-run only does not mutate document")


def test_marker_runtime_apply_split_for_appendix_table() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    import guides.coursework_kfu_2025.table_markers as tm

    doc = Document()
    doc.add_paragraph("Приложение А")
    doc.add_paragraph("Трудозатраты проекта")
    tbl = doc.add_table(rows=5, cols=3)
    tbl.rows[0].cells[0].text = "Исполнитель"
    tbl.rows[0].cells[1].text = "Работы"
    tbl.rows[0].cells[2].text = "Стоимость"
    for i in range(1, 5):
        tbl.rows[i].cells[0].text = f"r{i}c0"
        tbl.rows[i].cells[1].text = f"r{i}c1"
        tbl.rows[i].cells[2].text = f"r{i}c2"
    doc.add_paragraph("Источник: данные автора")

    diagnostic = tm.TableMarkerDiagnostic(
        table_index=0,
        rows_count=5,
        pages_detected=[53, 54],
        row_pages={0: 53, 1: 53, 2: 53, 3: 54, 4: 54},
        found_rows=[0, 1, 2, 3, 4],
        missing_rows=[],
        duplicate_rows={},
        candidate_for_split=False,
        page_spans=[tm.TablePageSpan(0, 2, 53), tm.TablePageSpan(3, 4, 54)],
        appendix_table=True,
        caption_detected=True,
        has_standard_table_caption=False,
        preceding_paragraph_text="Трудозатраты проекта",
    )

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "appendix_apply.docx"
        doc.save(path)

        old_enable = os.environ.get("KPFU_ENABLE_MARKER_SPLIT")
        old_apply = os.environ.get("KPFU_APPLY_MARKER_SPLIT")
        os.environ["KPFU_ENABLE_MARKER_SPLIT"] = "1"
        os.environ["KPFU_APPLY_MARKER_SPLIT"] = "1"

        old_diagnose_all = tm.diagnose_all_tables
        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tm.diagnose_all_tables = lambda _path, keep_temp=False: [diagnostic]
            tc.render_docx_to_pdf = lambda _path: (_ for _ in ()).throw(AssertionError("render path should not run after marker split apply"))
            tc.analyze_pdf_lines = lambda _path: (_ for _ in ()).throw(AssertionError("pdf analysis should not run after marker split apply"))
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tm.diagnose_all_tables = old_diagnose_all
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze
            if old_enable is None:
                os.environ.pop("KPFU_ENABLE_MARKER_SPLIT", None)
            else:
                os.environ["KPFU_ENABLE_MARKER_SPLIT"] = old_enable
            if old_apply is None:
                os.environ.pop("KPFU_APPLY_MARKER_SPLIT", None)
            else:
                os.environ["KPFU_APPLY_MARKER_SPLIT"] = old_apply

        out = Document(str(path))

    if n != 1:
        return _result(False, f"expected one appendix split mutation, got {n}")
    if len(out.tables) != 2:
        return _result(False, f"expected 2 tables after appendix split, got {len(out.tables)}")
    if [c.text for c in out.tables[0].rows[1].cells] != ["1", "2", "3"]:
        return _result(False, "numbered row missing in first appendix table")
    if [c.text for c in out.tables[1].rows[0].cells] != ["1", "2", "3"]:
        return _result(False, "numbered row missing in second appendix table")
    if any("Продолжение таблицы" in (p.text or "") for p in out.paragraphs):
        return _result(False, "appendix split inserted forbidden continuation paragraph")
    return _result(True, "eligible appendix table is split with numbered row and no continuation paragraph")


def test_marker_runtime_apply_split_for_ordinary_table() -> tuple[bool, str]:
    """
    Product rule: ordinary table continuation inserts a continuation marker
    and a numbered row without duplicating the textual table header.
    """
    import guides.coursework_kfu_2025.table_continuation as tc
    import guides.coursework_kfu_2025.table_markers as tm

    doc = Document()
    doc.add_paragraph("Таблица 1.1.1")
    tbl = doc.add_table(rows=5, cols=3)
    tbl.rows[0].cells[0].text = "Показатель"
    tbl.rows[0].cells[1].text = "Как влияет"
    tbl.rows[0].cells[2].text = "Последствия"
    for i in range(1, 5):
        tbl.rows[i].cells[0].text = f"r{i}c0"
        tbl.rows[i].cells[1].text = f"r{i}c1"
        tbl.rows[i].cells[2].text = f"r{i}c2"
    doc.add_paragraph("Источник: данные автора")

    diagnostic = tm.TableMarkerDiagnostic(
        table_index=0,
        rows_count=5,
        pages_detected=[12, 13],
        row_pages={0: 12, 1: 12, 2: 12, 3: 13, 4: 13},
        found_rows=[0, 1, 2, 3, 4],
        missing_rows=[],
        duplicate_rows={},
        candidate_for_split=False,
        page_spans=[tm.TablePageSpan(0, 2, 12), tm.TablePageSpan(3, 4, 13)],
        appendix_table=False,
        caption_detected=True,
        has_standard_table_caption=True,
        preceding_paragraph_text="Таблица 1.1.1",
    )

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "ordinary_apply.docx"
        doc.save(path)

        old_enable = os.environ.get("KPFU_ENABLE_MARKER_SPLIT")
        old_apply = os.environ.get("KPFU_APPLY_MARKER_SPLIT")
        os.environ["KPFU_ENABLE_MARKER_SPLIT"] = "1"
        os.environ["KPFU_APPLY_MARKER_SPLIT"] = "1"

        old_diagnose_all = tm.diagnose_all_tables
        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            tm.diagnose_all_tables = lambda _path, keep_temp=False: [diagnostic]
            tc.render_docx_to_pdf = lambda _path: (_ for _ in ()).throw(AssertionError("render path should not run after marker split apply"))
            tc.analyze_pdf_lines = lambda _path: (_ for _ in ()).throw(AssertionError("pdf analysis should not run after marker split apply"))
            n = tc.apply_rendered_table_continuation(path)
        finally:
            tm.diagnose_all_tables = old_diagnose_all
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze
            if old_enable is None:
                os.environ.pop("KPFU_ENABLE_MARKER_SPLIT", None)
            else:
                os.environ["KPFU_ENABLE_MARKER_SPLIT"] = old_enable
            if old_apply is None:
                os.environ.pop("KPFU_APPLY_MARKER_SPLIT", None)
            else:
                os.environ["KPFU_APPLY_MARKER_SPLIT"] = old_apply

        out = Document(str(path))

    if n != 1:
        return _result(False, f"expected one ordinary split mutation, got {n}")
    if len(out.tables) != 2:
        return _result(False, f"expected 2 tables after ordinary split, got {len(out.tables)}")
    if [c.text for c in out.tables[1].rows[0].cells] != ["1", "2", "3"]:
        return _result(False, "continuation table should start with numbered row only")
    continuation_paras = [p for p in out.paragraphs if p.text == "Продолжение таблицы 1.1.1"]
    if len(continuation_paras) != 1:
        return _result(False, "ordinary split did not insert continuation paragraph")
    pPr = continuation_paras[0]._element.find(qn("w:pPr"))
    page_break = pPr.find(qn("w:pageBreakBefore")) if pPr is not None else None
    jc = pPr.find(qn("w:jc")) if pPr is not None else None
    keep = pPr.find(qn("w:keepNext")) if pPr is not None else None
    if page_break is None:
        return _result(False, "ordinary continuation marker should start on a new page")
    if jc is None or jc.get(qn("w:val")) != "right":
        return _result(False, "ordinary continuation marker should be right-aligned")
    if keep is None:
        return _result(False, "ordinary continuation marker should keep with following table")
    if any(cell.text == "Показатель" for cell in out.tables[1].rows[0].cells):
        return _result(False, "text header leaked into continuation row")
    return _result(True, "eligible ordinary table is split with continuation paragraph")


def test_marker_runtime_apply_skips_ineligible_tables() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    import guides.coursework_kfu_2025.table_markers as tm

    cases = [
        ("duplicate", tm.TableMarkerDiagnostic(
            table_index=0,
            rows_count=4,
            pages_detected=[12, 13],
            row_pages={0: 12, 1: 12, 3: 13},
            found_rows=[0, 1, 3],
            missing_rows=[],
            duplicate_rows={2: [12, 13]},
            candidate_for_split=False,
            page_spans=[],
            appendix_table=False,
            caption_detected=True,
            has_standard_table_caption=True,
            preceding_paragraph_text="Таблица 2.1",
        )),
        ("missing", tm.TableMarkerDiagnostic(
            table_index=0,
            rows_count=5,
            pages_detected=[12, 13],
            row_pages={0: 12, 1: 12, 3: 13, 4: 13},
            found_rows=[0, 1, 3, 4],
            missing_rows=[2],
            duplicate_rows={},
            candidate_for_split=False,
            page_spans=[],
            appendix_table=False,
            caption_detected=True,
            has_standard_table_caption=True,
            preceding_paragraph_text="Таблица 2.2",
        )),
        ("threepage", tm.TableMarkerDiagnostic(
            table_index=0,
            rows_count=6,
            pages_detected=[12, 13, 14],
            row_pages={0: 12, 1: 12, 2: 13, 3: 13, 4: 14, 5: 14},
            found_rows=[0, 1, 2, 3, 4, 5],
            missing_rows=[],
            duplicate_rows={},
            candidate_for_split=False,
            page_spans=[],
            appendix_table=False,
            caption_detected=True,
            has_standard_table_caption=True,
            preceding_paragraph_text="Таблица 2.3",
        )),
    ]

    for label, diagnostic in cases:
        doc = Document()
        doc.add_paragraph("Таблица 2.1")
        tbl = doc.add_table(rows=4, cols=2)
        for i in range(4):
            tbl.rows[i].cells[0].text = f"r{i}c0"
            tbl.rows[i].cells[1].text = f"r{i}c1"

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / f"ineligible_{label}.docx"
            pdf_dir = Path(tmp) / "pdf"
            pdf_dir.mkdir()
            pdf_path = pdf_dir / f"ineligible_{label}.pdf"
            pdf_path.write_bytes(b"%PDF-1.4\n")
            doc.save(path)
            before = path.read_bytes()

            old_enable = os.environ.get("KPFU_ENABLE_MARKER_SPLIT")
            old_apply = os.environ.get("KPFU_APPLY_MARKER_SPLIT")
            os.environ["KPFU_ENABLE_MARKER_SPLIT"] = "1"
            os.environ["KPFU_APPLY_MARKER_SPLIT"] = "1"

            old_diagnose_all = tm.diagnose_all_tables
            old_render = tc.render_docx_to_pdf
            old_analyze = tc.analyze_pdf_lines
            try:
                tm.diagnose_all_tables = lambda _path, keep_temp=False, diag=diagnostic: [diag]
                tc.render_docx_to_pdf = lambda _path: pdf_path
                tc.analyze_pdf_lines = lambda _path: []
                n = tc.apply_rendered_table_continuation(path)
            finally:
                tm.diagnose_all_tables = old_diagnose_all
                tc.render_docx_to_pdf = old_render
                tc.analyze_pdf_lines = old_analyze
                if old_enable is None:
                    os.environ.pop("KPFU_ENABLE_MARKER_SPLIT", None)
                else:
                    os.environ["KPFU_ENABLE_MARKER_SPLIT"] = old_enable
                if old_apply is None:
                    os.environ.pop("KPFU_APPLY_MARKER_SPLIT", None)
                else:
                    os.environ["KPFU_APPLY_MARKER_SPLIT"] = old_apply

            after = path.read_bytes()

        if n != 0:
            return _result(False, f"ineligible case {label} should not mutate, got {n}")
        if before != after:
            return _result(False, f"ineligible case {label} changed document bytes")

    return _result(True, "ineligible duplicate/missing/3page cases do not mutate")


def test_marker_runtime_apply_is_idempotent_on_second_run() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    import guides.coursework_kfu_2025.table_markers as tm

    doc = Document()
    doc.add_paragraph("Таблица 7.1")
    tbl = doc.add_table(rows=5, cols=3)
    tbl.rows[0].cells[0].text = "A"
    tbl.rows[0].cells[1].text = "B"
    tbl.rows[0].cells[2].text = "C"
    for i in range(1, 5):
        tbl.rows[i].cells[0].text = f"r{i}c0"
        tbl.rows[i].cells[1].text = f"r{i}c1"
        tbl.rows[i].cells[2].text = f"r{i}c2"

    eligible = tm.TableMarkerDiagnostic(
        table_index=0,
        rows_count=5,
        pages_detected=[12, 13],
        row_pages={0: 12, 1: 12, 2: 12, 3: 13, 4: 13},
        found_rows=[0, 1, 2, 3, 4],
        missing_rows=[],
        duplicate_rows={},
        candidate_for_split=False,
        page_spans=[tm.TablePageSpan(0, 2, 12), tm.TablePageSpan(3, 4, 13)],
        appendix_table=False,
        caption_detected=True,
        has_standard_table_caption=True,
        preceding_paragraph_text="Таблица 7.1",
    )

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "idempotent.docx"
        pdf_dir = Path(tmp) / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "idempotent.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n")
        doc.save(path)

        old_enable = os.environ.get("KPFU_ENABLE_MARKER_SPLIT")
        old_apply = os.environ.get("KPFU_APPLY_MARKER_SPLIT")
        os.environ["KPFU_ENABLE_MARKER_SPLIT"] = "1"
        os.environ["KPFU_APPLY_MARKER_SPLIT"] = "1"

        old_diagnose_all = tm.diagnose_all_tables
        old_render = tc.render_docx_to_pdf
        old_analyze = tc.analyze_pdf_lines
        try:
            def fake_diagnose_all(docx_path, keep_temp=False):
                current = Document(str(docx_path))
                if len(current.tables) == 1:
                    return [eligible]
                return [
                    tm.TableMarkerDiagnostic(
                        table_index=0,
                        rows_count=len(current.tables[0].rows),
                        pages_detected=[12],
                        row_pages={0: 12, 1: 12, 2: 12, 3: 12},
                        found_rows=[0, 1, 2, 3],
                        missing_rows=[],
                        duplicate_rows={},
                        candidate_for_split=False,
                        page_spans=[tm.TablePageSpan(0, 3, 12)],
                        appendix_table=False,
                        caption_detected=True,
                        has_standard_table_caption=True,
                        preceding_paragraph_text="Таблица 7.1",
                    ),
                    tm.TableMarkerDiagnostic(
                        table_index=1,
                        rows_count=len(current.tables[1].rows),
                        pages_detected=[13],
                        row_pages={0: 13, 1: 13, 2: 13},
                        found_rows=[0, 1, 2],
                        missing_rows=[],
                        duplicate_rows={},
                        candidate_for_split=False,
                        page_spans=[tm.TablePageSpan(0, 2, 13)],
                        appendix_table=False,
                        caption_detected=True,
                        has_standard_table_caption=False,
                        preceding_paragraph_text="Продолжение таблицы 7.1",
                    ),
                ]

            tm.diagnose_all_tables = fake_diagnose_all
            tc.render_docx_to_pdf = lambda _path: pdf_path
            tc.analyze_pdf_lines = lambda _path: []
            first = tc.apply_rendered_table_continuation(path)
            second = tc.apply_rendered_table_continuation(path)
        finally:
            tm.diagnose_all_tables = old_diagnose_all
            tc.render_docx_to_pdf = old_render
            tc.analyze_pdf_lines = old_analyze
            if old_enable is None:
                os.environ.pop("KPFU_ENABLE_MARKER_SPLIT", None)
            else:
                os.environ["KPFU_ENABLE_MARKER_SPLIT"] = old_enable
            if old_apply is None:
                os.environ.pop("KPFU_APPLY_MARKER_SPLIT", None)
            else:
                os.environ["KPFU_APPLY_MARKER_SPLIT"] = old_apply

        out = Document(str(path))

    if first != 1 or second != 0:
        return _result(False, f"expected first run=1 and second run=0, got {first}/{second}")
    if len(out.tables) != 2:
        return _result(False, f"expected 2 tables after second run, got {len(out.tables)}")
    numbered_rows = sum(
        1 for row in out.tables[0].rows
        if [cell.text for cell in row.cells] == ["1", "2", "3"]
    )
    if numbered_rows != 1:
        return _result(False, f"first table should contain exactly one numbered row, got {numbered_rows}")
    continuation_count = sum(1 for p in out.paragraphs if p.text == "Продолжение таблицы 7.1")
    if continuation_count != 1:
        return _result(False, f"expected one continuation paragraph after two runs, got {continuation_count}")
    return _result(True, "active marker split is idempotent on second run")


def test_split_prototype_simple_table() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    doc = Document()
    tbl = doc.add_table(rows=5, cols=2)
    for i in range(5):
        tbl.rows[i].cells[0].text = f"r{i}c0"
        tbl.rows[i].cells[1].text = f"r{i}c1"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "simple.docx"
        doc.save(src)
        result = prototype_split_table_copy(src, 0, 3, header_rows=1, keep_temp=True)
        out = Document(str(result.output_docx_path))

    if result.total_tables_after != 2:
        return _result(False, f"expected 2 tables after split, got {result.total_tables_after}")
    if result.first_table_rows_count != 3:
        return _result(False, f"expected 3 rows in first table, got {result.first_table_rows_count}")
    if result.second_table_rows_count != 3:
        return _result(False, f"expected 3 rows in second table, got {result.second_table_rows_count}")
    if out.tables[1].rows[0].cells[0].text != "r0c0":
        return _result(False, f"header row not copied into second table: {out.tables[1].rows[0].cells[0].text!r}")
    if out.tables[1].rows[1].cells[0].text != "r3c0":
        return _result(False, f"tail rows not moved to second table: {out.tables[1].rows[1].cells[0].text!r}")
    return _result(True, "simple table split produced two clone-based tables")


def test_split_prototype_source_note_stays_after_second_table() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    doc = Document()
    tbl = doc.add_table(rows=5, cols=1)
    for i in range(5):
        tbl.rows[i].cells[0].text = f"row{i}"
    doc.add_paragraph("Источник: данные автора")

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "source_note.docx"
        doc.save(src)
        result = prototype_split_table_copy(src, 0, 3, header_rows=1, keep_temp=True)
        out = Document(str(result.output_docx_path))

    if result.source_note_after_second is not True:
        return _result(False, f"source note did not stay after second table: {result.source_note_after_second!r}")

    body = list(out.element.body)
    def _local(node):
        return node.tag.split("}")[-1] if "}" in node.tag else node.tag

    tags = [_local(node) for node in body]
    try:
        first_tbl_idx = tags.index("tbl")
        second_tbl_idx = tags.index("tbl", first_tbl_idx + 1)
        note_idx = next(
            i for i, node in enumerate(body)
            if _local(node) == "p" and "Источник:" in "".join(t.text or "" for t in node.findall('.//' + qn('w:t')))
        )
    except Exception as exc:
        return _result(False, f"failed to inspect body ordering: {exc}")
    if not (first_tbl_idx < second_tbl_idx < note_idx):
        return _result(False, f"source note ordering invalid: first={first_tbl_idx}, second={second_tbl_idx}, note={note_idx}")
    return _result(True, "source note remains after second table")


def test_split_prototype_original_document_unchanged() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    doc = Document()
    tbl = doc.add_table(rows=4, cols=1)
    for i in range(4):
        tbl.rows[i].cells[0].text = f"row{i}"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "original.docx"
        doc.save(src)
        before = src.read_bytes()
        prototype_split_table_copy(src, 0, 2, header_rows=1, keep_temp=True)
        after = src.read_bytes()

    if before != after:
        return _result(False, "source docx changed after prototype split")
    return _result(True, "prototype split leaves source document unchanged")


def test_split_prototype_invalid_table_index() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    doc = Document()
    doc.add_table(rows=2, cols=1)
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "invalid_idx.docx"
        doc.save(src)
        try:
            prototype_split_table_copy(src, 3, 1, header_rows=1, keep_temp=False)
        except ValueError:
            return _result(True, "invalid table index rejected")
        except Exception as exc:
            return _result(False, f"unexpected exception type: {exc}")
    return _result(False, "expected ValueError for invalid table index")


def test_split_prototype_invalid_split_before_row() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    doc = Document()
    doc.add_table(rows=3, cols=1)
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "invalid_split.docx"
        doc.save(src)
        try:
            prototype_split_table_copy(src, 0, 0, header_rows=1, keep_temp=False)
        except ValueError:
            return _result(True, "invalid split_before_row rejected")
        except Exception as exc:
            return _result(False, f"unexpected exception type: {exc}")
    return _result(False, "expected ValueError for invalid split_before_row")


def test_split_prototype_no_continuation_paragraph_inserted() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    doc = Document()
    doc.add_paragraph("Приложение А")
    doc.add_paragraph("Длинная таблица приложения")
    tbl = doc.add_table(rows=4, cols=1)
    for i in range(4):
        tbl.rows[i].cells[0].text = f"row{i}"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "appendix_split.docx"
        doc.save(src)
        result = prototype_split_table_copy(src, 0, 2, header_rows=1, keep_temp=True)
        out = Document(str(result.output_docx_path))

    paragraph_texts = [p.text for p in out.paragraphs]
    if any("Продолжение таблицы" in (text or "") for text in paragraph_texts):
        return _result(False, "unexpected continuation paragraph inserted")
    if any(text == "Продолжение" for text in paragraph_texts):
        return _result(False, "unexpected generic continuation paragraph inserted")
    return _result(True, "appendix split does not insert continuation paragraph")


def test_split_prototype_numbered_ordinary_continuation_row_only() -> tuple[bool, str]:
    """
    Product rule: ordinary table split continuation uses "Продолжение таблицы"
    plus the numbered row only; the original title/header row is not duplicated.
    """
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    doc = Document()
    doc.add_paragraph("Таблица 1.1.1")
    tbl = doc.add_table(rows=5, cols=3)
    tbl.rows[0].cells[0].text = "Показатель"
    tbl.rows[0].cells[1].text = "Как влияет"
    tbl.rows[0].cells[2].text = "Последствия"
    for i in range(1, 5):
        tbl.rows[i].cells[0].text = f"r{i}c0"
        tbl.rows[i].cells[1].text = f"r{i}c1"
        tbl.rows[i].cells[2].text = f"r{i}c2"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "ordinary_numbered.docx"
        doc.save(src)
        result = prototype_split_table_copy(
            src,
            0,
            3,
            header_rows=1,
            numbered_header=True,
            appendix_table=False,
            keep_temp=True,
        )
        out = Document(str(result.output_docx_path))

    if result.continuation_text != "Продолжение таблицы 1.1.1":
        return _result(False, f"unexpected continuation text: {result.continuation_text!r}")
    if result.continuation_paragraph_inserted is not True:
        return _result(False, "ordinary table should insert continuation paragraph")
    if result.column_count != 3:
        return _result(False, f"unexpected column_count: {result.column_count!r}")

    first_row_texts = [cell.text for cell in out.tables[0].rows[1].cells]
    second_row_texts = [cell.text for cell in out.tables[1].rows[0].cells]
    if first_row_texts != ["1", "2", "3"]:
        return _result(False, f"unexpected numbered row in first table: {first_row_texts!r}")
    if second_row_texts != ["1", "2", "3"]:
        return _result(False, f"unexpected continuation numbered row: {second_row_texts!r}")
    if [cell.text for cell in out.tables[1].rows[1].cells] != ["r3c0", "r3c1", "r3c2"]:
        return _result(False, "tail rows not moved under numbered continuation row")
    if any(cell.text == "Показатель" for cell in out.tables[1].rows[0].cells):
        return _result(False, "text header leaked into continuation numbered row")
    if "Продолжение таблицы 1.1.1" not in [p.text for p in out.paragraphs]:
        return _result(False, "continuation paragraph missing from ordinary numbered split")
    return _result(True, "ordinary numbered split uses continuation text and numbered row only")


def test_split_prototype_numbered_ordinary_split_caption_before_title() -> tuple[bool, str]:
    """
    Product rule: ordinary captions may be split as "Таблица X.Y.Z" plus a
    separate title paragraph before the table. Continuation uses the table
    number and does not duplicate the title/header text.
    """
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    title = "Комитеты НС Сбера: состав мандата и фокус надзора"

    doc = Document()
    doc.add_paragraph("Таблица 2.1.2")
    doc.add_paragraph(title)
    tbl = doc.add_table(rows=5, cols=3)
    tbl.rows[0].cells[0].text = "Показатель"
    tbl.rows[0].cells[1].text = "Как влияет"
    tbl.rows[0].cells[2].text = "Последствия"
    for i in range(1, 5):
        tbl.rows[i].cells[0].text = f"r{i}c0"
        tbl.rows[i].cells[1].text = f"r{i}c1"
        tbl.rows[i].cells[2].text = f"r{i}c2"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "ordinary_split_caption_before_title.docx"
        doc.save(src)
        result = prototype_split_table_copy(
            src,
            0,
            3,
            header_rows=1,
            numbered_header=True,
            appendix_table=False,
            keep_temp=True,
        )
        out = Document(str(result.output_docx_path))

    if result.continuation_text != "Продолжение таблицы 2.1.2":
        return _result(False, f"unexpected continuation text: {result.continuation_text!r}")
    if result.continuation_paragraph_inserted is not True:
        return _result(False, "ordinary table should insert continuation paragraph")
    if [cell.text for cell in out.tables[1].rows[0].cells] != ["1", "2", "3"]:
        return _result(False, "continuation table should start with numbered row")
    if any(cell.text == "Показатель" for cell in out.tables[1].rows[0].cells):
        return _result(False, "text header leaked into continuation numbered row")
    paragraph_texts = [p.text for p in out.paragraphs]
    if paragraph_texts.count(title) != 1:
        return _result(False, f"table title duplicated or removed: {paragraph_texts!r}")
    if paragraph_texts.count("Продолжение таблицы 2.1.2") != 1:
        return _result(False, f"continuation paragraph missing or duplicated: {paragraph_texts!r}")
    return _result(True, "ordinary split caption before title uses continuation number")


def test_split_prototype_numbered_appendix_has_no_continuation_text() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    doc = Document()
    doc.add_paragraph("Приложение А")
    doc.add_paragraph("Таблица приложения")
    tbl = doc.add_table(rows=4, cols=3)
    tbl.rows[0].cells[0].text = "Колонка А"
    tbl.rows[0].cells[1].text = "Колонка Б"
    tbl.rows[0].cells[2].text = "Колонка В"
    for i in range(1, 4):
        tbl.rows[i].cells[0].text = f"r{i}c0"
        tbl.rows[i].cells[1].text = f"r{i}c1"
        tbl.rows[i].cells[2].text = f"r{i}c2"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "appendix_numbered.docx"
        doc.save(src)
        result = prototype_split_table_copy(
            src,
            0,
            2,
            header_rows=1,
            numbered_header=True,
            appendix_table=True,
            keep_temp=True,
        )
        out = Document(str(result.output_docx_path))

    if result.continuation_paragraph_inserted:
        return _result(False, "appendix numbered split must not insert continuation paragraph")
    if result.continuation_text is not None:
        return _result(False, f"appendix continuation text must be None, got {result.continuation_text!r}")
    second_row_texts = [cell.text for cell in out.tables[1].rows[0].cells]
    if second_row_texts != ["1", "2", "3"]:
        return _result(False, f"unexpected appendix continuation row: {second_row_texts!r}")
    if any("Продолжение таблицы" in (text or "") for text in [p.text for p in out.paragraphs]):
        return _result(False, "appendix numbered split inserted continuation paragraph")
    return _result(True, "appendix numbered split keeps numbered row without continuation text")


def test_split_prototype_numbered_existing_row_reused_without_duplicate() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    doc = Document()
    doc.add_paragraph("Таблица 2.4")
    tbl = doc.add_table(rows=5, cols=3)
    tbl.rows[0].cells[0].text = "A"
    tbl.rows[0].cells[1].text = "B"
    tbl.rows[0].cells[2].text = "C"
    tbl.rows[1].cells[0].text = "1"
    tbl.rows[1].cells[1].text = "2"
    tbl.rows[1].cells[2].text = "3"
    for i in range(2, 5):
        tbl.rows[i].cells[0].text = f"r{i}c0"
        tbl.rows[i].cells[1].text = f"r{i}c1"
        tbl.rows[i].cells[2].text = f"r{i}c2"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "reuse_numbered.docx"
        doc.save(src)
        result = prototype_split_table_copy(
            src,
            0,
            4,
            header_rows=1,
            numbered_header=True,
            appendix_table=False,
            keep_temp=True,
        )
        out = Document(str(result.output_docx_path))

    if result.numbered_row_reused is not True:
        return _result(False, f"expected numbered row reuse, got {result.numbered_row_reused!r}")
    first_table_numbered_rows = sum(
        1 for row in out.tables[0].rows
        if [cell.text for cell in row.cells] == ["1", "2", "3"]
    )
    if first_table_numbered_rows != 1:
        return _result(False, f"expected exactly one numbered row in first table, got {first_table_numbered_rows}")
    second_row_texts = [cell.text for cell in out.tables[1].rows[0].cells]
    if second_row_texts != ["1", "2", "3"]:
        return _result(False, "reused numbered row not copied to continuation table")
    return _result(True, "existing numbered row is reused without duplication")


def test_split_prototype_numbered_malformed_existing_row_fails_safely() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    doc = Document()
    doc.add_paragraph("Таблица 3.1")
    tbl = doc.add_table(rows=4, cols=3)
    tbl.rows[0].cells[0].text = "A"
    tbl.rows[0].cells[1].text = "B"
    tbl.rows[0].cells[2].text = "C"
    tbl.rows[1].cells[0].text = "1"
    tbl.rows[1].cells[1].text = "3"
    tbl.rows[1].cells[2].text = "4"
    tbl.rows[2].cells[0].text = "r2c0"
    tbl.rows[2].cells[1].text = "r2c1"
    tbl.rows[2].cells[2].text = "r2c2"
    tbl.rows[3].cells[0].text = "r3c0"
    tbl.rows[3].cells[1].text = "r3c1"
    tbl.rows[3].cells[2].text = "r3c2"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "malformed_numbered.docx"
        doc.save(src)
        try:
            prototype_split_table_copy(
                src,
                0,
                3,
                header_rows=1,
                numbered_header=True,
                appendix_table=False,
                keep_temp=False,
            )
        except ValueError as exc:
            if "malformed" not in str(exc):
                return _result(False, f"unexpected ValueError text: {exc}")
            return _result(True, "malformed numbered row fails safely")
        except Exception as exc:
            return _result(False, f"unexpected exception type: {exc}")
    return _result(False, "expected ValueError for malformed numbered row")


def test_split_prototype_numbered_source_note_after_second_table() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    doc = Document()
    doc.add_paragraph("Таблица 4.2")
    tbl = doc.add_table(rows=5, cols=3)
    tbl.rows[0].cells[0].text = "A"
    tbl.rows[0].cells[1].text = "B"
    tbl.rows[0].cells[2].text = "C"
    for i in range(1, 5):
        tbl.rows[i].cells[0].text = f"r{i}c0"
        tbl.rows[i].cells[1].text = f"r{i}c1"
        tbl.rows[i].cells[2].text = f"r{i}c2"
    doc.add_paragraph("Источник: данные автора")

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "numbered_source_note.docx"
        doc.save(src)
        result = prototype_split_table_copy(
            src,
            0,
            3,
            header_rows=1,
            numbered_header=True,
            appendix_table=False,
            keep_temp=True,
        )
        out = Document(str(result.output_docx_path))

    if result.source_note_after_second is not True:
        return _result(False, "source note did not remain after continuation table")
    paragraph_texts = [p.text for p in out.paragraphs]
    if "Продолжение таблицы 4.2" not in paragraph_texts:
        return _result(False, "continuation paragraph missing in numbered source-note case")
    return _result(True, "numbered split keeps source note after second table")


def test_split_prototype_numbered_original_document_unchanged() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    doc = Document()
    doc.add_paragraph("Таблица 5.1")
    tbl = doc.add_table(rows=4, cols=3)
    tbl.rows[0].cells[0].text = "A"
    tbl.rows[0].cells[1].text = "B"
    tbl.rows[0].cells[2].text = "C"
    for i in range(1, 4):
        tbl.rows[i].cells[0].text = f"r{i}c0"
        tbl.rows[i].cells[1].text = f"r{i}c1"
        tbl.rows[i].cells[2].text = f"r{i}c2"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "numbered_original.docx"
        doc.save(src)
        before = src.read_bytes()
        prototype_split_table_copy(
            src,
            0,
            2,
            header_rows=1,
            numbered_header=True,
            appendix_table=False,
            keep_temp=True,
        )
        after = src.read_bytes()

    if before != after:
        return _result(False, "source docx changed after numbered prototype split")
    return _result(True, "numbered prototype split leaves source document unchanged")


def test_split_prototype_numbered_row_has_no_numpr_and_no_calibri() -> tuple[bool, str]:
    from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy

    doc = Document()
    doc.add_paragraph("Таблица 6.1")
    tbl = doc.add_table(rows=4, cols=3)
    tbl.rows[0].cells[0].text = "A"
    tbl.rows[0].cells[1].text = "B"
    tbl.rows[0].cells[2].text = "C"
    for i in range(1, 4):
        for j in range(3):
            tbl.rows[i].cells[j].text = f"r{i}c{j}"

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "numbered_markup.docx"
        doc.save(src)
        result = prototype_split_table_copy(
            src,
            0,
            2,
            header_rows=1,
            numbered_header=True,
            appendix_table=False,
            keep_temp=True,
        )
        out = Document(str(result.output_docx_path))

    numbered_row = out.tables[1].rows[0]
    for cell in numbered_row.cells:
        for paragraph in cell.paragraphs:
            p_pr = paragraph._element.find(qn("w:pPr"))
            if p_pr is not None and p_pr.find(qn("w:numPr")) is not None:
                return _result(False, "generated numbered row has w:numPr")
            for run in paragraph.runs:
                r_pr = run._element.find(qn("w:rPr"))
                fonts = r_pr.find(qn("w:rFonts")) if r_pr is not None else None
                ascii_font = fonts.get(qn("w:ascii")) if fonts is not None else None
                if ascii_font != "Times New Roman":
                    return _result(False, f"generated numbered row font is {ascii_font!r}, expected Times New Roman")
    return _result(True, "generated numbered row has no numPr and no Calibri fallback")


def test_marker_runtime_flags_do_not_change_headings() -> tuple[bool, str]:
    import guides.coursework_kfu_2025.table_continuation as tc
    import guides.coursework_kfu_2025.table_markers as tm

    def heading_snapshot(doc: Document):
        out = []
        for p in doc.paragraphs:
            text = (p.text or "").strip()
            if text in {"ВВЕДЕНИЕ", "1. ГЛАВА", "1.1. Подраздел"}:
                p_pr = p._element.find(qn("w:pPr"))
                num_pr = p_pr.find(qn("w:numPr")) if p_pr is not None else None
                fonts = []
                for run in p.runs[:2]:
                    r_pr = run._element.find(qn("w:rPr"))
                    r_fonts = r_pr.find(qn("w:rFonts")) if r_pr is not None else None
                    fonts.append(r_fonts.get(qn("w:ascii")) if r_fonts is not None else None)
                out.append((text, p.style.name if p.style else None, num_pr is not None, tuple(fonts)))
        return out

    base = Document()
    p = base.add_paragraph("ВВЕДЕНИЕ")
    p.style = "Heading 1"
    p.runs[0].font.name = "Times New Roman"
    p = base.add_paragraph("1. ГЛАВА")
    p.style = "Heading 1"
    p.runs[0].font.name = "Times New Roman"
    p = base.add_paragraph("1.1. Подраздел")
    p.style = "Heading 2"
    p.runs[0].font.name = "Times New Roman"
    base.add_paragraph("Таблица 1.1.1")
    tbl = base.add_table(rows=5, cols=3)
    tbl.rows[0].cells[0].text = "A"
    tbl.rows[0].cells[1].text = "B"
    tbl.rows[0].cells[2].text = "C"
    for i in range(1, 5):
        for j in range(3):
            tbl.rows[i].cells[j].text = f"r{i}c{j}"

    diagnostic = tm.TableMarkerDiagnostic(
        table_index=0,
        rows_count=5,
        pages_detected=[12, 13],
        row_pages={0: 12, 1: 12, 2: 12, 3: 13, 4: 13},
        found_rows=[0, 1, 2, 3, 4],
        missing_rows=[],
        duplicate_rows={},
        candidate_for_split=False,
        page_spans=[tm.TablePageSpan(0, 2, 12), tm.TablePageSpan(3, 4, 13)],
        appendix_table=False,
        caption_detected=True,
        has_standard_table_caption=True,
        preceding_paragraph_text="Таблица 1.1.1",
    )

    expected = heading_snapshot(base)

    for mode_name, env in [
        ("flags_off", {}),
        ("dry_run", {"KPFU_ENABLE_MARKER_SPLIT": "1"}),
        ("apply", {"KPFU_ENABLE_MARKER_SPLIT": "1", "KPFU_APPLY_MARKER_SPLIT": "1"}),
    ]:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / f"{mode_name}.docx"
            base.save(path)
            old_env = {k: os.environ.get(k) for k in ["KPFU_ENABLE_MARKER_SPLIT", "KPFU_APPLY_MARKER_SPLIT"]}
            old_diagnose_all = tm.diagnose_all_tables
            old_render = tc.render_docx_to_pdf
            old_analyze = tc.analyze_pdf_lines
            try:
                for k in old_env:
                    os.environ.pop(k, None)
                os.environ.update(env)
                tm.diagnose_all_tables = lambda _path, keep_temp=False: [diagnostic]
                tc.render_docx_to_pdf = lambda _path: (_ for _ in ()).throw(AssertionError("render path should not run in heading regression test"))
                tc.analyze_pdf_lines = lambda _path: []
                tc.apply_rendered_table_continuation(path)
            finally:
                tm.diagnose_all_tables = old_diagnose_all
                tc.render_docx_to_pdf = old_render
                tc.analyze_pdf_lines = old_analyze
                for k, v in old_env.items():
                    if v is None:
                        os.environ.pop(k, None)
                    else:
                        os.environ[k] = v

            out = Document(str(path))
            if heading_snapshot(out) != expected:
                return _result(False, f"heading snapshot changed in mode {mode_name}: {heading_snapshot(out)!r}")

    return _result(True, "flags off, dry-run, and apply do not change headings outside target table")


def test_marker_runtime_real_rybakov_target_applies_split() -> tuple[bool, str]:
    asset = next(ASSETS.glob("*Рыбаков*.docx"), None)
    if asset is None:
        return _result(True, "Рыбаков asset missing, skipped")

    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "rybakov_apply.docx"
        old_enable = os.environ.get("KPFU_ENABLE_MARKER_SPLIT")
        old_apply = os.environ.get("KPFU_APPLY_MARKER_SPLIT")
        try:
            os.environ["KPFU_ENABLE_MARKER_SPLIT"] = "1"
            os.environ["KPFU_APPLY_MARKER_SPLIT"] = "1"
            format_docx(str(asset), str(out))
        finally:
            if old_enable is None:
                os.environ.pop("KPFU_ENABLE_MARKER_SPLIT", None)
            else:
                os.environ["KPFU_ENABLE_MARKER_SPLIT"] = old_enable
            if old_apply is None:
                os.environ.pop("KPFU_APPLY_MARKER_SPLIT", None)
            else:
                os.environ["KPFU_APPLY_MARKER_SPLIT"] = old_apply

        doc = Document(str(out))

    if len(doc.tables) != 11:
        return _result(False, f"expected 11 tables after Рыбаков split, got {len(doc.tables)}")
    if len(doc.tables[9].rows) != 6 or len(doc.tables[10].rows) != 17:
        return _result(False, f"unexpected table row counts after split: {len(doc.tables[9].rows)}/{len(doc.tables[10].rows)}")
    if [c.text for c in doc.tables[9].rows[1].cells] != ["1", "2", "3", "4", "5"]:
        return _result(False, "first split table missing numbered row")
    if [c.text for c in doc.tables[10].rows[0].cells] != ["1", "2", "3", "4", "5"]:
        return _result(False, "second split table missing numbered row")
    continuations = [p.text for p in doc.paragraphs if "Продолжение таблицы" in (p.text or "")]
    if any(text == "Продолжение таблицы 10" for text in continuations):
        return _result(False, f"unexpected appendix continuation paragraph inserted: {continuations!r}")
    return _result(True, "Рыбаков target is split with numbered rows in active mode")


def test_marker_runtime_real_bondarev_keeps_headings_safe() -> tuple[bool, str]:
    asset = ASSETS / "курсовая_Бондарев_Никита_2_курс.docx"
    if not asset.exists():
        return _result(True, "Бондарев asset missing, skipped")

    def snapshot(path: Path):
        doc = Document(str(path))
        out = []
        for p in doc.paragraphs:
            text = (p.text or "").strip()
            if text in {
                "ВВЕДЕНИЕ",
                "1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ ФУНКЦИЙ И ОРГАНОВ САМОУПРАВЛЕНИЯ В ОРГАНИЗАЦИИ",
                "1.1. Понятие, сущность и классификация органов самоуправления в организациях",
            }:
                p_pr = p._element.find(qn("w:pPr"))
                num_pr = p_pr.find(qn("w:numPr")) if p_pr is not None else None
                fonts = []
                for run in p.runs[:2]:
                    r_pr = run._element.find(qn("w:rPr"))
                    r_fonts = r_pr.find(qn("w:rFonts")) if r_pr is not None else None
                    fonts.append(r_fonts.get(qn("w:ascii")) if r_fonts is not None else None)
                out.append((text, p.style.name if p.style else None, num_pr is not None, tuple(fonts)))
        return out

    with tempfile.TemporaryDirectory() as tmp:
        off = Path(tmp) / "bond_off.docx"
        on = Path(tmp) / "bond_on.docx"

        old_enable = os.environ.get("KPFU_ENABLE_MARKER_SPLIT")
        old_apply = os.environ.get("KPFU_APPLY_MARKER_SPLIT")
        try:
            os.environ.pop("KPFU_ENABLE_MARKER_SPLIT", None)
            os.environ.pop("KPFU_APPLY_MARKER_SPLIT", None)
            format_docx(str(asset), str(off))
            os.environ["KPFU_ENABLE_MARKER_SPLIT"] = "1"
            os.environ["KPFU_APPLY_MARKER_SPLIT"] = "1"
            format_docx(str(asset), str(on))
        finally:
            if old_enable is None:
                os.environ.pop("KPFU_ENABLE_MARKER_SPLIT", None)
            else:
                os.environ["KPFU_ENABLE_MARKER_SPLIT"] = old_enable
            if old_apply is None:
                os.environ.pop("KPFU_APPLY_MARKER_SPLIT", None)
            else:
                os.environ["KPFU_APPLY_MARKER_SPLIT"] = old_apply

        off_snapshot = snapshot(off)
        on_snapshot = snapshot(on)

    if off_snapshot != on_snapshot:
        return _result(False, f"Бондарев heading snapshot changed under active split: off={off_snapshot!r} on={on_snapshot!r}")
    if any(item[2] for item in on_snapshot):
        return _result(False, f"Бондарев headings unexpectedly gained w:numPr: {on_snapshot!r}")
    if any("Calibri" in (font or "") for item in on_snapshot for font in item[3]):
        return _result(False, f"Бондарев headings unexpectedly use Calibri: {on_snapshot!r}")
    return _result(True, "Бондарев active mode leaves headings unchanged")


# ── Runner ────────────────────────────────────────────────────────────────────

def run_all() -> None:
    # Default suite: fast synthetic/XML checks for confirmed product rules.
    # Real asset formatting is useful as a smoke check, but it is slower and
    # can preserve broken historical output; keep it opt-in below.
    tests = [
        # Figure/paragraph preservation.
        ("A  | rule4 does not delete images",          test_a_rule4_does_not_delete_images),
        ("A  | _para_has_image helper",                test_a_para_has_image_helper),
        ("A  | rule4 preserves section breaks",        test_a_rule4_preserves_front_matter_section_breaks),
        # Table continuation and split behavior.
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
        ("C  | rendered first-row spill move",         test_c_rendered_start_page_first_row_spill_moves_whole_table),
        ("C  | rendered skip existing page break",     test_c_rendered_start_page_skips_existing_page_break_candidate),
        ("C  | rendered disabled page break",          test_c_rendered_start_page_upgrades_disabled_page_break),
        ("C  | rendered start-page ambiguity skip",    test_c_rendered_start_page_skips_ambiguous_usability),
        ("C  | rendered first-row spill weak skip",    test_c_rendered_start_page_first_row_spill_needs_strong_next_page_evidence),
        ("C  | rendered first-row spill prose skip",   test_c_rendered_start_page_first_row_spill_ignores_later_prose_token_reuse),
        ("C  | rendered decision logging",             test_c_rendered_decision_logging_for_ambiguous_skip),
        ("C  | rendered start-page complete row",      test_c_rendered_start_page_keeps_table_with_clear_complete_data_row),
        ("C  | vMerge guard",                          test_c_vmerge_guard_rejects_boundary_inside_merge_zone),
        # General DOCX formatting invariants used by Phase 3 output.
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
        # Heading product rules: no Word autonumbering, manual text numbering remains.
        ("T2 | 'Глава N' without title → heading1", test_t2_chapter_heading_without_title),
        ("T2 | manual heading2 still works", test_t2_manual_heading2_still_promoted),
        ("T2 | Word-autonumbered heading2 still works", test_t2_word_autonumbered_heading2_with_style_still_promoted),
        ("T2 | Word-autonumbered heading1 still works", test_t2_word_autonumbered_heading1_with_style_still_promoted),
        ("T2 | heading style numbering removed", test_t2_heading_style_numbering_is_removed),
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
        # Marker split diagnostics and runtime decisions.
        ("M1 | source unchanged after instrumentation", test_marker_instrumentation_keeps_source_unchanged),
        ("M1 | only target table instrumented", test_marker_instrumentation_only_targets_selected_table),
        ("M1 | inline marker parsing", test_marker_extract_handles_inline_text_and_missing_rows),
        ("M1 | keep_temp mapping result", test_marker_map_rows_to_pages_keep_temp_debug_paths),
        ("M1 | 1pt fallback to 2pt", test_marker_map_rows_to_pages_falls_back_to_2pt_and_returns_debug_info),
        ("M1 | invalid table index", test_marker_instrumentation_rejects_invalid_table_index),
        ("M1 | row page span summary", test_marker_page_span_summary),
        ("M1 | diagnose all tables summary", test_marker_diagnose_all_tables_summary),
        ("M1 | diagnose table error handling", test_marker_diagnose_table_handles_mapping_error),
        ("M1 | appendix/caption metadata", test_marker_appendix_and_caption_metadata),
        ("M1 | dry-run eligible boundary", test_marker_runtime_dry_run_clean_two_page_table_is_eligible),
        ("M1 | dry-run duplicate skip", test_marker_runtime_dry_run_skips_duplicate_rows),
        ("M1 | dry-run missing skip", test_marker_runtime_dry_run_skips_missing_rows_outside_header),
        ("M1 | dry-run 3-page skip", test_marker_runtime_dry_run_skips_three_page_tables),
        ("M1 | dry-run eligible logging", test_marker_runtime_dry_run_logs_eligible_candidate),
        ("M1 | dry-run flag off", test_marker_runtime_dry_run_feature_flag_off_skips_detection_hook),
        ("M1 | dry-run no mutation", test_marker_runtime_dry_run_only_does_not_mutate_document),
        ("M1 | apply appendix split", test_marker_runtime_apply_split_for_appendix_table),
        ("M1 | apply ordinary split", test_marker_runtime_apply_split_for_ordinary_table),
        ("M1 | apply ineligible skip", test_marker_runtime_apply_skips_ineligible_tables),
        ("M1 | apply idempotent", test_marker_runtime_apply_is_idempotent_on_second_run),
        # Prototype split rules.
        ("S1 | prototype simple table split", test_split_prototype_simple_table),
        ("S1 | source note after second table", test_split_prototype_source_note_stays_after_second_table),
        ("S1 | original doc unchanged", test_split_prototype_original_document_unchanged),
        ("S1 | invalid table index", test_split_prototype_invalid_table_index),
        ("S1 | invalid split_before_row", test_split_prototype_invalid_split_before_row),
        ("S1 | no continuation paragraph", test_split_prototype_no_continuation_paragraph_inserted),
        ("S1 | numbered ordinary continuation", test_split_prototype_numbered_ordinary_continuation_row_only),
        ("S1 | numbered ordinary split caption", test_split_prototype_numbered_ordinary_split_caption_before_title),
        ("S1 | numbered appendix continuation", test_split_prototype_numbered_appendix_has_no_continuation_text),
        ("S1 | numbered row reused", test_split_prototype_numbered_existing_row_reused_without_duplicate),
        ("S1 | numbered malformed row", test_split_prototype_numbered_malformed_existing_row_fails_safely),
        ("S1 | numbered source note", test_split_prototype_numbered_source_note_after_second_table),
        ("S1 | numbered original unchanged", test_split_prototype_numbered_original_document_unchanged),
        ("S1 | numbered row safe markup", test_split_prototype_numbered_row_has_no_numpr_and_no_calibri),
        ("M1 | headings unchanged across flags", test_marker_runtime_flags_do_not_change_headings),
    ]

    if os.environ.get("KPFU_RUN_LONG_PHASE3_TESTS") == "1":
        tests.extend([
            ("M1 | real Рыбаков split", test_marker_runtime_real_rybakov_target_applies_split),
            ("M1 | real Бондарев headings", test_marker_runtime_real_bondarev_keeps_headings_safe),
        ])
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
