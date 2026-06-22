"""Structure regression gate tests: table/page-break code must never drop the
TOC or required sections. Pins `evaluate_document_structure`.

Run: python3 tests/test_document_structure.py
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from guides.coursework_kfu_2025.pdf_layout_analyzer import PdfLine  # noqa: E402
from guides.coursework_kfu_2025.document_structure_validation import (  # noqa: E402
    evaluate_document_structure,
    source_has_appendix,
    source_has_toc,
)


def _result(ok: bool, message: str) -> tuple[bool, str]:
    return ok, message


def _line(text, page, top=80.0):
    return PdfLine(text=text, page_num=page, top=top, bottom=top + 12)


def _well_formed():
    return [
        _line("Министерство науки и высшего образования", 1, 60),  # title page
        _line("СОДЕРЖАНИЕ", 2, 58),
        _line("ВВЕДЕНИЕ ............................... 3", 2, 95),  # TOC entry
        _line("ЗАКЛЮЧЕНИЕ ............................. 50", 2, 120),
        _line("ВВЕДЕНИЕ", 3, 58),
        _line("1. Теоретические основы темы", 4, 58),
        _line("2. Анализ предметной области", 20, 58),
        _line("ЗАКЛЮЧЕНИЕ", 50, 58),
        _line("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", 52, 58),
        _line("ПРИЛОЖЕНИЯ", 55, 58),
        _line("ПРИЛОЖЕНИЕ А", 56, 58),
    ]


def test_well_formed_document_has_no_structure_issues() -> tuple[bool, str]:
    issues = evaluate_document_structure(_well_formed(), expect_toc=True, expect_appendix=True)
    if issues:
        return _result(False, f"unexpected structure issues: {[i.issue_type for i in issues]}")
    return _result(True, "well-formed document passes the structure gate")


def test_missing_toc_is_fail() -> tuple[bool, str]:
    lines = [l for l in _well_formed() if l.text != "СОДЕРЖАНИЕ"]
    issues = evaluate_document_structure(lines, expect_toc=True, expect_appendix=True)
    if not any(i.issue_type == "missing_toc" and i.severity == "fail" for i in issues):
        return _result(False, f"missing TOC not flagged: {[i.issue_type for i in issues]}")
    return _result(True, "missing СОДЕРЖАНИЕ is a hard structure fail")


def test_toc_after_intro_is_fail() -> tuple[bool, str]:
    # TOC heading rendered AFTER the real intro (page order inverted)
    lines = [
        _line("ВВЕДЕНИЕ", 2, 58),
        _line("СОДЕРЖАНИЕ", 4, 58),
        _line("1. Глава", 5, 58),
        _line("ЗАКЛЮЧЕНИЕ", 9, 58),
        _line("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", 10, 58),
    ]
    issues = evaluate_document_structure(lines, expect_toc=True, expect_appendix=False)
    if not any(i.issue_type == "toc_after_intro" for i in issues):
        return _result(False, f"TOC-after-intro not flagged: {[i.issue_type for i in issues]}")
    return _result(True, "TOC after ВВЕДЕНИЕ is a hard structure fail")


def test_missing_required_sections_flagged() -> tuple[bool, str]:
    lines = [
        _line("СОДЕРЖАНИЕ", 2, 58),
        _line("ВВЕДЕНИЕ", 3, 58),
        _line("1. Глава", 4, 58),
        # no ЗАКЛЮЧЕНИЕ, no СПИСОК
    ]
    issues = evaluate_document_structure(lines, expect_toc=True, expect_appendix=False)
    types = {i.issue_type for i in issues}
    if "missing_conclusion" not in types or "missing_references" not in types:
        return _result(False, f"missing sections not all flagged: {types}")
    return _result(True, "missing ЗАКЛЮЧЕНИЕ / СПИСОК are hard fails")


def test_toc_entries_not_mistaken_for_headings() -> tuple[bool, str]:
    # Only TOC entries (with dot leaders) for intro/conclusion — no real headings.
    lines = [
        _line("СОДЕРЖАНИЕ", 2, 58),
        _line("ВВЕДЕНИЕ ......................... 3", 2, 95),
        _line("ЗАКЛЮЧЕНИЕ ....................... 9", 2, 120),
        _line("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ ... 10", 2, 140),
        _line("1. Глава", 3, 58),
    ]
    issues = evaluate_document_structure(lines, expect_toc=True, expect_appendix=False)
    types = {i.issue_type for i in issues}
    # the real intro/conclusion/references headings are absent → must be flagged
    if not {"missing_intro", "missing_conclusion", "missing_references"} <= types:
        return _result(False, f"TOC entries wrongly accepted as headings: {types}")
    return _result(True, "TOC dot-leader entries are not counted as real headings")


def test_appendix_not_expected_when_source_has_none() -> tuple[bool, str]:
    lines = [
        _line("СОДЕРЖАНИЕ", 2, 58),
        _line("ВВЕДЕНИЕ", 3, 58),
        _line("1. Глава", 4, 58),
        _line("ЗАКЛЮЧЕНИЕ", 9, 58),
        _line("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", 10, 58),
    ]
    issues = evaluate_document_structure(lines, expect_toc=True, expect_appendix=False)
    if any(i.issue_type == "missing_appendices" for i in issues):
        return _result(False, "appendix wrongly required when source has none")
    return _result(True, "appendix not required when source has no appendices")


def test_source_appendix_detection() -> tuple[bool, str]:
    # real appendix: two standalone labels (А and Б, mixed case)
    real = "ВВЕДЕНИЕ\n1. Глава\nПРИЛОЖЕНИЯ\nПриложение А\nПриложение Б"
    # ambiguous: a single 'Приложение 1' that is actually a reference entry
    ref_only = "ВВЕДЕНИЕ\nСПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ\nПриложение 1"
    if not source_has_appendix(real):
        return _result(False, "real appendix section (А+Б) not detected")
    if source_has_appendix(ref_only):
        return _result(False, "single reference-like 'Приложение 1' wrongly treated as appendix")
    if not source_has_toc("СОДЕРЖАНИЕ\nВВЕДЕНИЕ"):
        return _result(False, "TOC not detected")
    return _result(True, "source appendix/TOC detection distinguishes real sections from references")


def main() -> int:
    tests = [
        ("well-formed passes", test_well_formed_document_has_no_structure_issues),
        ("source appendix detection", test_source_appendix_detection),
        ("missing TOC is fail", test_missing_toc_is_fail),
        ("TOC after intro is fail", test_toc_after_intro_is_fail),
        ("missing required sections", test_missing_required_sections_flagged),
        ("TOC entries not headings", test_toc_entries_not_mistaken_for_headings),
        ("appendix not over-required", test_appendix_not_expected_when_source_has_none),
    ]
    failed = 0
    for name, fn in tests:
        ok, msg = fn()
        print(f"[{'PASS' if ok else 'FAIL'}] {name} — {msg}")
        if not ok:
            failed += 1
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
