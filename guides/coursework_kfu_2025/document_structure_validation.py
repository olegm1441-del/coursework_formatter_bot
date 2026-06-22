"""
Document-structure regression gate.

Tables / page-break / continuation code must never destroy global document
structure. This evaluates the *rendered* PDF lines and asserts that the required
KFU sections survive, in order:

  - СОДЕРЖАНИЕ (TOC) exists when the source has a TOC, and appears before the
    real ВВЕДЕНИЕ heading;
  - ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ exist;
  - at least one numbered chapter heading survives;
  - ПРИЛОЖЕНИЯ exists when the source has appendices.

A *heading* is an exact standalone line (e.g. ``ВВЕДЕНИЕ``); TOC entries
(``ВВЕДЕНИЕ ....... 3``) are not exact matches and are therefore ignored. Pure
function over ``pdf_lines`` so the smoke and tests share one implementation.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from .pdf_layout_analyzer import PdfLine


@dataclass(frozen=True)
class StructureIssue:
    issue_type: str
    severity: str  # "fail" | "needs_human_review"
    page: int | None
    evidence: dict[str, object] = field(default_factory=dict)


_CHAPTER_HEADING_RE = re.compile(r"^\d+(?:\.\d+)*\.?\s+[A-Za-zА-Яа-яЁё]")
_REFERENCES_RE = re.compile(r"^список\s+использован", re.IGNORECASE)
_APPENDIX_LABEL_RE = re.compile(r"^приложение\s+\S+$", re.IGNORECASE)


def _norm(text: str) -> str:
    return " ".join((text or "").replace("\xa0", " ").split())


def _exact_heading_pages(pdf_lines: list[PdfLine], target: str) -> list[int]:
    t = target.upper()
    return sorted({l.page_num for l in pdf_lines if _norm(l.text).upper() == t})


def _references_pages(pdf_lines: list[PdfLine]) -> list[int]:
    # exact-ish: a standalone references heading (no dot-leader / page number)
    out = []
    for l in pdf_lines:
        norm = _norm(l.text)
        if _REFERENCES_RE.match(norm) and ".." not in norm and not re.search(r"\d\s*$", norm):
            out.append(l.page_num)
    return sorted(set(out))


def _chapter_heading_pages(pdf_lines: list[PdfLine]) -> list[int]:
    out = []
    for l in pdf_lines:
        norm = _norm(l.text)
        if ".." in norm:  # TOC entry, skip
            continue
        if _CHAPTER_HEADING_RE.match(norm) and not re.search(r"\d\s*$", norm):
            out.append(l.page_num)
    return sorted(set(out))


def _appendix_label_pages(pdf_lines: list[PdfLine]) -> list[int]:
    out = []
    for l in pdf_lines:
        norm = _norm(l.text)
        if norm.lower() == "приложения":
            continue
        if _APPENDIX_LABEL_RE.match(norm) and ".." not in norm:
            out.append(l.page_num)
    return sorted(set(out))


def source_has_toc(source_text: str) -> bool:
    return "СОДЕРЖАНИЕ" in (source_text or "").upper()


def source_has_appendix(source_text: str) -> bool:
    """True only when the source has a real appendix SECTION — a standalone
    ``ПРИЛОЖЕНИЯ`` heading, an UPPERCASE ``ПРИЛОЖЕНИЕ X`` label, or ≥2 standalone
    ``Приложение X`` labels (e.g. А and Б). A single mixed-case ``Приложение 1``
    is ambiguous with a reference entry, so it is not enough — this avoids a
    false ``missing_appendices`` on docs that merely cite an appendix."""
    lines = [" ".join((ln or "").split()) for ln in (source_text or "").splitlines()]
    if any(ln.lower() == "приложения" for ln in lines):
        return True
    if any(re.match(r"^ПРИЛОЖЕНИ[ЯЕ]\b", ln) for ln in lines):
        return True
    labels = [ln for ln in lines if re.match(r"^приложение\s+\S+$", ln, re.IGNORECASE)]
    return len(labels) >= 2


def evaluate_document_structure(
    pdf_lines: list[PdfLine],
    *,
    expect_toc: bool = True,
    expect_appendix: bool = False,
) -> list[StructureIssue]:
    issues: list[StructureIssue] = []

    toc_pages = _exact_heading_pages(pdf_lines, "СОДЕРЖАНИЕ")
    intro_pages = _exact_heading_pages(pdf_lines, "ВВЕДЕНИЕ")
    concl_pages = _exact_heading_pages(pdf_lines, "ЗАКЛЮЧЕНИЕ")
    ref_pages = _references_pages(pdf_lines)
    chapter_pages = _chapter_heading_pages(pdf_lines)
    appendix_section = _exact_heading_pages(pdf_lines, "ПРИЛОЖЕНИЯ")
    appendix_labels = _appendix_label_pages(pdf_lines)

    if expect_toc and not toc_pages:
        issues.append(StructureIssue("missing_toc", "fail", None,
                                     {"detail": "СОДЕРЖАНИЕ heading not found in rendered output"}))
    if not intro_pages:
        issues.append(StructureIssue("missing_intro", "fail", None,
                                     {"detail": "ВВЕДЕНИЕ heading not found"}))
    if expect_toc and toc_pages and intro_pages and min(toc_pages) >= min(intro_pages):
        issues.append(StructureIssue("toc_after_intro", "fail", min(toc_pages),
                                     {"toc_page": min(toc_pages), "intro_page": min(intro_pages)}))
    if not concl_pages:
        issues.append(StructureIssue("missing_conclusion", "fail", None,
                                     {"detail": "ЗАКЛЮЧЕНИЕ heading not found"}))
    if not ref_pages:
        issues.append(StructureIssue("missing_references", "fail", None,
                                     {"detail": "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ not found"}))
    if not chapter_pages:
        issues.append(StructureIssue("missing_chapters", "fail", None,
                                     {"detail": "no numbered chapter heading found"}))
    if expect_appendix and not (appendix_section or appendix_labels):
        issues.append(StructureIssue("missing_appendices", "fail", None,
                                     {"detail": "source has appendices but none rendered"}))

    return issues
