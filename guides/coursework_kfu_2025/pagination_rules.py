"""
Phase 2 — Pagination rules via DOCX paragraph properties.
Runs AFTER safe_formatter (which resets all keep_with_next / keep_together to False).

Rules implemented:
    Rule 3  — Table caption / title must stay with the table below.
    Rule 5  — Section heading must not be the last line of a page without body text.
    Rule 6  — Figure must stay with its caption.

Rules NOT implemented here (require real page rendering → Phase 3):
    Rule 1  — Table continuation header ("Продолжение таблицы X.Y.Z")
    Rule 2  — No more than 3 empty lines at the bottom of a page
    Rule 4  — No empty first line of a page
"""

import logging
from docx import Document
from docx.oxml.ns import qn

from .classifier import classify_paragraph, paragraph_text

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

def _set_keep_with_next(paragraph, value: bool = True) -> None:
    paragraph.paragraph_format.keep_with_next = value


def _has_image(paragraph) -> bool:
    """True if paragraph contains an inline drawing or picture."""
    elem = paragraph._element
    return bool(
        elem.findall(".//" + qn("w:drawing"))
        or elem.findall(".//" + qn("w:pict"))
    )


def _classify_all(paragraphs) -> list[str]:
    """Return classification list aligned with paragraphs list."""
    kinds: list[str] = []
    for p in paragraphs:
        text = paragraph_text(p)
        prev_kind = kinds[-1] if kinds else None
        kinds.append(classify_paragraph(text, prev_kind=prev_kind))
    return kinds


# ---------------------------------------------------------------------------
# Rule 3 — Table caption / title → keep with table
# ---------------------------------------------------------------------------

def _apply_rule3(paragraphs: list, kinds: list[str]) -> int:
    """
    Set keep_with_next=True on every table_caption and table_title paragraph.
    The last paragraph in the caption/title chain will keep with the <w:tbl>
    element that follows it in the document body — Word honours this.
    """
    count = 0
    for p, kind in zip(paragraphs, kinds):
        if kind in ("table_caption", "table_title"):
            _set_keep_with_next(p)
            count += 1
    return count


# ---------------------------------------------------------------------------
# Rule 5 — Section headings → keep with first body paragraph below
# ---------------------------------------------------------------------------

def _apply_rule5(paragraphs: list, kinds: list[str]) -> int:
    """
    Set keep_with_next=True on heading1 and heading2 paragraphs.
    Word will not allow the heading to be the last line of a page.
    """
    count = 0
    for p, kind in zip(paragraphs, kinds):
        if kind in ("heading1", "heading2"):
            _set_keep_with_next(p)
            count += 1
    return count


# ---------------------------------------------------------------------------
# Rule 6 — Figure → keep with its caption
# ---------------------------------------------------------------------------

def _apply_rule6(paragraphs: list, kinds: list[str]) -> int:
    """
    For each paragraph that contains an image, set keep_with_next=True so the
    caption that immediately follows stays on the same page as the figure.

    Also: if a figure_caption is immediately preceded by a non-image paragraph
    that has keep_with_next already set — no extra action needed.
    If a figure_caption has NO preceding image paragraph (rare edge case), skip.
    """
    count = 0
    n = len(paragraphs)
    for i, (p, kind) in enumerate(zip(paragraphs, kinds)):
        # Case A: paragraph contains an image → chain it to the next paragraph
        if _has_image(p):
            _set_keep_with_next(p)
            count += 1
            continue

        # Case B: figure_caption but the paragraph just above is NOT an image
        # (e.g. image is wrapped in a frame or textbox — rare).
        # As a safety measure, set keep_with_next on the caption itself so it
        # at least chains to whatever is after it (avoids caption being totally
        # isolated at the top of a page in these edge cases).
        if kind == "figure_caption":
            prev_kind = kinds[i - 1] if i > 0 else None
            if prev_kind != "figure_caption" and not _has_image(paragraphs[i - 1] if i > 0 else p):
                _set_keep_with_next(p)
                count += 1

    return count


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def apply_pagination_rules(doc: Document) -> None:
    """
    Apply all Phase-2 pagination rules to an already-formatted Document object.
    Modifies the document in place; caller is responsible for saving.
    """
    paragraphs = doc.paragraphs
    kinds = _classify_all(paragraphs)

    r3 = _apply_rule3(paragraphs, kinds)
    r5 = _apply_rule5(paragraphs, kinds)
    r6 = _apply_rule6(paragraphs, kinds)

    logger.info(
        "pagination_rules applied: rule3=%d rule5=%d rule6=%d paragraphs",
        r3, r5, r6,
    )
