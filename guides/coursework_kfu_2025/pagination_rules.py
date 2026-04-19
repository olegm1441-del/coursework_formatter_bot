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
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

from .classifier import classify_paragraph, paragraph_text
from .docx_utils import is_source_or_note_line
from .rules import FIRST_LINE_INDENT_CM

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

def _set_keep_with_next(paragraph, value: bool = True) -> None:
    paragraph.paragraph_format.keep_with_next = value


def _set_keep_together(paragraph, value: bool = True) -> None:
    """Set keep_together (w:keepLines) — prevents a paragraph from being split mid-text."""
    paragraph.paragraph_format.keep_together = value


def _set_keep_next_xml(p_elem) -> None:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)
    if pPr.find(qn("w:keepNext")) is None:
        pPr.append(OxmlElement("w:keepNext"))


def _normalise_source_note(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.first_line_indent = Cm(FIRST_LINE_INDENT_CM)


def _has_image(paragraph) -> bool:
    """True if paragraph contains an inline drawing or picture."""
    from .docx_utils import xml_has_image
    return xml_has_image(paragraph._element)


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
            # keep_together prevents a long multi-line title from being split
            # across pages by Word's line-breaker (keepWithNext alone won't stop
            # that — it only keeps the paragraph joined to the NEXT element).
            _set_keep_together(p)
            count += 1
    return count


# ---------------------------------------------------------------------------
# Rule 5 — Section headings → keep with first body paragraph below
# ---------------------------------------------------------------------------

def _apply_rule5(paragraphs: list, kinds: list[str]) -> int:
    """
    Set keep_with_next=True on heading1/heading2 AND any immediately following
    empty paragraphs, so the chain reaches the first real body paragraph.

    Without propagating through empties, Word keeps heading+empty together but
    the empty paragraph still breaks before the body text — heading appears to
    hang alone at the bottom of the page.
    """
    count = 0
    n = len(paragraphs)
    for i, (p, kind) in enumerate(zip(paragraphs, kinds)):
        if kind not in ("heading1", "heading2"):
            continue

        _set_keep_with_next(p)
        # keep_together prevents a multi-line heading from being split mid-text
        # across page boundaries (keepWithNext alone only chains to next element).
        _set_keep_together(p)
        count += 1

        # Propagate through trailing empty paragraphs so the chain reaches
        # the first non-empty paragraph below the heading.
        j = i + 1
        while j < n and kinds[j] == "empty_paragraph":
            _set_keep_with_next(paragraphs[j])
            count += 1
            j += 1

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
        # Case A: paragraph contains an image → chain it to the caption below.
        # Propagate keepWithNext through any intervening empty paragraphs so the
        # chain reaches the figure_caption even when blank lines separate them.
        if _has_image(p):
            _set_keep_with_next(p)
            count += 1
            j = i + 1
            while j < n and kinds[j] == "empty_paragraph":
                _set_keep_with_next(paragraphs[j])
                count += 1
                j += 1
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
# Rule E — Source / note line must stay directly below its table
# ---------------------------------------------------------------------------

def _apply_rule_source_note(doc: Document) -> int:
    """
    Walk the document body.  When a paragraph starting with 'Источник:' or
    'Примечание:' immediately follows a table (with at most one intervening
    empty paragraph), set keep_with_next=True on the last paragraph in the
    last cell of the table's last row.

    This signals Word: "don't break between the table tail and the following
    source/note line."  Word then keeps them on the same page.

    Returns the number of tables whose last row was tagged.
    """
    body = doc.element.body
    children = list(body)
    n = len(children)
    count = 0

    para_map = {p._element: p for p in doc.paragraphs}

    i = 0
    while i < n:
        elem = children[i]
        local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag

        if local != "tbl":
            i += 1
            continue

        # Look ahead: skip at most one empty paragraph, then check for source/note
        j = i + 1
        skipped_empty = 0
        while j < n and skipped_empty <= 1:
            nelem = children[j]
            nlocal = nelem.tag.split("}")[-1] if "}" in nelem.tag else nelem.tag
            if nlocal != "p":
                break
            # Get paragraph text
            texts = nelem.findall(".//" + qn("w:t"))
            text = "".join(t.text or "" for t in texts).strip()
            if not text:
                skipped_empty += 1
                j += 1
                continue
            # Non-empty paragraph found
            if is_source_or_note_line(text):
                # Tag last paragraph in last cell of last row of this table
                tr_elems = elem.findall(qn("w:tr"))
                if tr_elems:
                    last_tr = tr_elems[-1]
                    tc_elems = last_tr.findall(qn("w:tc"))
                    if tc_elems:
                        last_tc = tc_elems[-1]
                        p_elems = last_tc.findall(qn("w:p"))
                        if p_elems:
                            _set_keep_next_xml(p_elems[-1])
                            count += 1
                            logger.debug(
                                "rule_source_note: keepWithNext on last row of table "
                                "before '%s'", text[:40]
                            )

                source_note_paras = []
                k = j
                while k < n:
                    pelem = children[k]
                    plocal = pelem.tag.split("}")[-1] if "}" in pelem.tag else pelem.tag
                    if plocal != "p":
                        break
                    p_text = "".join(t.text or "" for t in pelem.findall(".//" + qn("w:t"))).strip()
                    if not is_source_or_note_line(p_text):
                        break
                    para = para_map.get(pelem)
                    if para is None:
                        break
                    source_note_paras.append(para)
                    k += 1

                for idx, para in enumerate(source_note_paras):
                    _normalise_source_note(para)
                    if idx < len(source_note_paras) - 1:
                        _set_keep_with_next(para)
            break  # done with this table regardless
        i += 1

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
    rE = _apply_rule_source_note(doc)

    logger.info(
        "pagination_rules applied: rule3=%d rule5=%d rule6=%d rule_source_note=%d paragraphs",
        r3, r5, r6, rE,
    )
