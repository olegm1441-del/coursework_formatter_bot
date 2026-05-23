from __future__ import annotations

import logging
import re
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from .classifier import clean_spaces, find_body_start_index, parse_heading1, parse_heading2
from .layout_render import render_docx_to_pdf
from .page_breaks import apply_page_breaks
from .page_numbering import apply_page_numbering_policy
from .pdf_layout_analyzer import analyze_pdf_lines

logger = logging.getLogger(__name__)

_CONTENTS_HEADING_RE = re.compile(r"^\s*(содержание|оглавление)\s*$", re.IGNORECASE)
_SOURCE_NOTE_RE = re.compile(r"^\s*(источник|составлено по|рассчитано по|примечание)\s*:", re.IGNORECASE)
_TABLE_RE = re.compile(r"^\s*таблица\s+\d+(?:\.\d+){0,2}\b", re.IGNORECASE)
_FIGURE_RE = re.compile(r"^\s*(рис\.|рисунок)\s*\d+(?:\.\d+){0,2}\b", re.IGNORECASE)

_STRUCTURAL_HEADINGS = {
    "введение": "ВВЕДЕНИЕ",
    "заключение": "ЗАКЛЮЧЕНИЕ",
    "список использованных источников": "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    "список использованной литературы": "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    "приложения": "ПРИЛОЖЕНИЯ",
}

_PAGE_PLACEHOLDER = "000"


@dataclass(frozen=True)
class TocEntry:
    title: str
    bookmark_name: str
    paragraph_index: int


def _norm(text: str) -> str:
    return clean_spaces(text).lower().rstrip(".")


def _is_contents_heading(text: str) -> bool:
    return bool(_CONTENTS_HEADING_RE.match(clean_spaces(text)))


def _is_intro_heading(text: str) -> bool:
    return _norm(text) == "введение"


def _find_body_start_index_for_contents(document: Document) -> int | None:
    """
    Find the real body introduction for TOC rebuild.

    The generic classifier intentionally treats a standalone "ВВЕДЕНИЕ" as the
    body start. Old hand-made TOCs sometimes contain exactly that standalone
    entry before the real body, so for contents replacement we skip everything
    after the last pre-body contents/oglavlenie heading and use the last
    standalone intro after it.
    """
    paragraphs = document.paragraphs
    contents_indices = [
        idx
        for idx, paragraph in enumerate(paragraphs)
        if _is_contents_heading(paragraph.text)
    ]
    if contents_indices:
        last_contents_idx = max(contents_indices)
        intro_candidates = [
            idx
            for idx in range(last_contents_idx + 1, len(paragraphs))
            if _is_intro_heading(paragraphs[idx].text)
        ]
        if intro_candidates:
            return intro_candidates[-1]

    return find_body_start_index(document)


def _is_appendix_local_heading(text: str) -> bool:
    return bool(re.match(r"^\s*приложение\s+(?:\d{1,3}|[a-zа-яё])\b", text, re.IGNORECASE))


def _is_excluded_service_line(text: str) -> bool:
    return bool(_SOURCE_NOTE_RE.match(text) or _TABLE_RE.match(text) or _FIGURE_RE.match(text))


def _style_name(paragraph) -> str:
    try:
        return (paragraph.style.name or "").lower()
    except Exception:
        return ""


def _is_heading1_style(paragraph) -> bool:
    return _style_name(paragraph) in {"heading 1", "заголовок 1"}


def _is_heading2_style(paragraph) -> bool:
    return _style_name(paragraph) in {"heading 2", "заголовок 2"}


def _collect_body_entries(document: Document, body_start: int) -> list[TocEntry]:
    entries: list[TocEntry] = []

    for idx, paragraph in enumerate(document.paragraphs[body_start:], start=body_start):
        text = clean_spaces(paragraph.text)
        if not text or _is_excluded_service_line(text):
            continue

        low = _norm(text)
        if low in _STRUCTURAL_HEADINGS:
            canonical = _STRUCTURAL_HEADINGS[low]
            entries.append(TocEntry(canonical, f"kpfu_toc_{len(entries) + 1}", idx))
            if canonical == "ПРИЛОЖЕНИЯ":
                break
            continue

        if _is_appendix_local_heading(text):
            continue

        parsed_h1 = parse_heading1(text)
        if parsed_h1 and parsed_h1["kind"] == "heading1_chapter" and _is_heading1_style(paragraph):
            entries.append(TocEntry(text, f"kpfu_toc_{len(entries) + 1}", idx))
            continue

        if parse_heading2(text) and _is_heading2_style(paragraph):
            entries.append(TocEntry(text, f"kpfu_toc_{len(entries) + 1}", idx))

    return entries


def _find_existing_contents_start(document: Document, body_start: int) -> int | None:
    for idx, paragraph in enumerate(document.paragraphs[:body_start]):
        if _is_contents_heading(paragraph.text):
            return idx
    return None


def _remove_body_children_range(document: Document, start_elem, end_elem) -> None:
    body = document.element.body
    children = list(body)
    start_idx = children.index(start_elem)
    end_idx = children.index(end_elem)
    for child in children[start_idx:end_idx]:
        body.remove(child)


def _set_run_font(run, *, bold: bool) -> None:
    run.font.name = "Times New Roman"
    run.font.size = Pt(14)
    run.bold = bold
    run.italic = False
    run.underline = False
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:cs"), "Times New Roman")


def _run_pr_xml(*, bold: bool = False) -> OxmlElement:
    r_pr = OxmlElement("w:rPr")
    r_fonts = OxmlElement("w:rFonts")
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:cs"), "Times New Roman")
    r_pr.append(r_fonts)

    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "28")
    r_pr.append(size)

    size_cs = OxmlElement("w:szCs")
    size_cs.set(qn("w:val"), "28")
    r_pr.append(size_cs)

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "000000")
    r_pr.append(color)

    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "none")
    r_pr.append(underline)

    if bold:
        r_pr.append(OxmlElement("w:b"))
        r_pr.append(OxmlElement("w:bCs"))

    return r_pr


def _append_hyperlink_run(hyperlink: OxmlElement, text: str | None = None, *, tab: bool = False) -> None:
    run = OxmlElement("w:r")
    run.append(_run_pr_xml(bold=False))
    if tab:
        run.append(OxmlElement("w:tab"))
    else:
        text_el = OxmlElement("w:t")
        text_el.text = text or ""
        if text and (text.startswith(" ") or text.endswith(" ")):
            text_el.set(qn("xml:space"), "preserve")
        run.append(text_el)
    hyperlink.append(run)


def _set_internal_hyperlink_text(paragraph, title: str, page: str, anchor: str) -> None:
    p = paragraph._element
    p_pr = p.get_or_add_pPr()
    for child in list(p):
        if child is not p_pr:
            p.remove(child)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    hyperlink.set(qn("w:history"), "1")
    _append_hyperlink_run(hyperlink, title)
    _append_hyperlink_run(hyperlink, tab=True)
    _append_hyperlink_run(hyperlink, page)
    p.append(hyperlink)


def _clear_tab_stops(paragraph) -> None:
    p_pr = paragraph._element.get_or_add_pPr()
    for tabs in list(p_pr.findall(qn("w:tabs"))):
        p_pr.remove(tabs)


def _set_toc_entry_tab_stop(paragraph) -> None:
    p_pr = paragraph._element.get_or_add_pPr()
    _clear_tab_stops(paragraph)
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:leader"), "dot")
    tab.set(qn("w:pos"), str(Cm(16).twips))
    tabs.append(tab)
    p_pr.append(tabs)


def _format_contents_heading(paragraph) -> None:
    paragraph.text = "СОДЕРЖАНИЕ"
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.left_indent = Cm(0)
    paragraph.paragraph_format.right_indent = Cm(0)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1
    paragraph.paragraph_format.page_break_before = False
    _clear_tab_stops(paragraph)
    for run in paragraph.runs:
        _set_run_font(run, bold=True)


def _format_toc_entry(paragraph, title: str, page: str) -> None:
    paragraph.text = f"{title}\t{page}"
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.left_indent = Cm(0)
    paragraph.paragraph_format.right_indent = Cm(0)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.5
    paragraph.paragraph_format.page_break_before = False
    _set_toc_entry_tab_stop(paragraph)
    for run in paragraph.runs:
        _set_run_font(run, bold=False)


def _format_toc_entry_link(paragraph, entry: TocEntry, page: str) -> None:
    _format_toc_entry(paragraph, entry.title, page)
    _set_internal_hyperlink_text(paragraph, entry.title, page, entry.bookmark_name)


def _move_paragraphs_before(document: Document, paragraphs: list, reference_elem) -> None:
    body = document.element.body
    insert_idx = list(body).index(reference_elem)
    for paragraph in paragraphs:
        elem = paragraph._element
        body.remove(elem)
        body.insert(insert_idx, elem)
        insert_idx += 1


def _insert_contents_block(document: Document, entries: list[TocEntry], pages: dict[str, int | str]) -> int:
    body_start = _find_body_start_index_for_contents(document)
    if body_start is None:
        raise ValueError("real ВВЕДЕНИЕ heading not found")

    paragraphs = document.paragraphs
    intro_elem = paragraphs[body_start]._element
    contents_start = _find_existing_contents_start(document, body_start)
    removed_count = 0
    if contents_start is not None:
        removed_count = sum(
            1
            for paragraph in paragraphs[contents_start:body_start]
            if _is_contents_heading(paragraph.text)
        )
        _remove_body_children_range(document, paragraphs[contents_start]._element, intro_elem)

    new_paragraphs = [document.add_paragraph()]
    _format_contents_heading(new_paragraphs[0])
    for entry in entries:
        page = str(pages.get(entry.bookmark_name, _PAGE_PLACEHOLDER))
        p = document.add_paragraph()
        _format_toc_entry_link(p, entry, page)
        new_paragraphs.append(p)

    _move_paragraphs_before(document, new_paragraphs, intro_elem)
    return removed_count


def _line_norm(text: str) -> str:
    return clean_spaces(text).lower()


def _match_norm(text: str) -> str:
    text = _line_norm(text).replace("ё", "е")
    return re.sub(r"[^0-9a-zа-я]+", " ", text).strip()


def _is_heading_window_match(target: str, rendered: str) -> bool:
    target_norm = _match_norm(target)
    rendered_norm = _match_norm(rendered)
    if not target_norm or not rendered_norm:
        return False
    if rendered_norm == target_norm:
        return True
    if target_norm.startswith(rendered_norm) and len(rendered_norm) >= min(24, len(target_norm)):
        return True
    if rendered_norm.startswith(target_norm) and len(target_norm) >= 12:
        return True
    target_tokens = target_norm.split()
    rendered_tokens = rendered_norm.split()
    if len(target_tokens) >= 4 and len(rendered_tokens) >= 4:
        return rendered_tokens[:4] == target_tokens[:4] and len(rendered_norm) >= min(24, len(target_norm))
    return False


def _find_rendered_page(title: str, lines, *, min_page: int) -> int | None:
    target = _line_norm(title)
    if not target:
        return None
    for line in lines:
        if line.page_num < min_page:
            continue
        if _line_norm(line.text) == target:
            return line.page_num

    eligible = [line for line in lines if line.page_num >= min_page]
    for idx, line in enumerate(eligible):
        if _is_heading_window_match(title, line.text):
            return line.page_num
        joined = line.text
        for next_line in eligible[idx + 1:idx + 3]:
            if next_line.page_num != line.page_num:
                break
            joined = f"{joined} {next_line.text}"
            if _is_heading_window_match(title, joined):
                return line.page_num
    return None


def _validate_resolved_pages(entries: list[TocEntry], pages: dict[str, int]) -> None:
    ordered = [pages[entry.bookmark_name] for entry in entries]
    if not ordered:
        raise ValueError("no resolved TOC pages")

    if ordered[0] != 3:
        raise ValueError(f"resolved intro page must be 3, got {ordered[0]}")

    for prev, current in zip(ordered, ordered[1:]):
        if current < prev:
            raise ValueError(f"resolved TOC pages are not nondecreasing: {ordered}")

    if len(ordered) > 2 and len(set(ordered)) == 1:
        raise ValueError(f"degenerate resolved TOC pages: {ordered}")


def _resolve_display_pages(entries: list[TocEntry], lines) -> dict[str, int]:
    intro_pages = [
        line.page_num
        for line in lines
        if _line_norm(line.text) == "введение"
    ]
    if not intro_pages:
        raise ValueError("rendered ВВЕДЕНИЕ page not found")

    rendered_intro_page = max(intro_pages)
    offset = 3 - rendered_intro_page
    resolved: dict[str, int] = {}
    unresolved: list[str] = []
    min_page = rendered_intro_page
    for entry in entries:
        if _norm(entry.title) == "введение":
            page = rendered_intro_page
        else:
            page = _find_rendered_page(entry.title, lines, min_page=min_page)
        if page is None:
            unresolved.append(entry.title)
            continue
        resolved[entry.bookmark_name] = page + offset
        min_page = page
    if unresolved:
        logger.warning("contents_rebuild_unresolved headings=%s", unresolved)
        raise ValueError(f"rendered page not found for TOC entries: {unresolved}")
    _validate_resolved_pages(entries, resolved)
    return resolved


def _reapply_front_matter_layout(document: Document) -> None:
    from .safe_formatter import ensure_appendices_section_layout, ensure_front_matter_layout, normalize_sections

    body_start = _find_body_start_index_for_contents(document)
    if body_start is None:
        raise ValueError("real ВВЕДЕНИЕ heading not found")
    normalize_sections(document)
    ensure_front_matter_layout(document, body_start)
    apply_page_breaks(document, body_start)
    ensure_appendices_section_layout(document, body_start)
    apply_page_numbering_policy(document)


def _replace_contents_entry_pages(document: Document, entries: list[TocEntry], pages: dict[str, int]) -> None:
    body_start = _find_body_start_index_for_contents(document)
    if body_start is None:
        raise ValueError("real ВВЕДЕНИЕ heading not found")

    contents_start = _find_existing_contents_start(document, body_start)
    if contents_start is None:
        raise ValueError("draft contents block not found")

    entry_paragraphs = document.paragraphs[contents_start + 1:body_start]
    if len(entry_paragraphs) != len(entries):
        raise ValueError("draft contents entry count changed")

    for paragraph, entry in zip(entry_paragraphs, entries):
        _format_toc_entry_link(paragraph, entry, str(pages[entry.bookmark_name]))


def _add_bookmark(paragraph, name: str, bookmark_id: int) -> None:
    p = paragraph._element
    for existing in p.findall(qn("w:bookmarkStart")):
        if existing.get(qn("w:name")) == name:
            return

    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))

    insert_at = 1 if p.pPr is not None else 0
    p.insert(insert_at, start)
    p.append(end)


def _add_body_bookmarks(document: Document, entries: list[TocEntry]) -> None:
    paragraphs = document.paragraphs
    for idx, entry in enumerate(entries, start=1000):
        if 0 <= entry.paragraph_index < len(paragraphs):
            _add_bookmark(paragraphs[entry.paragraph_index], entry.bookmark_name, idx)


def rebuild_static_contents_page(docx_path: str | Path) -> bool:
    """
    Rebuild static KFU contents page. The source DOCX is replaced only after all
    TOC page numbers are resolved from a rendered draft; failures leave it intact.
    """
    source_path = Path(docx_path)
    workdir = Path(tempfile.mkdtemp(prefix="kpfu_contents_"))
    pdf_path: Path | None = None
    try:
        work_path = workdir / source_path.name
        shutil.copy2(source_path, work_path)

        document = Document(str(work_path))
        body_start = _find_body_start_index_for_contents(document)
        if body_start is None:
            logger.warning("contents_rebuild_skipped reason=intro_not_found path=%s", source_path)
            return False

        entries = _collect_body_entries(document, body_start)
        if not entries:
            logger.warning("contents_rebuild_skipped reason=no_entries path=%s", source_path)
            return False
        logger.info("contents_rebuild_entries_collected path=%s count=%d", source_path, len(entries))

        _add_body_bookmarks(document, entries)
        removed_count = _insert_contents_block(
            document,
            entries,
            {entry.bookmark_name: _PAGE_PLACEHOLDER for entry in entries},
        )
        logger.info("contents_rebuild_old_toc_removed path=%s count=%d", source_path, removed_count)
        _reapply_front_matter_layout(document)
        document.save(str(work_path))

        pdf_path = render_docx_to_pdf(work_path)
        lines = analyze_pdf_lines(pdf_path)
        pages = _resolve_display_pages(entries, lines)
        logger.info("contents_rebuild_render_resolved path=%s count=%d", source_path, len(pages))

        document = Document(str(work_path))
        _replace_contents_entry_pages(document, entries, pages)
        _reapply_front_matter_layout(document)
        document.save(str(work_path))

        shutil.copy2(work_path, source_path)
        logger.info("contents_rebuild_applied path=%s entries=%s", source_path, len(entries))
        return True
    except Exception as exc:
        logger.warning("contents_rebuild_skipped reason=%s path=%s", exc, source_path)
        return False
    finally:
        if pdf_path is not None:
            shutil.rmtree(pdf_path.parent, ignore_errors=True)
        shutil.rmtree(workdir, ignore_errors=True)
