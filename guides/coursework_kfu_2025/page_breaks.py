import re

from docx.oxml.ns import qn

from .classifier import clean_spaces, parse_heading1, parse_heading2


_CHAPTER_PREFIX_RE = re.compile(r"^\d+\.\s+\S")


def _is_styled_heading1_chapter(paragraph, text: str) -> bool:
    """A paragraph already promoted to Heading 1 that starts with `N.` is a chapter
    by construction — even when its title is too messy for parse_heading1 (e.g. a
    spawned `1. … Типа 1.Чето там …`). Such chapters still need a page break."""
    try:
        style = (paragraph.style.name or "").strip().lower()
    except Exception:
        style = ""
    if style not in {"heading 1", "заголовок 1"}:
        return False
    return bool(_CHAPTER_PREFIX_RE.match(clean_spaces(text)))


EXACT_PAGEBREAK_HEADINGS = {
    "заключение",
    "список использованных источников",
    "список использованной литературы",
    "приложения",
    "приложение",
}

REFERENCES_HEADINGS = {
    "список использованных источников",
    "список использованной литературы",
}

APPENDIX_HEADINGS = {
    "приложения",
    "приложение",
}


def _remove_page_breaks_from_run(run):
    """
    Удаляет только page-break'и из run, не трогая обычные переносы строк.
    """
    r = run._element
    for br in list(r.findall(qn("w:br"))):
        br_type = br.get(qn("w:type"))
        if br_type == "page":
            r.remove(br)


def _cleanup_existing_page_break_artifacts(document, body_start):
    """
    Чистим последствия старой версии page_breaks:
    - убираем явные page-break элементы из runs;
    - сбрасываем page_break_before у всех абзацев рабочей части.
    """
    for idx, p in enumerate(document.paragraphs):
        if idx < body_start:
            continue

        p.paragraph_format.page_break_before = False

        for run in p.runs:
            _remove_page_breaks_from_run(run)


_APPENDIX_START_LABEL_RE = re.compile(r"^приложение\s*\S+$", re.IGNORECASE)


def _is_appendix_start_label(text: str) -> bool:
    """A standalone appendix label that begins a NEW appendix, e.g. ``ПРИЛОЖЕНИЕ
    Б`` / ``Приложение 2``. Excludes the plural section heading ``ПРИЛОЖЕНИЯ``
    and body phrases like ``приложение к договору ...`` (more than one token)."""
    t = clean_spaces(text)
    low = t.lower()
    if low in {"приложения", "приложение"}:
        return False
    return bool(_APPENDIX_START_LABEL_RE.match(t))


def _needs_page_break_before(text: str) -> bool:
    t = clean_spaces(text)
    low = t.lower()

    if not t:
        return False

    if low in EXACT_PAGEBREAK_HEADINGS:
        return True

    # Every appendix START label begins a new appendix and must start on a new
    # page — not only the first one after the references block. Without this,
    # a second appendix (ПРИЛОЖЕНИЕ Б) after appendix-A content stays mid-page.
    if _is_appendix_start_label(t):
        return True

    parsed_h1 = parse_heading1(t)
    if parsed_h1 and parsed_h1["kind"] == "heading1_chapter":
        return True

    # ВАЖНО: перед heading2 разрыв страницы НЕ нужен
    if parse_heading2(t):
        return False

    return False


def apply_page_breaks(document, body_start):
    """
    Ставит page_break_before только перед:
    - ВВЕДЕНИЕ
    - новой главой
    - ЗАКЛЮЧЕНИЕ
    - СПИСКОМ ИСТОЧНИКОВ
    - ПРИЛОЖЕНИЯМИ

    ВНУТРИ списка источников разрывы страниц не ставит.
    """
    _cleanup_existing_page_break_artifacts(document, body_start)

    in_references = False

    for idx, paragraph in enumerate(document.paragraphs):
        if idx < body_start:
            continue

        text = clean_spaces(paragraph.text)
        low = text.lower()

        # Начало блока литературы
        if low in REFERENCES_HEADINGS:
            in_references = True
            paragraph.paragraph_format.page_break_before = True
            continue

        # Конец блока литературы
        if in_references and low in APPENDIX_HEADINGS:
            in_references = False
            paragraph.paragraph_format.page_break_before = True
            continue

        # Внутри списка литературы НИЧЕГО не разрываем
        if in_references:
            paragraph.paragraph_format.page_break_before = False
            continue

        if _needs_page_break_before(text) or _is_styled_heading1_chapter(paragraph, text):
            paragraph.paragraph_format.page_break_before = True
        else:
            paragraph.paragraph_format.page_break_before = False
