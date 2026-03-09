from docx.oxml.ns import qn

from .classifier import clean_spaces, parse_heading1, parse_heading2


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
        if br_type in (None, "page"):
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


def _needs_page_break_before(text: str) -> bool:
    t = clean_spaces(text)
    low = t.lower()

    if not t:
        return False

    if low in EXACT_PAGEBREAK_HEADINGS:
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

        if _needs_page_break_before(text):
            paragraph.paragraph_format.page_break_before = True
        else:
            paragraph.paragraph_format.page_break_before = False
