from docx.oxml.ns import qn

from .classifier import clean_spaces, parse_heading1, parse_heading2


EXACT_PAGEBREAK_HEADINGS = {
    "введение",
    "заключение",
    "список использованных источников",
    "список использованной литературы",
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
    - новой главой (heading1 chapter)
    - ЗАКЛЮЧЕНИЕ
    - СПИСКОМ ИСТОЧНИКОВ
    - ПРИЛОЖЕНИЯМИ

    Не ставит разрыв страницы перед heading2.
    """
    _cleanup_existing_page_break_artifacts(document, body_start)

    for idx, paragraph in enumerate(document.paragraphs):
        if idx < body_start:
            continue

        text = clean_spaces(paragraph.text)
        if not _needs_page_break_before(text):
            continue

        paragraph.paragraph_format.page_break_before = True
