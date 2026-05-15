import re
from .rules import INTRO_HEADING, REFERENCE_SUBHEADINGS


H1_EXACT = {
    "содержание",
    "введение",
    "заключение",
    "список использованных источников",
    "список использованной литературы",
    "приложения",
}

CHAPTER_RE = re.compile(r"^\s*глава\s+(\d+)\s*\.?\s*(.{0,140}?)\s*$", re.IGNORECASE)
NORMALIZED_H1_RE = re.compile(r"^\s*(\d+)\.\s+(.+?)\s*$")
H2_RE = re.compile(r"^\s*(\d+)\.(\d+)\.?\s+(.+?)\s*$")
BROKEN_H2_RE = re.compile(r"^\s*\.\s+(.+?)\s*$")

TABLE_CAPTION_RE = re.compile(
    r"^\s*(таблица|table)\s+\d+(?:\.\d+){0,2}\.?(?:\s*(?:[-—–]\s*)?.+)?\s*$",
    re.IGNORECASE,
)
TABLE_CONTINUATION_RE = re.compile(
    r"^\s*(продолжение\s+таблицы|continuation(?:\s+of)?\s+table)\b",
    re.IGNORECASE,
)
FIGURE_CAPTION_RE = re.compile(
    r"^\s*(рис\.|рисунок|figure|fig\.)\s*\d+(?:\.\d+){0,2}(?:\s*[.\-—–]?\s+.+)?\s*$",
    re.IGNORECASE,
)
SOURCE_LINE_RE = re.compile(r"^\s*(источник|составлено по|рассчитано по|примечание)\s*:\s*.+$", re.IGNORECASE)

# Russian reference-prose verbs that indicate "Таблица N …" / "Рис. N …" is body
# text REFERRING to a table/figure, not a caption title.
# Vocabulary kept in sync with the verb list already trusted by
# safe_formatter._is_figure_caption_text (introduced in commit be38951).
_CAPTION_REFERENCE_PROSE_RE = re.compile(
    r"^(показыва|отража|содерж|представлен|представля|"
    r"демонстрир|иллюстрир|свидетельств)\w*\b",
    re.IGNORECASE,
)


def caption_tail_is_reference_prose(tail: str) -> bool:
    """
    Return True when the text following a 'Таблица N' / 'Рис. N' prefix starts
    with a Russian reference verb (e.g. 'показывает', 'демонстрирует'). Such
    paragraphs are body prose referring to the table/figure, not caption titles.
    """
    if not tail:
        return False
    t = clean_spaces(tail).lstrip(".:—–-").lstrip()
    if not t:
        return False
    return bool(_CAPTION_REFERENCE_PROSE_RE.match(t))


def clean_spaces(text: str) -> str:
    if text is None:
        return ""
    text = text.replace("\u00A0", " ")
    text = text.replace("\u2007", " ")
    text = text.replace("\u202F", " ")
    text = text.replace("\t", " ")
    text = re.sub(r"[ ]{2,}", " ", text)
    text = re.sub(r"\s+([,.;:!?])", r"\1", text)
    return text.strip()


def paragraph_text(paragraph) -> str:
    return clean_spaces(paragraph.text)


def is_table_continuation_line(text: str) -> bool:
    t = clean_spaces(text)
    if not t:
        return False

    if len(t) > 100:
        return False

    return bool(TABLE_CONTINUATION_RE.match(t))








def is_probable_numbered_heading1_title(title: str) -> bool:
    t = clean_spaces(title)
    if not t:
        return False

    if len(t) > 140:
        return False

    if t.endswith((".", ":", ";", "!", "?")):
        return False

    if re.search(r"\.{2,}", t):
        return False

    if re.search(r"\d{1,4}\s*$", t):
        return False

    if t.count(".") >= 2:
        return False

    if re.search(r"\(\d{1,3}\)\s*$", t):
        return False

    if " - " in t or " — " in t or " – " in t:
        return False

    if TABLE_CAPTION_RE.match(t) or is_table_continuation_line(t) or FIGURE_CAPTION_RE.match(t):
        return False

    return True

def is_intro_heading_text(text: str) -> bool:
    t = clean_spaces(text).lower()
    t = t.rstrip(".")
    return t == INTRO_HEADING


def find_body_start_index(document):
    for idx, p in enumerate(document.paragraphs):
        if is_intro_heading_text(paragraph_text(p)):
            return idx
    return None


def parse_heading1(text: str):
    t = clean_spaces(text)
    low = t.lower()
    low_exact = low.rstrip(".")

    if low_exact in H1_EXACT:
        canonical = low_exact.upper()
        return {"kind": "heading1_exact", "chapter_num": None, "title": canonical}

    m = CHAPTER_RE.match(t)
    if m:
        raw_title = clean_spaces(m.group(2))
        # Empty title ("Глава 1") is valid — the chapter name is on the next paragraph.
        # Non-empty title must pass the heading-title sanity check.
        if not raw_title or is_probable_numbered_heading1_title(raw_title):
            return {
                "kind": "heading1_chapter",
                "chapter_num": int(m.group(1)),
                "title": raw_title,
            }

    m = NORMALIZED_H1_RE.match(t)
    if m:
        title = clean_spaces(m.group(2))
        if title and is_probable_numbered_heading1_title(title):
            return {
                "kind": "heading1_chapter",
                "chapter_num": int(m.group(1)),
                "title": title,
            }

    return None


def parse_heading2(text: str):
    t = clean_spaces(text)
    m = H2_RE.match(t)
    if not m:
        return None

    return {
        "chapter_num": int(m.group(1)),
        "paragraph_num": int(m.group(2)),
        "title": clean_spaces(m.group(3)),
    }


def parse_broken_heading2(text: str):
    t = clean_spaces(text)
    m = BROKEN_H2_RE.match(t)
    if not m:
        return None

    title = clean_spaces(m.group(1))
    if not title:
        return None

    return {"title": title}


def classify_paragraph(text: str, prev_kind=None) -> str:
    t = clean_spaces(text)
    if not t:
        return "empty_paragraph"

    low = t.lower()

    if low in REFERENCE_SUBHEADINGS:
        return "reference_subheading"

    if TABLE_CAPTION_RE.match(t):
        prefix = re.match(r"^\s*(таблица|table)\s+\d+(?:\.\d+){0,2}\.?", t, re.IGNORECASE)
        tail = t[prefix.end():] if prefix else ""
        if not caption_tail_is_reference_prose(tail):
            return "table_caption"

    if is_table_continuation_line(t):
        return "table_continuation"

    if FIGURE_CAPTION_RE.match(t):
        prefix = re.match(r"^\s*(рис\.|рисунок|figure|fig\.)\s*\d+(?:\.\d+){0,2}", t, re.IGNORECASE)
        tail = t[prefix.end():] if prefix else ""
        if not caption_tail_is_reference_prose(tail):
            return "figure_caption"

    if SOURCE_LINE_RE.match(t):
        return "source_line"

    parsed_h1 = parse_heading1(t)
    if parsed_h1:
        if parsed_h1["kind"] == "heading1_exact" and low == "содержание":
            return "toc_heading"
        return "heading1"

    if parse_heading2(t):
        return "heading2"

    if parse_broken_heading2(t):
        return "broken_heading2"

    if prev_kind in {"table_caption", "table_continuation"}:
        return "table_title"

    return "body_text"
