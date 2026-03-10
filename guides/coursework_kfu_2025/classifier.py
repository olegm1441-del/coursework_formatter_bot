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

CHAPTER_RE = re.compile(r"^\s*глава\s+(\d+)\s*\.?\s*(.+?)\s*$", re.IGNORECASE)
NORMALIZED_H1_RE = re.compile(r"^\s*(\d+)\.\s*(?::|[-—–])?\s+(.+?)\s*$")
H2_RE = re.compile(r"^\s*(\d+)\.(\d+)\.?\s*(?::|[-—–])?\s+(.+?)\s*$")
BROKEN_H2_RE = re.compile(r"^\s*\.\s+(.+?)\s*$")

TABLE_CAPTION_RE = re.compile(
    r"^\s*таблица\s+\d+(?:\.\d+){0,2}\.?\s*(?:[-—–].*)?$",
    re.IGNORECASE,
)
TABLE_CONTINUATION_RE = re.compile(
    r"^\s*продолжение\s+таблицы",
    re.IGNORECASE,
)
FIGURE_CAPTION_RE = re.compile(
    r"^\s*(рис\.|рисунок)\s*\d+(?:\.\d+){0,2}\s*[.\-—–]?\s+.+$",
    re.IGNORECASE,
)
SOURCE_LINE_RE = re.compile(r"^\s*(источник|составлено по|рассчитано по|примечание)\s*:\s*.+$", re.IGNORECASE)


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


def normalize_heading_candidate(text: str) -> str:
    """Normalize minor visual separators often produced by Word list exports."""
    t = clean_spaces(text)
    # Examples: "1. · Заголовок", "3.1 • Подраздел"
    t = re.sub(r"\s*[·•▪]\s*", " ", t)
    return clean_spaces(t)





def is_intro_heading_line(text: str) -> bool:
    """
    Robust detection of the body-start heading "Введение".

    Accepts minor punctuation variants like:
      - "ВВЕДЕНИЕ"
      - "Введение:"
      - "Введение."

    Does NOT treat TOC lines like "ВВЕДЕНИЕ........3" as body start.
    """
    t = clean_spaces(text).lower()
    if not t:
        return False

    # Remove trailing punctuation/spacing; keep internal text intact.
    t = re.sub(r"[\s\.:;!?]+$", "", t)

    return t == INTRO_HEADING


def find_body_start_index(document):
    for idx, p in enumerate(document.paragraphs):
        if is_intro_heading_line(paragraph_text(p)):
            return idx
    return None


def parse_heading1(text: str):
    t = normalize_heading_candidate(text)
    low = t.lower()

    if low in H1_EXACT:
        return {"kind": "heading1_exact", "chapter_num": None, "title": t}

    m = CHAPTER_RE.match(t)
    if m:
        return {
            "kind": "heading1_chapter",
            "chapter_num": int(m.group(1)),
            "title": clean_spaces(m.group(2)),
        }

    m = NORMALIZED_H1_RE.match(t)
    if m:
        title = clean_spaces(m.group(2))
        if title:
            return {
                "kind": "heading1_chapter",
                "chapter_num": int(m.group(1)),
                "title": title,
            }

    return None


def parse_heading2(text: str):
    t = normalize_heading_candidate(text)
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
        return "table_caption"

    if TABLE_CONTINUATION_RE.match(t):
        return "table_continuation"

    if FIGURE_CAPTION_RE.match(t):
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
