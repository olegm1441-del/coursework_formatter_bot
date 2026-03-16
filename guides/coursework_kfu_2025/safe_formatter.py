FORMULA_NUMBER_RE = re.compile(r"\((\d+\.\d+\.\d+|\d+\.\d+)\)\s*$")
FORMULA_EXPLANATION_RE = re.compile(r"^\s*где\b", re.IGNORECASE)

MATH_TOKEN_RE = re.compile(r"[=+\-*/×÷^(){}\[\]<>]|[A-Za-zА-Яа-яЁё]\s*=")


def is_formula_paragraph_text(text: str) -> bool:
    t = clean_spaces(text)
    if not t:
        return False

    if not FORMULA_NUMBER_RE.search(t):
        return False

    # До номера должен быть не обычный текст, а выражение
    left = FORMULA_NUMBER_RE.sub("", t).strip()
    if len(left) > 120:
        return False

    # Формула должна содержать математический маркер
    return bool(MATH_TOKEN_RE.search(left))


def is_formula_explanation_start(text: str) -> bool:
    return bool(FORMULA_EXPLANATION_RE.match(clean_spaces(text)))


def is_formula_explanation_continuation(text: str) -> bool:
    t = clean_spaces(text)
    if not t:
        return False
    if is_formula_explanation_start(t):
        return True
    # строка расшифровки символов: "V - ...", "R – ..."
    return bool(re.match(r"^[A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё0-9]*\s*[-–—=]\s*.+$", t))

def format_formula_paragraph(paragraph):
    text = clean_spaces(paragraph.text)
    m = FORMULA_NUMBER_RE.search(text)
    if not m:
        return

    number = m.group(0)
    expr = text[:m.start()].rstrip()

    if expr and not expr.endswith(","):
        expr = expr + ","

    replace_paragraph_text(paragraph, f"{expr}\t{number}")

    hard_reset_paragraph_format(paragraph, first_line_indent_cm=None)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.first_line_indent = Cm(0)

    tabs = paragraph.paragraph_format.tab_stops
    tabs.clear_all()
    tabs.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT)

    for run in paragraph.runs:
        set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=False, italic=False)

def format_formula_explanation_paragraph(paragraph, is_first=False):
    hard_reset_paragraph_format(paragraph, first_line_indent_cm=None)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.first_line_indent = Cm(0)

    text = clean_spaces(paragraph.text)

    if is_first:
        text = re.sub(r"^\s*где\s*:\s*", "где ", text, flags=re.IGNORECASE)
        text = re.sub(r"^\s*где\s+", "где ", text, flags=re.IGNORECASE)
        replace_paragraph_text(paragraph, text)

    for run in paragraph.runs:
        set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=False, italic=False)

def normalize_formula_blocks(document, body_start):
    changed = False
    paragraphs = document.paragraphs
    idx = max(body_start, 0)

    while idx < len(paragraphs):
        p = paragraphs[idx]
        text = clean_spaces(p.text)

        if not is_formula_paragraph_text(text):
            idx += 1
            continue

        # 1. Форматируем строку формулы
        format_formula_paragraph(p)

        # 2. Перед формулой должна быть ровно одна пустая строка
        if idx > body_start:
            prev_idx = idx - 1
            if not is_empty_paragraph(paragraphs[prev_idx]):
                new_p = insert_paragraph_after(paragraphs[prev_idx], "")
                format_empty_paragraph(new_p)
                changed = True
                paragraphs = document.paragraphs
                idx += 1
                p = paragraphs[idx]
            else:
                while prev_idx - 1 >= body_start and is_empty_paragraph(paragraphs[prev_idx - 1]):
                    remove_paragraph(paragraphs[prev_idx - 1])
                    changed = True
                    paragraphs = document.paragraphs
                    idx -= 1
                    prev_idx -= 1
                format_empty_paragraph(paragraphs[prev_idx])

        # 3. Форматируем блок "где ..."
        j = idx + 1
        first_expl = True
        while j < len(paragraphs):
            t = clean_spaces(paragraphs[j].text)

            if not t:
                break

            if first_expl and is_formula_explanation_start(t):
                format_formula_explanation_paragraph(paragraphs[j], is_first=True)
                first_expl = False
                j += 1
                continue

            if not first_expl and is_formula_explanation_continuation(t):
                format_formula_explanation_paragraph(paragraphs[j], is_first=False)
                j += 1
                continue

            break

        # 4. После формулы/пояснений должна быть ровно одна пустая строка
        tail_idx = j - 1 if j > idx + 1 else idx
        paragraphs = document.paragraphs

        if tail_idx + 1 >= len(paragraphs):
            new_p = insert_paragraph_after(paragraphs[tail_idx], "")
            format_empty_paragraph(new_p)
            changed = True
            paragraphs = document.paragraphs
        elif not is_empty_paragraph(paragraphs[tail_idx + 1]):
            new_p = insert_paragraph_after(paragraphs[tail_idx], "")
            format_empty_paragraph(new_p)
            changed = True
            paragraphs = document.paragraphs
        else:
            format_empty_paragraph(paragraphs[tail_idx + 1])
            while tail_idx + 2 < len(paragraphs) and is_empty_paragraph(paragraphs[tail_idx + 2]):
                remove_paragraph(paragraphs[tail_idx + 2])
                changed = True
                paragraphs = document.paragraphs

        idx = tail_idx + 2
        paragraphs = document.paragraphs

    return changed
    
def clear_paragraph_outline_level(paragraph):
    try:
        pPr = paragraph._element.get_or_add_pPr()
        outline = pPr.find(qn("w:outlineLvl"))
        if outline is not None:
            pPr.remove(outline)
    except Exception:
        pass



def set_paragraph_style_safe(paragraph, *style_names):
    for name in style_names:
        try:
            paragraph.style = name
            return True
        except Exception:
            pass
    return False




# ===== STRUCTURAL SPACING FIX =====

STRUCTURAL_HEADINGS = {
    "ВВЕДЕНИЕ",
    "ЗАКЛЮЧЕНИЕ",
}

def enforce_structural_spacing(doc):

    paragraphs = doc.paragraphs
    i = 0

    while i < len(paragraphs):

        p = paragraphs[i]
        text = p.text.strip().upper()

        if text in STRUCTURAL_HEADINGS:

            j = i + 1
            blank_count = 0

            while j < len(paragraphs) and not paragraphs[j].text.strip():
                blank_count += 1
                j += 1

            if blank_count == 0:
                new = insert_paragraph_after(p, "")
                new.paragraph_format.space_before = 0
                new.paragraph_format.space_after = 0
                i += 2
                paragraphs = doc.paragraphs
                continue

            if blank_count > 1:
                for k in range(i + 2, i + 1 + blank_count):
                    remove_paragraph(paragraphs[i + 2])

        i += 1

# ===== END STRUCTURAL SPACING FIX =====




# ===== AUTO PATCH: robust heading2 detection =====

def auto_detect_heading2(paragraph, current_chapter_num, next_paragraph_num, prev_kind=None):
    if current_chapter_num is None or next_paragraph_num is None:
        return False

    text = clean_spaces(paragraph.text)
    if not text:
        return False

    low = text.lower()

    forbidden_prefixes = (
        "таблица ",
        "рисунок ",
        "рис. ",
        "продолжение таблицы",
        "продолжение табл.",
        "источник:",
        "составлено по:",
        "рассчитано по:",
        "примечание:",
    )
    if low.startswith(forbidden_prefixes):
        return False

    if parse_heading1(text) or parse_heading2(text) or parse_broken_heading2(text):
        return False

    if is_table_continuation_text(text):
        return False

    if not looks_like_heading2_title(text):
        return False

    if paragraph_has_numbering(paragraph):
        return True

    if is_probable_center_bold_heading(paragraph):
        return True

    # Частый кейс: сразу после главы идёт подпункт, но Word потерял номер
    if prev_kind in {"heading1", "empty_paragraph"}:
        return True

    return False

# ===== END PATCH =====

# ===== FINAL PATCH: enforce spacing after structural headings =====

STRUCTURAL_HEADING_TEXTS = {
    "ВВЕДЕНИЕ",
    "ЗАКЛЮЧЕНИЕ",
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
}

def is_structural_heading_paragraph(paragraph):
    t = clean_spaces(paragraph.text).upper()
    return t in STRUCTURAL_HEADING_TEXTS

def enforce_single_blank_after_structural_headings(doc, body_start_idx=0):
    paragraphs = doc.paragraphs
    i = max(body_start_idx, 0)

    while i < len(paragraphs):
        p = paragraphs[i]
        if not is_structural_heading_paragraph(p):
            i += 1
            continue

        j = i + 1
        blank_idxs = []

        while j < len(paragraphs) and not clean_spaces(paragraphs[j].text):
            blank_idxs.append(j)
            j += 1

        if not blank_idxs:
            insert_paragraph_after(p, "")
            paragraphs = doc.paragraphs
            i += 2
            continue

        first_blank_idx = blank_idxs[0]
        for idx in reversed(blank_idxs[1:]):
            remove_paragraph(paragraphs[idx])

        paragraphs = doc.paragraphs
        format_empty_spacing_paragraph(paragraphs[first_blank_idx])
        i = first_blank_idx + 1

# ===== END FINAL PATCH =====



from pathlib import Path
import re
from copy import deepcopy
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, Mm, RGBColor, RGBColor
from docx.text.paragraph import Paragraph

from .rules import (
    FONT_NAME,
    BODY_FONT_SIZE_PT,
    TABLE_FONT_SIZE_PT,
    LINE_SPACING_BODY,
    LINE_SPACING_TABLE,
    FIRST_LINE_INDENT_CM,
    LEFT_MARGIN_MM,
    RIGHT_MARGIN_MM,
    TOP_MARGIN_MM,
    BOTTOM_MARGIN_MM,
)
from .classifier import (
    find_body_start_index,
    classify_paragraph,
    clean_spaces,
    parse_heading1,
    parse_heading2,
    parse_broken_heading2,
)
from .page_numbering import apply_page_numbering_policy
from .page_breaks import apply_page_breaks

MAX_NORMALIZATION_PASSES = 35

def run_with_pass_limit(step_name, func, document, body_start):
    """
    Re-run a normalization step until it stabilizes, but avoid full-text snapshots
    of the whole document on every pass.

    Preferred contract: a step may return:
      - True / positive int  -> document changed, run another pass
      - False / 0 / None     -> no changes, step is stable

    Backward compatibility: if a step returns None, we fall back to a cheap
    structural signature based on paragraph count and lengths.
    """
    previous_signature = None

    for _ in range(MAX_NORMALIZATION_PASSES):
        paragraphs = document.paragraphs
        before_signature = (
            len(paragraphs),
            sum(len(p.text) for p in paragraphs),
        )

        result = func(document, body_start)

        if isinstance(result, bool):
            if not result:
                return
            previous_signature = None
            continue

        if isinstance(result, int):
            if result <= 0:
                return
            previous_signature = None
            continue

        paragraphs_after = document.paragraphs
        after_signature = (
            len(paragraphs_after),
            sum(len(p.text) for p in paragraphs_after),
        )

        if after_signature == before_signature:
            return

        if after_signature == previous_signature:
            raise RuntimeError(f"Formatter step stuck: {step_name}")

        previous_signature = after_signature

    raise RuntimeError(f"Formatter step stuck: {step_name}")


TABLE_NUM_RE = re.compile(r"^\s*таблица\s*(\d+(?:\.\d+){0,2})\.?\s*(.*?)\s*$", re.IGNORECASE)
DASH_LINE_RE = re.compile(r"^\s*[—–\-•]\s*.+$")
FIG_RE = re.compile(
    r"^\s*(рисунок|рис\.)\s*(\d+(?:\.\d+){0,2})(?:\s*[.\-—–]?\s*(.+?))?\s*$",
    re.IGNORECASE,
)
HEADING2_ARTIFACT_RE = re.compile(r"^\s*[•·▪■◆►→\-–—]*\s*(\d+\.\d+\.?)\s*[•·▪■◆►→\-–—]*\s*(.+?)\s*$")

TABLE_CONTINUATION_RE = re.compile(r"^\s*продолжение\s+табл(?:ицы)?\.?\s*\d+(?:\.\d+){1,2}\.?\s*$", re.IGNORECASE)

def is_table_continuation_text(text: str) -> bool:
    t = clean_spaces(text)
    if not t:
        return False

    if TABLE_CONTINUATION_RE.match(t):
        return True

    # Защита от уже испорченных вариантов: "1.2. Продолжение таблицы 1.1.1"
    t2 = re.sub(r'^\s*\d+\.\d+\.?\s*', '', t, count=1)
    return bool(TABLE_CONTINUATION_RE.match(t2))



REFERENCE_SUBHEADINGS_CANON = {
    "официальные материалы": "Официальные материалы",
    "статистические материалы": "Статистические материалы",
    "справочные и архивные материалы": "Справочные и архивные материалы",
    "монографии и статьи": "Монографии и статьи",
    "учебники, учебные пособия и материалы": "Учебники, учебные пособия и материалы",
    "электронные ресурсы": "Электронные ресурсы",
    "материалы на иностранных языках": "Материалы на иностранных языках",
}


def insert_paragraph_after(paragraph, text=""):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    return new_para


def remove_paragraph(paragraph):
    p = paragraph._element
    parent = p.getparent()
    if parent is not None:
        parent.remove(p)


def replace_paragraph_text(paragraph, new_text: str):
    p = paragraph._element
    for child in list(p):
        if child.tag.endswith("}r") or child.tag.endswith("}hyperlink"):
            p.remove(child)
    paragraph.add_run(new_text)


def paragraph_has_drawing(paragraph) -> bool:
    p = paragraph._element
    return bool(
        p.xpath(
            ".//*[local-name()='drawing' or local-name()='pict' or local-name()='object']"
        )
    )


def is_empty_paragraph(paragraph):
    return clean_spaces(paragraph.text) == "" and not paragraph_has_drawing(paragraph)

def paragraph_has_drawing(paragraph) -> bool:
    p = paragraph._element
    return bool(
        p.xpath(
            ".//*[local-name()='drawing' or local-name()='pict' or local-name()='object']"
        )
    )

def ensure_empty_run(paragraph):
    if not paragraph.runs:
        paragraph.add_run("")
    return paragraph.runs[0]


def force_paragraph_xml_spacing(paragraph, line_rule="auto"):
    pPr = paragraph._element.get_or_add_pPr()

    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)

    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:beforeAutospacing"), "0")
    spacing.set(qn("w:afterAutospacing"), "0")

    if line_rule == "auto":
        spacing.set(qn("w:lineRule"), "auto")
        spacing.set(qn("w:line"), "360")
    elif line_rule == "exact":
        spacing.set(qn("w:lineRule"), "exact")
    elif line_rule == "atLeast":
        spacing.set(qn("w:lineRule"), "atLeast")

    snap = pPr.find(qn("w:snapToGrid"))
    if snap is None:
        snap = OxmlElement("w:snapToGrid")
        pPr.append(snap)
    snap.set(qn("w:val"), "0")


def hard_reset_paragraph_format(paragraph, first_line_indent_cm=None):
    force_paragraph_xml_spacing(paragraph, line_rule="auto")
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = LINE_SPACING_BODY
    fmt.left_indent = Cm(0)
    fmt.right_indent = Cm(0)

    if first_line_indent_cm is None:
        fmt.first_line_indent = Cm(0)
    else:
        fmt.first_line_indent = Cm(first_line_indent_cm)

    fmt.keep_together = False
    fmt.keep_with_next = False
    fmt.page_break_before = False
    fmt.widow_control = False


def set_run_font(run, font_name=FONT_NAME, size_pt=BODY_FONT_SIZE_PT, bold=None, italic=False, all_caps=None):
    run.font.name = font_name
    run.font.size = Pt(size_pt)

    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.rFonts
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)

    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)

    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if all_caps is not None:
        run.font.all_caps = all_caps

    try:
        run.font.color.rgb = RGBColor(0, 0, 0)
    except Exception:
        pass

    color = rPr.find(qn("w:color"))
    if color is None:
        color = OxmlElement("w:color")
        rPr.append(color)

    color.set(qn("w:val"), "000000")

    for attr in ("w:themeColor", "w:themeTint", "w:themeShade"):
        try:
            if color.get(qn(attr)) is not None:
                del color.attrib[qn(attr)]
        except Exception:
            pass

def set_section_margins(document):
    # Глобально отключаем зеркальные поля на уровне settings.xml
    try:
        settings_el = document.settings._element
        mirror = settings_el.find(qn("w:mirrorMargins"))
        if mirror is not None:
            settings_el.remove(mirror)
    except Exception:
        pass

    for section in document.sections:
        section.left_margin = Mm(LEFT_MARGIN_MM)
        section.right_margin = Mm(RIGHT_MARGIN_MM)
        section.top_margin = Mm(TOP_MARGIN_MM)
        section.bottom_margin = Mm(BOTTOM_MARGIN_MM)

        # На уровне секции убираем gutter/переплёт и следы зеркалинга
        try:
            sectPr = section._sectPr

            pgMar = sectPr.find(qn("w:pgMar"))
            if pgMar is not None:
                pgMar.set(qn("w:left"), str(Mm(LEFT_MARGIN_MM)._emu))
                pgMar.set(qn("w:right"), str(Mm(RIGHT_MARGIN_MM)._emu))
                pgMar.set(qn("w:top"), str(Mm(TOP_MARGIN_MM)._emu))
                pgMar.set(qn("w:bottom"), str(Mm(BOTTOM_MARGIN_MM)._emu))
                pgMar.set(qn("w:gutter"), "0")

            gutter = sectPr.find(qn("w:gutter"))
            if gutter is not None:
                sectPr.remove(gutter)
        except Exception:
            pass
def normalize_simple_paragraph_spaces(paragraph):
    if len(paragraph.runs) == 1 and "\n" not in paragraph.runs[0].text and "\v" not in paragraph.runs[0].text:
        old = paragraph.runs[0].text
        new = clean_spaces(old)
        if new != old:
            paragraph.runs[0].text = new

QUOTE_CHARS_DOUBLE = {
    '"',      # ASCII
    '“', '”', # curly double
    '„', '‟', # low/high double
    '«', '»', # уже правильные, но учитываем в общем потоке
    '″', '‟', '〝', '〞', '＂',
}

def _normalize_quotes_in_text_fragment(text: str, quote_state: dict) -> str:
    """
    Меняет все двойные кавычки на «» по принципу открытия/закрытия.
    Не трогает одинарные апострофы и штрихи — это сознательное ограничение
    ради безопасности обычных курсовых.
    """
    if not text:
        return text

    out = []
    for ch in text:
        if ch in QUOTE_CHARS_DOUBLE:
            if quote_state["open"]:
                out.append("«")
            else:
                out.append("»")
            quote_state["open"] = not quote_state["open"]
        else:
            out.append(ch)
    return "".join(out)


def normalize_quotes_in_paragraph_runs(paragraph, quote_state: dict):
    """
    Нормализует кавычки в run-ах абзаца без пересборки абзаца,
    чтобы не ломать гиперссылки, разметку и прочую структуру Word.
    """
    for run in paragraph.runs:
        old = run.text
        new = _normalize_quotes_in_text_fragment(old, quote_state)
        if new != old:
            run.text = new


def normalize_quotes_in_document(document, body_start=0):
    """
    Проходит по рабочей части документа последовательно сверху вниз.
    Состояние открытия/закрытия кавычек сохраняется между абзацами.
    """
    quote_state = {"open": True}

    for idx, paragraph in enumerate(document.paragraphs):
        if idx < body_start:
            continue
        normalize_quotes_in_paragraph_runs(paragraph, quote_state)

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    normalize_quotes_in_paragraph_runs(paragraph, quote_state)

def canonical_reference_subheading_text(text: str):
    t = clean_spaces(text)
    if not t:
        return None

    t = re.sub(r'^\s*[•·▪■◆►→\-–—]+\s*', '', t)
    t = re.sub(r'^\s*\d+\.\s*', '', t)
    t = clean_spaces(t)

    return REFERENCE_SUBHEADINGS_CANON.get(t.lower())



# ===== Reference list case normalization =====
_REF_URL_RE = re.compile(r'https?://\S+', re.IGNORECASE)
_REF_TOKEN_RE = re.compile(r'([A-Za-zА-ЯЁа-яё]+(?:[-–—][A-Za-zА-ЯЁа-яё]+)*)')
_REF_ACRONYM_KEEP = {
    'ФНС', 'РФ', 'РБК', 'ТТС', 'ЭДО', 'СЭД', 'СМК', 'ГОСТ', 'ИСО', 'ЕС', 'АО', 'ООО', 'ПАО',
    'ISO', 'IEC', 'IEEE', 'OECD', 'EU', 'USA', 'UK', 'UN', 'PDF', 'HTML', 'URL', 'DOI', 'ISBN',
    'CRM', 'ERP', 'API', 'XML', 'JSON', 'UPD', 'B2B', 'B2G', 'B2C', 'ID', 'IT', 'AI', 'FTS',
}
_REF_CANONICAL_TOKEN_MAP = {
    'EIDAS': 'eIDAS',
    'BUSINESSTAT': 'BusinesStat',
    'CONSULTANTPLUS': 'КонсультантПлюс',
    'КОНСУЛЬТАНТПЛЮС': 'КонсультантПлюс',
}

def _looks_like_shouting_reference(text: str) -> bool:
    letters = [ch for ch in text if ch.isalpha()]
    if len(letters) < 12:
        return False
    uppers = sum(1 for ch in letters if ch.isupper())
    return (uppers / len(letters)) >= 0.65

def _normalize_reference_token(token: str) -> str:
    if not token:
        return token

    upper = token.upper()
    if upper in _REF_CANONICAL_TOKEN_MAP:
        return _REF_CANONICAL_TOKEN_MAP[upper]

    # Сохраняем общеупотребимые аббревиатуры и короткие токены с цифрами
    if upper in _REF_ACRONYM_KEEP:
        return upper
    if any(ch.isdigit() for ch in token):
        return token
    if len(token) <= 3 and token.isupper():
        return upper

    # Полностью верхний регистр -> нормальный Title Case
    if token.isupper():
        if '-' in token or '–' in token or '—' in token:
            parts = re.split(r'([-–—])', token)
            return ''.join(_normalize_reference_token(part) if part not in '-–—' else part for part in parts)
        low = token.lower()
        return low[:1].upper() + low[1:]

    return token

def _normalize_reference_case_fragment(fragment: str) -> str:
    return _REF_TOKEN_RE.sub(lambda m: _normalize_reference_token(m.group(0)), fragment)

def smart_normalize_reference_line_case(text: str) -> str:
    clean = clean_spaces(text)
    if not clean:
        return clean

    m = re.match(r'^(\d+\.\s+)(.+)$', clean)
    prefix = ''
    body = clean
    if m:
        prefix, body = m.group(1), m.group(2)

    if not _looks_like_shouting_reference(body):
        return clean

    urls = []
    def _url_repl(match):
        urls.append(match.group(0).lower())
        return f'__REFURL{len(urls)-1}__'

    body = _REF_URL_RE.sub(_url_repl, body)
    body = _normalize_reference_case_fragment(body)

    for i, url in enumerate(urls):
        body = body.replace(f'__REFURL{i}__', url)

    return f'{prefix}{body}' if prefix else body

def strip_leading_heading_garbage(text: str) -> str:
    t = clean_spaces(text)
    if not t:
        return t

    # Убираем ведущие маркеры/мусор, которые Word мог оставить как текст
    t = re.sub(r'^\s*[•·▪■◆►→◦●○\-–—]+\s*', '', t)

    # Убираем лишние пробелы после очистки
    t = clean_spaces(t)
    return t

def auto_detect_numbered_heading1(paragraph, current_chapter_num=None, next_paragraph=None):
    text = clean_spaces(paragraph.text)
    if not text:
        return False

    low = text.lower()

    # Уже распознанный heading1 не трогаем
    if parse_heading1(text):
        return False

    # Не трогаем подписи таблиц/рисунков и служебные строки
    forbidden_prefixes = (
        "таблица",
        "табл.",
        "рисунок",
        "рис.",
        "источник:",
        "составлено по:",
        "рассчитано по:",
        "примечание:",
        "продолжение таблицы",
        "продолжение табл.",
    )
    if low.startswith(forbidden_prefixes):
        return False

    # Нужна именно Word-автонумерация / numbering
    if not paragraph_has_numbering(paragraph):
        return False

    # Если это уже похоже на heading2, не считаем heading1
    if parse_heading2(text) or parse_broken_heading2(text):
        return False

    # Запрещённые финальные знаки
    if text.endswith((":", ";", "?", "!")):
        return False

    words = text.split()
    word_limit = 12 if "." in text else 15
    if len(words) < 1 or len(words) > word_limit:
        return False

    # Если следующий абзац тоже numbered и тоже короткий,
    # это больше похоже на список, а не на heading1
    if next_paragraph is not None:
        next_text = clean_spaces(next_paragraph.text)
        if next_text and paragraph_has_numbering(next_paragraph):
            if not parse_heading1(next_text) and not parse_heading2(next_text):
                next_words = next_text.split()
                next_limit = 12 if "." in next_text else 15
                if 1 <= len(next_words) <= next_limit and not next_text.endswith((":", ";", "?", "!")):
                    return False

    return True
    
def normalize_heading2_artifacts(paragraph):
    text = clean_spaces(paragraph.text)
    if not text:
        return

    m = HEADING2_ARTIFACT_RE.match(text)
    if not m:
        return

    num = m.group(1)
    title = clean_spaces(m.group(2))
    if not parse_heading2(f"{num} {title}"):
        return

    if not num.endswith("."):
        num += "."
    replace_paragraph_text(paragraph, f"{num} {title}")


def is_probable_center_bold_heading(paragraph):
    if paragraph.alignment != WD_ALIGN_PARAGRAPH.CENTER:
        return False
    if not paragraph.runs:
        return False

    non_empty_runs = [r for r in paragraph.runs if r.text and r.text.strip()]
    if not non_empty_runs:
        return False

    bold_runs = sum(1 for r in non_empty_runs if r.bold)
    return bold_runs >= max(1, len(non_empty_runs) // 2)


def paragraph_has_numbering(paragraph):
    pPr = paragraph._element.pPr
    if pPr is None:
        return False
    return pPr.find(qn("w:numPr")) is not None


def remove_paragraph_numbering(paragraph):
    pPr = paragraph._element.get_or_add_pPr()
    numPr = pPr.find(qn("w:numPr"))
    if numPr is not None:
        pPr.remove(numPr)

def remove_page_break_artifacts_from_paragraph(paragraph):
    paragraph.paragraph_format.page_break_before = False
    paragraph.paragraph_format.keep_with_next = False
    paragraph.paragraph_format.keep_together = False
    paragraph.paragraph_format.widow_control = False

    for run in paragraph.runs:
        r = run._element

        # Удаляем явные разрывы страницы внутри runs
        for br in list(r.findall(qn("w:br"))):
            br_type = br.get(qn("w:type"))
            if br_type in (None, "page"):
                r.remove(br)

        # На всякий случай убираем lastRenderedPageBreak
        for lrp in list(r.findall(qn("w:lastRenderedPageBreak"))):
            r.remove(lrp)

def is_references_heading_text(text: str) -> bool:
    low = clean_spaces(text).lower()
    return low in {
        "список использованных источников",
        "список использованной литературы",
    }


def is_appendix_heading_text(text: str) -> bool:
    low = clean_spaces(text).lower()
    return low in {"приложение", "приложения"}
            
def format_empty_paragraphs_in_body(document, body_start):
    for idx, paragraph in enumerate(document.paragraphs):
        if idx < body_start:
            continue
        if is_empty_paragraph(paragraph):
            format_empty_paragraph(paragraph)

def format_body(paragraph):
    set_paragraph_style_safe(paragraph, "Normal", "Обычный")
    clear_paragraph_outline_level(paragraph)
    remove_paragraph_numbering(paragraph)

    hard_reset_paragraph_format(paragraph, first_line_indent_cm=FIRST_LINE_INDENT_CM)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for run in paragraph.runs:
        set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=False, all_caps=False)


def format_heading1(paragraph):
    remove_page_break_artifacts_from_paragraph(paragraph)
    remove_paragraph_numbering(paragraph)

    text = clean_spaces(paragraph.text)
    if text:
        replace_paragraph_text(paragraph, text.upper())

    set_paragraph_style_safe(paragraph, "Heading 1", "Заголовок 1")
    clear_paragraph_outline_level(paragraph)
    hard_reset_paragraph_format(paragraph, first_line_indent_cm=None)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for run in paragraph.runs:
        set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=True, italic=False, all_caps=False)

def format_heading2(paragraph):
    remove_page_break_artifacts_from_paragraph(paragraph)
    remove_paragraph_numbering(paragraph)

    set_paragraph_style_safe(paragraph, "Heading 2", "Заголовок 2")
    hard_reset_paragraph_format(paragraph, first_line_indent_cm=None)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for run in paragraph.runs:
        set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=True, all_caps=False)

def format_table_caption(paragraph):
    text = clean_spaces(paragraph.text)
    m = TABLE_NUM_RE.match(text)
    if m:
        number = m.group(1)
        replace_paragraph_text(paragraph, f"Таблица {number}")

    set_paragraph_style_safe(paragraph, "Normal", "Обычный")
    clear_paragraph_outline_level(paragraph)
    remove_paragraph_numbering(paragraph)

    hard_reset_paragraph_format(paragraph, first_line_indent_cm=None)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    for run in paragraph.runs:
        set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=False, italic=False, all_caps=False)


def format_table_title(paragraph):
    set_paragraph_style_safe(paragraph, "Normal", "Обычный")
    clear_paragraph_outline_level(paragraph)
    remove_paragraph_numbering(paragraph)

    hard_reset_paragraph_format(paragraph, first_line_indent_cm=None)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for run in paragraph.runs:
        set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=False, all_caps=False)


def format_source_line(paragraph):
    set_paragraph_style_safe(paragraph, "Normal", "Обычный")
    clear_paragraph_outline_level(paragraph)
    remove_paragraph_numbering(paragraph)

    hard_reset_paragraph_format(paragraph, first_line_indent_cm=FIRST_LINE_INDENT_CM)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for run in paragraph.runs:
        set_run_font(run, size_pt=TABLE_FONT_SIZE_PT, bold=False, italic=False, all_caps=False)

def format_reference_subheading(paragraph):
    # Обязательно делаем обычным абзацем, а не заголовком
    set_paragraph_style_safe(paragraph, "Normal", "Обычный")
    clear_paragraph_outline_level(paragraph)
    remove_paragraph_numbering(paragraph)

    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.first_line_indent = Cm(FIRST_LINE_INDENT_CM)
    paragraph.paragraph_format.left_indent = Cm(0)
    paragraph.paragraph_format.right_indent = Cm(0)
    paragraph.paragraph_format.page_break_before = False
    paragraph.paragraph_format.keep_with_next = False
    paragraph.paragraph_format.keep_together = False
    paragraph.paragraph_format.widow_control = False

    for run in paragraph.runs:
        set_run_font(
            run,
            size_pt=BODY_FONT_SIZE_PT,
            bold=True,
            italic=False,
            all_caps=False,
        )

def format_figure_caption(paragraph):
    hard_reset_paragraph_format(paragraph, first_line_indent_cm=FIRST_LINE_INDENT_CM)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in paragraph.runs:
        set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=False, all_caps=False)


def set_cell_border(cell, color="000000", size="4", space="0"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = tcBorders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            tcBorders.append(element)

        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), space)
        element.set(qn("w:color"), color)

def force_table_outer_borders_single(table, color="000000", size="4", space="0"):
    """
    Жестко задает таблице одинарные границы и убирает cell spacing,
    из-за которого в Word могут визуально появляться двойные линии.
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)

    tblBorders = tblPr.find(qn("w:tblBorders"))
    if tblBorders is None:
        tblBorders = OxmlElement("w:tblBorders")
        tblPr.append(tblBorders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = tblBorders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            tblBorders.append(element)

        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), space)
        element.set(qn("w:color"), color)

    tblCellSpacing = tblPr.find(qn("w:tblCellSpacing"))
    if tblCellSpacing is None:
        tblCellSpacing = OxmlElement("w:tblCellSpacing")
        tblPr.append(tblCellSpacing)

    tblCellSpacing.set(qn("w:w"), "0")
    tblCellSpacing.set(qn("w:type"), "dxa")

def apply_table_borders(table):
    force_table_outer_borders_single(table, size="4")

    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, size="4")

def force_zero_indent_in_table_paragraph(paragraph):
    """
    Жестко сбрасывает любые абзацные отступы внутри таблицы:
    первая строка, левый/правый отступ, а также XML-атрибуты w:ind.
    Это узкая функция только для абзацев внутри ячеек таблицы.
    """
    fmt = paragraph.paragraph_format
    fmt.first_line_indent = Cm(0)
    fmt.left_indent = Cm(0)
    fmt.right_indent = Cm(0)

    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)

    # Полностью обнуляем ключевые виды отступов Word
    ind.set(qn("w:firstLine"), "0")
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:right"), "0")
    ind.set(qn("w:start"), "0")
    ind.set(qn("w:end"), "0")
    ind.set(qn("w:hanging"), "0")


def force_table_run_plain(run):
    """
    Жестко убирает жирность у run внутри таблицы.
    Обычного run.bold = False иногда недостаточно, поэтому
    дополнительно прибиваем XML-свойства жирности.
    """
    set_run_font(
        run,
        size_pt=TABLE_FONT_SIZE_PT,
        bold=False,
        italic=False,
        all_caps=False,
    )

    run.bold = False
    run.font.bold = False

    rPr = run._element.get_or_add_rPr()

    for tag in ("w:b", "w:bCs"):
        node = rPr.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            rPr.append(node)
        node.set(qn("w:val"), "0")
        
def format_tables(document):
    for table in document.tables:
        apply_table_borders(table)

        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    force_paragraph_xml_spacing(paragraph, line_rule="auto")

                    fmt = paragraph.paragraph_format
                    fmt.first_line_indent = Cm(0)
                    fmt.left_indent = Cm(0)
                    fmt.right_indent = Cm(0)
                    fmt.line_spacing = LINE_SPACING_TABLE
                    fmt.space_before = Pt(0)
                    fmt.space_after = Pt(0)
                    fmt.keep_together = False
                    fmt.keep_with_next = False
                    fmt.page_break_before = False
                    fmt.widow_control = False

                    # Жестко обнуляем отступы и на уровне XML тоже
                    force_zero_indent_in_table_paragraph(paragraph)

                    for run in paragraph.runs:
                        # Жестко убираем жирность и нормализуем шрифт
                        force_table_run_plain(run)
                        
def smart_repair_heading1(paragraph, text: str):
    cleaned = strip_leading_heading_garbage(text)
    parsed = parse_heading1(cleaned)
    if not parsed:
        return False

    if parsed["kind"] == "heading1_chapter":
        chapter_num = parsed["chapter_num"]
        title = parsed["title"].upper()
        replace_paragraph_text(paragraph, f"{chapter_num}. {title}")
        remove_paragraph_numbering(paragraph)
        format_heading1(paragraph)
        return True

    if parsed["kind"] == "heading1_exact":
        replace_paragraph_text(paragraph, cleaned.upper())
        remove_paragraph_numbering(paragraph)
        format_heading1(paragraph)
        return True

    return False


def smart_repair_broken_heading2(paragraph, current_chapter_num, next_paragraph_num):
    if current_chapter_num is None or next_paragraph_num is None:
        return None

    text = clean_spaces(paragraph.text)
    parsed = parse_broken_heading2(text)
    if not parsed:
        return None

    if not is_probable_center_bold_heading(paragraph):
        return None

    title = parsed["title"].lstrip(". ").strip()
    new_text = f"{current_chapter_num}.{next_paragraph_num}. {title}"
    replace_paragraph_text(paragraph, new_text)
    remove_paragraph_numbering(paragraph)
    format_heading2(paragraph)
    return new_text


def looks_like_heading2_title(text: str) -> bool:
    t = clean_spaces(text)
    if not t:
        return False

    low = t.lower()

    if low.startswith("таблица "):
        return False
    if low.startswith("рисунок "):
        return False
    if low.startswith("рис. "):
        return False
    if low.startswith("продолжение таблицы"):
        return False
    if low.startswith("продолжение табл."):
        return False

    if is_table_continuation_text(t):
        return False

    if low in REFERENCE_SUBHEADINGS_CANON:
        return False
    if parse_heading1(t) or parse_heading2(t) or parse_broken_heading2(t):
        return False
    if TABLE_NUM_RE.match(t) or FIG_RE.match(t) or DASH_LINE_RE.match(t):
        return False
    if re.match(r"^\s*(источник|составлено по|рассчитано по|примечание)\s*:", t, re.IGNORECASE):
        return False
    if t.endswith((".", ":", ";", "?", "!")):
        return False
    if len(t) > 220:
        return False

    return True


def is_likely_numbered_heading2_candidate(paragraph, current_chapter_num, next_paragraph_num, prev_kind=None):
    if current_chapter_num is None or next_paragraph_num is None:
        return False
    if not paragraph_has_numbering(paragraph):
        return False

    text = clean_spaces(paragraph.text)
    if not looks_like_heading2_title(text):
        return False

    if is_probable_center_bold_heading(paragraph):
        return True

    # Частый кейс: Word-автонумерация у первого параграфа после названия главы.
    if prev_kind in {"heading1", "empty_paragraph"}:
        return True

    return True


    
def normalize_heading2_numbering(paragraph, current_chapter_num, next_paragraph_num):
    if current_chapter_num is None or next_paragraph_num is None:
        return None

    text = strip_leading_heading_garbage(paragraph.text)
    text = clean_spaces(text)
    if not text:
        return None

    has_num = paragraph_has_numbering(paragraph)
    parsed = parse_heading2(text)

    if parsed:
        normalized = f"{parsed['chapter_num']}.{parsed['paragraph_num']}. {parsed['title']}"
        if text != normalized:
            replace_paragraph_text(paragraph, normalized)
        if has_num:
            remove_paragraph_numbering(paragraph)
        return normalized

    if has_num and looks_like_heading2_title(text):
        title = text.lstrip(". ").strip()
        new_text = f"{current_chapter_num}.{next_paragraph_num}. {title}"
        replace_paragraph_text(paragraph, new_text)
        remove_paragraph_numbering(paragraph)
        format_heading2(paragraph)
        return new_text

    return None

def normalize_table_continuation_text(paragraph):
    text = clean_spaces(paragraph.text)
    low = text.lower()

    if "продол" in low and "таб" in low:
        m = re.search(r"(\d+(?:\.\d+){1,2})", text)
        if m:
            replace_paragraph_text(paragraph, f"Продолжение таблицы {m.group(1)}")


def normalize_figure_caption_text(paragraph):
    text = clean_spaces(paragraph.text)
    if not text:
        return

    m = FIG_RE.match(text)
    if not m:
        return

    number = m.group(2)
    title = clean_spaces(m.group(3) or "")

    if title:
        normalized = f"Рис. {number}. {title}"
    else:
        normalized = f"Рис. {number}"

    if text != normalized:
        replace_paragraph_text(paragraph, normalized)


def normalize_toc_line(text: str) -> str:
    t = clean_spaces(text.replace("\t", " "))

    # Убираем хвосты содержания:
    # ..... 12
    # ……… 12
    # . . . 12
    # смешанные лидеры и пробелы перед номером страницы
    t = re.sub(r'[\s\.\u2024\u2025\u2026·•]+(\d+)\s*$', "", t).strip()

    # Дополнительно убираем хвосты вида "………………" без номера,
    # если Word уже отдельно разорвал страницу/табуляцию
    t = re.sub(r'[\s\.\u2024\u2025\u2026·•]+$', "", t).strip()

    return t


def build_toc_heading_maps(document, body_start):
    h1_map = {}
    h2_map = {}

    if body_start is None:
        return h1_map, h2_map

    for idx, p in enumerate(document.paragraphs):
        if idx >= body_start:
            break

        text = normalize_toc_line(p.text)
        if not text:
            continue

        parsed_h1 = parse_heading1(text)
        if parsed_h1 and parsed_h1["kind"] == "heading1_chapter":
            h1_map[parsed_h1["chapter_num"]] = f'{parsed_h1["chapter_num"]}. {parsed_h1["title"]}'
            continue

        parsed_h2 = parse_heading2(text)
        if parsed_h2:
            key = (parsed_h2["chapter_num"], parsed_h2["paragraph_num"])
            h2_map[key] = f'{parsed_h2["chapter_num"]}.{parsed_h2["paragraph_num"]}. {parsed_h2["title"]}'

    return h1_map, h2_map

def detect_kind_from_paragraph_object(paragraph, text: str, prev_kind=None) -> str:
    t = clean_spaces(text)
    low = t.lower()

    parsed_h1 = parse_heading1(t)
    if parsed_h1:
        if parsed_h1["kind"] == "heading1_exact" and low == "содержание":
            return "toc_heading"
        return "heading1"

    if parse_heading2(t):
        return "heading2"

    if parse_broken_heading2(t):
        return "broken_heading2"

    if TABLE_NUM_RE.match(t):
        return "table_caption"

    if is_table_continuation_text(t):
        return "table_continuation"

    if FIG_RE.match(t):
        return "figure_caption"

    if re.match(r"^\s*(источник|составлено по|рассчитано по|примечание)\s*:", t, re.IGNORECASE):
        return "source_line"

    style_name = ""
    try:
        style_name = (paragraph.style.name or "").strip().lower()
    except Exception:
        style_name = ""

    if style_name in {"heading 1", "заголовок 1"}:
        return "heading1"

    if style_name in {"heading 2", "заголовок 2"}:
        return "heading2"


    if prev_kind in {"table_caption", "table_continuation"}:
        return "table_title"

    return "body_text"
    
def split_manual_dash_lists(document, body_start):
    changed = True
    while changed:
        changed = False
        paragraphs = document.paragraphs

        for idx, p in enumerate(paragraphs):
            if idx < body_start:
                continue

            raw = p.text.replace("\r", "\n").replace("\v", "\n")
            if "\n" not in raw:
                continue

            parts = [x.strip() for x in re.split(r"[\n]+", raw) if x.strip()]
            if len(parts) < 2:
                continue

            if not all(DASH_LINE_RE.match(x) for x in parts[1:]):
                continue

            replace_paragraph_text(p, parts[0])
            prev = p
            for item in parts[1:]:
                prev = insert_paragraph_after(prev, item)

            changed = True
            break


def split_table_captions_prepass(document, body_start):
    changed = True
    while changed:
        changed = False
        paragraphs = document.paragraphs

        for idx, p in enumerate(paragraphs):
            if idx < body_start:
                continue

            text = clean_spaces(p.text)
            if not text or not text.lower().startswith("таблица"):
                continue

            m = TABLE_NUM_RE.match(text)
            if not m:
                continue

            number = m.group(1)
            tail = clean_spaces(m.group(2))
            if not tail:
                continue

            replace_paragraph_text(p, f"Таблица {number}")
            title_p = insert_paragraph_after(p, tail)

            format_table_caption(p)
            format_table_title(title_p)

            changed = True
            break


def convert_reference_numbering_to_plain_text(document, body_start):
    in_references = False
    ref_counter = 1

    for idx, paragraph in enumerate(document.paragraphs):
        if idx < body_start:
            continue

        text = clean_spaces(paragraph.text)
        low = text.lower()
        canonical = canonical_reference_subheading_text(text)

        if low in {
            "список использованных источников",
            "список использованной литературы",
        }:
            in_references = True
            ref_counter = 1
            continue

        if not in_references:
            continue

        if low in {"приложения", "приложение"}:
            in_references = False
            continue

        # Подзаголовки внутри списка источников
        if canonical:
            replace_paragraph_text(paragraph, canonical)
            remove_paragraph_numbering(paragraph)
            remove_page_break_artifacts_from_paragraph(paragraph)

            format_reference_subheading(paragraph)
            continue

        if is_empty_paragraph(paragraph):
            continue

        # Любой обычный источник в блоке литературы
        remove_paragraph_numbering(paragraph)
        remove_page_break_artifacts_from_paragraph(paragraph)
        set_paragraph_style_safe(paragraph, "Normal", "Обычный")
        clear_paragraph_outline_level(paragraph)

        clean = clean_spaces(paragraph.text)

        # Если видимый номер уже есть — сохраняем его
        m = re.match(r"^\s*(\d+)\.\s+(.+)$", clean)
        if m:
            number = int(m.group(1))
            source_text = clean_spaces(m.group(2))
            normalized = f"{number}. {source_text}"
            ref_counter = number + 1
        else:
            normalized = f"{ref_counter}. {clean}"
            ref_counter += 1

        normalized = smart_normalize_reference_line_case(normalized)
        replace_paragraph_text(paragraph, normalized)
        format_body(paragraph)

def compact_references_block(document, body_start):
    in_references = False
    changed = True

    while changed:
        changed = False
        paragraphs = document.paragraphs

        for idx, p in enumerate(paragraphs):
            if idx < body_start:
                continue

            text = clean_spaces(p.text)
            low = text.lower()
            canonical = canonical_reference_subheading_text(text)

            if is_references_heading_text(text):
                in_references = True
                continue

            if not in_references:
                continue

            if is_appendix_heading_text(text):
                in_references = False
                continue

            # Полностью убираем пустые абзацы внутри блока литературы
            if is_empty_paragraph(p):
                remove_paragraph(p)
                changed = True
                break

            # Сначала снимаем весь мусор разрывов / списков / заголовков
            remove_page_break_artifacts_from_paragraph(p)
            remove_paragraph_numbering(p)
            set_paragraph_style_safe(p, "Normal", "Обычный")
            clear_paragraph_outline_level(p)

            # Подзаголовки разделов внутри литературы
            if canonical:
                replace_paragraph_text(p, canonical)
                format_reference_subheading(p)

                p.paragraph_format.page_break_before = False
                p.paragraph_format.keep_with_next = False
                p.paragraph_format.keep_together = False
                p.paragraph_format.widow_control = False

                continue

            # Обычный источник
            clean = clean_spaces(p.text)

            m = re.match(r"^\s*(\d+)\.\s+(.+)$", clean)
            if m:
                number = int(m.group(1))
                source_text = clean_spaces(m.group(2))
                normalized = f"{number}. {source_text}"
            else:
                normalized = clean

            normalized = smart_normalize_reference_line_case(normalized)

            if clean != normalized:
                replace_paragraph_text(p, normalized)

            format_body(p)

            # Финальный добивающий reset именно после format_body
            p.paragraph_format.page_break_before = False
            p.paragraph_format.keep_with_next = False
            p.paragraph_format.keep_together = False
            p.paragraph_format.widow_control = False

def ensure_single_blank_after_references_heading(document, body_start):
    any_changes = False
    changed = True
    while changed:
        changed = False
        paragraphs = document.paragraphs

        for idx, p in enumerate(paragraphs):
            if idx < body_start:
                continue

            text = clean_spaces(p.text)
            if not is_references_heading_text(text):
                continue

            # гарантируем 1 пустую строку после заголовка списка
            if idx + 1 >= len(paragraphs):
                new_p = insert_paragraph_after(p, "")
                format_empty_paragraph(new_p)
                changed = True
                any_changes = True
                break

            next_p = paragraphs[idx + 1]

            if not is_empty_paragraph(next_p):
                new_p = insert_paragraph_after(p, "")
                format_empty_paragraph(new_p)
                changed = True
                any_changes = True
                break

            # если пустых строк больше одной — удаляем лишние
            while idx + 2 < len(paragraphs) and is_empty_paragraph(paragraphs[idx + 2]):
                remove_paragraph(paragraphs[idx + 2])
                paragraphs = document.paragraphs
                changed = True
                any_changes = True

            format_empty_paragraph(next_p)
            break

    return any_changes
    
def ensure_single_blank_after_headings(document, body_start):
    paragraphs = document.paragraphs
    prev_kind = None
    changed = False
    in_references = False

    idx = max(body_start, 0)

    while idx < len(paragraphs):
        p = paragraphs[idx]
        text = clean_spaces(p.text)

        if is_references_heading_text(text):
            in_references = True
        elif in_references and is_appendix_heading_text(text):
            in_references = False

        kind = classify_paragraph(text, prev_kind=prev_kind)
        parsed_h1 = parse_heading1(text)

        need_blank_after = False

        # После параграфов 1.1 / 1.2 / 2.1 и т.д. нужна одна пустая строка
        if kind == "heading2":
            need_blank_after = True

        # После ВВЕДЕНИЯ / ЗАКЛЮЧЕНИЯ / СПИСКА ИСТОЧНИКОВ нужна одна пустая строка
        # После названий глав 1 / 2 / 3 пустая строка НЕ нужна
        elif parsed_h1:
            if parsed_h1["kind"] == "heading1_exact":
                need_blank_after = True
            elif parsed_h1["kind"] == "heading1_chapter":
                need_blank_after = False

        if not need_blank_after:
            prev_kind = kind
            idx += 1
            continue

        if idx + 1 >= len(paragraphs):
            new_p = insert_paragraph_after(p, "")
            format_empty_paragraph(new_p)
            changed = True
            break

        next_p = paragraphs[idx + 1]

        if not is_empty_paragraph(next_p):
            new_p = insert_paragraph_after(p, "")
            format_empty_paragraph(new_p)
            changed = True
            break

        while idx + 2 < len(paragraphs) and is_empty_paragraph(paragraphs[idx + 2]):
            remove_paragraph(paragraphs[idx + 2])
            paragraphs = document.paragraphs
            changed = True

        if idx + 1 < len(paragraphs) and is_empty_paragraph(paragraphs[idx + 1]):
            format_empty_paragraph(paragraphs[idx + 1])

        prev_kind = kind
        idx += 1

    return changed
    
def collapse_empty_paragraphs_in_body(paragraphs, body_start):
    empty_count = 0
    for idx, p in enumerate(list(paragraphs)):
        if idx < body_start:
            continue

        if is_empty_paragraph(p):
            empty_count += 1
            if empty_count > 1:
                remove_paragraph(p)
        else:
            empty_count = 0

def remove_single_empty_between_body_paragraphs(document, body_start):
    changed = True
    while changed:
        changed = False
        paragraphs = document.paragraphs

        for idx, p in enumerate(paragraphs):
            if idx < body_start:
                continue

            if not is_empty_paragraph(p):
                continue

            # Ищем ближайший непустой абзац слева
            prev_idx = idx - 1
            while prev_idx >= body_start and is_empty_paragraph(paragraphs[prev_idx]):
                prev_idx -= 1

            # Ищем ближайший непустой абзац справа
            next_idx = idx + 1
            while next_idx < len(paragraphs) and is_empty_paragraph(paragraphs[next_idx]):
                next_idx += 1

            if prev_idx < body_start or next_idx >= len(paragraphs):
                continue

            prev_text = clean_spaces(paragraphs[prev_idx].text)
            next_text = clean_spaces(paragraphs[next_idx].text)

            prev_prev_kind = None
            for j in range(body_start, prev_idx):
                t = clean_spaces(paragraphs[j].text)
                if not t:
                    continue
                prev_prev_kind = classify_paragraph(t, prev_kind=prev_prev_kind)

            prev_kind = classify_paragraph(prev_text, prev_kind=prev_prev_kind)
            next_kind = classify_paragraph(next_text, prev_kind=prev_kind)

            # Удаляем только случайную пустую строку между двумя обычными абзацами текста
            if prev_kind == "body_text" and next_kind == "body_text":
                remove_paragraph(p)
                changed = True
                break

    return changed

def ensure_empty_after_source_and_note(document, body_start):
    changed = True
    while changed:
        changed = False
        paragraphs = document.paragraphs
        prev_kind = None
        in_references = False

        for idx, p in enumerate(paragraphs):
            if idx < body_start:
                continue

            text = clean_spaces(p.text)
            low = text.lower()

            if is_references_heading_text(text):
                in_references = True
                prev_kind = "heading1"
                continue

            if in_references and is_appendix_heading_text(text):
                in_references = False

            # Внутри списка источников ничего не разрежаем
            if in_references:
                prev_kind = "body_text"
                continue

            kind = classify_paragraph(text, prev_kind=prev_kind)
            is_note_line = bool(re.match(r"^\s*примечание\s*:", text, re.IGNORECASE))

            # ===== SOURCE LINE LOGIC =====
            if kind == "source_line":

                # 1) Если сразу после источника идёт пустая строка,
                # а после неё Примечание: -> удалить эту пустую строку
                if (
                    idx + 2 < len(paragraphs)
                    and is_empty_paragraph(paragraphs[idx + 1])
                    and re.match(r"^\s*примечание\s*:", clean_spaces(paragraphs[idx + 2].text), re.IGNORECASE)
                ):
                    remove_paragraph(paragraphs[idx + 1])
                    changed = True
                    break

                # 2) Если сразу после источника идёт Примечание: -> ничего не вставляем
                if idx + 1 < len(paragraphs):
                    next_text = clean_spaces(paragraphs[idx + 1].text)
                    if re.match(r"^\s*примечание\s*:", next_text, re.IGNORECASE):
                        prev_kind = kind
                        continue

                # 3) Во всех остальных случаях после источника должна быть ровно одна пустая строка
                if idx + 1 >= len(paragraphs):
                    new_p = insert_paragraph_after(p, "")
                    format_empty_paragraph(new_p)
                    changed = True
                    break

                if not is_empty_paragraph(paragraphs[idx + 1]):
                    new_p = insert_paragraph_after(p, "")
                    format_empty_paragraph(new_p)
                    changed = True
                    break

                if (
                    idx + 2 < len(paragraphs)
                    and is_empty_paragraph(paragraphs[idx + 1])
                    and is_empty_paragraph(paragraphs[idx + 2])
                ):
                    remove_paragraph(paragraphs[idx + 2])
                    changed = True
                    break

                format_empty_paragraph(paragraphs[idx + 1])
                prev_kind = kind
                continue

            # ===== NOTE LINE LOGIC =====
            if is_note_line:
                # Ищем предыдущий непустой абзац
                prev_nonempty_idx = idx - 1
                while prev_nonempty_idx >= body_start and is_empty_paragraph(paragraphs[prev_nonempty_idx]):
                    prev_nonempty_idx -= 1

                prev_nonempty_kind = None
                if prev_nonempty_idx >= body_start:
                    prev_nonempty_text = clean_spaces(paragraphs[prev_nonempty_idx].text)
                    prev_nonempty_kind = classify_paragraph(prev_nonempty_text, prev_kind=None)

                # Логику после Примечания применяем только если перед ним реально был Источник
                if prev_nonempty_kind == "source_line":

                    # После примечания должна быть ровно одна пустая строка
                    if idx + 1 >= len(paragraphs):
                        new_p = insert_paragraph_after(p, "")
                        format_empty_paragraph(new_p)
                        changed = True
                        break

                    if not is_empty_paragraph(paragraphs[idx + 1]):
                        new_p = insert_paragraph_after(p, "")
                        format_empty_paragraph(new_p)
                        changed = True
                        break

                    if (
                        idx + 2 < len(paragraphs)
                        and is_empty_paragraph(paragraphs[idx + 1])
                        and is_empty_paragraph(paragraphs[idx + 2])
                    ):
                        remove_paragraph(paragraphs[idx + 2])
                        changed = True
                        break

                    format_empty_paragraph(paragraphs[idx + 1])

                prev_kind = "body_text"
                continue

            prev_kind = kind
            
def ensure_one_empty_after(paragraphs, index):
    """Ensure exactly one empty paragraph right after paragraphs[index]."""
    if index >= len(paragraphs):
        return False

    changed = False
    p = paragraphs[index]

    if index + 1 >= len(paragraphs):
        new_p = insert_paragraph_after(p, "")
        format_empty_paragraph(new_p)
        return True

    next_p = paragraphs[index + 1]
    if not is_empty_paragraph(next_p):
        new_p = insert_paragraph_after(p, "")
        format_empty_paragraph(new_p)
        return True

    format_empty_paragraph(next_p)

    while index + 2 < len(paragraphs) and is_empty_paragraph(paragraphs[index + 2]):
        remove_paragraph(paragraphs[index + 2])
        paragraphs = p._parent.paragraphs
        changed = True

    return changed

def ensure_empty_between_heading1_and_heading2(document, body_start):
    changed = True
    while changed:
        changed = False
        paragraphs = document.paragraphs
        prev_kind = None

        for idx, p in enumerate(paragraphs):
            if idx < body_start:
                continue

            kind = classify_paragraph(clean_spaces(p.text), prev_kind=prev_kind)

            if kind == "heading1" and idx + 2 < len(paragraphs):
                next_p = paragraphs[idx + 1]
                next2_p = paragraphs[idx + 2]
                next2_kind = classify_paragraph(clean_spaces(next2_p.text), prev_kind="empty_paragraph")

                if is_empty_paragraph(next_p) and next2_kind == "heading2":
                    remove_paragraph(next_p)
                    changed = True
                    break

            prev_kind = kind


def ensure_compact_heading2_spacing(document, body_start):
    """
    Normalize spacing around heading2 in one pass.

    Rules:
      - no empty paragraph immediately before heading2;
      - exactly one empty paragraph immediately after heading2.

    Returns True if any changes were made, otherwise False.
    """
    paragraphs = document.paragraphs
    prev_kind = None
    changed = False
    idx = max(body_start, 0)
    in_references = False

    while idx < len(paragraphs):
        p = paragraphs[idx]
        text = clean_spaces(p.text)

        if is_references_heading_text(text):
            in_references = True
            prev_kind = "heading1"
            idx += 1
            continue

        if in_references and is_appendix_heading_text(text):
            in_references = False

        # Внутри списка источников этот проход не должен ничего вставлять/удалять
        if in_references:
            prev_kind = "body_text"
            idx += 1
            continue

        kind = classify_paragraph(text, prev_kind=prev_kind)

        if kind != "heading2":
            prev_kind = kind
            idx += 1
            continue

        while idx - 1 >= body_start and is_empty_paragraph(paragraphs[idx - 1]):
            remove_paragraph(paragraphs[idx - 1])
            paragraphs = document.paragraphs
            idx -= 1
            changed = True
            
        next_is_heading2 = False
        if idx + 1 < len(paragraphs):
            next_text = clean_spaces(paragraphs[idx + 1].text)
            next_is_heading2 = classify_paragraph(next_text, prev_kind="heading2") == "heading2"



        if next_is_heading2:
            while idx + 1 < len(paragraphs) and is_empty_paragraph(paragraphs[idx + 1]):
                remove_paragraph(paragraphs[idx + 1])
                paragraphs = document.paragraphs
                changed = True
        else:
            if idx + 1 >= len(paragraphs) or not is_empty_paragraph(paragraphs[idx + 1]):
                new_p = OxmlElement("w:p")
                p._element.addnext(new_p)
                paragraphs = document.paragraphs
                changed = True
                
            while idx + 2 < len(paragraphs) and is_empty_paragraph(paragraphs[idx + 2]):
                remove_paragraph(paragraphs[idx + 2])
                paragraphs = document.paragraphs
                changed = True

            if idx + 1 < len(paragraphs) and is_empty_paragraph(paragraphs[idx + 1]):
                hard_reset_paragraph_format(paragraphs[idx + 1], first_line_indent_cm=None)


        prev_kind = kind
        idx += 1

    return changed



STRUCTURAL_HEADING_TEXTS_V2 = {
    "ВВЕДЕНИЕ",
    "ЗАКЛЮЧЕНИЕ",
}

def normalize_structural_heading_spacing_v2(document, body_start):
    changed = True
    while changed:
        changed = False
        paragraphs = document.paragraphs

        for idx, p in enumerate(paragraphs):
            if idx < body_start:
                continue

            text = clean_spaces(p.text).upper()
            if text not in STRUCTURAL_HEADING_TEXTS_V2:
                continue

            # Сразу после ВВЕДЕНИЕ / ЗАКЛЮЧЕНИЕ должна быть ровно одна пустая строка
            if idx + 1 >= len(paragraphs):
                new_p = insert_paragraph_after(p, "")
                hard_reset_paragraph_format(new_p, first_line_indent_cm=None)
                changed = True
                break

            next_p = paragraphs[idx + 1]

            if not is_empty_paragraph(next_p):
                new_p = insert_paragraph_after(p, "")
                hard_reset_paragraph_format(new_p, first_line_indent_cm=None)
                changed = True
                break

            # Если пустых строк больше одной — сжимаем до одной
            if idx + 2 < len(paragraphs) and is_empty_paragraph(paragraphs[idx + 2]):
                remove_paragraph(paragraphs[idx + 2])
                changed = True
                break

            # Нормализуем единственную пустую строку
            hard_reset_paragraph_format(next_p, first_line_indent_cm=None)


def ensure_empty_before_table_caption(document, body_start):
    changed = True
    while changed:
        changed = False
        paragraphs = document.paragraphs
        prev_kind = None

        for idx, p in enumerate(paragraphs):
            if idx < body_start:
                continue

            kind = classify_paragraph(clean_spaces(p.text), prev_kind=prev_kind)

            if kind in {"table_caption", "table_continuation"}:
                if idx - 1 >= body_start:
                    prev_p = paragraphs[idx - 1]
                    if not is_empty_paragraph(prev_p):
                        new_p = OxmlElement("w:p")
                        prev_p._element.addnext(new_p)
                        changed = True
                        break
                    if idx - 2 >= body_start and is_empty_paragraph(paragraphs[idx - 2]):
                        remove_paragraph(prev_p)
                        changed = True
                        break

            prev_kind = kind


def remove_extra_empty_after_service_lines(document, body_start):
    target_kinds = {
        "table_caption",
        "table_title",
        "table_continuation",
        "reference_subheading",
    }

    changed = True
    while changed:
        changed = False
        paragraphs = document.paragraphs
        prev_kind = None

        for idx, p in enumerate(paragraphs):
            if idx < body_start:
                continue

            kind = classify_paragraph(clean_spaces(p.text), prev_kind=prev_kind)

            if kind in target_kinds:
                if idx + 1 < len(paragraphs) and is_empty_paragraph(paragraphs[idx + 1]):
                    remove_paragraph(paragraphs[idx + 1])
                    changed = True
                    break

            prev_kind = kind


def cleanup_reference_subheadings_layout(document, body_start):
    in_references = False
    changed = True

    while changed:
        changed = False
        paragraphs = document.paragraphs

        for idx, p in enumerate(paragraphs):
            if idx < body_start:
                continue

            text = clean_spaces(p.text)
            low = text.lower()

            if low in {
                "список использованных источников",
                "список использованной литературы",
            }:
                in_references = True
                continue

            if not in_references:
                continue

            if low in {"приложения", "приложение"}:
                in_references = False
                continue

            canonical = canonical_reference_subheading_text(text)
            if canonical:
                replace_paragraph_text(p, canonical)
                remove_paragraph_numbering(p)
                p.paragraph_format.page_break_before = False
                format_reference_subheading(p)

                if idx - 1 >= body_start and is_empty_paragraph(paragraphs[idx - 1]):
                    # Если пустая строка стоит сразу после
                    # "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
                    # её сохраняем — она нужна по шаблону.
                    if idx - 2 >= body_start:
                        prev_prev_text = clean_spaces(paragraphs[idx - 2].text)
                        if is_references_heading_text(prev_prev_text):
                            pass
                        else:
                            remove_paragraph(paragraphs[idx - 1])
                            changed = True
                            break
                    else:
                        remove_paragraph(paragraphs[idx - 1])
                        changed = True
                        break

                if idx + 1 < len(paragraphs) and is_empty_paragraph(paragraphs[idx + 1]):
                    remove_paragraph(paragraphs[idx + 1])
                    changed = True
                    break


def format_empty_paragraph(paragraph):
    set_paragraph_style_safe(paragraph, "Normal", "Обычный")
    clear_paragraph_outline_level(paragraph)
    remove_paragraph_numbering(paragraph)

    hard_reset_paragraph_format(paragraph, first_line_indent_cm=None)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = ensure_empty_run(paragraph)
    set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=False, all_caps=False)
    
def normalize_sections(document):
    """
    Удаляет секционные разрывы из абзацев внутри документа,
    чтобы потом можно было заново поставить один правильный
    разрыв секции перед ВВЕДЕНИЕМ.
    """
    for p in document.paragraphs:
        pPr = p._element.pPr
        if pPr is None:
            continue

        sectPr = pPr.find(qn("w:sectPr"))
        if sectPr is not None:
            pPr.remove(sectPr)

def ensure_section_break_before_introduction(document, body_start):
    """
    Ставит разрыв секции типа Next Page перед абзацем 'ВВЕДЕНИЕ'.
    Это нужно, чтобы:
    - 1-я страница имела свой футер,
    - 2-я страница была пустой,
    - с 3-й страницы можно было включить нумерацию отдельной секцией.
    """
    if body_start is None:
        return

    paragraphs = document.paragraphs
    if body_start <= 0 or body_start >= len(paragraphs):
        return

    intro_p = paragraphs[body_start]
    prev_p = paragraphs[body_start - 1]

    prev_pPr = prev_p._element.get_or_add_pPr()

    # Если секционный разрыв уже стоит — второй раз не добавляем
    existing_sectPr = prev_pPr.find(qn("w:sectPr"))
    if existing_sectPr is not None:
        return

    next_pPr = intro_p._element.pPr
    if next_pPr is not None and next_pPr.find(qn("w:sectPr")) is not None:
        return

    body = document._body._element
    body_sectPr = body.sectPr
    if body_sectPr is None:
        return

    new_sectPr = deepcopy(body_sectPr)

    # Делаем разрыв секции "со следующей страницы"
    type_el = new_sectPr.find(qn("w:type"))
    if type_el is None:
        type_el = OxmlElement("w:type")
        new_sectPr.insert(0, type_el)
    type_el.set(qn("w:val"), "nextPage")

    prev_pPr.append(new_sectPr)

def _append_next_page_section_break_after(paragraph, body_sectpr):
    pPr = paragraph._element.get_or_add_pPr()

    old = pPr.find(qn("w:sectPr"))
    if old is not None:
        pPr.remove(old)

    new_sectPr = deepcopy(body_sectpr)

    # Не тащим старые ссылки на футеры/заголовки и старый старт нумерации
    for tag in ("w:pgNumType", "w:footerReference", "w:headerReference"):
        for el in list(new_sectPr.findall(qn(tag))):
            new_sectPr.remove(el)

    type_el = new_sectPr.find(qn("w:type"))
    if type_el is None:
        type_el = OxmlElement("w:type")
        new_sectPr.insert(0, type_el)
    type_el.set(qn("w:val"), "nextPage")

    pPr.append(new_sectPr)


def ensure_front_matter_layout(document, body_start):
    """
    Целевая модель:
    если есть содержание:
        секция 1 = титул
        секция 2 = содержание
        секция 3 = введение и далее
    если содержания нет:
        секция 1 = титул
        секция 2 = введение и далее
    """
    if body_start is None or body_start <= 0:
        return

    paragraphs = document.paragraphs
    if body_start >= len(paragraphs):
        return

    body = document._body._element
    body_sectpr = body.sectPr
    if body_sectpr is None:
        return

    # Полная очистка page-break артефактов до введения
    for i in range(body_start):
        p = paragraphs[i]
        p.paragraph_format.page_break_before = False

        for run in p.runs:
            r = run._element
            for br in list(r.findall(qn("w:br"))):
                br_type = br.get(qn("w:type"))
                if br_type in (None, "page"):
                    r.remove(br)

    # Ищем содержание до введения
    contents_idx = None
    for i in range(body_start):
        t = clean_spaces(paragraphs[i].text).upper()
        if ("СОДЕРЖАН" in t) or ("ОГЛАВЛЕН" in t):
            contents_idx = i
            break

    # На самом абзаце введения обычный page break не нужен
    paragraphs[body_start].paragraph_format.page_break_before = False

    if contents_idx is not None and contents_idx > 0:
        # титул -> содержание
        _append_next_page_section_break_after(paragraphs[contents_idx - 1], body_sectpr)
        # содержание -> введение
        _append_next_page_section_break_after(paragraphs[body_start - 1], body_sectpr)
    else:
        # титул -> введение
        _append_next_page_section_break_after(paragraphs[body_start - 1], body_sectpr)

def remove_all_italic(doc):
    """
    Убирает курсив, highlight, цвет текста и XML-заливку из всего документа.
    """

    def clear_run(run):
        run.italic = False

        try:
            run.font.highlight_color = None
        except Exception:
            pass

        try:
            run.font.color.rgb = RGBColor(0, 0, 0)
        except Exception:
            pass

        rPr = run._element.get_or_add_rPr()

        for tag in ("w:highlight", "w:shd"):
            node = rPr.find(qn(tag))
            if node is not None:
                rPr.remove(node)

        color = rPr.find(qn("w:color"))
        if color is None:
            color = OxmlElement("w:color")
            rPr.append(color)
        color.set(qn("w:val"), "000000")

        for attr in ("w:themeColor", "w:themeTint", "w:themeShade"):
            qname = qn(attr)
            if qname in color.attrib:
                del color.attrib[qname]

    for p in doc.paragraphs:
        for r in p.runs:
            clear_run(r)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        clear_run(r)


def process_document(input_path: Path, output_path: Path):
    doc = Document(str(input_path))

    # Сразу чистим визуальный мусор по всему документу
    remove_all_italic(doc)
    set_section_margins(doc)

    body_start = find_body_start_index(doc)
    if body_start is None:
        raise RuntimeError("Не найден заголовок 'Введение'; файл пропущен из соображений безопасности.")

    toc_h1_map, toc_h2_map = build_toc_heading_maps(doc, body_start)

    split_manual_dash_lists(doc, body_start)
    split_table_captions_prepass(doc, body_start)
    normalize_quotes_in_document(doc, body_start or 0)
    # Преднормализация только тела работы; содержание не трогаем
    for idx, paragraph in enumerate(doc.paragraphs):
        if idx < body_start:
            continue
        normalize_simple_paragraph_spaces(paragraph)
        normalize_heading2_artifacts(paragraph)

    paragraphs = doc.paragraphs
    prev_kind = None
    current_chapter_num = None
    next_paragraph_num = None

    # Основной проход по телу документа
    for idx, paragraph in enumerate(doc.paragraphs):
        if idx < body_start:
            continue

        text = clean_spaces(paragraph.text)
        if not text:
            prev_kind = "empty_paragraph"
            continue

        text = strip_leading_heading_garbage(text)
        if text != clean_spaces(paragraph.text):
            replace_paragraph_text(paragraph, text)

        kind = detect_kind_from_paragraph_object(paragraph, text, prev_kind=prev_kind)

        parsed_h1 = parse_heading1(text)
        if parsed_h1:
            if parsed_h1["kind"] == "heading1_chapter":
                toc_text = toc_h1_map.get(parsed_h1["chapter_num"])
                current_text = f'{parsed_h1["chapter_num"]}. {parsed_h1["title"]}'

                if toc_text and len(current_text) < len(toc_text):
                    replace_paragraph_text(paragraph, toc_text)
                    text = clean_spaces(paragraph.text)
                    parsed_h1 = parse_heading1(text)

                current_chapter_num = parsed_h1["chapter_num"]
                next_paragraph_num = 1
                smart_repair_heading1(paragraph, text)
                kind = "heading1"

            elif parsed_h1["kind"] == "heading1_exact":
                current_chapter_num = None
                next_paragraph_num = None
                smart_repair_heading1(paragraph, text)
                kind = "heading1"

        parsed_h2_existing = parse_heading2(text)
        if parsed_h2_existing:
            toc_text = toc_h2_map.get(
                (parsed_h2_existing["chapter_num"], parsed_h2_existing["paragraph_num"])
            )
            current_text = (
                f'{parsed_h2_existing["chapter_num"]}.'
                f'{parsed_h2_existing["paragraph_num"]}. '
                f'{parsed_h2_existing["title"]}'
            )

            if toc_text and len(current_text) < len(toc_text):
                replace_paragraph_text(paragraph, toc_text)
                text = clean_spaces(paragraph.text)
                kind = "heading2"

        if kind == "broken_heading2":
            repaired = smart_repair_broken_heading2(
                paragraph,
                current_chapter_num,
                next_paragraph_num,
            )
            if repaired:
                text = clean_spaces(paragraph.text)
                kind = "heading2"

        if kind != "table_continuation" and (
            kind == "heading2"
            or auto_detect_heading2(
                paragraph,
                current_chapter_num,
                next_paragraph_num,
                prev_kind,
            )
            or is_likely_numbered_heading2_candidate(
                paragraph,
                current_chapter_num,
                next_paragraph_num,
                prev_kind=prev_kind,
            )
        ):
            normalized_text = normalize_heading2_numbering(
                paragraph,
                current_chapter_num,
                next_paragraph_num,
            )
            if normalized_text:
                kind = "heading2"
                parsed_h2 = parse_heading2(clean_spaces(paragraph.text))
                if parsed_h2:
                    current_chapter_num = parsed_h2["chapter_num"]
                    next_paragraph_num = parsed_h2["paragraph_num"] + 1

        if kind not in {
            "heading1",
            "heading2",
            "table_caption",
            "table_continuation",
            "table_title",
            "figure_caption",
            "source_line",
            "reference_subheading",
        }:
            if auto_detect_numbered_heading1(
                paragraph,
                current_chapter_num=current_chapter_num,
                next_paragraph=doc.paragraphs[idx + 1] if idx + 1 < len(doc.paragraphs) else None,
            ):
                inferred_chapter_num = 1 if current_chapter_num is None else current_chapter_num + 1
                heading_text = clean_spaces(paragraph.text)
                replace_paragraph_text(paragraph, f"{inferred_chapter_num}. {heading_text}")
                kind = "heading1"
                current_chapter_num = inferred_chapter_num
                next_paragraph_num = 1

        if kind == "table_continuation":
            normalize_table_continuation_text(paragraph)

        if kind == "figure_caption":
            normalize_figure_caption_text(paragraph)

        if kind == "heading1":
            format_heading1(paragraph)

        elif kind == "heading2":
            format_heading2(paragraph)

        elif kind == "table_caption":
            format_table_caption(paragraph)

        elif kind == "table_continuation":
            format_table_caption(paragraph)

        elif kind == "table_title":
            format_table_title(paragraph)

        elif kind == "figure_caption":
            format_figure_caption(paragraph)

        elif kind == "source_line":
            format_source_line(paragraph)

        elif kind == "reference_subheading":
            canonical = canonical_reference_subheading_text(text)
            if canonical:
                replace_paragraph_text(paragraph, canonical)
            format_reference_subheading(paragraph)

        else:
            format_body(paragraph)

        prev_kind = kind

    format_tables(doc)

    convert_reference_numbering_to_plain_text(doc, body_start)

    run_with_pass_limit(
        "compact_references_block",
        compact_references_block,
        doc,
        body_start,
    )

    run_with_pass_limit(
        "ensure_single_blank_after_references_heading",
        ensure_single_blank_after_references_heading,
        doc,
        body_start,
    )

    collapse_empty_paragraphs_in_body(doc.paragraphs, body_start)

    run_with_pass_limit(
        "ensure_compact_heading2_spacing",
        ensure_compact_heading2_spacing,
        doc,
        body_start,
    )

    run_with_pass_limit(
        "ensure_empty_before_table_caption",
        ensure_empty_before_table_caption,
        doc,
        body_start,
    )

    run_with_pass_limit(
        "remove_extra_empty_after_service_lines",
        remove_extra_empty_after_service_lines,
        doc,
        body_start,
    )

    run_with_pass_limit(
        "ensure_empty_after_source_and_note",
        ensure_empty_after_source_and_note,
        doc,
        body_start,
    )

    run_with_pass_limit(
        "cleanup_reference_subheadings_layout",
        cleanup_reference_subheadings_layout,
        doc,
        body_start,
    )

    collapse_empty_paragraphs_in_body(doc.paragraphs, body_start)

    run_with_pass_limit(
        "ensure_single_blank_after_headings",
        ensure_single_blank_after_headings,
        doc,
        body_start,
    )

    run_with_pass_limit(
        "normalize_structural_heading_spacing_v2",
        normalize_structural_heading_spacing_v2,
        doc,
        body_start,
    )
    run_with_pass_limit(
        "remove_single_empty_between_body_paragraphs",
        remove_single_empty_between_body_paragraphs,
        doc,
        body_start,
    )

    # Финальный жёсткий проход:
    # добиваем заголовки, таблицы и обычный текст уже после всех структурных вставок/удалений
    prev_nonempty_kind = None
    for idx, paragraph in enumerate(doc.paragraphs):
        if idx < body_start:
            continue

        text = clean_spaces(paragraph.text)

        if not text:
            format_empty_paragraph(paragraph)
            continue

        text = strip_leading_heading_garbage(text)
        if text != clean_spaces(paragraph.text):
            replace_paragraph_text(paragraph, text)

        if parse_heading1(text):
            smart_repair_heading1(paragraph, text)
            format_heading1(paragraph)
            prev_nonempty_kind = "heading1"
            continue

        if parse_heading2(text):
            format_heading2(paragraph)
            prev_nonempty_kind = "heading2"
            continue

        if TABLE_NUM_RE.match(text):
            format_table_caption(paragraph)
            prev_nonempty_kind = "table_caption"
            continue

        if is_table_continuation_text(text):
            normalize_table_continuation_text(paragraph)
            format_table_caption(paragraph)
            prev_nonempty_kind = "table_continuation"
            continue

        if prev_nonempty_kind in {"table_caption", "table_continuation"}:
            format_table_title(paragraph)
            prev_nonempty_kind = "table_title"
            continue

        if FIG_RE.match(text):
            normalize_figure_caption_text(paragraph)
            format_figure_caption(paragraph)
            prev_nonempty_kind = "figure_caption"
            continue

        if re.match(r"^\s*(источник|составлено по|рассчитано по|примечание)\s*:", text, re.IGNORECASE):
            format_source_line(paragraph)
            prev_nonempty_kind = "source_line"
            continue

        canonical = canonical_reference_subheading_text(text)
        if canonical:
            replace_paragraph_text(paragraph, canonical)
            format_reference_subheading(paragraph)
            prev_nonempty_kind = "reference_subheading"
            continue

        format_body(paragraph)
        prev_nonempty_kind = "body_text"

    normalize_sections(doc)
    ensure_front_matter_layout(doc, body_start)
    apply_page_breaks(doc, body_start)
    apply_page_numbering_policy(doc)

    # И ещё раз дочищаем цвет / highlight в самом конце
    remove_all_italic(doc)

    doc.save(str(output_path))


