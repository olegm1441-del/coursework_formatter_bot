
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

def looks_like_heading2_title(text: str) -> bool:
    t = clean_spaces(text)
    if not t:
        return False

    if TABLE_CONTINUATION_RE.match(t):
        return False

    if t.endswith((".", ":", ";", "!", "?")):
        return False

    if len(t) > 220:
        return False

    return True

def auto_detect_heading2(paragraph, current_chapter_num, next_paragraph_num, prev_kind=None):
    text = clean_spaces(paragraph.text)

    low = text.lower()

    if low.startswith("таблица "):
        return None
    if low.startswith("рисунок "):
        return None
    if low.startswith("рис. "):
        return None
    if low.startswith("продолжение таблицы"):
        return None
    if low.startswith("продолжение табл."):
        return None

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
FIG_RE = re.compile(r"^\s*(рисунок|рис\.)\s*(\d+(?:\.\d+){0,2})\s*[.\-—–]?\s*(.+?)\s*$", re.IGNORECASE)
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


def is_empty_paragraph(paragraph):
    return clean_spaces(paragraph.text) == ""


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
    for section in document.sections:
        section.left_margin = Mm(LEFT_MARGIN_MM)
        section.right_margin = Mm(RIGHT_MARGIN_MM)
        section.top_margin = Mm(TOP_MARGIN_MM)
        section.bottom_margin = Mm(BOTTOM_MARGIN_MM)


def normalize_simple_paragraph_spaces(paragraph):
    if len(paragraph.runs) == 1 and "\n" not in paragraph.runs[0].text and "\v" not in paragraph.runs[0].text:
        old = paragraph.runs[0].text
        new = clean_spaces(old)
        if new != old:
            paragraph.runs[0].text = new


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
            
def format_empty_paragraph(paragraph):
    hard_reset_paragraph_format(paragraph, first_line_indent_cm=None)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = ensure_empty_run(paragraph)
    set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=False, all_caps=False)

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

    set_paragraph_style_safe(paragraph, "Heading 1", "Заголовок 1")
    hard_reset_paragraph_format(paragraph, first_line_indent_cm=None)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for run in paragraph.runs:
        set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=True, all_caps=True)

def format_heading2(paragraph):
    remove_page_break_artifacts_from_paragraph(paragraph)
    remove_paragraph_numbering(paragraph)

    set_paragraph_style_safe(paragraph, "Heading 2", "Заголовок 2")
    hard_reset_paragraph_format(paragraph, first_line_indent_cm=None)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for run in paragraph.runs:
        set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=True, all_caps=False)

def format_table_caption(paragraph):
    set_paragraph_style_safe(paragraph, "Normal", "Обычный")
    clear_paragraph_outline_level(paragraph)
    remove_paragraph_numbering(paragraph)

    hard_reset_paragraph_format(paragraph, first_line_indent_cm=None)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    for run in paragraph.runs:
        set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=False, all_caps=False)


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
        set_run_font(run, size_pt=BODY_FONT_SIZE_PT, bold=False, all_caps=False)

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


def apply_table_borders(table):
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, size="4")


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

                    for run in paragraph.runs:
                        set_run_font(run, size_pt=TABLE_FONT_SIZE_PT, bold=False, italic=False, all_caps=False)


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
    title = clean_spaces(m.group(3))
    replace_paragraph_text(paragraph, f"Рис. {number}. {title}")


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

            # После "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" должна быть ровно одна пустая строка
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

            # Если пустых строк больше одной — оставляем только одну
            if idx + 2 < len(paragraphs) and is_empty_paragraph(paragraphs[idx + 2]):
                remove_paragraph(paragraphs[idx + 2])
                changed = True
                break

            format_empty_paragraph(next_p)
            break

    return
    
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
            
def ensure_single_blank_after_headings(doc, body_start):
    paragraphs = doc.paragraphs

    for i in range(body_start, len(paragraphs) - 1):
        p = paragraphs[i]
        text = clean_spaces(p.text).upper()

        # Проверяем стиль
        style_name = ""
        try:
            style_name = p.style.name.lower()
        except Exception:
            pass

        is_heading1 = "heading 1" in style_name
        is_heading2 = "heading 2" in style_name

        # ----- правила -----

        # параграфы (1.1, 1.2 и т.п.)
        if is_heading2:
            ensure_one_empty_after(paragraphs, i)
            continue

        # специальные разделы
        if is_heading1 and text in {
            "ВВЕДЕНИЕ",
            "ЗАКЛЮЧЕНИЕ",
            "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
        }:
            ensure_one_empty_after(paragraphs, i)
            continue

        # главы — ничего не делаем

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
        if "СОДЕРЖАН" in t:
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

def process_document(input_path: Path, output_path: Path):
    doc = Document(str(input_path))

    body_start = find_body_start_index(doc)
    if body_start is None:
        raise RuntimeError("Не найден заголовок 'Введение'; файл пропущен из соображений безопасности.")

    set_section_margins(doc)

    split_manual_dash_lists(doc, body_start)
    split_table_captions_prepass(doc, body_start)

    for idx, paragraph in enumerate(doc.paragraphs):
        if idx < body_start:
            continue
        normalize_simple_paragraph_spaces(paragraph)
        normalize_heading2_artifacts(paragraph)

    prev_kind = None
    current_chapter_num = None
    next_paragraph_num = None

    for idx, paragraph in enumerate(doc.paragraphs):
        if idx < body_start:
            continue

        text = clean_spaces(paragraph.text)
        kind = classify_paragraph(text, prev_kind=prev_kind)

        if kind == "empty_paragraph":
            prev_kind = kind
            continue

        parsed_h1 = parse_heading1(text)
        if parsed_h1:
            if parsed_h1["kind"] == "heading1_chapter":
                current_chapter_num = parsed_h1["chapter_num"]
                next_paragraph_num = 1
                smart_repair_heading1(paragraph, text)
                kind = "heading1"
            elif parsed_h1["kind"] == "heading1_exact":
                current_chapter_num = None
                next_paragraph_num = None
                remove_paragraph_numbering(paragraph)
                kind = "heading1"

        if kind != "table_continuation" and (
            kind == "heading2"
            or auto_detect_heading2(paragraph, current_chapter_num, next_paragraph_num, prev_kind)
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

        if kind == "broken_heading2":
            repaired = smart_repair_broken_heading2(
                paragraph,
                current_chapter_num,
                next_paragraph_num,
            )
            if repaired:
                kind = "heading2"
                next_paragraph_num = (next_paragraph_num or 0) + 1

        if kind == "table_continuation":
            normalize_table_continuation_text(paragraph)

        if kind == "figure_caption":
            normalize_figure_caption_text(paragraph)

        if kind == "heading1":
            remove_paragraph_numbering(paragraph)
            format_heading1(paragraph)
        elif kind != "table_continuation" and (
            kind == "heading2"
            or auto_detect_heading2(paragraph, current_chapter_num, next_paragraph_num, prev_kind)
        ):
            remove_paragraph_numbering(paragraph)
            format_heading2(paragraph)
        elif kind == "table_caption":
            format_table_caption(paragraph)
        elif kind == "table_continuation":
            format_table_caption(paragraph)
        elif kind == "table_title":
            format_table_title(paragraph)
        elif kind == "source_line":
            format_source_line(paragraph)
        elif kind == "reference_subheading":
            format_reference_subheading(paragraph)
        elif kind == "figure_caption":
            format_figure_caption(paragraph)
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

    apply_page_breaks(doc, body_start)
    normalize_sections(doc)
    ensure_front_matter_layout(doc, body_start)
    apply_page_numbering_policy(doc)
    remove_all_italic(doc)

    run_with_pass_limit(
        "ensure_single_blank_after_references_heading_final",
        ensure_single_blank_after_references_heading,
        doc,
        body_start,
    )

    doc.save(str(output_path))
    
def remove_all_italic(doc):
    """
    Убирает курсив из всего документа
    """
    for p in doc.paragraphs:
        for r in p.runs:
            r.italic = False

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.italic = False
