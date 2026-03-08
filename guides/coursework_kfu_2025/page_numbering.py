from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


TITLE_FOOTER_TEXT = "Казань – 2026 г."


def _clear_paragraph(paragraph):
    p = paragraph._element
    for child in list(p):
        p.remove(child)


def _clear_footer(footer):
    for p in footer.paragraphs:
        _clear_paragraph(p)


def _set_run_font(run):
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0, 0, 0)

    rPr = run._element.get_or_add_rPr()

    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)

    rFonts.set(qn("w:ascii"), "Times New Roman")
    rFonts.set(qn("w:hAnsi"), "Times New Roman")
    rFonts.set(qn("w:cs"), "Times New Roman")
    rFonts.set(qn("w:eastAsia"), "Times New Roman")

    sz = rPr.find(qn("w:sz"))
    if sz is None:
        sz = OxmlElement("w:sz")
        rPr.append(sz)
    sz.set(qn("w:val"), "24")

    szCs = rPr.find(qn("w:szCs"))
    if szCs is None:
        szCs = OxmlElement("w:szCs")
        rPr.append(szCs)
    szCs.set(qn("w:val"), "24")


def _add_text_to_paragraph(paragraph, text):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    _set_run_font(run)


def _add_page_field_to_paragraph(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = paragraph.add_run()
    _set_run_font(run)

    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"

    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")

    run._element.append(fld_char_begin)
    run._element.append(instr_text)
    run._element.append(fld_char_end)


def _prepare_footer(section, use_first_page=False):
    footer = section.first_page_footer if use_first_page else section.footer
    footer.is_linked_to_previous = False
    _clear_footer(footer)

    if not footer.paragraphs:
        p = footer.add_paragraph()
    else:
        p = footer.paragraphs[0]
    _clear_paragraph(p)
    return footer, p


def apply_page_numbering_policy(document):
    """
    Логика:
    - секция 1:
        1-я страница -> "Казань – 2026 г."
        остальные страницы секции -> пусто
    - секция 2 и далее:
        первая страница секции -> номер страницы
        остальные страницы секции -> номер страницы
    """
    if not document.sections:
        return

    sections = list(document.sections)

    # --- Первая секция: титул + содержание ---
    first_section = sections[0]
    first_section.different_first_page_header_footer = True

    _, p1 = _prepare_footer(first_section, use_first_page=True)
    _add_text_to_paragraph(p1, TITLE_FOOTER_TEXT)

    _, p2 = _prepare_footer(first_section, use_first_page=False)
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # специально ничего не пишем -> страница содержания пустая внизу

    # --- Все следующие секции: обычная нумерация ---
    for section in sections[1:]:
        section.different_first_page_header_footer = True

        _, fp = _prepare_footer(section, use_first_page=True)
        _add_page_field_to_paragraph(fp)

        _, dp = _prepare_footer(section, use_first_page=False)
        _add_page_field_to_paragraph(dp)
