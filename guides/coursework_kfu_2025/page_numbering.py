from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


TITLE_FOOTER_TEXT = "Казань – 2026 г."


def _clear_paragraph(paragraph):
    p = paragraph._element
    for child in list(p):
        p.remove(child)


def _clear_footer_obj(footer):
    try:
        footer.is_linked_to_previous = False
    except Exception:
        pass

    root = footer._element
    for child in list(root):
        root.remove(child)

    p = OxmlElement("w:p")
    root.append(p)


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
    _clear_paragraph(paragraph)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    _set_run_font(run)


def _add_page_field_to_paragraph(paragraph):
    _clear_paragraph(paragraph)
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


def _get_footer_paragraph(footer):
    _clear_footer_obj(footer)
    return footer.paragraphs[0]


def _remove_pg_num_type(section):
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn("w:pgNumType"))
    if pgNumType is not None:
        sectPr.remove(pgNumType)


def _set_page_number_start(section, start_value):
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn("w:pgNumType"))
    if pgNumType is None:
        pgNumType = OxmlElement("w:pgNumType")
        sectPr.append(pgNumType)
    pgNumType.set(qn("w:start"), str(start_value))


def _reset_all_footer_state(document):
    try:
        document.settings.odd_and_even_pages_header_footer = False
    except Exception:
        pass

    for section in document.sections:
        section.different_first_page_header_footer = True

        _remove_pg_num_type(section)

        _clear_footer_obj(section.footer)
        _clear_footer_obj(section.first_page_footer)

        try:
            _clear_footer_obj(section.even_page_footer)
        except Exception:
            pass


def _blank_section(section):
    section.different_first_page_header_footer = True

    p1 = _get_footer_paragraph(section.first_page_footer)
    _clear_paragraph(p1)

    p2 = _get_footer_paragraph(section.footer)
    _clear_paragraph(p2)

    _remove_pg_num_type(section)


def _number_section(section, start_value=None):
    section.different_first_page_header_footer = True

    if start_value is not None:
        _set_page_number_start(section, start_value)

    fp = _get_footer_paragraph(section.first_page_footer)
    _add_page_field_to_paragraph(fp)

    dp = _get_footer_paragraph(section.footer)
    _add_page_field_to_paragraph(dp)


def apply_page_numbering_policy(document):
    """
    Целевая логика:
    - 3 секции: титул / содержание / основная часть
    - 2 секции: титул / основная часть
    """
    sections = list(document.sections)
    if not sections:
        return

    _reset_all_footer_state(document)

    if len(sections) >= 3:
        first_section = sections[0]
        first_section.different_first_page_header_footer = True
        p1 = _get_footer_paragraph(first_section.first_page_footer)
        _add_text_to_paragraph(p1, TITLE_FOOTER_TEXT)
        p2 = _get_footer_paragraph(first_section.footer)
        _clear_paragraph(p2)

        _blank_section(sections[1])
        _number_section(sections[2], start_value=3)
        for section in sections[3:]:
            _number_section(section)
        return

    first_section = sections[0]
    first_section.different_first_page_header_footer = True
    p1 = _get_footer_paragraph(first_section.first_page_footer)
    _add_text_to_paragraph(p1, TITLE_FOOTER_TEXT)
    p2 = _get_footer_paragraph(first_section.footer)
    _clear_paragraph(p2)

    if len(sections) >= 2:
        _number_section(sections[1], start_value=2)
        for section in sections[2:]:
            _number_section(section)
