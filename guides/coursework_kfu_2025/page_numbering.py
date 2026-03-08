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

    for p in footer.paragraphs:
        _clear_paragraph(p)

    if not footer.paragraphs:
        footer.add_paragraph()


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
    run._element.append(instrText)
    run._element.append(fld_char_end)


def _get_footer_paragraph(footer):
    _clear_footer_obj(footer)
    if not footer.paragraphs:
        return footer.add_paragraph()
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


def apply_page_numbering_policy(document):
    """
    - стр. 1 -> "Казань – 2026 г."
    - стр. 2 -> пусто
    - стр. 3 -> номер 3 (если есть содержание)
    """
    sections = list(document.sections)
    if not sections:
        return

    _reset_all_footer_state(document)

    body_texts = [p.text.strip().upper() for p in document.paragraphs]
    has_contents = any("СОДЕРЖАН" in t for t in body_texts)

    first_section = sections[0]
    first_section.different_first_page_header_footer = True

    p1 = _get_footer_paragraph(first_section.first_page_footer)
    _add_text_to_paragraph(p1, TITLE_FOOTER_TEXT)

    p2 = _get_footer_paragraph(first_section.footer)
    _clear_paragraph(p2)

    start_num = 3 if has_contents else 2

    for idx, section in enumerate(sections[1:], start=1):
        section.different_first_page_header_footer = True

        if idx == 1:
            _set_page_number_start(section, start_num)

        fp = _get_footer_paragraph(section.first_page_footer)
        _add_page_field_to_paragraph(fp)

        dp = _get_footer_paragraph(section.footer)
        _add_page_field_to_paragraph(dp)
