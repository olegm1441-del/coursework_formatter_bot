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
    rFonts = rPr.rFonts
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)

    rFonts.set(qn("w:ascii"), "Times New Roman")
    rFonts.set(qn("w:hAnsi"), "Times New Roman")
    rFonts.set(qn("w:cs"), "Times New Roman")
    rFonts.set(qn("w:eastAsia"), "Times New Roman")


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


def apply_page_numbering_policy(document):
    """
    Логика:
    - 1 страница: в футере по центру "Казань – 2026 г."
    - 2 страница: пусто
    - с 3 страницы: номер страницы по центру внизу
    """

    if not document.sections:
        return

    first_section = document.sections[0]

    # Разные колонтитулы для первой страницы секции
    first_section.different_first_page_header_footer = True

    # Отвязываем футеры от возможных предыдущих секций
    first_section.footer.is_linked_to_previous = False
    first_section.first_page_footer.is_linked_to_previous = False

    # Полностью очищаем футеры
    _clear_footer(first_section.first_page_footer)
    _clear_footer(first_section.footer)

    # --- 1-я страница ---
    # Титул: "Казань – 2026 г."
    first_page_footer = first_section.first_page_footer
    if not first_page_footer.paragraphs:
        p1 = first_page_footer.add_paragraph()
    else:
        p1 = first_page_footer.paragraphs[0]
    _clear_paragraph(p1)
    _add_text_to_paragraph(p1, TITLE_FOOTER_TEXT)

    # --- 2-я и далее страницы этой секции ---
    # По умолчанию обычный footer применяется ко 2-й, 3-й и т.д. странице.
    # Чтобы 2-я была пустой, а с 3-й началась нумерация,
    # вставляем сначала пустой абзац, а затем PAGE.
    default_footer = first_section.footer
    if not default_footer.paragraphs:
        p2 = default_footer.add_paragraph()
    else:
        p2 = default_footer.paragraphs[0]
    _clear_paragraph(p2)
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Если есть ещё старые абзацы — очищаем и их тоже
    for extra_p in default_footer.paragraphs[1:]:
        _clear_paragraph(extra_p)

    # Создаём отдельный абзац с полем PAGE
    page_paragraph = default_footer.add_paragraph()
    _clear_paragraph(page_paragraph)
    _add_page_field_to_paragraph(page_paragraph)

    # Важный момент:
    # Word сам не умеет по одному футеру показать пусто только на 2-й странице
    # и номер с 3-й без секционного разрыва.
    # Поэтому для железобетонной логики нужен разрыв секции перед "ВВЕДЕНИЕ".
