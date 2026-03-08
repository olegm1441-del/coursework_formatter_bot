from docx.enum.text import WD_ALIGN_PARAGRAPH


def apply_page_numbering_policy(document):
    """
    Минимальная безопасная политика нумерации страниц.
    Если в документе есть секции, нумерация ставится по центру футера.
    """

    for section in document.sections:
        footer = section.footer

        if not footer.paragraphs:
            p = footer.add_paragraph()
        else:
            p = footer.paragraphs[0]

        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = p.add_run()
        from docx.shared import Pt

        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        fld = run._element
