from __future__ import annotations

import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx import Document
from docx.oxml.ns import qn

from guides.coursework_kfu_2025.safe_formatter import (
    canonical_numbered_reference_subheading_text,
    canonical_reference_subheading_text,
    cleanup_reference_subheadings_layout,
    ensure_blank_before_reference_subheadings,
    ensure_single_blank_after_references_heading,
    process_document,
)


def _paragraph_texts(doc: Document) -> list[str]:
    return [p.text for p in doc.paragraphs]


def test_reference_subheading_spacing() -> tuple[bool, str]:
    doc = Document()
    doc.add_paragraph("Список использованных источников")
    doc.add_paragraph("нормативные правовые акты")
    doc.add_paragraph("1. Федеральный закон ...")
    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph("3. статьи")
    doc.add_paragraph("2. Иванов И.И. Статья ...")
    doc.add_paragraph("3. Петров П.П. Источник ...")
    doc.add_paragraph("ДИССЕРТАЦИИ")
    doc.add_paragraph("4. Сидоров С.С. Диссертация ...")

    body_start = 0
    ensure_blank_before_reference_subheadings(doc, body_start)
    ensure_single_blank_after_references_heading(doc, body_start)
    cleanup_reference_subheadings_layout(doc, body_start)

    expected = [
        "Список использованных источников",
        "",
        "Нормативные правовые акты",
        "1. Федеральный закон ...",
        "",
        "Статьи",
        "2. Иванов И.И. Статья ...",
        "3. Петров П.П. Источник ...",
        "",
        "Диссертации",
        "4. Сидоров С.С. Диссертация ...",
    ]
    actual = _paragraph_texts(doc)
    if actual != expected:
        return False, f"unexpected paragraph layout:\nexpected={expected!r}\nactual={actual!r}"

    return True, "reference subheadings have exactly one blank before them"


def test_reference_subheading_detection_is_strict() -> tuple[bool, str]:
    cases = [
        "1. Статьи",
        "• Статьи",
        "- Статьи",
        "Статьи и монографии",
        "Материалы интернет-сайтов: сайты",
        "статьи в периодических изданиях",
    ]
    for text in cases:
        if canonical_reference_subheading_text(text) is not None:
            return False, f"false reference subheading detected: {text!r}"

    if canonical_reference_subheading_text("статьи") != "Статьи":
        return False, "exact case-insensitive subheading was not detected"
    if (
        canonical_reference_subheading_text("статьи в периодических изданиях и сборниках")
        != "Статьи в периодических изданиях и сборниках"
    ):
        return False, "new exact reference subheading was not detected"
    if (
        canonical_numbered_reference_subheading_text("1. диссертации, авторефераты диссертаций")
        != "Диссертации, авторефераты диссертаций"
    ):
        return False, "new numbered reference subheading was not recovered"

    return True, "reference subheading detection is exact-match only"


def test_numbered_reference_entries_are_not_headings() -> tuple[bool, str]:
    false_heading = (
        "1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ КОММУНИКАЦИОННОЙ ПОЛИТИКИ "
        "В СИСТЕМЕ МАРКЕТИНГА ПРЕДПРИЯТИЯ"
    )

    doc = Document()
    doc.add_paragraph("ВВЕДЕНИЕ")
    doc.add_paragraph("Краткий текст введения.")
    doc.add_paragraph("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph("1. Монографии и учебники")
    doc.add_paragraph("2. Бельзецкий А. И. Маркетология: монография.")
    doc.add_paragraph(false_heading)
    doc.add_paragraph(
        "4. Закон РФ от 07.02.1992 № 2300-1 «О защите прав потребителей» "
        "[Электронный ресурс]. — URL: https://example.com/(дата обращения: 06.03.2026)."
    )
    doc.add_paragraph("")
    doc.add_paragraph("5. статьи")
    doc.add_paragraph("6. Иванов И. И. Название статьи.")

    with tempfile.TemporaryDirectory() as tmp:
        input_path = Path(tmp) / "in.docx"
        output_path = Path(tmp) / "out.docx"
        doc.save(str(input_path))

        process_document(input_path, output_path)
        out_doc = Document(str(output_path))

    texts = _paragraph_texts(out_doc)
    try:
        refs_idx = texts.index("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    except ValueError:
        return False, "references heading missing after formatting"

    if texts[refs_idx + 1] != "":
        return False, "missing single blank after references heading"
    if texts[refs_idx + 2] != "Монографии и учебники":
        return False, f"numbered block heading was not recovered: {texts[refs_idx + 2]!r}"
    if not texts[refs_idx + 3].startswith("1. "):
        return False, f"first real source lost numbering: {texts[refs_idx + 3]!r}"
    if not texts[refs_idx + 4].startswith("2. "):
        return False, f"numbered reference entry lost numbering: {texts[refs_idx + 4]!r}"
    if "коммуникационной политики" not in texts[refs_idx + 4].lower():
        return False, f"numbered reference entry text changed unexpectedly: {texts[refs_idx + 4]!r}"
    if texts[refs_idx + 5].startswith("3. ") is False:
        return False, "next numbered reference entry changed unexpectedly"
    if "https://example.com/ (дата обращения" not in texts[refs_idx + 5]:
        return False, f"URL spacing was not normalized: {texts[refs_idx + 5]!r}"
    if texts[refs_idx + 6] != "":
        return False, "missing single blank before real reference subheading"
    if texts[refs_idx + 7] != "Статьи":
        return False, f"numbered reference subheading was not canonicalized: {texts[refs_idx + 7]!r}"
    if texts[refs_idx + 8] == "":
        return False, "unexpected blank after reference subheading"

    block_heading_para = out_doc.paragraphs[refs_idx + 2]
    if block_heading_para.alignment != 1:
        return False, "recovered block heading is not centered"

    false_heading_para = out_doc.paragraphs[refs_idx + 4]
    style_name = (false_heading_para.style.name or "").lower()
    if "heading" in style_name or "заголовок" in style_name:
        return False, f"numbered reference entry got heading style: {false_heading_para.style.name!r}"

    for offset in (3, 4, 5):
        pPr = out_doc.paragraphs[refs_idx + offset]._element.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        attrs = ind.attrib if ind is not None else {}
        if attrs.get(qn("w:left")) != "0":
            return False, f"reference entry has non-zero left indent: {attrs}"
        if attrs.get(qn("w:firstLine")) != "709":
            return False, f"reference entry does not have first-line indent 1.25 cm: {attrs}"
        if attrs.get(qn("w:hanging")) is not None:
            return False, f"reference entry still has hanging indent: {attrs}"

    hyperlinks = out_doc.paragraphs[refs_idx + 5]._element.findall(".//" + qn("w:hyperlink"))
    if not hyperlinks:
        return False, "plain URL was not converted to a DOCX hyperlink"

    return True, "numbered reference entries stay body text inside references"


def main() -> int:
    tests = [
        ("reference subheading spacing", test_reference_subheading_spacing),
        ("strict reference subheading detection", test_reference_subheading_detection_is_strict),
        ("numbered reference entries", test_numbered_reference_entries_are_not_headings),
    ]
    failed = 0
    for name, fn in tests:
        ok, msg = fn()
        status = "PASS" if ok else "FAIL"
        print(f"[{status}] {name} — {msg}")
        if not ok:
            failed += 1
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
