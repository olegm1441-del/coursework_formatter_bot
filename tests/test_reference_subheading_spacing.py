from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx import Document

from guides.coursework_kfu_2025.safe_formatter import (
    cleanup_reference_subheadings_layout,
    ensure_blank_before_reference_subheadings,
    ensure_single_blank_after_references_heading,
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
    doc.add_paragraph("статьи")
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


def main() -> int:
    ok, msg = test_reference_subheading_spacing()
    status = "PASS" if ok else "FAIL"
    print(f"[{status}] reference subheading spacing — {msg}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
