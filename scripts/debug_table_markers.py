from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from guides.coursework_kfu_2025.table_markers import diagnose_all_tables


def _format_row_pages(row_pages: dict[int, int]) -> str:
    if not row_pages:
        return "{}"
    parts = [f"{row}:{page}" for row, page in sorted(row_pages.items())]
    return "{ " + ", ".join(parts) + " }"


def _format_page_spans(page_spans) -> list[str]:
    if not page_spans:
        return ["none"]
    lines = []
    for span in page_spans:
        lines.append(f"rows {span.start_row}–{span.end_row} -> page {span.page_num}")
    return lines


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Diagnose marker-based table row->page mapping for every table in a DOCX."
    )
    parser.add_argument("docx_path", help="Path to source .docx")
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Preserve instrumented DOCX/PDF artifacts for manual inspection.",
    )
    args = parser.parse_args()

    docx_path = Path(args.docx_path)
    if not docx_path.exists():
        print(f"ERROR: DOCX not found: {docx_path}")
        return 2

    diagnostics = diagnose_all_tables(docx_path, keep_temp=args.keep_temp)
    print(f"Document: {docx_path}")
    print(f"Tables found: {len(diagnostics)}")

    candidate_indexes = [d.table_index for d in diagnostics if d.candidate_for_split]
    for item in diagnostics:
        print("")
        print(f"table_index: {item.table_index}")
        print(f"rows_count: {item.rows_count}")
        print(f"pages_detected: {item.pages_detected}")
        print(f"row_pages: {_format_row_pages(item.row_pages)}")
        print(f"missing_rows: {item.missing_rows}")
        print(f"duplicate_rows: {item.duplicate_rows}")
        print(f"appendix_table: {'yes' if item.appendix_table else 'no'}")
        print(f"caption_detected: {'yes' if item.caption_detected else 'no'}")
        print(f"has_standard_table_caption: {'yes' if item.has_standard_table_caption else 'no'}")
        if item.preceding_paragraph_text is not None:
            print(f"preceding_paragraph_text: {item.preceding_paragraph_text}")
        print(f"candidate_for_split: {'yes' if item.candidate_for_split else 'no'}")
        print(f"marker_font_size_pt: {item.marker_font_size_pt}")
        if item.error_message is not None:
            print(f"error: {item.error_message}")
        print("page_spans:")
        for line in _format_page_spans(item.page_spans):
            print(f"  {line}")
        if item.instrumented_docx_path is not None:
            print(f"instrumented_docx_path: {item.instrumented_docx_path}")
        if item.pdf_path is not None:
            print(f"pdf_path: {item.pdf_path}")

    print("")
    print(f"candidate_tables: {candidate_indexes}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
