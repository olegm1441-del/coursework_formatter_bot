from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from guides.coursework_kfu_2025.table_split_prototype import prototype_split_table_copy


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run isolated prototype DOCX table split on a temp copy."
    )
    parser.add_argument("docx_path", help="Path to source .docx")
    parser.add_argument("--table-index", type=int, required=True)
    parser.add_argument("--split-before-row", type=int, required=True)
    parser.add_argument("--header-rows", type=int, default=1)
    parser.add_argument("--numbered-header", action="store_true")
    parser.add_argument("--appendix", action="store_true")
    parser.add_argument("--keep-temp", action="store_true")
    args = parser.parse_args()

    result = prototype_split_table_copy(
        args.docx_path,
        args.table_index,
        args.split_before_row,
        header_rows=args.header_rows,
        numbered_header=args.numbered_header,
        appendix_table=args.appendix,
        keep_temp=args.keep_temp,
    )

    print(f"output_docx_path: {result.output_docx_path}")
    print(f"workdir_path: {result.workdir_path}")
    print(f"table_index: {result.table_index}")
    print(f"second_table_index: {result.second_table_index}")
    print(f"total_tables_before: {result.total_tables_before}")
    print(f"total_tables_after: {result.total_tables_after}")
    print(f"original_rows_count: {result.original_rows_count}")
    print(f"rows_in_first_table: {result.first_table_rows_count}")
    print(f"rows_in_second_table: {result.second_table_rows_count}")
    print(f"numbered_header_enabled: {result.numbered_header_enabled}")
    print(f"numbered_row_reused: {result.numbered_row_reused}")
    print(f"column_count: {result.column_count}")
    print(f"continuation_paragraph_inserted: {result.continuation_paragraph_inserted}")
    print(f"continuation_text: {result.continuation_text}")
    if result.source_note_after_second is None:
        print("source_note_after_second: n/a")
    else:
        print(f"source_note_after_second: {result.source_note_after_second}")
    print(f"source_note_text: {result.source_note_text}")
    print(f"diagnostics: {result.diagnostics}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
