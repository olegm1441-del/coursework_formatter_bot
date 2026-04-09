import logging
from pathlib import Path

from docx import Document

from .safe_formatter import process_document
from .pagination_rules import apply_pagination_rules
from .table_continuation import apply_table_continuation, apply_rule4_empty_first_lines

logger = logging.getLogger(__name__)


def format_docx(input_path: str, output_path: str) -> str:
    input_path = Path(input_path)
    output_path = Path(output_path)

    if not input_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    if input_path.suffix.lower() != ".docx":
        raise ValueError("Поддерживаются только .docx файлы")

    # Phase 1: structural formatting
    process_document(input_path, output_path)

    if not output_path.exists():
        raise RuntimeError("Файл не был создан после Phase 1")

    # Phase 2: pagination rules (keep_with_next flags)
    try:
        doc = Document(str(output_path))
        apply_pagination_rules(doc)
        doc.save(str(output_path))
        logger.info("format_docx: phase2 pagination rules applied")
    except Exception:
        logger.exception("format_docx: phase2 failed, skipping (phase1 result preserved)")

    # Phase 3: table continuation + Rule 4 (geometry-based, no LibreOffice)
    try:
        doc = Document(str(output_path))
        n_splits = apply_table_continuation(doc)
        n_removed = apply_rule4_empty_first_lines(doc)
        if n_splits > 0 or n_removed > 0:
            doc.save(str(output_path))
            logger.info(
                "format_docx: phase3 splits=%d empty_first_lines_removed=%d",
                n_splits, n_removed,
            )
        else:
            logger.info("format_docx: phase3 no changes")
    except Exception:
        logger.exception("format_docx: phase3 failed, skipping (phase2 result preserved)")

    return str(output_path)
