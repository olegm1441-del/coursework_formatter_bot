import logging
from pathlib import Path

from docx import Document

from .safe_formatter import process_document
from .pagination_rules import apply_pagination_rules
from .table_continuation import (
    apply_table_merging,
    apply_table_continuation,
    apply_rendered_table_continuation,
    apply_rule3_table_orphan,
    apply_rule4_empty_first_lines,
    apply_rule6_figure_orphan,
    remove_empty_before_figure_captions,
)
from .docx_utils import FormattingReport

logger = logging.getLogger(__name__)


def format_docx(input_path: str, output_path: str) -> tuple[str, list[str]]:
    """
    Format *input_path* and write the result to *output_path*.

    Returns:
        (output_path_str, warnings) where *warnings* is a (possibly empty)
        list of short Russian strings describing issues the user should
        check manually (e.g. tables that could not be auto-split).
    """
    input_path = Path(input_path)
    output_path = Path(output_path)

    if not input_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    if input_path.suffix.lower() != ".docx":
        raise ValueError("Поддерживаются только .docx файлы")

    report = FormattingReport()

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

    # Phase 3: DOCX-only cleanup/normalisation, then rendered table split entry.
    try:
        doc = Document(str(output_path))
        n_merged  = apply_table_merging(doc)
        n_tables  = apply_table_continuation(doc, report=report)
        n_rule3   = apply_rule3_table_orphan(doc)
        n_rule4   = apply_rule4_empty_first_lines(doc)
        n_rule6   = apply_rule6_figure_orphan(doc)
        n_gap     = remove_empty_before_figure_captions(doc)
        if n_merged > 0 or n_tables > 0 or n_rule3 > 0 or n_rule4 > 0 or n_rule6 > 0 or n_gap > 0:
            doc.save(str(output_path))
            logger.info(
                "format_docx: phase3 merged=%d tables=%d rule3=%d rule4=%d rule6=%d gap=%d",
                n_merged, n_tables, n_rule3, n_rule4, n_rule6, n_gap,
            )
        else:
            logger.info("format_docx: phase3 no changes")

        n_rendered = apply_rendered_table_continuation(output_path, report=report)
        if n_rendered:
            logger.info("format_docx: rendered table continuation splits=%d", n_rendered)
    except Exception:
        logger.exception("format_docx: phase3 failed, skipping (phase2 result preserved)")

    return str(output_path), report.warnings
