import logging
import os
import shutil
import tempfile
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
    restore_docx_if_same_page_continuation_markers,
    remove_same_page_continuation_markers_inplace,
)
from .contents_builder import rebuild_static_contents_page, strip_obsolete_toc_blocks_inplace
from .docx_utils import FormattingReport

logger = logging.getLogger(__name__)


def _flag_value(name: str) -> str:
    return os.getenv(name, "<unset>")


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

    # Phase 3: DOCX-only cleanup/normalisation.
    try:
        doc = Document(str(output_path))
        logger.info(
            "format_docx: phase3_start output_path=%s tables=%d marker_split_enabled=%s marker_split_apply=%s",
            output_path,
            len(doc.tables),
            _flag_value("KPFU_ENABLE_MARKER_SPLIT"),
            _flag_value("KPFU_APPLY_MARKER_SPLIT"),
        )
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
    except Exception:
        logger.exception("format_docx: phase3 failed, skipping (phase2 result preserved)")

    # Strip obsolete TOC artifacts (Word SDT TOC and plain-text old TOC block)
    # BEFORE the rendered-continuation backup.  By stripping first the backup
    # becomes: DOCX-only continuation markers + canonical TOC (rebuilt below).
    # If the rendered-continuation gate fires and restores this backup the user
    # always sees exactly one canonical СОДЕРЖАНИЕ — never a stale hand-made
    # copy that was present in the source or from a prior formatting run.
    try:
        report_strip = strip_obsolete_toc_blocks_inplace(output_path)
        if report_strip["sdt_removed"] or report_strip["plain_toc_removed"]:
            logger.info(
                "format_docx: pre-backup TOC strip sdt=%d plain_paragraphs=%d",
                report_strip["sdt_removed"],
                report_strip["plain_toc_removed"],
            )
    except Exception:
        logger.exception("format_docx: pre-backup TOC strip failed, continuing")

    # Rebuild the canonical contents page BEFORE taking the rendered-continuation
    # backup.  The backup therefore contains a freshly resolved СОДЕРЖАНИЕ, so
    # gate restoration gives the user a clean, correct document instead of one
    # that either lacks a TOC or retains the old hand-made copy.
    try:
        if rebuild_static_contents_page(output_path):
            logger.info("format_docx: pre-backup canonical TOC rebuilt")
    except Exception:
        logger.exception("format_docx: pre-backup canonical TOC rebuild failed, continuing")

    # Rendered table continuation entry.  The backup is taken from the
    # stripped + rebuilt state so that any gate restoration returns a document
    # that already has exactly one canonical СОДЕРЖАНИЕ and zero same-page
    # continuation violations.
    table_gate_backup_dir: Path | None = None
    table_gate_backup_path: Path | None = None
    try:
        table_gate_backup_dir = Path(tempfile.mkdtemp(prefix="kpfu_format_table_gate_"))
        table_gate_backup_path = table_gate_backup_dir / output_path.name
        shutil.copy2(output_path, table_gate_backup_path)
        n_rendered = apply_rendered_table_continuation(output_path, report=report)
        if n_rendered:
            logger.info("format_docx: rendered table continuation splits=%d", n_rendered)
    except Exception:
        logger.exception("format_docx: rendered table continuation failed")

    # Safety gate: if the rendered-continuation markers land on the same page
    # as the surrounding table segments, revert to the pre-rendered backup
    # (which already has the canonical TOC and zero same-page violations).
    _gate_restored = False
    if table_gate_backup_path is not None:
        try:
            if restore_docx_if_same_page_continuation_markers(
                output_path,
                table_gate_backup_path,
                report=report,
                context="format_docx_final",
            ):
                _gate_restored = True
                logger.warning(
                    "format_docx: final same-page continuation marker gate restored pre-rendered table state"
                )
        except Exception:
            logger.exception("format_docx: final same-page continuation marker validation failed")
        finally:
            if table_gate_backup_dir is not None:
                shutil.rmtree(table_gate_backup_dir, ignore_errors=True)

    # When the gate restored the canonical backup, DOCX-only markers that were
    # calibrated for the old student TOC layout may now be same-page in the
    # canonical layout (old TOC was typically larger by 1 page).  Remove them:
    # these markers are stale artefacts of the old layout — the table actually
    # fits on one page with the canonical TOC, so no continuation header is
    # needed.
    if _gate_restored:
        try:
            n_removed = remove_same_page_continuation_markers_inplace(output_path)
            if n_removed:
                logger.info(
                    "format_docx: removed %d same-page DOCX-only markers from canonical backup after gate",
                    n_removed,
                )
        except Exception:
            logger.exception(
                "format_docx: failed to remove same-page markers from canonical backup"
            )

    return str(output_path), report.warnings
