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
    repair_manual_chain_overflow_before_marker,
    normalize_exact_grid_same_page_repeated_fragments_inplace,
)
from .contents_builder import rebuild_static_contents_page, strip_obsolete_toc_blocks_inplace
from .docx_utils import FormattingReport
from .layout_render import render_docx_to_pdf
from .pdf_layout_analyzer import analyze_pdf_lines
from .rendered_table_validation import (
    RenderedContinuationViolation,
    build_rendered_table_identities,
    validate_rendered_continuations,
)

logger = logging.getLogger(__name__)


def _flag_value(name: str) -> str:
    return os.getenv(name, "<unset>")


def _rendered_continuation_violations_for_docx(
    docx_path: Path,
    *,
    source_table_identities: list | None = None,
) -> list[RenderedContinuationViolation]:
    pdf_path: Path | None = None
    try:
        pdf_path = render_docx_to_pdf(Path(docx_path))
        pdf_lines = analyze_pdf_lines(pdf_path)
        doc = Document(str(docx_path))
        identities = build_rendered_table_identities(doc)
        return validate_rendered_continuations(
            pdf_lines,
            identities,
            source_table_identities=source_table_identities,
        )
    finally:
        if pdf_path is not None:
            shutil.rmtree(pdf_path.parent, ignore_errors=True)


def _rendered_continuation_warning(
    violation: RenderedContinuationViolation,
) -> str | None:
    table_num = violation.table_num or f"#{violation.table_index}"
    if (
        violation.violation_type == "missing_continuation_marker"
        and violation.confidence == "high"
    ):
        return (
            f"Проверьте перенос таблицы {table_num}: "
            f"стр. {violation.page} без маркера продолжения."
        )
    if (
        violation.violation_type == "suspected_missing_continuation_marker"
        and violation.confidence == "medium"
    ):
        return (
            f"Проверьте возможный перенос таблицы {table_num}: "
            f"стр. {violation.page} без маркера продолжения."
        )
    if (
        violation.violation_type == "late_continuation_marker"
        and violation.confidence == "high"
    ):
        return (
            f"Проверьте перенос таблицы {table_num}: "
            f"на стр. {violation.page} строки идут до маркера продолжения."
        )
    if (
        violation.violation_type == "same_page_repeated_fragment"
        and violation.confidence == "high"
    ):
        return (
            f"Проверьте таблицу {table_num}: "
            f"на стр. {violation.page} повторный фрагмент виден на той же странице."
        )
    if (
        violation.violation_type == "same_page_adjacent_fragment"
        and violation.confidence == "high"
    ):
        return (
            f"Проверьте таблицу {table_num}: "
            f"на стр. {violation.page} соседний фрагмент выглядит как часть той же таблицы."
        )
    if violation.violation_type == "missing_or_late_continuation_marker":
        return (
            f"Проверьте перенос таблицы {table_num}: "
            f"стр. {violation.page} без корректного маркера продолжения."
        )
    if violation.violation_type == "ambiguous_adjacent_tables":
        return (
            f"Проверьте таблицу {table_num}: "
            "соседние таблицы похожи на фрагменты, но связь не доказана."
        )
    if violation.violation_type == "source_bad_duplicated_content_rows":
        return (
            f"Проверьте таблицу {table_num}: "
            "в исходном файле есть повторяющиеся содержательные строки."
        )
    return None


def _append_rendered_continuation_warnings(
    report: FormattingReport,
    violations: list[RenderedContinuationViolation],
) -> None:
    for violation in violations:
        message = _rendered_continuation_warning(violation)
        if message is None:
            continue
        report.warn(message)
        logger.warning(
            "rendered_continuation_violation table_num=%s table_index=%s page=%s type=%s confidence=%s evidence=%s",
            violation.table_num,
            violation.table_index,
            violation.page,
            violation.violation_type,
            violation.confidence,
            violation.evidence,
        )


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
    try:
        source_table_identities = build_rendered_table_identities(Document(str(input_path)))
    except Exception:
        source_table_identities = None

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
            n_removed = remove_same_page_continuation_markers_inplace(
                output_path,
                report=None,
            )
            if n_removed:
                logger.info(
                    "format_docx: removed %d same-page DOCX-only markers from canonical backup after gate",
                    n_removed,
                )
        except Exception:
            logger.exception(
                "format_docx: failed to remove same-page markers from canonical backup"
            )

    try:
        n_same_page_exact = normalize_exact_grid_same_page_repeated_fragments_inplace(
            output_path,
            source_docx_path=input_path,
            report=report,
        )
        if n_same_page_exact:
            logger.info(
                "format_docx: normalized %d exact-grid same-page table fragment(s)",
                n_same_page_exact,
            )
    except Exception:
        logger.exception("format_docx: exact-grid same-page fragment normalization failed")

    try:
        rendered_violations = _rendered_continuation_violations_for_docx(
            output_path,
            source_table_identities=source_table_identities,
        )
        if rendered_violations:
            n_repaired = repair_manual_chain_overflow_before_marker(
                output_path,
                rendered_violations,
                report=report,
            )
            if n_repaired:
                logger.info(
                    "format_docx: repaired %d manual-chain overflow continuation(s)",
                    n_repaired,
                )
                try:
                    remove_same_page_continuation_markers_inplace(output_path, report=None)
                except Exception:
                    logger.exception(
                        "format_docx: same-page marker cleanup after manual-chain repair failed"
                    )
                rendered_violations = _rendered_continuation_violations_for_docx(
                    output_path,
                    source_table_identities=source_table_identities,
                )
        if rendered_violations:
            _append_rendered_continuation_warnings(report, rendered_violations)
            hard_count = sum(
                1
                for violation in rendered_violations
                if violation.violation_type == "missing_continuation_marker"
                and violation.confidence == "high"
            )
            suspected_count = sum(
                1
                for violation in rendered_violations
                if violation.violation_type == "suspected_missing_continuation_marker"
                and violation.confidence == "medium"
            )
            logger.warning(
                "format_docx: final rendered continuation validation review_needed hard=%d suspected=%d",
                hard_count,
                suspected_count,
            )
    except Exception:
        logger.exception("format_docx: final rendered continuation validation failed")
        report.warn(
            "Автопроверка переносов таблиц по PDF не выполнена. Проверьте таблицы вручную."
        )

    return str(output_path), report.warnings
