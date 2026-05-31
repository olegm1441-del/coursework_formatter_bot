from __future__ import annotations

import json
import shutil
from dataclasses import asdict, dataclass
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


@dataclass(frozen=True)
class SplitDecision:
    eligible: bool
    reason: str


@dataclass(frozen=True)
class RowRole:
    row_index: int
    role: str


@dataclass(frozen=True)
class NumericRowSafety:
    safe: bool
    reason: str | None = None


@dataclass(frozen=True)
class ContinuationMarkerDiagnostic:
    text: str
    marker_page: int | None
    marker_kind: str
    verdict: str
    same_page_violation: bool | None
    previous_table_page: int | None
    following_table_page: int | None
    confidence: str


@dataclass(frozen=True)
class DiagnosticStage:
    name: str
    docx_path: str
    tables: list[dict]


@dataclass(frozen=True)
class DiagnosticRunResult:
    artifact_root: str
    summary_json_path: str
    stages: list[DiagnosticStage]


def _cell_texts(row_xml) -> list[str]:
    out = []
    for cell in row_xml.findall(qn("w:tc")):
        text = " ".join((node.text or "") for node in cell.findall(".//" + qn("w:t")))
        out.append(" ".join(text.split()))
    return out


def _is_numeric_row(values: list[str]) -> bool:
    return len(values) >= 2 and values == [str(i) for i in range(1, len(values) + 1)]


def _row_has_gridspan(row_xml) -> bool:
    return any(cell.find(qn("w:tcPr")) is not None and cell.find(qn("w:tcPr")).find(qn("w:gridSpan")) is not None for cell in row_xml.findall(qn("w:tc")))


def _row_has_vmerge_continue(row_xml) -> bool:
    for cell in row_xml.findall(qn("w:tc")):
        tc_pr = cell.find(qn("w:tcPr"))
        if tc_pr is None:
            continue
        vm = tc_pr.find(qn("w:vMerge"))
        if vm is not None and vm.get(qn("w:val")) in {None, "continue"}:
            return True
    return False


def classify_table_rows(tbl_xml, *, header_rows: int = 1) -> list[RowRole]:
    roles: list[RowRole] = []
    for idx, row in enumerate(tbl_xml.findall(qn("w:tr"))):
        values = _cell_texts(row)
        joined = " ".join(values).strip().lower()
        if _is_numeric_row(values):
            role = "numeric"
        elif joined.startswith(("источник:", "примечание:")):
            role = "source_note"
        elif idx < header_rows:
            role = "header"
        else:
            role = "body"
        roles.append(RowRole(idx, role))
    return roles


def numeric_row_synthesis_safety(tbl_xml, *, header_rows: int = 1) -> NumericRowSafety:
    rows = tbl_xml.findall(qn("w:tr"))
    if not rows or header_rows <= 0:
        return NumericRowSafety(False, "unsafe_numeric_row_synthesis")
    for row in rows[:header_rows]:
        if _row_has_gridspan(row):
            return NumericRowSafety(False, "unsafe_numeric_row_synthesis")
    return NumericRowSafety(True)


def evaluate_split_eligibility(
    tbl_xml,
    *,
    rendered_row_pages: dict[int, int] | None,
    split_before_row: int,
    header_rows: int = 1,
    min_body_rows_per_fragment: int = 1,
) -> SplitDecision:
    if rendered_row_pages is None:
        return SplitDecision(False, "render_boundary_unmapped")
    if tbl_xml.find(qn("w:tblGrid")) is None:
        return SplitDecision(False, "malformed_grid")
    rows = tbl_xml.findall(qn("w:tr"))
    if split_before_row < header_rows + min_body_rows_per_fragment:
        return SplitDecision(False, "fragment_too_small")
    if split_before_row >= len(rows):
        return SplitDecision(False, "fragment_too_small")
    if _row_has_vmerge_continue(rows[split_before_row]):
        return SplitDecision(False, "unsafe_vmerge_boundary")
    roles = classify_table_rows(tbl_xml, header_rows=header_rows)
    if roles[split_before_row].role == "source_note":
        return SplitDecision(False, "source_note_boundary_uncertain")
    first_pages = {rendered_row_pages.get(i) for i in range(0, split_before_row)}
    second_pages = {rendered_row_pages.get(i) for i in range(split_before_row, len(rows))}
    if None in first_pages or None in second_pages or not first_pages or not second_pages:
        return SplitDecision(False, "render_boundary_unmapped")
    if max(first_pages) >= min(second_pages):
        return SplitDecision(False, "render_boundary_unmapped")
    return SplitDecision(True, "eligible")


def _xml(el) -> str | None:
    return el.xml if el is not None else None


def _tbl_pr(tbl_xml):
    return tbl_xml.find(qn("w:tblPr"))


def _tbl_w(tbl_xml) -> dict:
    tbl_pr = _tbl_pr(tbl_xml)
    node = tbl_pr.find(qn("w:tblW")) if tbl_pr is not None else None
    return {
        "w": node.get(qn("w:w")) if node is not None else None,
        "type": node.get(qn("w:type")) if node is not None else None,
    }


def _tbl_layout(tbl_xml) -> str | None:
    tbl_pr = _tbl_pr(tbl_xml)
    node = tbl_pr.find(qn("w:tblLayout")) if tbl_pr is not None else None
    return node.get(qn("w:type")) if node is not None else None


def _grid_widths(tbl_xml) -> list[str | None]:
    grid = tbl_xml.find(qn("w:tblGrid"))
    if grid is None:
        return []
    return [col.get(qn("w:w")) for col in grid.findall(qn("w:gridCol"))]


def _tcw(tbl_xml) -> list[dict]:
    out = []
    for cell in tbl_xml.findall(".//" + qn("w:tc")):
        tc_pr = cell.find(qn("w:tcPr"))
        node = tc_pr.find(qn("w:tcW")) if tc_pr is not None else None
        out.append({
            "w": node.get(qn("w:w")) if node is not None else None,
            "type": node.get(qn("w:type")) if node is not None else None,
        })
    return out


def _geometry_policy(tbl_xml) -> str:
    try:
        from guides.coursework_kfu_2025.table_continuation import classify_table_geometry_policy

        return classify_table_geometry_policy(tbl_xml)
    except Exception:
        return "unknown"


def snapshot_table(tbl_xml, *, table_index: int = 0) -> dict:
    rows = tbl_xml.findall(qn("w:tr"))
    return {
        "table_index": table_index,
        "tblPr_xml": _xml(_tbl_pr(tbl_xml)),
        "tblGrid_xml": _xml(tbl_xml.find(qn("w:tblGrid"))),
        "gridCol_widths": _grid_widths(tbl_xml),
        "row_count": len(rows),
        "raw_tc_counts": [len(row.findall(qn("w:tc"))) for row in rows],
        "tblW": _tbl_w(tbl_xml),
        "tblLayout": _tbl_layout(tbl_xml),
        "tcW": _tcw(tbl_xml),
        "gridSpan": [cell.find(qn("w:tcPr")).find(qn("w:gridSpan")).get(qn("w:val"))
                     for cell in tbl_xml.findall(".//" + qn("w:tc"))
                     if cell.find(qn("w:tcPr")) is not None and cell.find(qn("w:tcPr")).find(qn("w:gridSpan")) is not None],
        "vMerge": [cell.find(qn("w:tcPr")).find(qn("w:vMerge")).get(qn("w:val"))
                   for cell in tbl_xml.findall(".//" + qn("w:tc"))
                   if cell.find(qn("w:tcPr")) is not None and cell.find(qn("w:tcPr")).find(qn("w:vMerge")) is not None],
        "geometry_policy": _geometry_policy(tbl_xml),
    }


def diff_table_geometry(before: dict, after: dict) -> dict:
    keys = (
        "tblPr_xml", "tblGrid_xml", "gridCol_widths", "row_count",
        "raw_tc_counts", "tblW", "tblLayout", "tcW", "gridSpan", "vMerge",
        "geometry_policy",
    )
    return {key: {"before": before.get(key), "after": after.get(key)} for key in keys if before.get(key) != after.get(key)}


def _stage(name: str, path: Path, previous: DiagnosticStage | None = None) -> DiagnosticStage:
    doc = Document(str(path))
    tables = []
    prev_tables = previous.tables if previous is not None else []
    for idx, table in enumerate(doc.tables):
        snap = snapshot_table(table._tbl, table_index=idx)
        if idx < len(prev_tables):
            snap["geometry_diff_from_previous_stage"] = diff_table_geometry(prev_tables[idx], snap)
        else:
            snap["geometry_diff_from_previous_stage"] = {}
        tables.append(snap)
    return DiagnosticStage(name=name, docx_path=str(path), tables=tables)


def run_table_engine_diagnostics(
    docx_path: Path,
    *,
    artifact_root: Path,
    render: bool = False,
) -> DiagnosticRunResult:
    from guides.coursework_kfu_2025.safe_formatter import process_document
    from guides.coursework_kfu_2025.table_continuation import apply_table_continuation

    artifact_root.mkdir(parents=True, exist_ok=True)
    source = artifact_root / "00_source.docx"
    safe = artifact_root / "01_after_safe_formatter.docx"
    phase3 = artifact_root / "02_after_table_continuation.docx"
    final = artifact_root / "03_final.docx"

    shutil.copy2(docx_path, source)
    process_document(source, safe)
    shutil.copy2(safe, phase3)
    doc = Document(str(phase3))
    apply_table_continuation(doc)
    doc.save(str(phase3))
    shutil.copy2(phase3, final)

    stages: list[DiagnosticStage] = []
    for name, path in [
        ("00_source", source),
        ("01_after_safe_formatter", safe),
        ("02_after_table_continuation", phase3),
        ("03_final", final),
    ]:
        stages.append(_stage(name, path, stages[-1] if stages else None))

    summary = artifact_root / "summary.json"
    summary.write_text(
        json.dumps({"stages": [asdict(stage) for stage in stages]}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return DiagnosticRunResult(str(artifact_root), str(summary), stages)


def _page(block: dict) -> int:
    return int(block.get("page") or 0)


def _y0(block: dict) -> float:
    return float(block.get("y0") or 0.0)


def _y1(block: dict) -> float:
    return float(block.get("y1") or 0.0)


def validate_continuation_markers_from_blocks(blocks: list[dict]) -> list[ContinuationMarkerDiagnostic]:
    from guides.coursework_kfu_2025.rendered_table_validation import classify_continuation_marker_line

    tables = sorted(
        [block for block in blocks if (block.get("kind") or "").lower() == "table"],
        key=lambda block: (_page(block), _y0(block)),
    )
    out: list[ContinuationMarkerDiagnostic] = []
    for block in blocks:
        if (block.get("kind") or "").lower() != "text":
            continue
        marker = classify_continuation_marker_line(str(block.get("text") or ""))
        if marker.marker_kind == "none":
            continue
        marker_page = _page(block) or None
        if marker.marker_kind == "source_inline_marker_text":
            out.append(ContinuationMarkerDiagnostic(marker.text, marker_page, marker.marker_kind, marker.marker_kind, None, None, None, "low"))
            continue
        previous = [table for table in tables if _page(table) < _page(block) or (_page(table) == _page(block) and _y1(table) <= _y0(block) + 2.0)]
        following = [table for table in tables if _page(table) > _page(block) or (_page(table) == _page(block) and _y0(table) >= _y1(block) - 2.0)]
        prev_page = _page(previous[-1]) if previous else None
        next_page = _page(following[0]) if following else None
        if prev_page is None or next_page is None:
            out.append(ContinuationMarkerDiagnostic(marker.text, marker_page, marker.marker_kind, "unknown", None, prev_page, next_page, "low"))
        elif prev_page == marker_page:
            out.append(ContinuationMarkerDiagnostic(marker.text, marker_page, marker.marker_kind, "fail", True, prev_page, next_page, "high" if next_page == marker_page else "medium"))
        else:
            out.append(ContinuationMarkerDiagnostic(marker.text, marker_page, marker.marker_kind, "pass", False, prev_page, next_page, "high"))
    return out


def write_continuation_marker_reports(artifact_root: Path, stages: dict[str, list[ContinuationMarkerDiagnostic]]) -> dict[str, str]:
    artifact_root.mkdir(parents=True, exist_ok=True)
    json_path = artifact_root / "continuation_markers.json"
    md_path = artifact_root / "continuation_markers.md"
    payload = {name: [asdict(item) for item in items] for name, items in stages.items()}
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = []
    for name, items in stages.items():
        lines.append(f"## {name}")
        for item in items:
            lines.append(f"- {item.text}: page {item.marker_page}, violation={item.same_page_violation}, verdict={item.verdict}")
    md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return {
        "continuation_markers_json_path": str(json_path),
        "markdown_summary_path": str(md_path),
    }
