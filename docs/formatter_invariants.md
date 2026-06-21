# Formatter Invariants

Operational memory for KFU coursework formatter work. Keep this file short, conservative, and current.

## Stable Baseline

| Commit | Scope stabilized |
| --- | --- |
| `31a612a` | Front-matter freeze, ordinary soft-break preservation, merged body-heading separation after real intro. |
| `9712da2` | Appendix title normalization immediately after `ПРИЛОЖЕНИЕ N`. |
| `edb10a1` | Static visual layout stabilization for `СОДЕРЖАНИЕ` / TOC entries. |

## Core Rules

- Work only on `test-bot`; never merge or push to `main`.
- One issue -> one patch.
- Stability > cleverness.
- Minimal diffs only.
- No broad cleanup refactors.
- Preserve regression-gold truth cases before improving stress cases.
- If a fix requires another subsystem, stop and split the batch.

## Workflow Modes

- `Mode: inspect`: read/analyze only.
- `Mode: plan`: plan only.
- `Mode: patch`: edit one approved issue only.
- `Mode: test`: run/check tests only.
- `Mode: deploy-check`: git/push only, no code edits.

## Structural Invariants

- Never mutate text before the real standalone `ВВЕДЕНИЕ`.
- Do not confuse `ВВЕДЕНИЕ` inside TOC with the real intro heading.
- Ordinary soft breaks are not page breaks.
- Remove only real page breaks, not `w:br` line wrapping.
- Chapter/section headings after intro must not merge with headings or body text.
- Appendix labels stay uppercase/right-aligned; continuation labels must not trigger title formatting.
- Appendix titles immediately after labels are centered, not bold, no trailing dot, with one blank after.
- Reference block headings are not numbered sources.
- Table captions require an adjacent real table.
- Source/note service lines must be structural boundaries.

## Table Continuation Rules

The Phase 3 marker-driven table-split subsystem (`guides/coursework_kfu_2025/table_continuation.py` + `table_split_prototype.py` + `table_markers.py`) must obey the rules below. These are product-level rules; tests and patches must enforce them.

### 1. Split trigger

- A table may be split only if it **genuinely spans pages** after Phase 1 + Phase 2 formatting, verified by LibreOffice marker render evidence.
- Do not split a table merely because it is "long" in the DOCX. Row count is a candidate filter (`_MIN_ROWS_FOR_SPLIT_CANDIDACY`), not a split trigger.
- The marker-render row→page map drives the split decision; PDF/visual evidence is authoritative.

### 2. Minimum valid first fragment and table-start orphan

A first-page fragment of an actual split/continued table is valid only if it contains, in this order:

- Table caption (`Таблица X.Y.Z`).
- Table title (if present).
- Semantic header row (if present in source).
- Numeric column-numbering row (`1 2 3 …`) if this table is actually split/continued — see rule 3.
- At least **ONE real data row**.

**The criterion is not** "avoid header + one small row with a huge blank below". A first fragment that contains the structural prefix above + one real data row is valid even if the page below has visible whitespace.

Invalid split/continuation first fragments:

- caption only;
- caption + title only;
- caption + title + header only;
- caption + title + header + numeric row, zero data rows;
- first data row pushed to the next page by inserting the synthetic numeric row (see rule 3 — split point must compensate).

If an ordinary non-split table starts near the bottom of a page and that page contains only the caption/title/header and optionally a source numeric row, but zero complete real data rows while the first real data row starts on the next page, this is a **table-start orphan**. The MVP repair is to move the whole table start to the next page by inserting exactly **two blank paragraphs** before the table block, before the caption when a caption exists. Do not split the table, do not insert `Продолжение таблицы`, do not synthesize a numeric row, and do not convert the repair into a page-break-before rule.

### 3. Numeric column-numbering row

The methodical KFU continuation format requires a numeric column-index row "1 2 3 … N" before data rows in actual split/continued table fragments. This row is **not** a "repaired header" — it is a separate numeric index row.

Required behaviour:

- **Ordinary non-split table**: do not synthesize a numeric row.
- **First / original split fragment**: must contain the numeric row directly above the data rows.
- **Each continuation fragment**: must repeat the numeric row before its data rows.
- **Preservation**: if the source table already has a numeric row at row index 1, reuse it; do not duplicate.
- **Synthesis**: if the formatter synthesizes a numeric row for the continuation, it must also ensure the first fragment has the corresponding numeric row.
- **No-push invariant**: adding the synthetic numeric row to the first fragment must **not** push the first real data row off that fragment. If unavoidable, the split point must be adjusted (E3 NUM-row compensation: `split_before_row - 1` under guards), or the split skipped safely.

### 4. Continuation label format

The continuation page must contain, in this exact structural sequence:

```
Продолжение таблицы X.Y.Z

<continuation table fragment>
```

- One paragraph "Продолжение таблицы X.Y.Z" (right-aligned, page-break-before, keep-with-next).
- Exactly **ONE** blank paragraph immediately after the marker.
- The continuation `<w:tbl>` directly after the blank.

Forbidden in the continuation block:

- Repeating the original caption "Таблица X.Y.Z".
- Repeating the original title.
- Any duplicate "Таблица X" line.

### 5. Header and numbering behaviour for continuation fragments

- Repeat the numeric column row (rule 3).
- Do not unconditionally repeat the full semantic header. Repeat only when the source structure or methodical norm requires it.
- Preserve merged cells, column widths, and table borders byte-equivalent across both fragments.
- Do not invent or modify semantic header text. If the source has a complex multi-row header, keep the original header structure on the first fragment and emit at least the numeric row on the continuation; never synthesise a "repaired header" that differs from the source.

### 6. Source and note placement

`Источник:` and `Примечание:` paragraphs:

- Belong to the **final** table fragment only.
- Must not remain after a non-final fragment.
- Must not be duplicated across fragments.
- Order: `Источник:` first, then `Примечание:` (if both present). Both attach to the final fragment.

### 7. Valid manual continuation chains

If the source DOCX already contains a valid manual continuation chain (`Таблица X` → `Продолжение таблицы X` paragraph → continuation table), it must be preserved exactly unless clearly malformed. `_valid_manual_continuation_table_ids` is the authoritative detector; both the strict `tbl→p→tbl` shape and the auto-inserted `tbl→p→blank→tbl` shape count as valid.

Do not rebuild or re-split a recognised manual chain.

### 8. Split blockers

A split is invalid (must be skipped, with logged reason) if any of the following holds:

- A table spans pages in the rendered output but no continuation label was inserted.
- The continuation label carries the wrong table number.
- The continuation label is detached from its continuation table by anything other than one blank.
- The first fragment has zero data rows.
- A synthetic numeric row pushed the first data row off the first fragment.
- Numeric row missing from the first split/continued fragment.
- Numeric row missing from a continuation fragment.
- Any original data row is missing across fragments.
- Any original data row is duplicated across fragments.
- `Источник:` or `Примечание:` is duplicated.
- `Источник:` or `Примечание:` left after a non-final fragment.
- Column widths change between fragments.
- Merged cells break across the split.
- Table borders break.
- The table extends outside page margins after split.
- A page becomes almost empty (orphan single-row continuation) because of the split point.
- A previously valid (gold) split regresses.

### Acceptance criteria for future table-split patches

A patch that changes table-split behaviour is accepted only if **all** of the following hold:

- Бондарев problematic tables (see `docs/truth_cases.md`) are either fixed or explicitly skipped with logged reason.
- Бондарев `Таблица 1.3.2` remains correct (gold case; no regression).
- `курсовая пример 1` has no regressions.
- `нейромаркетинг` (Рыбаков) has no regressions.
- No global `render_budget_exceeded` regression — the candidate-mode path must still enter the marker pass on multi-table docs.
- No user warning is emitted unless real candidate tables were skipped or failed (budget overflow, hard timeout, or per-candidate diagnose error).
- All formatter test suites pass (`test_phase3.py`, `test_reference_subheading_spacing.py`, `test_ux_texts.py`).
- PDF visual smoke is provided side-by-side (before/after) for at least Бондарев, нейромаркетинг, and курсовая пример 1.

## Rendered Table Layout Acceptance Gate (Stage 0 + Stage A)

Structural success (row preservation, cell padding, warning counts) is **not** layout success. The gate makes visible rendered table defects explicit.

### Stage 0 — conservative table mode
- `KPFU_RENDERED_TABLE_CONTINUATION` (default unset/false): the risky rendered continuation insertion + exact/compatible same-page merge passes are skipped, so the formatter never **creates** a same-page `Продолжение таблицы` split or a synthesized numeric row. Tables are left whole; a whole table that flows across a page is less wrong than a bad same-page split.
- Set to `1`/`true` to re-enable the experimental path (still subject to the gate).
- Only the rendered path (`apply_rendered_table_continuation`) inserts markers; the DOCX-only `apply_table_continuation` and the orphan-move guard do not, so gating that one entry point is sufficient.

### Stage A — `evaluate_table_layout_acceptance` (in `rendered_table_validation.py`)
Pure function over rendered `pdf_lines` + DOCX identities (+ optional `doc`). Blocker severities: `fail` (NO-GO) / `needs_human_review` (not provably clean). `format_docx` surfaces blockers as report warnings + structured logs and still returns the file; smoke/deploy reads the blockers for GO/NO-GO. Detectors:
- `same_page_continuation` (fail) — `Продолжение таблицы N` on the same page as a fragment of N.
- `single_table_crosses_pages_without_marker` (fail) — a table's data rows span >1 page with no valid continuation marker (a *marked* split, e.g. Rybakov, is accepted).
- `orphaned_header_row` (fail) — a page carries the header/numeric of N but zero data rows while data appears on a later page.
- `appendix_label_not_on_new_page` (fail) — `ПРИЛОЖЕНИЕ X` with substantial non-appendix content above it on its page.
- `fragment_grid_mismatch` (fail on column-count change; `needs_human_review` on >2% per-column width drift) — adjacent fragments must preserve the grid (Rule 8).
- `cell_text_overflow_or_illegible_squeeze` (`needs_human_review`) — ≥5 lone non-numeric ≤2-char fragment lines concentrated on one page of a table region; numeric/page-number lines excluded so clean tables do not false-fire.

This gate detects defects only; the actual table repairs are Stage B–E.

## Forbidden Dangerous Operations

- No global paragraph merge/delete passes without structural guards.
- No formatter changes before real intro unless the batch is explicitly TOC/front-matter scoped.
- No broad regex-only cleanup across the whole document.
- No heading/table/list/reference fixes in the same patch.
- No page-numbering changes inside table, reference, or caption batches.
- No runtime, payment, handler, or deployment edits during formatter batches.

## Regression Philosophy

- Tests protect product rules, not current bugs.
- Prefer deterministic DOCX/XML checks over PDF visual checks.
- Use PDF smoke for truth-case confidence, not as the only proof.
- Always report remaining known risks instead of silently expanding scope.
- Keep malformed stress cases documented even when a later patch is needed.
