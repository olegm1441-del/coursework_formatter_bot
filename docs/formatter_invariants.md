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
