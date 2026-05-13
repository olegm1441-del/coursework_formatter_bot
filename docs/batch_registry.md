# Batch Registry

Operational registry for completed and planned formatter/UX batches.

## Completed

| Batch | Commit | Scope | Forbidden scope | Main acceptance | Smoke docs |
| --- | --- | --- | --- | --- | --- |
| B1 references stabilization | `39718f5` | Flexible reference block-heading detection. | No split-reference merging, page numbering, appendices, tables, lists. | Reference headings unnumbered; sources sequentially numbered; old numbering cleanup preserved. | Roman/побитая, Rybakov, Bondarev. |
| B2 contents stabilization | `edb10a1` | Static visual TOC layout: tabs, dot leaders, entry cleanup. | No page numbering, appendices, references, table split logic. | `СОДЕРЖАНИЕ` clean; TOC entries use one right tab with leader; no duplicate tabs. | Rybakov, Bondarev, bad2/example_coursework_bad2. |
| B2.1 appendix title normalization | `9712da2` | Immediate title after `ПРИЛОЖЕНИЕ N`. | No page numbering, TOC, references, table continuation subsystem. | Label behavior unchanged; immediate title centered/not bold/no trailing dot; one blank after. | Rybakov, Bondarev, Neuromarketing. |
| P0/P1 structural soft-break preservation | `31a612a` | Front-matter freeze, preserve ordinary soft breaks, split merged body headings after intro. | No TOC recovery, table captions, source/note ordering, references. | No mutation before real intro; soft breaks not treated as page breaks; body headings separated. | Rybakov, Neuromarketing, Bondarev. |
| UX eta text patch | `1f76e38` | Change processing ETA copy to `через минуту`. | No formatter/runtime behavior changes beyond text. | Old phrase absent; UX text test passes; runtime files compile. | Not formatter-smoked. |

## Planned

### Table Caption Adjacency Hardening

- Scope: classify `Таблица N` as a caption only when an actual table is adjacent.
- Forbidden scope: inline dash normalization, source/note ordering, TOC, lists, appendices.
- Risks: missing real captions; preserving false positives in malformed docs.
- Acceptance:
  - `Таблица 1.1.1 показывает...` remains body text without adjacent table.
  - Real caption + title before table still formats as caption/title.
  - Appendix immediate table-like title behavior from B2.1 remains intact.
- Smoke docs: Neuromarketing, Rybakov, Bondarev.

### Inline Dash Table Caption Normalization

- Scope: normalize real `Таблица N — Title` captions when a real table follows.
- Forbidden scope: caption classification broadening, body dash cleanup, TOC, references.
- Risks: stripping meaningful dashes from body text.
- Acceptance:
  - Real inline dash caption becomes `Таблица N` + clean centered title.
  - No dangling dash title.
  - Ordinary body dashes untouched.
- Smoke docs: Neuromarketing, coursework_unformatted2.

### Source/Note Figure Ordering

- Scope: keep `Источник:` and `Примечание:` as service lines around figures/tables.
- Forbidden scope: table caption classification, TOC, references, lists.
- Risks: moving service lines across unrelated paragraphs.
- Acceptance:
  - `Источник:` and `Примечание:` are separate paragraphs.
  - No `.Примечание` merge.
  - Figure source/note order remains before `Рис.` where required.
- Smoke docs: Neuromarketing, Rybakov.

### Lists as Word Bullets

- Scope: convert approved dash-list patterns to Word bullet formatting.
- Forbidden scope: headings, tables, references, TOC, page numbering.
- Risks: converting prose dashes or bibliography lines.
- Acceptance:
  - Real list items become Word bullets with en dash style where required.
  - Body punctuation and references remain unchanged.
- Smoke docs: Neuromarketing, coursework_unformatted2.

### TOC Recovery

- Scope: repair malformed static TOC paragraphization, leaders, and page refs.
- Forbidden scope: body headings, table captions, references, appendices unless explicitly required.
- Risks: mutating front matter incorrectly; confusing TOC `ВВЕДЕНИЕ` with real intro.
- Acceptance:
  - Each TOC entry is a separate paragraph.
  - Leaders/page numbers are stable.
  - TOC stops before real intro.
  - Front-matter freeze remains respected outside the TOC-owned range.
- Smoke docs: bad2/example_coursework_bad2, Neuromarketing, Rybakov.

### Blank-Page Audit

- Scope: audit and remove duplicated/empty page or section breaks.
- Forbidden scope: content rewriting, caption/reference/list changes.
- Risks: removing intentional page starts before major sections.
- Acceptance:
  - No unexpected blank pages before conclusion, references, or appendices.
  - Intentional section/page breaks remain.
  - Page numbering behavior from A1/P0 remains stable.
- Smoke docs: Neuromarketing, Bondarev, Rybakov.
