# Truth Cases

Smoke/truth cases used to protect formatter behavior. Paths may vary by workstation; identify by file name and content.

## Case Matrix

| Case | Tags | Purpose | Stresses | Representative pages |
| --- | --- | --- | --- | --- |
| Rybakov | regression-gold | Protect already-good formatter output. | TOC, intro page 3, chapter headings, references, appendix first/continuation pages. | TOC, first chapter, appendix first page, appendix continuation. |
| Neuromarketing | stress-case, no-title-page, tables | Exposes structural regressions in real student work. | Real intro detection without title page, TOC fragility, table captions, source/note lines, appendix A. | TOC, first chapter/`1.1`, `Таблица 1.1.1`, source/note + figure area, appendix A. |
| Bondarev | appendix-heavy | Protect appendix label/title and continuation behavior. | Appendix A/B, appendix tables, page numbering after appendices, TOC stability. | TOC, first chapter, appendix A, appendix B/continuation. |
| bad2 / example_coursework_bad2 | malformed TOC | Stress malformed static contents and front matter. | TOC leaders/tabs, bad spacing, title/contents boundaries. | TOC, intro, first body heading. |
| coursework_unformatted2 | malformed body | Broad unformatted coursework stress case. | Headings, tables, page breaks, references, appendices. | TOC, intro, first chapter, first table, references. |
| Roman / побитая | broken references | Protect B1 reference block-heading behavior. | Reference subheadings, old numbering cleanup, malformed bibliography layout. | References start, each reference block heading, first numbered source after each block. |

## Expected Stable Behavior

### Rybakov

- Real `ВВЕДЕНИЕ` is detected as the body start and keeps visible page number 3.
- TOC/front matter are not destructively rewritten.
- Chapter and section headings remain separate.
- Appendix label/title formatting remains stable.
- References do not regress.

### Neuromarketing

- Formatting succeeds even when the title page is missing or malformed.
- Real intro remains page 3 when formatter rules require it.
- Front matter before real intro is preserved under P0 policy.
- `1.` and `1.1.` headings are not merged after intro.
- Table caption false positives and inline-dash captions are known planned fixes.
- Source/note/figure ordering is a known planned fix.
- Appendix A label/title should remain readable and normalized.

### Bondarev

- Appendix labels are uppercase/right-aligned.
- Immediate appendix titles are centered, not bold, and have no trailing dot.
- Appendix continuation logic remains unchanged.
- First appendix page numbering and later appendix-page numbering rules remain stable.

#### Bondarev table-split truth targets

These are the per-table expectations for the marker-driven table-split subsystem (see `docs/formatter_invariants.md` → "Table Continuation Rules"). Any patch that touches table continuation must pass smoke against this list.

Known **bad / needs fixing** (current marker split produces invalid layout):

- `Таблица 1.1.3` — did not split correctly.
- `Таблица 1.2.1` — did not split correctly.
- `Таблица 2.1.5` — did not split correctly.
- `Таблица 2.3.1` — layout broke.
- `Таблица 2.3.3` — layout broke.

Known **gold / regression-protected** (must remain correct):

- `Таблица 1.3.2` — split looked correct under the d19e6ea baseline. Do not regress.

Any future table-split patch must either fix the "bad" tables above or explicitly skip them with a logged reason; in either case it must not regress 1.3.2.

#### Demo table-start orphan truth target

`Пример_че_может_бот.docx`, `Таблица 1.1.3` is the current table-start orphan truth case. If the rendered start page contains only `Таблица 1.1.3`, the title/header, and zero complete real data rows while the first row (`Кейс 1`) starts on the next page, the formatter must move the whole table start by inserting exactly two blank paragraphs before the caption. It must not split the table, insert `Продолжение таблицы 1.1.3`, or synthesize a numeric row for this ordinary non-split table.

### bad2 / example_coursework_bad2

- Useful for TOC checks only when the active batch touches contents/front matter.
- Do not use it to justify body/table/reference changes.
- Current policy: do not mutate pre-intro text unless the batch explicitly owns TOC recovery.

### coursework_unformatted2

- Use as a broad smoke document after structural batches.
- Inspect for unexpected blank pages and over-aggressive cleanup.

### Roman / побитая

- Reference block headings remain unnumbered.
- Real reference entries remain sequentially numbered.
- No duplicate numbering such as `3. 3.`.

## Smoke Checklist

- Real intro page and visible page number.
- TOC exists and entries are not newly collapsed.
- `1.` separate from `1.1.` and `2.` separate from `2.1.`.
- Table captions are not promoted from analytical prose.
- Source/note lines are not merged.
- Appendix labels/titles are stable.
- References B1 behavior is unchanged.
- No unexpected blank pages.
