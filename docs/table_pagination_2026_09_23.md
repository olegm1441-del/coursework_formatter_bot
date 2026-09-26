# Table pagination repair — 23 September 2026

Base: `cc962e07e784a40e1ee091f83342f6d45c8b001d` (`main`).

## Requirements

KFU IUEF 2025 methodology, section 6.1, pp. 25–26: a continued table needs
`Продолжение таблицы N`; the first page must contain data, not just the caption
and header. The example repeats column numbers, with the source after the last
fragment. Source rows, including duplicated student rows, must be preserved.

## Defects and repair

- The acceptance gate excluded the entire page of the next caption, concealing
  a preceding table's tail on that page. Evidence is now bounded by individual
  lines, including caption-less continuation fragments.
- A single continuation marker could mask a missing marker on a third page.
  Every consecutive data page is now checked, including appendix tables with
  `ПРОДОЛЖЕНИЕ ПРИЛОЖЕНИЯ` labels. Appendix row matches are restricted to
  the appendix rather than matching reused body-table text.
- Existing two-page repair paths skip repeated source rows. A final bounded
  repair now splits proven spills without deduplicating data. It compares the
  complete ordered non-numeric row payload before/after, re-renders, and rolls
  back if the target spill remains or a new layout failure appears.
- Adding column numbers can overflow the first fragment. Rows are moved into
  the newly created continuation rather than generating a near-empty middle
  page and another hard page break.
- Diagnostic markers at the end of narrow cells could wrap and disappear from
  row mapping. They are now placed first in the diagnostic copy only.
- Synthesized column-number rows no longer retain data-row line breaks, fields,
  hyperlinks, or fixed/minimum row heights.

No payment, credit, referral, bot handler, deployment, dependency, or environment
configuration changes. Main is not modified or deployed.

## Verification

- Formatter scripts: 15 test files passed; `test_phase3.py`: 460 passed,
  0 failed.
- `python tests/test_table_page_boundaries.py`: 11 regression cases, including
  next-caption boundaries, caption-less tails, all-page marker coverage, source
  preservation and rollback after a new neighbouring failure.
- `python tests/test_remaining_table_spills_rendered.py`: real LibreOffice
  integration with 20 fixed-height rows and internal line breaks; 3 splits,
  ordered content preserved, clean rendered gate, byte-identical retry. A
  separate appendix run verifies another 3 splits and appendix labels. Source
  rows identical to the semantic header are included in this fixture.
- Additional five-page rendered fixture: all 20 data rows retained in order;
  four correctly labelled continuation pages, source after the final fragment.
- Real corpus verification is recorded below. All 13 student/report DOCX
  fixtures completed; PDF references and the methodology appendix are reference
  materials, not student inputs.

## Limits

Uncertain mappings, unsafe merged-cell boundaries, rows taller than a printable
page, or exhausted render budgets are not guessed: the trial is rolled back and
remaining layout warnings remain visible. This is not a promise that every
arbitrary malformed DOCX will become perfect. The existing final gate reports
warnings rather than blocking file delivery. Word and LibreOffice can paginate
differently; rendered checks in this change use LibreOfficeDev 26.8.

The late repair uses the existing 90-second cleanup budget; clean documents reuse
the existing render cache and do not instrument every table. Original duplicate
student content remains in place and can still produce a source-content warning.

## Initial corpus audit (before Word feedback)

All supplied coursework/report DOCX examples were run through the formatter;
the methodology appendix file is a reference, not a student test case.
The table below records a separate late-pass and preservation audit.

| Case | New late splits | Hard layout blockers | Lost / duplicated rows |
| --- | ---: | ---: | ---: |
| Гаянов_Амир_Ленарович_Разработка проектного решения по автоматизации документооборота в организации.docx | 0 | 0 | 0 / 0 |
| coursework_bad_kpfu_2025.docx | 0 | 0 | 0 / 0 |
| example_notbad_coursework_kpfu_2025.docx | 0 | 0 | 0 / 0 |
| побитая_курсовая_Роман.docx | 0 | 0 | 0 / 0 |
| before_курсова 17. Критерии и показатели конкурентоспособности организации.docx | 0 | 0 | 1 / 1 |
| нейромаркетинг_Рыбаков.docx | 0 | 0 | 0 / 0 |
| Пример_че_может_бот.docx | 0 | 0 | 0 / 0 |
| 1_example_unformatted_coursework_kpfu_2025.docx | 0 | 0 | 0 / 0 |
| Рыбаков_Олег_Дмитриевич_курсовая_3_курс.docx | 2 | 0 | 0 / 0 |
| example_coursework_bad2_kpfu_2025.docx | 0 | 0 | 0 / 0 |
| example_otchet_2025_3_kurs.docx | 0 | 0 | 0 / 0 |

`before_курсова 17…` has a pre-existing content-normalization discrepancy:
`«Сибур»` becomes `»Сибур«`, and the source-note count differs (7 vs 6).
The new late table pass made zero edits to this case. This case must not be
reported as a fully clean content-preservation result. This change is scoped to
pagination and does not alter the existing quote/source normalization stages.

`coursework_unformatted2_kpfu_2025.docx` also completed with zero formatting
warnings. The delivery case is Bondarev: all 82 data rows retained, no added or
duplicated data rows. The initial delivery split table 1.2.1 across pages 16–17; the Word feedback
below showed that this was an unnecessary split. The source's repeated rows
are retained intentionally.


## Follow-up after Word screenshots

The earlier clean acceptance result was insufficient: it checked continuation
labels but did not prove that the split was necessary. Word moved the first
fragment of Bondarev 1.2.1 to a fresh page while retaining the hard break before
the continuation, leaving most of the first page empty.

- Short ordinary tables and compatible numeric-led fragment chains are now
  tried whole on a fresh page. Estimates only shortlist; a real render must
  place all rows and the caption together, preserve ordered data, and introduce
  no new layout failures. Downstream tables and their existing continuation
  boundaries are rechecked after reflow. Rejected trials restore exact bytes.
- First appendix label follows the `ПРИЛОЖЕНИЯ` heading on the same page,
  including after TOC rebuilding; later appendices retain their new-page break.
  TOC text `ПРИЛОЖЕНИЯ 54` no longer counts as an appendix anchor for body tables.
- Rendered validation detects isolated appendix headings and continuation
  labels in the middle of a page. Compatible short appendix fragments merge;
  longer fragments retain a verified new-page continuation. Rybakov's appendix
  previously had two continuation labels on one page despite zero gate errors.
- Oversized table grids are reduced proportionally to the section's printable
  width before height calculation. Cell widths, spans, table width and layout
  are updated consistently. Valid landscape grids remain unchanged. Six tables
  in `coursework_bad_kpfu_2025` previously clipped their right-hand content;
  the corrected output keeps all six within the page.
- Short uncaptioned body tables receive a separate measured keep-together pass.
  The report fixture's tables at DOCX indices 4 and 8 previously crossed pages;
  both now fit whole. No table numbers or captions are invented.
- TOC page numbers are refreshed after all layout repairs without rebuilding
  body page breaks. The trial must render with stable numbers before acceptance.

### Follow-up evidence

All 13 student/report fixtures completed the full formatter again. The five
width-affected cases were rerun after the width fix. Table-page contact sheets
were visually inspected across the corpus; Bondarev's entire 65-page output was
inspected, with full-size checks of table 1.2.1 and the appendix boundary.
The final report fixture was rerun after the uncaptioned-table fix, and the
`example_notbad` short-table cascade was rerendered after its boundary fix.
These are LibreOffice checks, not a native Microsoft Word verification.

- Bondarev: table 1.2.1 is whole on page 17 (header + 12 original rows); no
  continuation label. All 82 distinct source data rows retained, zero lost,
  added or duplicated rows; 17 captions retained. Final rendered blockers: none.
- Bondarev: heading and appendix A are on page 63; continuation on page 64;
  appendix B on page 65. TOC entry updated to page 63. Table 2.3.3 already
  consists of incompatible-grid fragments in the source and is retained as
  requested. Original duplicate content remains and can produce a content warning.
- Rybakov (third year): table 2.1.2 whole on page 27, appendix continuation
  correctly separated; 70 source data rows preserved.
- Additional preservation audits passed for the other fixtures except the
  previously recorded quote/source-note discrepancy in `before_курсова 17…`.
  Zero pagination-gate failures does not mean that unrelated source/content
  defects have disappeared. Deliberately narrow source columns in the malformed
  examples can still wrap words and numbers; this change does not redesign
  every source grid.
- `test_whole_table_pagination.py`: real-render tests for intact and presplit
  short tables, long-table rollback even with a misleading height estimate,
  source numeric rows, appendix merge versus page break, first-appendix/TOC
  interaction, uncaptioned short/long tables, byte-identical retry, oversized
  grids with merged cells and valid landscape sections. All pass.
- Existing phase-3 suite: 460 passed, 0 failed. Existing boundary, rendered
  continuation, acceptance and source-preservation tests also passed.

The code is in the feature branch; production and main have not been changed.

## 2026-09-24: balance whole-table moves against page utilization

The user's next Word review accepted 1.2.1 and appendix A, but identified a
large gap before Bondarev 1.3.2. Moving every short table whole was too broad.

- A late measured trial releases the whole-table keep chain when the preceding
  page has at least 180 pt of unused body space. It accepts a refill only when
  at least two data rows and at least half the table's data rows fit there,
  including the required column-number header. Otherwise the whole-table
  layout remains. Merged tables, appendix tables and existing continuation
  chains are excluded from this ordinary-table refill trial.
- The existing rendered splitter creates continuation labels and rebalances
  for the inserted number row. Exact ordered table payload and the rendered
  gate must remain valid, or the trial restores the original bytes.
- Reflow exposed two additional adjacency defects in the real examples:
  an empty paragraph before a forced table caption could occupy an otherwise
  blank page, and a source/note could detach from a table or split internally.
  The late adjacency pass removes only truly empty caption spacers, binds the
  last row to its source/note chain, keeps each note paragraph whole, then
  renders and repairs resulting spills. Fields, bookmarks, drawings, section
  breaks and authored page breaks are never removed by this cleanup.
- Adjacency runs before and after gap refilling, followed by the existing final
  static-TOC refresh. No public interface, dependency, environment variable,
  billing, auth, Telegram routing, infrastructure or deployment changes.

### Focused acceptance evidence

- Bondarev 1.3.2: three data rows on page 25, two on page 26 with the continuation
  label and source. A suffix-only intermediate trial used 4 + 1. The final full
  pipeline restored previously clipped 1.1.3 rows and kept source/note chains
  together, changing upstream pagination. The final 3 + 2 fills page 25 down
  to its printable bottom. Table 1.2.1 remains whole; appendix A is unchanged.
- Manual-split example (`example_notbad_coursework_kpfu_2025`): table 1.1.1 uses
  three rows on page 5 and the final row/source on page 6. The spurious blank
  page before 1.1.2 is gone. Table 2.3.1 splits 3 + 1 on pages 27-28, with the
  complete source and note after the final row.
- Rybakov: table 2.1.2 uses available page-26 space and continues on page 27.
  Appendix 1 has numeric headers and explicit continuation labels on pages
  55-57; its last row and complete source are together on page 57.
- Oversized-grid example: all six table grids remain inside the printable
  width. Authored narrow columns can still wrap words/numbers; this is not a
  claim that every malformed source grid has been editorially redesigned.
- Source data-row counts retained: Bondarev 82, manual-split example 65,
  oversized-grid example 31, Rybakov 70. No lost, added or duplicated data rows
  in these four delivery cases. Original Bondarev 2.3.3 geometry remains as
  explicitly accepted by the user.

`tests/test_table_page_gap_refill.py` adds real-render regressions for useful
refill, insufficient space / minority fit, failed-split rollback, appendix
exclusion, ordered content, notes attached to the last row, whole note
paragraphs, empty caption spacers and byte-identical retries. All pass.

### Final rendering findings

- Visual review found two clipped 1.1.3 rows in Bondarev's previous delivery:
  they existed in DOCX but LibreOffice did not draw them. The whole-table pass
  now shortlists missing rendered tails (when the leading row is visible),
  checks diagnostic markers, and accepts a fresh-page placement only with a
  complete row map. All three original data rows now render on page 11.
- A repaired spill can push a later table across a page boundary. The bounded
  repair now tries the downstream short table whole, then repairs remaining
  downstream spills within the SAME deadline and a bounded recursion depth.
  The entire trial rolls back if any new failure remains. This fixes
  `example_coursework_bad2_kpfu_2025` table 2.3.1 and its downstream 2.3.3.
- Diagnostic markers now prefer a cell at least 30 pt wide, using grid width
  when cell width is automatic. A narrow year/index column previously wrapped
  the marker itself and made all rows appear unmappable. This changes only
  temporary diagnostic copies; a real-render regression covers the case.
- The four delivered DOCX were rebuilt from their original sources with the
  full formatter, then rendered and inspected. Their rendered gate reports
  zero blockers and their ordered source data rows are retained. Original
  malformed Bondarev 2.3.3 and narrow authored column proportions remain known
  source limitations. These are LibreOffice checks, not native Word tests.

Final delivery audit: Bondarev 64 pages / 82 data rows; Rybakov 59 / 70;
manual-split example 40 / 65; oversized-grid example 65 / 31. All four have
zero lost/added/duplicated table data rows, zero rendered blockers and zero
empty pages. The additional final full run of `example_coursework_bad2` retains
65 data rows and all 21 source/note paragraphs with zero rendered blockers.
The 13-input corpus retains the earlier documented non-pagination content
exception in `before_курсова 17…`; it must not be described as entirely clean.

Final checks: phase-3 460 passed / 0 failed; whole-table pagination, rendered
spill/appendix, cross-page marker insertion, acceptance, preservation and the
new eight-group page-gap/note/visibility/narrow-marker regressions passed.

## 2026-09-26: final number rows and malformed fragment grids

- Remove the 1..N row only from a captioned ordinary/appendix table whose data
  rows are all proven to render on one page. Protect marked and unlabelled
  continuation chains. Keep numeric data outside the header position intact.
  Re-render after removal; repair downstream spills within the same deadline
  and restore exact bytes on new failures or changed ordered content.
- Refill an early hard continuation when the first fragment leaves over 180 pt
  of unused page space. Merge the compatible ordinary chain, release its keep
  flags, and reuse measured splitting. Accept only if more original data rows
  fit on the original page, with an unchanged payload and no new failure.
- A caption directly after an appendix heading belongs to that appendix's
  evidence window. It must not terminate the window before any table rows.
- Reconcile redundant grid subdivisions before merging physical fragments.
  The real repository example had seven physical cells mapped to eight grid
  columns; a 45-twip extra boundary differed between its header and body rows.
  Copying those rows into a seven-column grid made the right edge overflow.
  Compatibility is proved on clones; actual cells/text are never discarded.
  Genuinely incompatible grids retain their authored fragments.
- Count each physical OOXML cell once in preservation checks. python-docx can
  return one grid-spanning cell multiple times; removing a redundant grid
  boundary must not appear to delete duplicated text. Distinct cells with the
  same text still count separately. The layout gate now detects row extents
  exceeding their table grid.
- Before repairing a proven spill, combine an immediately adjacent unlabelled
  numeric-led tail with an identical grid. This handles repository-example
  table 2.2.3, where the overflowing head and separate final row previously
  caused reciprocal rollback. Authored text, captions, page/section breaks,
  fields, drawings and merged/different grids stop this merge.

Focused checks: phase-3 460 passed / 0 failed; acceptance 20; preservation 7;
page-boundary 11; grid reconciliation 4 groups; rendered final cleanup 8 groups;
existing page-gap/source-note/narrow-marker checks passed. New render tests
cover one-page ordinary and appendix number removal, retained continuation
numbers, early-split refill, exact rollback/retry, an overflowing head with a
numeric tail, and a captioned multi-page appendix.

No database schema, connection, payment, authentication, Telegram routing,
Railway configuration, dependency or environment-variable change is included.
The branch checkpoint was pushed before corpus verification; main/release is
subject to the final rendered results below.


### Final full-pipeline corpus, 26 September

All 17 byte-distinct original coursework/report inputs available in the
repository and supplied files completed. Good/Perfect examples were excluded
from processing; methodology/reference documents were used as references.
Rows below count physical OOXML data rows rather than aliases of spanned cells;
this explains the difference from older Bondarev inventory counts above.
These results supersede earlier pagination/count snapshots.

| Input | Pages | Data rows source / output | Layout blockers | Content audit | Seconds |
| --- | ---: | ---: | ---: | --- | ---: |
| example_otchet_2025_3_kurs.docx (uploaded) | 43 | 37 / 37 | 0 | pass | 45 |
| Пример_че_может_бот.docx (uploaded) | 12 | 18 / 18 | 0 | pass | 9 |
| 1_example_unformatted_coursework_kpfu_2025.docx (uploaded) | 41 | 65 / 65 | 0 | pass | 232 |
| coursework_bad_kpfu_2025.docx (uploaded) | 65 | 31 / 31 | 0 | pass | 94 |
| курсовая_Бондарев_Никита_2_курс.docx (uploaded) | 64 | 81 / 81 | 0 | pass | 347 |
| before_курсова 17. Критерии и показатели конкурентоспособности организации.docx (uploaded) | 40 | 22 / 22 | 0 | known normalization exception | 24 |
| побитая_курсовая_Роман.docx (uploaded) | 35 | 25 / 25 | 0 | pass | 22 |
| example_coursework_bad2_kpfu_2025.docx (uploaded) | 41 | 65 / 65 | 0 | pass | 142 |
| example_notbad_coursework_kpfu_2025.docx (uploaded) | 40 | 65 / 65 | 0 | pass | 97 |
| coursework_bad_kpfu_2025.docx (repository) | 65 | 31 / 31 | 0 | pass | 92 |
| курсовая_Бондарев_Никита_2_курс.docx (repository) | 61 | 78 / 78 | 0 | pass | 282 |
| example_unformatted_coursework_kpfu_2025.docx (repository) | 42 | 65 / 65 | 0 | pass | 282 |
| 01-coursework_unformatted2_kpfu_2025.docx (uploaded) | 65 | 31 / 31 | 0 | pass | 91 |
| курсовая пример 1.docx (repository) | 69 | 112 / 112 | 0 | pass | 599 |
| Гаянов_Амир_Ленарович_Разработка проектного решения по автоматизации документооборота в организации.docx (uploaded) | 44 | 59 / 59 | 0 | pass | 325 |
| нейромаркетинг_Рыбаков.docx (uploaded) | 73 | 84 / 84 | 0 | pass | 383 |
| 06-_-_-_-_3_-.docx (uploaded) | 60 | 70 / 70 | 0 | pass | 176 |

Table-page contact sheets were inspected across all 17 outputs. All pages of
the three delivered cases (repository example 1, uploaded bad-width coursework,
neuromarketing) and Rybakov third-year coursework were also inspected. The three
deliverables were independently rendered using the document rendering tool.
Full-size checks covered example 1 table 2.2.3 and both halves of 2.3.1.

- Repository example 1: 112 original data rows and all 24 source/note paragraphs
  retained; table 2.2.3 is whole on page 46 without column numbers. Table 2.3.1
  has seven aligned columns on pages 50-51 and a complete source after its tail.
- Uploaded Bondarev: table 1.2.1 stays whole on page 17 without a numeric row;
  1.3.2 fills page 25 with three rows and continues on page 26 with two rows and
  its source. Appendix A shares page 62 with the section heading and continues
  on page 63. The explicitly accepted source geometry of 2.3.3 remains.
- Neuromarketing: all 84 data rows retained. Appendix 1 continues on page 68;
  appendix 2 fits page 71 without numbers; appendix 3 continues on page 73.
- Bad-width coursework: all 31 data rows retained; all six tables remain within
  the printable page width and have no unnecessary number row.
- A separate final numeric-header inventory found no independent ordinary
  single-page tables retaining the index header. Remaining appendix headers
  belong to visually verified multi-page continuations.

All rendered gates are clean. The existing Criteria-17 quote/source-note
normalization discrepancy remains the sole content-audit exception (16/17
content audits pass). Zero gate failures is not a claim of universal visual
perfection: authored very narrow year/number columns still wrap, and unsafe or
poorly mapped tables can retain whitespace because trials are conservative.
No native Microsoft Word or Railway-container render was available. Runtime
was 9-599 seconds per input in this environment, partly under three concurrent
jobs; this is not a production latency benchmark or a speedup claim.

Release scope is formatter/validation/tests/documentation only. There is no
schema migration and no change to DATABASE_URL, Postgres, persistent volumes,
billing, authentication, Telegram handlers, or Railway service configuration.
