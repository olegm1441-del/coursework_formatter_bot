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
