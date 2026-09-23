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

## Corpus audit

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
duplicated data rows. The final table 1.2.1 spans pages 16–17 and the appendix
table spans pages 63–64 with the required continuation labels. The source's repeated rows are retained intentionally.
