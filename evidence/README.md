# Source evidence index

`source_inventory.json` lists the 40 original input files, SHA-256 hashes, extraction status and per-file sheet/cell or physical-PDF-page counts. It also enumerates candidate rig-less/service tabs. This is an inventory, not a conclusion that their monetary scope or applicability has been reconciled.

To regenerate the **local** SQLite index (do not commit the generated database):

```sh
python -m venv /tmp/azns-index-venv
/tmp/azns-index-venv/bin/pip install -r tools/requirements-evidence.txt
/tmp/azns-index-venv/bin/python tools/build_evidence_index.py --source . --output /tmp/azns-evidence.sqlite
/tmp/azns-index-venv/bin/python tools/summarize_evidence_index.py /tmp/azns-evidence.sqlite evidence/source_inventory.json
/tmp/azns-index-venv/bin/python -m unittest discover -s tests -v
```

Choose a new output path if the index exists: the builder refuses to overwrite an existing database. The database generated during this revision was `/tmp/azns-evidence.sqlite`; `/tmp` may be cleared between sessions. `files`, `sheets`, `cells`, `pages` and `errors` tables preserve source identity, sheet name, cell coordinate and PDF physical page. `cells.raw_value` stores OOXML cached values, not evaluated formulas or normalized engineering units. `cells.formula_xml` preserves the formula element, including shared-formula attributes. Numeric values may be date serials. PDF text extraction does not provide OCR, preserve full layout, guarantee readable tables or validate interpretation; even a scanned page may have searchable header text. A run reporting zero extraction errors does not prove semantic completeness or source correctness.

Example search (PDF page is physical, starting at 1; printed page can differ):

```sql
SELECT f.name, p.page, p.text FROM pages p JOIN files f ON f.id=p.file_id
WHERE lower(p.text) LIKE '%formation%';
SELECT f.name, s.name, c.address, c.raw_value, c.formula_xml
FROM cells c JOIN sheets s ON s.id=c.sheet_id JOIN files f ON f.id=s.file_id
WHERE f.name='CBS AZNS-122WO1.xlsx' AND s.name='Cost Breakdown-122'
  AND c.address IN ('G190','H190');
```

Full extraction of raw content does not automatically deliver a geological model, verified formation tops, event attribution, calculated hydraulics, operational translations, approved parameter windows or reconciled costs. Those require separate source-specific parsing and engineering review.

## Formation-table snapshots

`formation_table_snapshots.json` transcribes one selected, dynamically located worksheet snapshot per pack. Regenerate it after creating the local index:

```sh
/tmp/azns-index-venv/bin/python tools/extract_formation_tables.py /tmp/azns-evidence.sqlite evidence/formation_table_snapshots.json
```

Selection is the last sheet in workbook order containing a recognized formation table with a numeric depth, or, if none has one, the last recognized table containing names. Workbook order is not necessarily chronological; for example the 127WO1 snapshot is sheet `88-2 Data` although later workbook sheets exist. The tool uses the table's actual labels and records populated unlabelled cells rather than guessing their meaning. It does not reconstruct all daily formation tables, interpret depth datum or resolve questionable row alignment. In well 10, `1-2 Data` and `83-2 Data` show a shifted formation-name sequence, requiring source reconciliation before assigning tops. The 116, 127WO2 and 166 snapshots contain names without numerical depths and must not be read as zero-depth tops. See the report appendices for the raw citations and limitations.

## 207 bit observations

`207_bit_observations.json` retains daily bit observations and source cell references from the 155 data sheets in `AZNS207Pack.xlsx`. Regenerate with:

```sh
/tmp/azns-index-venv/bin/python tools/extract_207_bit_observations.py /tmp/azns-evidence.sqlite evidence/207_bit_observations.json
```

It groups strictly numeric bit numbers and explicit rerun numbers; missing reruns remain unknown. The summary shows maxima of *reported cumulative counters*, not run-by-run incremental metres or hours. A cumulative counter can span reruns; it must not be summed across groups. Metadata conflicts, changing names/units, number reuse, and last nonblank dull values require manual reconciliation. This is specific to the 207 workbook structure and does not repair other wells' bit tables.

## Daily drilling and workover observations (interim)

Regenerate after building the raw index, from the repository root:

```sh
/tmp/azns-index-venv/bin/python tools/extract_drilling_tracks.py /tmp/azns-evidence.sqlite evidence/drilling_track_observations.json
```

The JSON preserves sheet, cell, raw value and numeric interpretation from recognized legacy and DDR sheet families in all 13 packs, including a shifted-column DDR family. Recognition is not full daily-sheet coverage: 12 recognizes 65/72 candidate pairs and 127WO1 88/137; some later `-2 Remark` sheets are not a data sheet. The index's cached values can be stale. Explicit numeric legacy ranges are separated into endpoints, while ambiguous slash values stay unparsed. Numeric duration cells with Excel time formats are converted from day fractions to hours; numeric-looking text is not converted merely because of its style. Status extraction stops at the first daily total, excluding the next block. Days count as complete only when recognized segments total 23.9–24.1 hours with known codes and numeric durations; omitted/missing segments and ambiguous codes invalidate a day. N/NPT, UP/unplanned and NPT-UNPLN remain separate source categories. Counts are not an independently validated breakdown or proof of actual schedule progress.

Four interim W3 drilling graphics plot observations at report-end MD against **unverified** formation snapshot bands and legacy shoes, not event depths or continuous logs. The excluded 207 7,500-gpm value remains in the evidence. Nine workover W3 graphics show report-number timelines, source-code statuses and incomplete totals. Neither type establishes approved operating limits, actual-vs-planned phase performance, a geological interpretation or verified NPT. All 13 new English W3 images were visually inspected; English and Persian editions share the corresponding image. Inspection is not rendered-page QA.

## CBS primary financial sheets

`cbs_cost_observations.json` records raw cached values, cell references, formulas, formula errors and missing formula caches for the primary cost, material, contingency and summary sheets recognized by name in all 13 CBS workbooks. Regenerate directly from the untouched originals:

```sh
/tmp/azns-index-venv/bin/python tools/extract_cbs_evidence.py --output evidence/cbs_cost_observations.json
```

This is a transcription of selected primary financial sheets, not an exhaustive extraction of every service-rate worksheet, a formula recalculation or a financial audit. `CBS AZNS-85.xlsx` has a schedule sheet but no sheet matching the selected cost-sheet names; this is **not** evidence that no budget exists elsewhere. The 122WO2 contingency-material sheet has 245 cached `#REF!` cells, and the separate 127WO1 side-tracking workbook's main cost sheet has 26. A cached numeric total may still omit error rows and is not automatically reliable. In 127ST2, `Cost Breakdown-127!F6` sums F7:F12 while F197 sums F5:F196, including the subtotal and its children; the source's intended cost basis and total must be reconciled rather than silently corrected. In 40, D5 × E5 does not equal F5. In 207, `Total Cost-EPD!F3` includes the material total while G4 is zero; the source uses the identifier F-20 and the identity/scope needs confirmation. In 116, the combined summary also contains 127; do not allocate the joint ceiling to one well. In 127WO2, source comments include unresolved material/service exclusions.

The provisional W5 and table 8-2 replacements use source-labelled category rows. Five layout variants identify a grand total over individual rows and a non-subtotal category header: their category figures are derived by summing numeric cached values in the cited row ranges. Other layouts use cached category rows directly. These categories can overlap, omit entries, or fail to match source totals; the calculations neither prove a cost ledger nor fix workbook formulas. Plan-labelled and extra-work columns are not presumed additive. The 85 W5 shows only the cited rig schedule and is not a monetary chart. Source-status M6 and financial-status M3 in the four field reports are likewise interim, not portfolio financial or contractual determinations.

### Preliminary bit/tool evidence across all 13 packs

`well_bit_observations.json` is produced by `tools/extract_well_bit_observations.py /tmp/azns-evidence.sqlite evidence/well_bit_observations.json`. It scans named daily data pages and recognizes distinct legacy H/K/M and DDR B/E/G or shifted A/D/F families using bit-label anchors. It records cell coordinates and raw fields, explicit numeric bit IDs, explicit reruns, metadata variants, maximal cumulative counters and unassigned entries. Ambiguous `RR#3`, `R#1`, mills and identifiers with no clear numeric bit are retained unassigned; blank rerun is unknown. Workbook sheet order does not prove chronology; maxima may include earlier runs, are not incremental performance, and may conflate reused identifiers. All conflicts require review against bit records. 127WO1 recognized 88 named bit data sheets vs 137 candidate pairs; 12 recognized 65 vs 72; 116 recognized 111 vs 112. These are not evidence of absence elsewhere. The previous 207-only JSON is retained as an earlier extraction, not the current renderer input.

`tools/replace_well_bit_evidence.py` replaces W7 and the bit tables in 26 well reports. Where no positive counter is available it plots extraction coverage instead of an empty panel; unidentified tool entries have a separate table. `tools/replace_unsupported_npt_sections.py` replaces unsubstantiated incident-mechanism and contractual-attribution sections and W4 with partial source status-code totals from `activity_observations.json` and cited example cells. The cumulative line is a share of displayed source-code hours only, not a total NPT definition; complete-day coverage is explicit. Neither W4 nor W7 establishes a cause, a liability, an as-built run history, or a financial result. The earlier `replace_w4_pareto.py` and `replace_207_bit_evidence.py` target legacy report contents and should not be rerun over the current reports.

### Programme cover summaries and selected formation tops

`program_summary_reviewed.json` is a manually reviewed transcription of selected well-information pages, one per programme, including 127 ST1 separately from ST2. `tools/build_program_summary.py` verifies key strings against the local PDF-text index, but is not a complete semantic reading of the PDFs. The cited page is the physical PDF page. Duration labels and ambiguous contingency meanings are retained, not blindly summed. Page 8 of the 207 programme labels 125.4 days as estimated well duration and 10 days as estimated moving time; 127WO1 page 7 states 60.4 main and 58.9 contingency days; 75 page 8 labels 141.4 main and 149.2 with contingency. These source figures conflict with some prior legacy report baselines and do not establish comparable actual durations.

`program_formation_tops_reviewed.json` contains selected clearly paired labels and measured-depth values from programme pages for 10, 116, 122WO1, 122WO2 and 40. Well 116 uses the source MD (log) column; 10 and 116 tables are called actual in their source but are historical, not independently validated. The 122WO2 figure is explicitly a schematic *before workover #1*, not evidence of the present WO2 condition. Well 40 values are expected. Spatially dislocated or unnamed rows are not inferred. `tools/add_program_formation_tables.py` checks source strings before inserting the selected tables. Existing 207 comparison is separate.

`tools/apply_program_and_safety_revision.py` updates the 13 well reports in both languages and all four field reports: W1 now shows source programme values and data-coverage counts, not an unverified plan/actual timeline. Unverified generic numerical operating envelopes and unsupported risk-ownership claims in those specific sections are withdrawn; example daily-source cells and review gates replace them. `tools/qualify_legacy_portfolio.py` and `tools/qualify_remaining_schedule_values.py` withdraw selected portfolio and schedule claims. These edits do not validate or translate the rest of the legacy content. The section-replacement scripts target the prior report structure and should not be rerun over the revised reports without making them idempotent.

### Independent activity-page transcription

`activity_observations.json` is regenerated from the source index with `/tmp/azns-venv/bin/python tools/extract_activity_observations.py /tmp/azns-evidence.sqlite evidence/activity_observations.json`. The extractor recognizes source status headers independently from data-sheet format; it captures source status, duration, code, narrative cell coordinates and unattached narrative continuation cells. It selects the first slot only if duplicate status pages for a day are found and flags all such cases. Coverage is not a verified rig timeline. In 127WO1 there are 137 recognized activity pages and 131 classified 24-hour records, even though only 88 data sheets were recognized; 12 has 72 activity pages and 70 complete records with 65 recognized data sheets. Note that complete classification requires recognized status/duration rows summing 23.9–24.1 hours and start/end cells present; overlap, time sequence and source classification are not independently validated.

`daily_observation_examples.json` contains 26 manually translated selective source observations with exact sheet/cell anchors and original raw cell text. `replace_daily_narrative.py` replaces old unsupported chronological conclusions with these examples and a W2 source-status daily log; `replace_unsupported_npt_sections.py` now uses the independent activity data for W4. `update_activity_coverage_figures.py` refreshes W1 and field M6; `explain_w3_activity_scope.py` records that W3 remains data-linked. These are interim analyses, not exhaustive incident reviews. Do not run older non-idempotent narrative/schedule revisions over the current reports.

### Selected programme phases, providers and casing intervals

`tools/build_program_phases_and_services.py` directly re-reads all 14 original programme PDFs and checks two selected phase anchors and the service-table excerpt on cited physical pages. Its `program_phases_services_reviewed.json` records 28 bilingual phase summaries and 56 transcribed assignments across four service classes, not the full programme service tables or evidence of actual mobilization. `tools/add_phase_service_review.py` adds the selected phase statements, two independently cited daily observations and service assignments to the 26 well editions; similar subjects do not prove a phase match, achievement, variance or liability. The 127 ST1 phase examples are separately cited in the four field reports. The identical interval wording in the 12 and 127WO1 programmes is flagged for review, not asserted as valid for either operation.

`tools/build_program_casing_intervals.py` reads 14 original PDFs directly and checks source anchors for 71 manually transcribed programme table intervals in `program_casing_intervals_reviewed.json`. Layout-mode PDF extraction was used to check the segmented columns for 85, 166, 127WO2 and the tables for 40 and 207; the page text anchors do not replace an engineering review. `replace_program_casing_evidence.py` replaces the legacy table 6-1 and W6 in 26 reports; the separately indexed 127 ST1 pre-operation table is reported at field level rather than being merged with ST2. At 207 the plotted intervals are **planned**; at 40 the first two are source-labelled existing and the deeper intervals are design. At 127ST2 the table is explicitly before ST2. Changes of specification are shown as individual segments, not assigned independent shoes. These are not final nesting, cement or barrier schematics. `update_w3_remove_legacy_shoes.py` removes old unvalidated shoe markers in the four drilling W3 plots. Do not rerun the older `replace_w6_intervals.py` or `replace_drilling_w3.py` main script over these documents; those target previous report content.

### Partial performance, fluid and BHA review

`fluid_observation_summary.json` is generated by `tools/replace_legacy_performance_fluid.py` from `drilling_track_observations.json`. Per recognized data sheet, it counts finite numeric observations and gives the source cells of the minimum and maximum. Mud weight is filtered to 30–200 in the source pcf convention; the other three fields accept nonnegative values. These are basic screening bounds, not validation of unit, date, value or source quality. Nonnumeric and rejected fields are counted separately; missing observations are not zero. The field names and source units must be checked sheet by sheet. Loss figures must not be summed until daily/cumulative basis and duplication are resolved. For example, 12 has 54 numeric loss entries, maximum 130 bbl; 85 has seven numeric zero entries, not evidence of no loss on other days; 127WO1 has no recognized numeric entries, not zero total loss. The old legacy performance activity-family percentages and total-loss assertions have been withdrawn from the 26 well reports. Source status codes are now tabulated using complete status pages independent of data-sheet availability, consistent with W4; this does not prove activity, productivity, chronology or causation.

`bha_observations.json` is an exploratory extraction of raw component, quantity, OD, ID, length and cumulative-length cells from nonempty blocks with recognized legacy or DDR headers on data-linked sheets. Rebuild after creating the local index:

```sh
/tmp/azns-index-venv/bin/python tools/extract_bha_observations.py /tmp/azns-evidence.sqlite evidence/bha_observations.json
```

A blank component terminates a block, only the first 15 subsequent rows are considered, and unrecognized layouts are listed, so coverage is incomplete. Multiple columns can be simultaneous; repeated BHA IDs and workbook ordering do not prove unique runs or chronology. Raw values have not been normalized or validated. `bha_selected_reviewed.json` contains one selected block per well and four bilingual selected component names, with explicit cells and programme navigation pages, inserted using `tools/add_bha_source_review.py`. It is not a complete assembly or a planned-vs-actual match. For 122WO1 `12-2 Data!I43` calls the scraper 7 inches while `K43` gives 5-3/8 inches; the discrepancy is unresolved. 207 programme pages 61–66 contain image-only figures, not fully parsed by text extraction. Programme page anchors are not a substitute for review of whole pages, operational phase, full string and source design. No hydraulic pressure drop or recommended operating envelope was calculated from incomplete geometry and unsynchronized observations. `tools/fix_bit_caption_order.py` places each 4-2 bit caption directly after its bit table, before any unassigned-tool table.
