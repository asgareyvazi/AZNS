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

### Field-level source-scoped figures and narrative

`tools/build_field_source_figures.py` builds English-labelled 300 dpi PNGs in `field_figures/` and `field_figure_observations.json` records the data and source anchors. `tools/replace_field_legacy_claims.py` embeds them into the four bilingual field reports, replaces selected unsupported portfolio assertions and reconstructs field tables as source-review gates. It uses fixed paragraph indices and table order: validate document layout before reapplying to a changed report.

* M1 shows 14 programme-stated main/estimated durations, excluding uncombined moving and contingency; this is not a fleet timeline or actual duration. The full source filename and page for each point are preserved in JSON and the programme table.
* M2 counts all workbook sheets and corrects the ambiguous 127WO label to 127WO2 while explicitly mapping it to `AZNS127WOPack.xlsx`; worksheet counts are not completeness or report-day counts.
* M4 independently compares recognized status pages, status pages whose classified durations sum to 23.9–24.1 h, and data pages. M5 shows BHA-block, PV and numeric-loss recognition per recognized data page. Neither is rig performance or evidence that missing data mean zero.
* M7 plots four operations' report-end MD samples against source report number; no lines, common zero, causality, rate, formation assignment or learning curve are inferred. 75 includes an observed depth reset requiring phase context. Other wells are not represented by their final depth as a substitute for continuous data.
* M8 is an abridged view of two selected planned phases from each of six well 122/127 programmes, not a verified lifecycle or execution sequence. M9 lists six traceable reconciliation examples, not a quantified risk matrix.
* M10 pairs same-report numeric loss and report-end MD where both are recognized; all source cells are recorded, explicit zeros retained and other wells with no pairs omitted. It does not imply simultaneous measurement, loss-event depth, total loss, formation or mechanism. Its symlog vertical axis includes an explicit zero line.

The old unsupported field-wide cost, annual exposure, rig-ranking, contract-performance, causal and procurement conclusions were removed or replaced with review gates, not established by caveat alone. Financial examples cite cached CBS cells; none is actual expenditure or automatically additive. Programme service assignments are shown as printed, not as evidence of contracts or actual presence. All eight new figures were individually inspected at PNG level; that does not prove readable Word page layout. Two further field figures M3 and M6 remain the earlier source-review and source-code displays.

### Mud-material source fields and 207 programme stratigraphy

`mud_material_observations.json` is rebuilt from the read-only SQLite index by:

```sh
/tmp/azns-index-venv/bin/python tools/extract_mud_material_observations.py /tmp/azns-evidence.sqlite evidence/mud_material_observations.json
```

The tool reads material headers dynamically on the 1,248 recognized data sheets and retains sheet, cell, cached raw value, formula XML, material name, source unit and distinct used, received, total-received and stock fields. In the examined extraction it recognized material headers on all 1,248 data sheets and retained 23,131 rows with at least one source field; this is header/layout recognition, not proof of complete or reliable inventory. DDR headers and used columns move: 207 day 2 uses P7 and day 72 Q7; 116 `13-2 Sheet1` uses L39 rather than the usual F39. The DDR mud-material block is bounded to rows 8–49 and legacy block rows 40–90; other tables, including cement/additive blocks, are not covered. Raw blanks, spreadsheet errors, cached formulas, duplicate materials and source units are not normalized to zero. Values are not summed or converted, and stock or cumulative receipts are not classified as consumption. `mud_material_selected_reviewed.json` contains two selected sheets per operation and two or three positive source-used fields per sheet, with exact headers and cell provenance; these are illustrative, not a material balance or evidence of downhole treatment. `tools/add_material_source_evidence.py` inserts the bilingual selected observations, and is intentionally guarded against duplicate insertion.

`program_207_stratigraphy_reviewed.json` preserves 16 programme-prognosed tops from physical p.17 of `AZNS-207 Program.pdf`, matched to `AZNS207Pack.xlsx` `155-2 DDR Data!P71:P86`. The programme is a forecast (p.19 says depths may be adjusted during work), while the daily table's actual/prognosed basis and datum are not verified. The displayed differences are arithmetic, not actual geological deviations. The source itself calls the Pabdeh interval 1,994–2,178 m with 165 m thickness on p.14; p.17 lists 184 m, matching the endpoint difference. This contradiction is not silently resolved. The 207 W3 plot now uses only the p.17 forecast as background; the lower edge at 4,180 m is programme TD, not an asserted formation contact. It does not locate loss events in formations. The PNG was visually inspected. `tools/add_207_program_stratigraphy.py` checks every programme row's page text; `tools/replace_207_program_w3.py` regenerates the plot.

`tools/review_well_front_matter.py` replaces legacy well cover inventories, unsupported KPI/loss totals and invented issue/reviewer/approval statements with source-coverage and draft-status tables; it does not certify identity, rig assignments, operations, dates or costs. The old 85 rig-less source table was moved from before the cover into the programme section and its selected Persian technical prose translated. The source plan is not an instruction, executed work or contractual finding. The script edits known report structures and should be revalidated before reapplication.

### Additional daily-event and CBS service-sheet spot checks

`tools/add_well_event_reviews.py` uses source-cell anchors in `activity_observations.json` to add three individually translated selected observations to each bilingual well report; `well_event_reviews_selected.json` retains raw source text, anchor, paraphrases, and any status segment. These are excerpts, not complete chronology or incident conclusions. The tool is guarded by its inserted heading; changes to source or translation after insertion need deliberate document updates.

`tools/extract_cbs_workbook_sheet_inventory.py /tmp/azns-evidence.sqlite evidence/cbs_workbook_sheet_inventory.json` inventories all indexed sheets, A1 text and cached formula/error counts, without interpreting each service or recalculating formulas. Current inventory has 13 books, 196 sheets and 331 error cells. Sheet name, absent cells, a cached zero and an error are not interchangeable evidence. `tools/add_cbs_service_spot_checks.py` reads original workbooks in formula and cached modes, preserving selected values in `cbs_service_spot_checks.json`; it adds coiled-tubing samples for seven wells plus a 40 rig-less breakdown and a 127WO1 schedule-reference warning. It is guarded by its heading and must not be used to append samples after reports change without checking the existing content. Service estimates are not execution or incremental costs, and the seven samples do not exhaust service sheets. The F-20/127 sheet-name conflict is explicitly flagged. `tests/test_event_reviews_and_cbs_services.py` re-reads selected original cells, sheet names and cached error counts and checks insertion into both languages.

### Visually reviewed programme schematic tops

`tools/add_formation_schematic_review.py` records 81 manually paired tops from seven source pages in `program_schematic_tops_reviewed.json`, and adds bilingual source-labelled tables. PDF text extraction rechecks individual numeric strings on text-based pages, but does not prove their spatial association: those pairs were visually read from the pages. The 12 p.20 and ST2 p.25 figures are embedded images, for which the JSON stores SHA-256 fingerprints; the ST1 p.62 table has the same image fingerprint as ST2 p.25. These checks prove source-image identity, not independent correctness. The ST2 image shows original-well data in distinct forecast, sample, log and subsea columns; only selected MD(log) values are transcribed. The 12, 75, 85, 166, 127WO1 and 127WO2 tables are pre-workover programme schematics, not current as-built values. On 85, slash punctuation is preserved as printed, not interpreted as decimal punctuation. Datum, branch, revision and actual event depth must be checked before drawing formation/event conclusions. The script is guarded by the inserted headings, so later changes to selected data require deliberate updates to both reports and evidence.

`tools/replace_unverified_formation_bands.py` regenerates the 40, 127WO1 and 127ST2 W3 PNGs using `replace_drilling_w3.plot(..., show_formations=False)`, retaining daily parameter samples but replacing unverified formation bands with the source report-number track. The output depth is report-end depth, not tool/event depth; absence of formation bands is intentional pending branch/datum reconciliation. The three 300 dpi PNGs were inspected visually and are byte-matched to the embedded figures in both editions by `tests/test_formation_schematic_review.py`. Do not run older W3 regeneration scripts over these reports without reapplying the programme-specific/source-scope review.

### 122 WO1 selected campaign reconstruction

`tools/review_122wo1_campaign.py` reads the original 122 WO1 pack, programme and CBS, records 17 selected multilingual observations with 90 exact source cells, two conflicting date notes, source formula/cache fields and nine completion-component depths in `122wo1_campaign_review.json`, then inserts the evidence in the bilingual well reports. It is guarded by a heading, not a full synchronization mechanism. Report 29's contradictory top-of-fish value is retained exactly; report 34 claims complete milling of the fish tail, which is not independent evidence of recovery of the entire fish. A 55-day 3-hour elapsed interval calculated from the report's explicit spud and release timestamps is a calendar interval, not proof of approved comparable billed rig days. Programme p.27 describes rig-less steps; report 56's CT brine displacement does not prove acid-wash execution. Financial numbers are cached estimates and do not prove spend. `tools/add_field_122wo1_campaign_note.py` adds a qualified summary to the field documents.

`tools/build_122wo1_completion_depths.py` creates `well_figures/122WO1_completion_depths.png` from the nine depth values in report 54 and embeds it in both well editions. The original image was visually inspected and is not a physical configuration or barrier schematic. `tests/test_122wo1_campaign_review.py` re-reads the 90 selected source cells and completion entries, checks timestamps and CBS formulas/caches, and byte-compares the PNG embedded in both DOCX files. None of these checks replaces interpretation of complete signed operational, pressure-test or financial records.
