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
