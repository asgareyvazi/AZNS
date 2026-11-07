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
