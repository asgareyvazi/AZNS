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
