#!/usr/bin/env python3
from pathlib import Path
from docx import Document
from replace_well_bit_evidence import DATA
for well in DATA:
 for lang in ('EN','FA'):
  path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
  table=next(t for t in d.tables if t.rows and t.rows[0].cells[0].text in ('Bit','شماره') and len(t.rows[0].cells)==6)
  cap=next(p for p in d.paragraphs if p.text.startswith(('Table 4-2 — Identified bit observations','جدول ۴-۲ ـ مشاهدات مته با شناسه')))
  table._tbl.addnext(cap._p);cap.alignment=1;d.save(path)
