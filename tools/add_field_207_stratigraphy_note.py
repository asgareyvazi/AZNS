#!/usr/bin/env python3
"""Add the 207 programme Pabdeh thickness discrepancy to field reports."""
from pathlib import Path
from docx import Document
from replace_well_bit_evidence import style
for kind in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'Field_{kind}_Report_{lang}.docx');d=Document(path)
  marker='207 programme Pabdeh thickness discrepancy:' if not fa else 'اختلاف ضخامت پابده در برنامه ۲۰۷:'
  if any(p.text.startswith(marker) for p in d.paragraphs):continue
  anchor=next(p for p in d.paragraphs if p.text.startswith(('Programme source discrepancies:','اختلاف‌های منابع برنامه:')))
  desc=(' صفحه فیزیکی ۱۴ فایل AZNS-207 Program.pdf بازه ۱۹۹۴ تا ۲۱۷۸ متر را با ضخامت ۱۶۵ متر می‌نویسد، اما جدول صفحه ۱۷ ضخامت ۱۸۴ متر می‌دهد که برابر اختلاف حسابی ابتدا و انتهاست. هر دو مقدار در منبع حفظ شده‌اند و مبنا باید بررسی شود.' if fa else ' AZNS-207 Program.pdf physical page 14 states the Pabdeh interval as 1,994–2,178 m with a thickness of 165 m, while the table on page 17 gives 184 m, equal to the arithmetic endpoint difference. Both source values are retained pending review.')
  p=d.add_paragraph(marker+desc);style(p,fa);anchor._p.addnext(p._p);d.save(path)
