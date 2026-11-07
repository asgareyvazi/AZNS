#!/usr/bin/env python3
from pathlib import Path
from docx import Document
from replace_well_bit_evidence import DATA,style
for well in DATA:
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
  for p in list(d.paragraphs):
   if p.text.startswith(('W3 uses the data-linked extraction:', 'نمودار W3 از استخراج وابسته به شیت داده استفاده می‌کند:')):p._p.getparent().remove(p._p)
  cap=next(p for p in d.paragraphs if p.text.startswith(('Figure W3','شکل W3')))
  p=d.add_paragraph('نمودار W3 از استخراج وابسته به شیت داده استفاده می‌کند: شمار و ساعت‌های آن ممکن است کمتر از W2 و W4 باشد، زیرا صفحات وضعیت مستقل از وجود شیت داده در W2 و W4 بررسی شده‌اند. نقاط فاقد شیت داده معتبر به W3 افزوده نشده‌اند.' if fa else 'W3 uses the data-linked extraction: its coverage and hours may be smaller than W2 and W4, which assess status pages independently of data sheets. Missing or unrecognized data sheets are not added as W3 observations.');style(p,fa);cap._p.addnext(p._p);d.save(path)
