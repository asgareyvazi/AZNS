#!/usr/bin/env python3
"""Refresh exploratory W3 plots without unverified legacy shoe markers."""
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from replace_drilling_w3 import files,plot
from replace_well_bit_evidence import text,style
for well,file in files.items():
 png=plot(well,file)
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path);i=next(i for i,p in enumerate(d.paragraphs) if p.text.startswith(('Figure W3','شکل W3')));p=d.paragraphs[i-1];bl=p._element.xpath('.//a:blip')[0];rid=bl.get(qn('r:embed'));d.part.related_parts[rid]._blob=png
  sh=next(s for s in d.inline_shapes if any(x.get(qn('r:embed'))==rid for x in s._inline.xpath('.//a:blip')));sh.height=int(sh.width*10/14.5)
  for p in d.paragraphs:
   if p.text.startswith('Points combine daily entries with report-end depth'):
    text(p,p.text.replace('Formation tops and shoe markers require reconciliation.','Formation bands are unverified and shoe markers are omitted pending reconciliation.'));style(p,fa)
   if p.text.startswith('مقادیر نقاط از داده روزانه'):
    text(p,p.text.replace('رأس سازند و کفشک‌ها به تطبیق نیاز دارند.','بازه‌های سازندی تأیید نشده‌اند و نشانه کفشک تا هنگام تطبیق حذف شده است.'));style(p,fa)
  d.save(path)
 print(well)
