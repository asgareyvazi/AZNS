#!/usr/bin/env python3
"""Replace unverified formation bands with actual source report-number samples."""
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from replace_drilling_w3 import files,plot
from replace_well_bit_evidence import text,style
for well in ('40','127WO1','127ST2'):
 png=plot(well,files[well],show_formations=False);Path(f'evidence/well_figures/{well}_W3.png').write_bytes(png)
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path);pars=d.paragraphs;i=next(i for i,p in enumerate(pars) if p.text.startswith(('Figure W3','شکل W3')))
  rid=pars[i-1]._p.xpath('.//a:blip')[0].get(qn('r:embed'));d.part.related_parts[rid]._blob=png
  sh=next(s for s in d.inline_shapes if any(b.get(qn('r:embed'))==rid for b in s._inline.xpath('.//a:blip')));sh.height=int(sh.width*10/14.5)
  text(pars[i],('شکل W3 ـ نمونه پارامترهای روزانه در عمق پایان گزارش؛ بدون نوار سازندی تأییدنشده' if fa else 'Figure W3 — Daily parameter samples at report-end depth; unverified formation bands omitted'));pars[i].alignment=1
  old=next(p for p in pars if p.text.startswith(('Points combine daily entries with report-end depth','مقادیر نقاط از داده روزانه')))
  text(old,('نقاط از داده روزانه در عمق پایان گزارش گرفته شده‌اند؛ این عمق، عمق ابزار یا رخداد و لاگ پیوسته نیست. ستون نخست شماره شیت است، نه شماره تأییدشده گزارش یا محور زمان پیوسته. نوارهای سازندی و کفشک‌ها تا تطبیق شاخه، مرجع عمق و وضعیت اجرا حذف شده‌اند. کمینه و بیشینه مقادیر تاریخی منبع‌اند، نه پنجره عملیاتی؛ RPM قالب قدیمی ممکن است شامل دور موتور باشد.' if fa else 'Points are daily entries at report-end depth, not tool/event depth or a continuous log. The first track is worksheet prefix, not verified report number or a continuous time axis. Formation bands and shoes are omitted until branch, depth datum and configuration are reconciled. Minima and maxima are source observations, not an operating window; legacy RPM may include motor RPM.'));style(old,fa);d.save(path)
 print(well)
