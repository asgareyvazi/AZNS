#!/usr/bin/env python3
"""Replace 207 W3 formation bands with the p.17 forecast, not a DDR snapshot."""
from pathlib import Path
import json
from docx import Document
from docx.oxml.ns import qn
from replace_well_bit_evidence import text
import replace_drilling_w3 as w3

j=json.loads(Path('evidence/program_207_stratigraphy_reviewed.json').read_text());name='AZNS207Pack.xlsx'
w3.formations[name]={'tables':[{'rows':[{'formation':r['name'],'cells':[{'header':'m MD','value':str(r['programme_md_for_plot'])}]} for r in j['rows']]}]}
png=w3.plot('207',name,programme_bottom=j['programme_limit_md_m'],programme_reference='AZNS-207 Program.pdf, physical p.17')
out=Path('evidence/well_figures/207_W3.png');out.parent.mkdir(exist_ok=True);out.write_bytes(png)
for lang in ('EN','FA'):
 fa=lang=='FA';path=Path(f'207_Drilling_Engineering_Report_{lang}.docx');d=Document(path);pars=d.paragraphs;i=next(i for i,p in enumerate(pars) if p.text.startswith(('Figure W3','شکل W3')))
 rid=pars[i-1]._p.xpath('.//a:blip')[0].get(qn('r:embed'));d.part.related_parts[rid]._blob=png
 sh=next(s for s in d.inline_shapes if any(b.get(qn('r:embed'))==rid for b in s._inline.xpath('.//a:blip')));sh.height=int(sh.width*10/14.5)
 text(pars[i],('شکل W3 ـ نمونه پارامترهای روزانه با نوار سازندی پیش‌بینی‌شده برنامه، نه انتساب رویداد' if fa else 'Figure W3 — Daily parameter samples with programme-prognosed bands, not event attribution'));pars[i].alignment=1
 old=next(p for p in pars if p.text.startswith(('Points combine daily entries with report-end depth','مقادیر نقاط از داده روزانه')))
 text(old,('نوارهای سازند از رأس‌های پیش‌بینی‌شده صفحه فیزیکی ۱۷ برنامه چاه ۲۰۷ گرفته شده‌اند و تا عمق نهایی برنامه ۴۱۸۰ متر محدودند، نه یک مرز اثبات‌شده سازند. نقاط پارامترها نمونه روزانه در عمق پایان گزارش‌اند، نه عمق ابزار، رویداد یا لاگ پیوسته. تطبیق واقعی/پیش‌بینی بودن، مبنای عمق، واحد و هم‌زمانی لازم است؛ کفشک‌ها نمایش داده نشده‌اند. مقدار مشکوک دبی ۷۵۰۰ گالن بر دقیقه از ترسیم حذف و در شواهد حفظ شده است.' if fa else 'Formation bands use programme-prognosed tops on physical page 17 and stop at the programme TD of 4,180 m, not a proven formation contact. Parameter points are daily entries at report-end depth, not tool/event depths or a continuous log. Actual versus forecast status, datum, units and simultaneity require reconciliation; shoes are not plotted. The suspect 7,500 gpm flow entry is excluded from the plot but retained in evidence.'))
 d.save(path)
