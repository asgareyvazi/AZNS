#!/usr/bin/env python3
"""Withdraw unsupported legacy field schedule/NPT totals and label duration basis."""
from pathlib import Path
from docx import Document
from replace_cbs_financials import celltext
from replace_well_bit_evidence import DATA,style
for well in DATA:
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
  for t in d.tables:
   if t.rows and t.cell(0,0).text in ('Reconciliation metric','شاخص تطبیق'):
    for row in t.rows[1:]:
     if row.cells[0].text.startswith(('Duration','مدت')):celltext(row.cells[0],'مدت ثبت‌شده در برنامه (روز)' if fa else 'Programme-stated duration (days)',fa)
  d.save(path)
for typ in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'Field_{typ}_Report_{lang}.docx');d=Document(path)
  for t in d.tables:
   if not t.rows:continue
   hs=[c.text.strip() for c in t.rows[0].cells]
   for j,h in enumerate(hs):
    if h in ('روز برنامه','روز واقعی','انحراف (روز)','درصد NPT'):
     for row in t.rows[1:]:celltext(row.cells[j],'تطبیق‌نشده',True)
   for row in t.rows[1:]:
    if row.cells[0].text in ('Total rig days','Total rig hours','Unplanned hours','New-hole metreage','مجموع روزهای دکل','مجموع ساعات دکل','ساعت غیربرنامه‌ریزی‌شده','متراژ حفره جدید','Rig days consumed','Unplanned time','روز دکل مصرف‌شده','زمان غیرمولد'):
     if hs[0] in ('Portfolio metric','شاخص سبد','Measure','شاخص'):
      for c in row.cells[1:]:celltext(c,'تطبیق‌نشده' if fa else 'Not reconciled',fa)
  if not any(p.text.startswith(('The 127 ST1 programme is distinct', 'برنامه انحراف اول چاه ۱۲۷ مستقل')) for p in d.paragraphs):
   t=next(t for t in d.tables if t.rows and any(c.text in ('Programme duration (days)','مدت برنامه (روز)') for c in t.rows[0].cells))
   p=d.add_paragraph('برنامه انحراف اول چاه ۱۲۷ مستقل از برنامه انحراف دوم است: در صفحه فیزیکی ۱۱ فایل AZNS -127 Side Track#01 Program Rev.03.pdf گشودن پنجره و انحراف از ۳۳۰۰ متر، حفاری حفره‌های ۸½ و ۵⅞ اینچ و تکمیل، با مدت ۸۰ روز برنامه اصلی، ذکر شده است. این برنامه به‌تنهایی اثبات اجرای عملیات یا گزارش روزانه مستقل نیست.' if fa else 'The 127 ST1 programme is distinct from ST2: physical page 11 of AZNS -127 Side Track#01 Program Rev.03.pdf states a window and sidetrack from 3,300 m MD, 8½- and 5⅞-inch holes, completion and an 80-day main plan. The programme alone does not establish execution or a separate daily-report record.')
   style(p,fa);t._tbl.addnext(p._p)
  d.save(path)
