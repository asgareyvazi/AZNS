#!/usr/bin/env python3
"""Withdraw unverified portfolio comparisons and use cited programme summaries."""
from pathlib import Path
from docx import Document
from docx.shared import Pt
from replace_cbs_financials import celltext,set_text
from apply_program_and_safety_revision import PROGS,DATA,TRACKS

for well in DATA:
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
  for t in d.tables:
   if t.rows and t.cell(0,0).text=='Reconciliation metric':
    for row in t.rows[1:]:
     if row.cells[0].text.startswith('Depth reached'):
      for c in row.cells[1:]:celltext(c,'Not reconciled',False)
  for p in d.paragraphs:
   if p.text in ('Recommendations and review requirements','Risk-review checklist; ownership and severity not assigned','توصیه‌ها و الزامات بازبینی','فهرست بررسی ریسک؛ مسئولیت و شدت تعیین نشده‌اند'):
    for r in p.runs:r.font.size=Pt(13 if p.style.name=='Heading 1' else 11)
  d.save(path)

for typ in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'Field_{typ}_Report_{lang}.docx');d=Document(path)
  if typ.startswith('Master'):
   count=0
   for para in d.paragraphs:
    x=para.text
    if not x.startswith('AZNS-') or not ('rig days against' in x or 'روز دکل در برابر برنامه' in x):continue
    tag=x.split(' (',1)[0].replace('AZNS-','').replace(' ','');assert tag in PROGS,(tag,x[:60]);p=PROGS[tag];w=DATA[tag];tr=TRACKS[w['source']]
    msg=(f"چاه {tag}: برنامه {p['file']} در صفحه فیزیکی {p['physical_page']} چنین هدفی را ذکر می‌کند: {p['objective_fa']} مدت برنامه {p['duration_days_fa']} روز ثبت شده است. {len(tr['days'])} شیت داده از {tr['paired_sheet_candidates']} جفت نامزد شناسایی شده؛ این شمار، مدت اجرای تأییدشده نیست. مقایسه زمان، هزینه و علت حوادث به تطبیق مستقل نیاز دارد." if fa else f"AZNS-{tag}: {p['file']}, physical page {p['physical_page']}, states: {p['objective_en']} Programme duration is {p['duration_days_en']} days. {len(tr['days'])} data sheets were recognized out of {tr['paired_sheet_candidates']} candidate pairs; this count is not verified execution duration. Schedule, cost and event causes require independent reconciliation.")
    set_text(para,msg);count+=1
   assert count==13,(lang,count)
  for t in d.tables:
   if not t.rows:continue
   hs=[c.text.strip() for c in t.rows[0].cells];head=hs[0]
   # Preserve new primary-source programme summary, which is not a legacy comparison.
   if 'Programme duration (days)' in hs or 'مدت برنامه (روز)' in hs:continue
   uncertain={'Rig days','Planned days','Actual days','Ratio','Slip (d)','Planned d','Planned','Schedule ratio','NPT %','Days','سهم غیرمولد','روز دکل','روزهای دکل','مدت برنامه‌ای','مدت واقعی','نسبت','انحراف زمانی','برنامه','نسبت زمانی','روزها','روز','درصد غیرمولد','درصد زمان غیرمولد'}
   if head in ('Technical cause','علت فنی'):
    for row in t.rows[1:]:
     for c in row.cells[1:]:celltext(c,'تطبیق‌نشده' if fa else 'Not reconciled',fa)
   else:
    for j,h in enumerate(hs):
     if h in uncertain:
      for row in t.rows[1:]:celltext(row.cells[j],'تطبیق‌نشده' if fa else 'Not reconciled',fa)
  d.save(path)
