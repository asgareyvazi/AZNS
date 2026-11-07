#!/usr/bin/env python3
"""Insert a limited source-backed phase and service comparison in 30 reports."""
import json
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from replace_well_bit_evidence import DATA,style
from docx.oxml.ns import qn
from pypdf import PdfReader

P={x['well']:x for x in json.loads(Path('evidence/program_phases_services_reviewed.json').read_text())}
E={x['well']:x for x in json.loads(Path('evidence/daily_observation_examples.json').read_text())}
TRANSLATE={'VIRA':'ویرا','Pars Hafaran Arvand':'پارس حفاران اروند','SEALAND':'سیلند','MAPSA KISH':'مپسا کیش','PANIZ':'پانیز','GEO DATA':'ژئودیتا','Pars Geo Data':'پارس ژئودیتا','TBA':'تعیین‌نشده در برنامه','APD':'ای‌پی‌دی','PANIZ/SEALAND':'پانیز / سیلند','PGD':'پی‌جی‌دی'}

def add_after(d,after,text,fa,heading=False):
 p=d.add_paragraph(style='Heading 2' if heading else None);p.add_run(text);style(p,fa)
 if heading:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
 after._p.addnext(p._p);return p

def ensure_12_warning():
 a=P['12'];b=P['127WO1'];aa=' '.join((PdfReader(a['file']).pages[27].extract_text() or '').split());bb=' '.join((PdfReader(b['file']).pages[57].extract_text() or '').split());phrase='Check and smooth collapse interval from 3889-3911m.';assert phrase in aa and phrase in bb
ensure_12_warning()

for well in DATA:
 plan=P[well];examples=E[well];sv=plan['service_assignment_as_printed']
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
  heading=('مراحل منتخب برنامه و شواهد اجرای محدود' if fa else 'Selected programme phases and limited execution evidence')
  # The first insertion is before the next top-level section; re-run removes
  # the entire subsection to preserve idempotence.
  old=next((p for p in d.paragraphs if p.text==heading),None)
  if old:
   e=old._p;end=next(p for p in d.paragraphs if p.text in ('Selected daily-report observations','روایت منتخب از گزارش روزانه'))._p
   while e is not end:n=e.getnext();e.getparent().remove(e);e=n
  following=next(p for p in d.paragraphs if p.text in ('Selected daily-report observations','روایت منتخب از گزارش روزانه'))
  # build in natural order before following
  content=[]
  content.append((heading,True))
  content.append((('مراحل زیر گزیده برنامه‌اند، نه اثبات تحقق یا برنامه مصوب نهایی. مشاهدات روزانه انتخابی مستقل از این مراحل نقل شده‌اند؛ مشابهت موضوع به‌تنهایی تطبیق زمانی، تکمیل مرحله یا انحراف از برنامه را ثابت نمی‌کند.' if fa else 'The phases below are selected programme statements, not proof of execution or of the final approved revision. The selected daily observations are quoted separately; a similar topic alone does not establish phase timing, completion or deviation.'),False))
  for ph in plan['selected_plan_phases']:
   content.append(((f"برنامه ـ {ph['persian']} [{plan['file']}، صفحه فیزیکی {ph['physical_page']}]" if fa else f"Programme — {ph['english']} [{plan['file']}, physical p. {ph['physical_page']}]"),False))
  for ex in examples['selected_examples']:
   content.append(((f"شیت منبع {ex['day']} ـ {ex['persian_summary']} [{examples['file']}، {ex['sheet']}!{ex['cell']}]" if fa else f"Source worksheet {ex['day']} — {ex['english_summary']} [{examples['file']}, {ex['sheet']}!{ex['cell']}]"),False))
  if well=='12':content.append((('هشدار تطبیق: عبارت بررسی فروریختگی ۳۸۸۹ تا ۳۹۱۱ متر در صفحه فیزیکی ۲۸ برنامه چاه ۱۲ و عیناً در صفحه ۵۸ برنامه تعمیر اول ۱۲۷ آمده است. این همسانی تعلق بازه به هر چاه را ثابت نمی‌کند؛ پیش از استفاده باید با سوابق چاه و نسخه تأییدشده تطبیق شود.' if fa else 'Reconciliation flag: the wording about checking a 3,889–3,911 m collapse interval appears both on physical p. 28 of the well 12 programme and p. 58 of the 127 WO1 programme. This duplication does not establish which well the interval applies to; verify against well-specific records and the approved revision before use.'),False))
  if fa:
   service=f"خدمات درج‌شده در برنامه: سیال حفاری و تکمیل: {TRANSLATE[sv['drilling_completion_fluids']]}؛ مدیریت پسماند: {TRANSLATE[sv['waste_management']]}؛ نمودارگیری الکتریکی و مشبک‌کاری: {TRANSLATE[sv['electrical_logging_perforation']]}؛ گل‌نگاری: {TRANSLATE[sv['mud_logging']]}. [ {plan['file']}، صفحه فیزیکی {plan['services_physical_page']} ] این فهرست انتساب برنامه است، نه اثبات اعزام، قرارداد قطعی، هزینه یا مسئولیت حادثه."
  else:
   service=f"Services listed in the programme: drilling and completion fluids: {sv['drilling_completion_fluids']}; waste management: {sv['waste_management']}; electrical logging and perforation: {sv['electrical_logging_perforation']}; mud logging: {sv['mud_logging']}. [{plan['file']}, physical p. {plan['services_physical_page']}] These are programme assignments, not evidence of actual deployment, contractual award, cost or incident liability."
  content.append((service,False))
  for text,isheading in content:
   p=d.add_paragraph(style='Heading 2' if isheading else None);p.add_run(text);style(p,fa)
   if isheading:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
   following._p.addprevious(p._p)
  d.save(path)

for typ in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'Field_{typ}_Report_{lang}.docx');d=Document(path)
  target=next(p for p in d.paragraphs if p.text.startswith(('The 127 ST1 programme is distinct','برنامه انحراف اول چاه ۱۲۷ مستقل')))
  for p in list(d.paragraphs):
   if p.text.startswith(('Additional ST1 phase references:','ارجاعات تکمیلی فاز انحراف اول:')):p._p.getparent().remove(p._p)
  ph=P['127ST1']['selected_plan_phases']
  text=(f"ارجاعات تکمیلی فاز انحراف اول: {ph[0]['persian']} [{P['127ST1']['file']}، صفحه فیزیکی ۷۰] {ph[1]['persian']} [صفحه فیزیکی ۷۱]. صفحه ۱۱ عمق آغاز انحراف را ۳۳۰۰ متر بیان می‌کند؛ رابطه این عدد با محل پنجره صفحه ۷۰ باید از نقشه و نسخه مصوب تطبیق شود. این موارد اجرای واقعی انحراف اول را ثابت نمی‌کنند." if fa else f"Additional ST1 phase references: {ph[0]['english']} [{P['127ST1']['file']}, physical p. 70] {ph[1]['english']} [physical p. 71]. Page 11 states a sidetrack from 3,300 m; reconcile its definition against the window position on p. 70 and the approved design. These statements do not prove actual ST1 execution.")
  add_after(d,target,text,fa);d.save(path)
print('Updated 26 well and four field reports')
