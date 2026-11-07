#!/usr/bin/env python3
"""Update W1 and field M6 for independently recognized status pages."""
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from replace_cbs_financials import set_text
from replace_well_bit_evidence import DATA,style
from apply_program_and_safety_revision import PROGS,TRACKS,ACTIVITY,plot
import replace_field_m3_m6 as field

def replace_image(d,caption,img,aspect):
 i=next(i for i,p in enumerate(d.paragraphs) if p.text.startswith(caption));b=d.paragraphs[i-1]._element.xpath('.//a:blip')[0];rid=b.get(qn('r:embed'));d.part.related_parts[rid]._blob=img
 sh=next(s for s in d.inline_shapes if any(x.get(qn('r:embed'))==rid for x in s._inline.xpath('.//a:blip')));sh.height=int(sh.width*aspect);return d.paragraphs[i]

for well,w in DATA.items():
 tr=TRACKS[w['source']];a=ACTIVITY[w['source']];image=plot(w,tr,PROGS[well]);complete=sum(d['activity']['complete'] for d in a['days'])
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path);cap=replace_image(d,('Figure W1','شکل W1'),image,6/11)
  for old in list(d.paragraphs):
   if old.text.startswith(('Status-page coverage is assessed independently:','پوشش صفحات وضعیت مستقل از شیت داده بررسی شده است:')):old._p.getparent().remove(old._p)
  note=d.add_paragraph((f"پوشش صفحات وضعیت مستقل از شیت داده بررسی شده است: {a['remark_days']} صفحه وضعیت و {len(tr['days'])} شیت داده شناسایی شد؛ {complete} صفحه وضعیت دارای جمع قابل طبقه‌بندی نزدیک ۲۴ ساعت‌اند. در برخی گزارش‌ها صفحه عملیات وجود دارد اما شیت داده قالب متفاوت دارد یا موجود نیست. این شمارها مدت تأییدشده دکل نیستند." if fa else f"Status-page coverage is assessed independently: {a['remark_days']} status pages and {len(tr['days'])} data sheets were recognized; {complete} status pages contain classified durations totalling approximately 24 hours. Some reports have a status page but an absent or differently formatted data page. These counts are not verified rig duration."));style(note,fa);cap._p.addnext(note._p);d.save(path)
 print(well,a['remark_days'],len(tr['days']),complete,flush=True)

field.tr={k:{**v,'days':ACTIVITY[k]['days']} for k,v in field.tr.items()};m6=field.m6()
for typ in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'Field_{typ}_Report_{lang}.docx');d=Document(path);cap=replace_image(d,('Figure M6','شکل M6'),m6,8/12);i=next(i for i,p in enumerate(d.paragraphs) if p._p is cap._p)
  p=d.paragraphs[i-2]
  set_text(p,'این شکل ساعت کدهای منبع را فقط برای گزارش‌های دارای ۲۴ ساعت طبقه‌بندی‌شده نشان می‌دهد. صفحات وضعیت مستقل از شیت داده بررسی شده‌اند؛ روزهای نامشخص صفر نیستند. کدهای غیرمولد و غیربرنامه‌ای جدا هستند و نمودار علت، مسئولیت قراردادی یا هزینه را تعیین نمی‌کند.' if fa else 'This figure shows source-code hours only for fully classified 24-hour reports. Status pages are identified independently of data pages; unknown days are not zero. NPT and unplanned codes remain distinct, and the figure does not determine cause, contractual liability or cost.')
  d.save(path)
