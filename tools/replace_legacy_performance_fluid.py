#!/usr/bin/env python3
"""Replace unsupported performance and fluid aggregates with traceable subsets."""
from pathlib import Path
import json,math
from docx import Document
from docx.oxml.ns import qn
from replace_well_bit_evidence import DATA,set_table,style,text

TRACKS={x['file']:x for x in json.loads(Path('evidence/drilling_track_observations.json').read_text())}
ACT={x['file']:x for x in json.loads(Path('evidence/activity_observations.json').read_text())}
KEYS=('mud_weight','pv','yp','downhole_loss')

def valid(k,x):
 if x is None or x.get('value') is None:return False
 v=x['value']
 if not math.isfinite(v):return False
 # Only basic filters; they do not validate unit, sample type or quality.
 return 30<=v<=200 if k=='mud_weight' else v>=0

def evidence(w):
 tr=TRACKS[w['source']];summ={}
 for key in KEYS:
  allx=[d['fields'].get(key) for d in tr['days'] if d['fields'].get(key)]
  xs=[x for x in allx if valid(key,x)]
  summ[key]={'numeric_accepted_count':len(xs),'present_but_rejected_or_non_numeric_count':len(allx)-len(xs),
   'minimum':min(xs,key=lambda x:x['value']) if xs else None,'maximum':max(xs,key=lambda x:x['value']) if xs else None}
 return {'well':w['well'],'file':w['source'],'recognized_data_sheets':len(tr['days']),'candidate_pairs':tr['paired_sheet_candidates'],'fields':summ,'method':'One observed source field per recognized data sheet; incomplete coverage. Finite numeric values >=0, except mud weight restricted to 30–200, retained; rejected and nonnumeric entries are not zero. No verified daily loss total, reservoir attribution, continuous measurements, current mud properties or design window is inferred. Full source entries and cells are in drilling_track_observations.json.'}

def disp(x,fa):
 if not x:return 'ثبت عددی شناسایی نشد' if fa else 'No numeric entry recognized'
 return f"{x['value']:g} [{x['sheet']}!{x['cell']}]"

def main():
 out=[]
 for well,w in DATA.items():
  ev=evidence(w);out.append(ev);a=ACT[w['source']];full=[d['activity'] for d in a['days'] if d['activity']['complete']];vals=[sum(x[k] for x in full) for k in ('planned_h','npt_h','unplanned_h','mixed_h')];tot=sum(vals)
  for lang in ('EN','FA'):
   fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
   # Replace the prior generated activity-family percentages, which mixed
   # unverified labels with partial daily data, with explicit source statuses.
   old=next(t for t in d.tables if t.rows and t.rows[0].cells[0].text in ('Activity family','گروه فعالیت','خانواده فعالیت','کد وضعیت منبع','Source status code'))
   rows=[(['کد وضعیت منبع','ساعت در گزارش‌های کامل','سهم از ساعت طبقه‌بندی‌شده'] if fa else ['Source status code','Hours in complete reports','Share of classified hours'])]
   labels=['برنامه‌ای ثبت‌شده','غیرمولد ثبت‌شده','غیربرنامه‌ای ثبت‌شده','کد ترکیبی غیرمولد–غیربرنامه‌ای'] if fa else ['Reported planned','N / NPT','UP / unplanned','NPT-UNPLN (mixed)']
   rows += [[name,f'{v:,.2f}',f'{100*v/tot:.1f}%' if tot else '—'] for name,v in zip(labels,vals)];set_table(old,rows,[7.0,4.4,4.6],fa)
   cap=next(p for p in d.paragraphs if p.text.startswith(('Table 4-1','جدول ۴-۱')));text(cap,'جدول ۴-۱ ـ کدهای وضعیت در گزارش‌های کامل شناسایی‌شده، نه تفکیک علل' if fa else 'Table 4-1 — Source status codes in recognized complete reports, not causal categories');cap.alignment=1
   h=next(p for p in d.paragraphs if p.text in ('Performance and Parameter Analysis','تحلیل عملکرد و پارامترهای عملیاتی'))
   # Replace only narrative paragraphs preceding W3; keep its image and notes.
   e=h._p.getnext()
   while e is not None and e is not cap._p:
    nxt=e.getnext()
    if e.tag==qn('w:p') and e.xpath('.//w:t'):e.getparent().remove(e)
    e=nxt
   prior=next((p for p in list(d.paragraphs) if p.text.startswith(('Performance coverage:', 'پوشش تحلیل عملکرد:'))),None)
   if prior:prior._p.getparent().remove(prior._p)
   p=d.add_paragraph((f"پوشش تحلیل عملکرد: {len(full)} صفحه وضعیت با جمع مدت‌های قابل طبقه‌بندی نزدیک ۲۴ ساعت از {a['remark_days']} صفحه وضعیت شناسایی‌شده استفاده شده است. برنامه‌ای به‌معنی بهره‌وری اثبات‌شده نیست و کدهای غیرمولد، غیربرنامه‌ای و ترکیبی جدا نگه داشته می‌شوند. ساعت‌های حذف‌شده صفر فرض نشده‌اند؛ نوع فعالیت، علل و مدت کل عملیات از این جدول تعیین نمی‌شوند." if fa else f"Performance coverage: {len(full)} status pages with classified durations totalling approximately 24 hours are used, out of {a['remark_days']} recognized status pages. Planned does not prove productive time; NPT, unplanned and mixed codes are distinct. Omitted hours are not zero; activity families, causes and total operation duration are not established by this table."));style(p,fa);old._tbl.addprevious(p._p)
   # Drop unsupported productivity/progress narrative between table caption and
   # W3 image, preserving other figures and the new W3 scope notes.
   el=cap._p.getnext()
   while el is not None:
    if el.tag==qn('w:p') and el.xpath('.//a:blip'):break
    nxt=el.getnext()
    if el.tag==qn('w:p') and el.xpath('.//w:t'):el.getparent().remove(el)
    el=nxt
   insert=d.add_paragraph('مقایسه نرخ پیشروی و فازها به تطبیق دامنه، زمان و ماهیت فعالیت نیاز دارد؛ عمق پایان گزارش، عمق ابزار یا نرخ برخط نیست.' if fa else 'Progress and phase-rate comparisons require reconciliation of scope, time and activity. Report-end depth is not tool depth or instantaneous rate.');style(insert,fa);cap._p.addnext(insert._p)
   # Replace entire legacy fluid section; unsupported sum of missing losses is
   # not a zero and unknown reporting basis precludes adding values.
   start=next(p for p in d.paragraphs if p.text in ('Drilling Fluid Engineering and Loss Accounting','مهندسی سیال حفاری و حساب‌داری افت'))
   end=next(p for p in d.paragraphs if p.text in ('Completion, Casing and Service-Module Findings','یافته‌های تکمیل، جداری و ماژول‌های خدماتی'))
   node=start._p.getnext()
   while node is not end._p:
    nxt=node.getnext()
    if nxt is None:raise ValueError((path,'end not found'))
    node.getparent().remove(node);node=nxt
   intro=(f"این جدول فقط مقادیر عددی بازیابی‌شده از {ev['recognized_data_sheets']} شیت داده شناخته‌شده از {ev['candidate_pairs']} زوج نامزد را خلاصه می‌کند. هر حد همراه نمونه سلول منبع است؛ پوشش، هم‌زمانی، صحت طبقه‌بندی واحد و کیفیت آزمایش باید تأیید شوند. حداقل و حداکثر، پنجره طراحی یا وضعیت جاری سیال نیستند." if fa else f"This table summarizes numeric entries recovered from {ev['recognized_data_sheets']} recognized data sheets out of {ev['candidate_pairs']} candidate pairs. Each extreme has a source-cell example; coverage, simultaneity, units and test quality require verification. Minima and maxima are not a design window or current fluid condition.")
   p=d.add_paragraph(intro);style(p,fa);end._p.addprevious(p._p)
   table=d.add_table(rows=1,cols=4);table.style='Table Grid';end._p.addprevious(table._tbl)
   en=['Mud weight (pcf)','Plastic viscosity (cP)','Yield point (lb/100 ft²)','Source downhole loss (bbl)'];fr=['وزن گل (پوند بر فوت مکعب)','گرانروی پلاستیک (سانتی‌پواز)','نقطه تسلیم (پوند بر ۱۰۰ فوت مربع)','هرزروی درون‌چاهی ثبت‌شده (بشکه)']
   rows=[(['مشاهده منبع','تعداد عددی','کمینه و سلول','بیشینه و سلول'] if fa else ['Source observation','Numeric entries','Minimum and cell','Maximum and cell'])]
   for label,key in zip(fr if fa else en,KEYS):
    x=ev['fields'][key];rows.append([label,str(x['numeric_accepted_count']),disp(x['minimum'],fa),disp(x['maximum'],fa)])
   set_table(table,rows,[4.0,2.2,4.9,4.9],fa)
   caption=d.add_paragraph('جدول ۵-۱ ـ گزیده آماری سلول‌های سیال و هرزروی با پوشش جزئی' if fa else 'Table 5-1 — Partial source-cell fluid and loss observations');style(caption,fa,True);end._p.addprevious(caption._p)
   note=('سلول خالی یا بدون مقدار عددی صفر نیست. برچسب منبع برای هرزروی، حجم درون‌چاهی یا سازندی گزارش‌شده است؛ بدون احراز مبنای ثبت، تجمعی‌بودن و تکرار، جمع یا انتساب به یک سازند محاسبه نشده است. اختلاف مقدار سلول با شرح عملیات و مصرف مواد باید جداگانه بررسی شود.' if fa else 'A blank or nonnumeric entry is not zero. The source loss field is labelled downhole or formation loss volume; no sum or formation attribution is calculated until reporting basis, cumulative versus daily treatment and duplication are reconciled. Reconcile cell values with operational remarks and materials use separately.')
   p=d.add_paragraph(note);style(p,fa);end._p.addprevious(p._p);d.save(path)
  print(well,len(full),ev['fields']['mud_weight']['numeric_accepted_count'],ev['fields']['downhole_loss']['numeric_accepted_count'],flush=True)
 Path('evidence/fluid_observation_summary.json').write_text(json.dumps(out,indent=2,ensure_ascii=False)+'\n')
if __name__=='__main__':main()
