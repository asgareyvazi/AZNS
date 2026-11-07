#!/usr/bin/env python3
"""Inject page-cited programme summaries, withdraw unsupported comparisons and envelopes."""
from pathlib import Path
from io import BytesIO
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from replace_well_bit_evidence import DATA,style,set_table,text
from replace_cbs_financials import celltext

PROGS={x['well']:x for x in json.loads(Path('evidence/program_summary_reviewed.json').read_text())}
TRACKS={x['file']:x for x in json.loads(Path('evidence/drilling_track_observations.json').read_text())}

def plot(w,track,p):
  complete=sum(d['activity']['complete'] for d in track['days']);vals=[track['paired_sheet_candidates'],len(track['days']),complete]
  fig,ax=plt.subplots(figsize=(11,6),dpi=300);fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white')
  ax.barh(range(3),vals,color=['#90aabd','#28688d','#d0a153'],height=.52);ax.set_yticks(range(3),['Candidate sheet pairs','Recognized data sheets','Complete classified days'],fontsize=10);ax.invert_yaxis();maxv=max(vals) or 1;ax.set_xlim(0,maxv*1.26)
  for i,v in enumerate(vals):ax.text(v+maxv*.02,i,str(v),va='center',fontsize=10)
  ax.set_xlabel('Pairs, sheets or days as labelled; not verified elapsed or rig days',fontsize=9);ax.grid(axis='x',alpha=.18);ax.set_axisbelow(True);ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0)
  fig.suptitle(f"AZNS-{w['well']} | Programme and report coverage",x=.05,y=.98,ha='left',fontsize=15,fontweight='bold',color='#223952')
  fig.text(.05,.92,f"Programme duration (days): {p['duration_days_en']}. Actual comparison is not established.",fontsize=9,color='#9b431c')
  fig.text(.05,.05,f"Sources: programme physical p. {p['physical_page']}; {w['source']}. Candidate sheets are not proof of rig duration.",fontsize=8.5,color='#9b431c')
  fig.subplots_adjust(left=.29,right=.96,top=.82,bottom=.18);fig.canvas.draw();renderer=fig.canvas.get_renderer()
  for x in [*fig.texts,*ax.get_yticklabels(),*ax.texts]:
    box=x.get_window_extent(renderer)
    if box.x0<0 or box.x1>fig.bbox.x1 or box.y0<0 or box.y1>fig.bbox.y1:raise ValueError((w['well'],x.get_text()))
  out=BytesIO();fig.savefig(out,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return out.getvalue()

def insert_before(anchor,d,content,fa,heading=None):
 p=d.add_paragraph(content,style=heading) if heading else d.add_paragraph(content)
 style(p,fa,bool(heading));anchor.addprevious(p._p);return p

def save_section(doc,well,fa,w,track,p):
  begin=next(x for x in doc.paragraphs if x.text.startswith(('Section-Wise Optimum Parameters','پارامترهای بهینه به‌تفکیک بخش')))
  end=next(x for x in doc.paragraphs if x.text.startswith(('Data Provenance and Cell-Level Traceability','منشأ داده و ردیابی در سطح سلول')))
  text(begin,'بازبینی پارامترهای عملیاتی و دروازه‌های تصمیم' if fa else 'Operational observations and engineering decision gates');begin.alignment=WD_ALIGN_PARAGRAPH.CENTER
  cursor=begin._p.getnext()
  while cursor is not end._p:
    nxt=cursor.getnext()
    if nxt is None:raise ValueError((well,'section boundary'))
    cursor.getparent().remove(cursor);cursor=nxt
  insert_before(end._p,doc,'دامنه‌های عددی توصیه‌شده قبلی تأیید نشده‌اند و کنار گذاشته شده‌اند. مقادیر زیر فقط نمونه‌های خام ثبت روزانه‌اند؛ محل ابزار، یکتایی فاز، واحد و اعتبار اندازه‌گیری باید بررسی شود. این مشاهدات محدوده مجاز، پنجره بهینه یا دستور عملیات نیستند.' if fa else 'Previous numerical operating envelopes were not validated and have been withdrawn. The entries below are examples of raw daily records only; tool location, phase, units and measurement validity require review. These observations are not safe limits, optimum windows or operating instructions.',fa)
  insert_before(end._p,doc,'جدول ۹-۱ ـ نمونه سلول‌های ثبت‌شده؛ واحد و مبنای اندازه‌گیری نیازمند تأیید.' if fa else 'Table 9-1 — Example recorded cells; units and measurement basis require verification.',fa)
  t=doc.add_table(rows=1,cols=3);t.style='Table Grid';end._p.addprevious(t._tbl)
  data=[(['پارامتر خام','نمونه مقدار ثبت‌شده','شیت و سلول'] if fa else ['Raw field','Example recorded value','Sheet and cell'])]
  fields=[('mud_weight','Mud weight','وزن گل'),('wob_a_min','Bit load / WOB','بار مته'),('rpm_a_min','RPM','سرعت دوران'),('flow_a_min','Flow','دبی'),('spp_a_min','Standpipe pressure','فشار لوله ایستاده')]
  for key,en,ar in fields:
    candidates=[d['fields'].get(key) for d in track['days'] if d['fields'].get(key) and d['fields'][key].get('value') is not None]
    item=candidates[0] if candidates else None
    data.append([ar if fa else en,str(item['raw']).strip() if item else ('ثبت عددی قابل تفسیر یافت نشد' if fa else 'No interpretable numeric sample'),f"{item['sheet']}!{item['cell']}" if item else '—'])
  set_table(t,data,[4.4,5.0,7.4],fa)
  insert_before(end._p,doc,'توصیه‌ها و الزامات بازبینی' if fa else 'Recommendations and review requirements',fa,'Heading 2')
  for en,fr in [
   ('Reconcile the programme revision, scope, start/end boundaries and rig/service time against daily reports and approved schedules before reporting a variance.','پیش از گزارش انحراف، نسخه و دامنه برنامه، مرزهای آغاز و پایان و زمان دکل و خدمات با گزارش‌های روزانه و برنامه مصوب تطبیق شوند.'),
   ('Have responsible specialists verify well barriers and formation, casing and equipment limits against current approved records; no numerical design setpoint is established by this report.','متخصصان مسئول، سدهای ایمنی چاه و حدود سازند، جداری و تجهیزات را با اسناد مصوب و جاری بررسی کنند؛ این گزارش هیچ نقطه تنظیم عددی طراحی را اثبات نمی‌کند.'),
   ('Reconcile bit identifiers, cumulative counters, tool assembly and failure evidence with original field and vendor records before selecting future equipment or diagnosing a mechanism.','پیش از انتخاب تجهیزات آینده یا تعیین سازوکار خرابی، شناسه مته، شمارنده تجمعی، آرایش ابزار و شواهد خرابی با سوابق میدانی و سازنده تطبیق شوند.'),
   ('Reconcile costs with signed contracts, approved scope and invoices; do not convert source status-code hours into financial liability without a contractual assessment.','هزینه‌ها با قرارداد امضاشده، دامنه مصوب و فاکتورها تطبیق شوند؛ ساعت کدهای منبع بدون ارزیابی قراردادی به مسئولیت مالی تبدیل نشود.')]:
    p1=insert_before(end._p,doc,fr if fa else en,fa);p1.style='List Bullet'
  insert_before(end._p,doc,'فهرست بررسی ریسک؛ مسئولیت و شدت تعیین نشده‌اند' if fa else 'Risk-review checklist; ownership and severity not assigned',fa,'Heading 1')
  t=doc.add_table(rows=1,cols=3);t.style='Table Grid';end._p.addprevious(t._tbl)
  h=['موضوع بررسی','شواهد در دسترس','اقدام مورد نیاز'] if fa else ['Review topic','Available evidence','Required follow-up']
  rows=[h]
  rows += ([['یکپارچگی چاه و فشار','هدف و وضعیت ذکرشده در برنامه؛ نه تأیید وضعیت فعلی','تطبیق آزمون‌ها، سدهای ایمنی و حدود مصوب جاری'],['حوادث و زمان ثبت‌شده','کدهای منبع با پوشش جزئی در جدول ۷-۱','بازبینی متن رویداد، ساعت، علت و قرارداد بدون انتساب فرضی'],['تجهیزات و عملیات','نمونه سلول‌های جدول ۹-۱ و مشاهدات ابزار جدول ۴-۲','تطبیق واحد، کالیبراسیون، شناسه ابزار و سابقه سازنده'],['برنامه و هزینه','صفحه برنامه و شواهد ساختار شکست هزینه','تطبیق دامنه، فاکتورها، مبنای زمانی و نسخه مصوب']] if fa else [['Well integrity and pressure','Programme objective and stated background, not proof of current condition','Reconcile tests, barriers and current approved limits'],['Incidents and recorded time','Partial source-code coverage in table 7-1','Review event narrative, duration, cause and contract without presumed attribution'],['Equipment and operations','Example cells in table 9-1 and tool observations in table 4-2','Verify units, calibration, tool identity and vendor records'],['Schedule and cost','Programme page and cost-breakdown evidence','Reconcile scope, invoices, time basis and approved revision']])
  set_table(t,rows,[4.0,6.4,6.4],fa)

def main():
 for well,w in DATA.items():
  p=PROGS[well];tr=TRACKS[w['source']];image=plot(w,tr,p)
  for lang in ('EN','FA'):
   fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
   for para in d.paragraphs:
    s=para.text
    if s.startswith(('The planned envelope of this operation','چارچوب برنامه‌ای این عملیات')):
     text(para,(f"برنامه {p['file']} در صفحه فیزیکی {p['physical_page']} چنین هدفی را ثبت کرده است: {p['objective_fa']} مدت ثبت‌شده برحسب روز: {p['duration_days_fa']}. اعداد احتیاطی مطابق برچسب منبع نقل شده‌اند و بدون تطبیق نباید جمع یا با اجرای واقعی مقایسه شوند." if fa else f"The programme {p['file']}, physical page {p['physical_page']}, states the following objective: {p['objective_en']} The stated duration in days is {p['duration_days_en']}. Contingency figures retain the source labels and should not be added or compared with execution without reconciliation."))
    elif s.startswith(('The execution consumed','اجرا ')) and ('rig days against' in s or 'روز دکل در برابر' in s):
     text(para,'مقایسه مدت اجرا و برنامه تا تطبیق نسخه، دامنه، گزارش روزانه و مرز زمانی معتبر نیست؛ نسبت و انتساب علت قبلی کنار گذاشته شده است.' if fa else 'A planned-versus-actual duration comparison is not established until revision, scope, daily reports and time boundaries are reconciled; the previous ratio and causal attribution are withdrawn.')
    elif s.startswith(('Modules reported as absent were genuinely unavailable','ماژول‌هایی که «موجود نیست» علامت خورده‌اند')):
     text(para,'وضعیت ماژول‌های متن قدیمی نیازمند بازبینی است؛ ناموجودبودن در یک استخراج، نبود آن در کل آرشیو را ثابت نمی‌کند.' if fa else 'The legacy module inventory requires review; absence from an extraction does not establish absence from the entire archive.')
    elif 'All datasets were parsed cell by cell' in s:
     text(para,f"This interim report references {w['source']} and {p['file']}. The archive was indexed, but semantic extraction and verification remain incomplete; see the evidence notes and source inventory.")
    elif 'تمام داده‌ها سلول‌به‌سلول بازخوانی شده‌اند' in s:
     text(para,f"این گزارش مقدماتی به {w['source']} و {p['file']} ارجاع می‌دهد. منابع نمایه شده‌اند، اما استخراج معنایی و تأیید آنها کامل نیست؛ به یادداشت‌های شواهد و فهرست منابع مراجعه شود.")
   for t in d.tables:
    if not t.rows:continue
    head=t.cell(0,0).text
    if head in ('KPI','شاخص') and len(t.columns)==2:
     for r in t.rows[1:]:
      label=r.cells[0].text
      if label in ('Planned vs actual duration','مدت برنامه‌ای در برابر واقعی'):
       celltext(r.cells[1],(f"برنامه: {p['duration_days_fa']} روز (صفحه {p['physical_page']})؛ مقایسه اجرا تطبیق‌نشده" if fa else f"Programme: {p['duration_days_en']} days (p. {p['physical_page']}); actual comparison pending"),fa)
      elif label in ('Productive vs NPT','زمان مولد در برابر غیرمولد'):
       celltext(r.cells[1],'تطبیق‌نشده؛ به جدول ۷-۱ مراجعه شود' if fa else 'Not reconciled; see table 7-1',fa)
    if head in ('Reconciliation metric','شاخص تطبیق') and len(t.columns)==4:
     for row in t.rows[1:]:
      label=row.cells[0].text
      if label.startswith(('Duration','مدت')):
       celltext(row.cells[1],p['duration_days_fa'] if fa else p['duration_days_en'],fa)
       for c in row.cells[2:]:celltext(c,'تطبیق‌نشده' if fa else 'Not reconciled',fa)
      elif ('NPT' in label or 'غیرمولد' in label or label.startswith(('Final depth','عمق نهایی'))):
       for c in row.cells[1:]:celltext(c,'تطبیق‌نشده' if fa else 'Not reconciled',fa)
   i=next(i for i,x in enumerate(d.paragraphs) if x.text.startswith(('Figure W1','شکل W1')));before=d.paragraphs[i-1];bl=before._element.xpath('.//a:blip')[0];rid=bl.get(qn('r:embed'));d.part.related_parts[rid]._blob=image
   shape=next(s for s in d.inline_shapes if any(b.get(qn('r:embed'))==rid for b in s._inline.xpath('.//a:blip')));shape.height=int(shape.width*6/11)
   text(d.paragraphs[i],'شکل W1 ـ پوشش شواهد روزانه و مدت ثبت‌شده در برنامه؛ مقایسه اجرا تأیید نشده است' if fa else 'Figure W1 — Daily-record coverage and stated programme duration; actual comparison unverified')
   save_section(d,well,fa,w,tr,p)
   marker='خلاصه مستند برنامه و محدودیت مقایسه' if fa else 'Programme source summary and comparison limitations'
   if not any(q.text==marker for q in d.paragraphs):
    h=d.add_heading(marker,level=1);h.alignment=WD_ALIGN_PARAGRAPH.CENTER
    q=d.add_paragraph((f"منبع: {p['file']}، صفحه فیزیکی {p['physical_page']}. هدف برنامه: {p['objective_fa']} مدت مندرج در برنامه به روز: {p['duration_days_fa']}. عمق اندازه‌گیری‌شده درج‌شده در برنامه {p['programme_stated_md_m']} متر و سازند تولیدی/مخزنی ذکرشده {p['formation_fa']} است؛ این مقادیر به‌تنهایی عمق حاصل‌شده یا شرایط فعلی را ثابت نمی‌کنند. مرز زمانی، دامنه و تفسیر احتیاط نیازمند تطبیق‌اند." if fa else f"Source: {p['file']}, physical page {p['physical_page']}. Programme objective: {p['objective_en']} Stated programme duration, in days: {p['duration_days_en']}. The programme states measured depth {p['programme_stated_md_m']} m and producer/reservoir formation {p['formation_en']}; these do not by themselves establish achieved depth or current condition. Time boundaries, scope and contingency interpretation require reconciliation."));style(q,fa)
   for t in d.tables:
    if t.rows:
     for c in t.rows[0].cells:
      for p0 in c.paragraphs:p0.alignment=WD_ALIGN_PARAGRAPH.CENTER
   d.save(path)
  print(well,'programme page',p['physical_page'],flush=True)
 for typ in ('Master_Drilling_Engineering','Executive_Management'):
  for lang in ('EN','FA'):
   fa=lang=='FA';path=Path(f'Field_{typ}_Report_{lang}.docx');d=Document(path);marker='مقایسه شواهد برنامه‌های ۱۴گانه' if fa else 'Fourteen source programmes: reviewed well-information pages'
   if any(x.text==marker for x in d.paragraphs):continue
   h=d.add_heading(marker,level=1);h.alignment=WD_ALIGN_PARAGRAPH.CENTER
   q=d.add_paragraph('جدول زیر نسخه برنامه و صفحه فیزیکی، مدت با برچسب منبع و عمق مندرج در همان صفحه را نشان می‌دهد؛ برنامه‌ها شامل انحراف اول ۱۲۷ نیز هستند، که با انحراف دوم یکی نیست. این مقادیر اثبات اجرا یا عمق حاصل‌شده نیستند؛ ارقام احتیاطی بدون تطبیق دامنه و مبنای زمانی جمع نشده‌اند.' if fa else 'The table identifies programme revision and physical page, duration with source labels and the depth stated on that page. It includes 127 ST1 as distinct from ST2. These figures do not establish execution or achieved depth; contingency values are not added without scope and time-basis reconciliation.');style(q,fa)
   t=d.add_table(rows=1,cols=4);t.style='Table Grid';rows=[(['عملیات','مدت برنامه (روز)','عمق درج‌شده (متر)','فایل برنامه / صفحه فیزیکی'] if fa else ['Operation','Programme duration (days)','Stated MD (m)','Programme file / physical page'])]
   for well,p in PROGS.items():rows.append([well,p['duration_days_fa'] if fa else p['duration_days_en'],p['programme_stated_md_m'],p['file']+' / '+str(p['physical_page'])])
   set_table(t,rows,[2.0,4.6,2.5,7.8],fa);d.save(path)
if __name__=='__main__':main()
