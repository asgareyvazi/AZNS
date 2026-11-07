#!/usr/bin/env python3
"""Correct provisional front matter and unsourced legacy well-summary metrics."""
from pathlib import Path
import json
from docx import Document
from replace_well_bit_evidence import DATA,set_table,text
from replace_cbs_financials import mapping

P={x['well']:x for x in json.loads(Path('evidence/program_summary_reviewed.json').read_text())};F={x['well']:x for x in json.loads(Path('evidence/fluid_observation_summary.json').read_text())};A={x['file']:x for x in json.loads(Path('evidence/activity_observations.json').read_text())}

def main():
 for well,w in DATA.items():
  p=P[well];fl=F[well];a=A[w['source']];full=sum(d['activity']['complete'] for d in a['days']);loss=fl['fields']['downhole_loss'];maxv=loss['maximum']
  for lang in ('EN','FA'):
   fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
   def table(*heads):return next(t for t in d.tables if t.rows and t.rows[0].cells[0].text in heads)
   t=table('Item','شرح')
   rows=([['شرح','مقدار'],['چاه / عملیات',well],['بسته گزارش روزانه',w['source']],['کتاب‌کار هزینه',mapping[well][0]],['برنامه چاه',p['file']],['وضعیت','پیش‌نویس مقدماتی؛ تطبیق و تأیید مهندسی لازم است']] if fa else [['Item','Value'],['Well / operation',well],['Daily-report pack',w['source']],['Cost workbook',mapping[well][0]],['Well programme',p['file']],['Status','Interim draft; source reconciliation and engineering approval required']])
   set_table(t,rows,[5.2,11.3],fa)
   t=table('Document control','کنترل سند')
   rows=([['کنترل سند','وضعیت'],['کد سند رسمی','احراز نشده'],['وضعیت','پیش‌نویس مقدماتی؛ صادر یا مصوب نشده'],['تاریخ صدور رسمی','احراز نشده'],['تهیه','پیش‌نویس تحلیلی ماشینی'],['بازبینی تخصصی','انجام و تأیید آن احراز نشده'],['تصویب و توزیع','احراز نشده']] if fa else [['Document control','Status'],['Official document code','Not established'],['Revision / status','Interim draft; not issued or approved'],['Official issue date','Not established'],['Preparation','Automated analytical draft'],['Specialist review','Completion and approval not established'],['Approval and distribution','Not established']])
   set_table(t,rows,[5.2,11.3],fa)
   t=table('KPI','شاخص')
   if maxv:
    maximum=f"{maxv['value']:g} [{maxv['sheet']}!{maxv['cell']}]"
   else:maximum='داده عددی شناسایی نشد' if fa else 'No numeric entry recognized'
   rows=([['شاخص','مقدار و محدوده'],['مدت درج‌شده در برنامه',p['duration_days_fa']+f" [صفحه {p['physical_page']}]"],['صفحات وضعیت شناسایی‌شده',str(a['remark_days'])],['صفحات وضعیت با جمع مدت نزدیک ۲۴ ساعت',str(full)],['صفحات داده شناسایی‌شده',str(fl['recognized_data_sheets'])],['تعداد مقادیر عددی هرزروی',str(loss['numeric_accepted_count'])],['بیشینه یک مقدار ثبت‌شده هرزروی، بشکه',maximum],['مدت واقعی، مجموع هرزروی و هزینه واقعی','تطبیق نشده و از این مشاهدات محاسبه نمی‌شوند']] if fa else [['KPI','Value and scope'],['Programme-stated duration',p['duration_days_en']+f" [physical p.{p['physical_page']}]"],['Recognized status pages',str(a['remark_days'])],['Status pages with approximately 24 h classified',str(full)],['Recognized data pages',str(fl['recognized_data_sheets'])],['Numeric downhole-loss entries',str(loss['numeric_accepted_count'])],['Maximum single source loss entry, bbl',maximum],['Actual duration, total loss and actual cost','Not reconciled; not inferred from these observations']])
   set_table(t,rows,[7.1,9.4],fa)
   t=table('Dataset module','ماژول داده')
   rows=([['منبع / استخراج','شناسایی و محدودیت'],['بسته روزانه',w['source']],['صفحات وضعیت',f"{a['remark_days']} صفحه؛ {full} صفحه دارای زمان‌های طبقه‌بندی‌شده نزدیک ۲۴ ساعت"],['صفحات داده',f"{fl['recognized_data_sheets']} شیت از {fl['candidate_pairs']} زوج نامزد"],['کتاب‌کار هزینه',mapping[well][0]+'؛ مقادیر ذخیره‌شده و دامنه نیازمند تطبیق'],['برنامه چاه',p['file']+'؛ صفحه اطلاعات '+str(p['physical_page'])],['محدودیت','شناسایی‌شدن یا نشدن یک بخش، وجود یا کامل‌بودن آن در کل آرشیو را ثابت نمی‌کند']] if fa else [['Source / extraction','Recognition and limitation'],['Daily pack',w['source']],['Status pages',f"{a['remark_days']} recognized; {full} with approximately 24 classified hours"],['Data pages',f"{fl['recognized_data_sheets']} of {fl['candidate_pairs']} candidate pairs"],['CBS workbook',mapping[well][0]+'; cached values and scope require reconciliation'],['Programme PDF',p['file']+f"; information page {p['physical_page']}"],['Limitation','Recognition or absence in an extraction does not establish complete archive coverage']])
   set_table(t,rows,[5.3,11.2],fa)
   # The old generic TOC names a complete day-by-day reconstruction that is
   # not delivered. Reflect actual top-level headings without page claims.
   toc=table('No.','ردیف');exclude={'Executive KPI Summary','Contents','خلاصه شاخص‌های کلیدی اجرا','فهرست مطالب'}
   headings=[q.text for q in d.paragraphs if q.style and q.style.name=='Heading 1' and q.text not in exclude]
   rows=[['ردیف','عنوان بخش'] if fa else ['No.','Section']]+[[str(i),h] for i,h in enumerate(headings,1)]
   set_table(toc,rows,[1.6,14.9],fa)
   # Static identity is legacy-compiled and is not certified.
   for q in d.paragraphs:
    if q.text.startswith('Prepared for the South Azadegan Field Development'):
     text(q,'Interim source review for South Azadegan; engineering approval pending')
    elif q.text.startswith('تهیه‌شده برای توسعه میدان آزادگان جنوبی'):
     text(q,'بازبینی مقدماتی منابع میدان آزادگان جنوبی؛ تأیید مهندسی در انتظار')
    elif q.text.startswith('Table 0-1  Document control and distribution'):
     text(q,'Table 0-1 — Document control and review status.')
    elif q.text.startswith('جدول ۰-۱  کنترل سند و توزیع'):
     text(q,'جدول ۰-۱ ـ کنترل سند و وضعیت بازبینی.')
   for q in d.paragraphs:
    if q.text.startswith('This interim report references'):
     text(q,q.text+' Static identity and rig assignments below are legacy-compiled and require verification against cited source cells and approved records.')
    elif q.text.startswith('این گزارش مقدماتی به'):
     text(q,q.text+' شناسنامه و تخصیص دکل در جدول بعد از گزارش قبلی گردآوری شده‌اند و باید با سلول‌های منبع و اسناد مصوب تطبیق شوند.')
   # Remove duplicate 207 table, retaining historical salinity note.
   if well=='207':
    old=next((t for t in d.tables if t.rows and t.rows[0].cells[0].text in ('Formation','سازند') and len(t.rows[0].cells)==4 and 'Pabdeh' in ' '.join(c.text for r in t.rows for c in r.cells) if not fa),None) if not fa else next((t for t in d.tables if t.rows and t.rows[0].cells[0].text=='سازند' and len(t.rows[0].cells)==4 and 'برنامه MD' in t.rows[0].cells[1].text),None)
    if old is not None:old._tbl.getparent().remove(old._tbl)
    for q in d.paragraphs:
     if q.text.startswith('Programme depths are transcribed from section 1.5.1'):
      text(q,'Programme pages 14–16 describe formation intervals; page 17 lists prognosed tops. The more complete programme-versus-DDR comparison follows, with limitations and source references.')
     elif q.text.startswith('ستون برنامه از بخش‌های ۱٫۵٫۱'):
      text(q,'صفحات ۱۴ تا ۱۶ برنامه بازه‌های سازندی را توصیف و صفحه ۱۷ رأس‌های پیش‌بینی‌شده را فهرست می‌کند. مقایسه کامل‌تر برنامه و جدول روزانه با ارجاعات و محدودیت‌ها در ادامه آمده است.')
   if well=='85':
    # Move the planned rig-less source table from before the cover into the
    # programme section; preserve the four elements and their order.
    plan=table('Source cell','سلول منبع');node=plan._tbl;intro=node.getprevious();head=intro.getprevious();note=node.getnext()
    if 'CBS AZNS-85.xlsx' in ''.join(intro.xpath('.//w:t/text()')):
     anchor=next(q for q in d.paragraphs if q.text in ('Selected programme phases and limited execution evidence','مراحل منتخب برنامه و شواهد اجرای محدود'))
     for element in (head,intro,node,note):anchor._p.addprevious(element)
    if fa:
     replacements={2:'راندن لوله مغزی سیار و اجرای اسیدشویی طبق برنامه و خروج رشته.',6:'راندن لوله مغزی سیار تا انتهای چاه، جابه‌جایی سیال با گازوئیل و خروج.',7:'راندن اسلیک‌لاین و نصب پلاگ و پرانگ در نشیمنگاه توقف و خروج.',9:'ایمن‌سازی چاه با نصب شیر فشار برگشتی در آویز لوله و بستن شیرهای درختک به‌جز شیر اصلی پایینی.'}
     for row,value in replacements.items():
      pcell=plan.rows[row].cells[1].paragraphs[0];text(pcell,value)
   d.save(path)
  print(well,full,loss['numeric_accepted_count'])
if __name__=='__main__':main()
