#!/usr/bin/env python3
"""Correct all sheet-prefix labels; never assert these are report numbers."""
from pathlib import Path
import json,re
from docx import Document
from docx.oxml.ns import qn
from replace_well_bit_evidence import DATA,style,text
from replace_daily_narrative import chart as w2chart
from replace_workover_w3 import chart as workoverchart,tracks as worktracks,files as workfiles
from apply_program_and_safety_revision import ACTIVITY
import replace_drilling_w3 as drilling

A={r['file']:r for r in json.loads(Path('evidence/sheet_report_number_audit.json').read_text())['workbooks']}
DIR=Path('evidence/well_figures');DIR.mkdir(exist_ok=True)

def embed(doc,caption,png):
 pars=doc.paragraphs;i=next(i for i,p in enumerate(pars) if p.text.startswith(caption));rid=pars[i-1]._p.xpath('.//a:blip')[0].get(qn('r:embed'));doc.part.related_parts[rid]._blob=png
 return pars[i]

def fix_labels(p):
 # Label only cited source observations, not free-form discussion.
 if 'Pack.xlsx' not in p.text:return
 for r in p.runs:
  old=r.text
  new=re.sub(r'\bDaily report (\d+)\s*—',r'Source worksheet \1 —',old)
  new=re.sub(r'\bReport (\d+)\s*—',r'Source worksheet \1 —',new)
  new=re.sub(r'^گزارش روزانه (\d+)\s*ـ',r'شیت منبع \1 ـ',new)
  new=re.sub(r'^گزارش (\d+)\s*ـ',r'شیت منبع \1 ـ',new)
  if new!=old:r.text=new

def main():
 for well,w in DATA.items():
  src=w['source'];audit=A[src];activity=ACTIVITY[src]
  w2=w2chart(well,w,activity);(DIR/f'{well}_W2.png').write_bytes(w2)
  w3=None;complete=None
  if well in workfiles:w3,complete=workoverchart(well,worktracks[src]);(DIR/f'{well}_W3.png').write_bytes(w3)
  elif well!='207':w3=drilling.plot(well,src,show_formations=False);(DIR/f'{well}_W3.png').write_bytes(w3)
  for lang in ('EN','FA'):
   fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
   embed(d,('Figure W2','شکل W2'),w2)
   if w3:cap=embed(d,('Figure W3','شکل W3'),w3)
   for p in d.paragraphs:
    fix_labels(p)
    if p.text.startswith('These three entries are ordered by report number;'):
     text(p,p.text.replace('ordered by report number;', 'ordered by worksheet prefix, not necessarily report number;'));style(p,fa)
    elif p.text.startswith('این سه مدخل به‌ترتیب شماره گزارش نمایش داده شده‌اند؛'):
     text(p,p.text.replace('به‌ترتیب شماره گزارش نمایش داده شده‌اند؛','به‌ترتیب شماره شیت، نه الزاماً شماره گزارش، نمایش داده شده‌اند؛'));style(p,fa)
    if p.text.startswith('The window record, leak entry and test relate to different report numbers'):
     text(p,p.text.replace('different report numbers','different source worksheets'));style(p,fa)
    elif p.text.startswith('ثبت پنجره، مدخل نشتی و آزمون به شماره گزارش‌ها'):
     text(p,p.text.replace('به شماره گزارش‌ها','به شیت‌های منبع'));style(p,fa)
    if well in workfiles and p.text.startswith(('Figure W3','شکل W3')):
     text(p,(f'شکل W3 ـ داده‌های روزانه تعمیری به تفکیک شماره شیت، چاه {well}' if fa else f'Figure W3 — Workover source observations by worksheet prefix, AZNS-{well}'));p.alignment=1
    if well in workfiles and p.text.startswith(('The workover log uses only','نمودار روزانه تعمیری فقط')):
     text(p,(f"این نمودار از {len(worktracks[src]['days'])} شیت داده و {complete} صفحه دارای مدت طبقه‌بندی‌شده نزدیک ۲۴ ساعت استفاده می‌کند. محور شماره شیت است، نه شماره گزارش یا زمان پیوسته؛ ترتیب آن برای برخی فایل‌ها زمانی نیست. تعداد فیلد عددی معادل کامل‌بودن داده نیست. کدهای غیرمولد و غیربرنامه‌ای جدا هستند؛ مقدار مفقود صفر نیست و عمق پایان گزارش لزوماً عمق ابزار یا رخداد نیست." if fa else f"This plot uses {len(worktracks[src]['days'])} recognized data sheets and {complete} pages with classified durations totaling approximately 24 hours. Its axis is worksheet prefix, not report number or continuous time; for some workbooks sheet order is not chronological. Numeric-field counts are not data completeness. NPT and unplanned codes remain distinct; missing is not zero and report-end depth is not necessarily tool or event depth."));style(p,fa)
    if well in ('40','127WO1','127ST2') and p.text.startswith(('Points are daily entries at report-end depth','نقاط از داده روزانه در عمق پایان گزارش')):
     x=p.text.replace('first track is report number, not a continuous time axis','first track is worksheet prefix, not necessarily report number or a continuous time axis').replace('ستون نخست شماره گزارش است، نه محور زمان پیوسته','ستون نخست شماره شیت است، نه الزاماً شماره گزارش یا محور زمان پیوسته').replace(' Triangles average three consecutive reports with ROP values.','').replace(' مثلث‌ها میانگین سه گزارش متوالی دارای ROP هستند.','');text(p,x);style(p,fa)
   title='کنترل شماره شیت و شماره گزارش منبع' if fa else 'Worksheet-prefix and report-header reconciliation'
   if not any(p.text==title for p in d.paragraphs):
    anchor=next(p for p in d.paragraphs if p.text in ('Selected daily-report observations','روایت منتخب از گزارش روزانه'))
    h=d.add_heading(title,level=2);h.alignment=1;anchor._p.addprevious(h._p)
    msg=(f"در {audit['status_pages']} صفحه وضعیت شناسایی‌شده، شماره مندرج در سرصفحه با پیشوند عددی نام شیت در {audit['mismatched']} صفحه متفاوت است و {audit['unreadable_header']} سرصفحه عددی قابل تشخیص ندارد. محورهای مبتنی بر پیشوند شیت، شماره رسمی گزارش یا گاه‌شمار نیستند. شماره، تاریخ و زمان هر مدخل باید از سرصفحه و اسناد امضاشده تطبیق شود. جزئیات سلول‌ها در evidence/sheet_report_number_audit.json آمده‌اند." if fa else f"Of {audit['status_pages']} recognized status pages, the reported header number differs from the worksheet-name prefix on {audit['mismatched']} pages; {audit['unreadable_header']} headers have no readable numeric value. Axes based on worksheet prefix are not official report numbers or chronologies. Reconcile each entry's number, date and time with the header and signed records. Cell-level details are in evidence/sheet_report_number_audit.json.")
    p=d.add_paragraph(msg);style(p,fa);anchor._p.addprevious(p._p)
    if well=='116':
     msg=('در شیت 101-1 Remark سلول‌های S5/N5 شماره ۱۰۷ و تاریخ ۲۱ سپتامبر ۲۰۲۴ را نشان می‌دهند؛ شیت 102-1 شماره ۱۱۷ و تاریخ ۱ اکتبر، شیت 103-1 شماره ۱۱۸ و تاریخ ۲ اکتبر و شیت‌های 104-1 تا 112-1 شماره‌های ۱۰۸ تا ۱۱۶ را نشان می‌دهند. پس ترتیب شیت‌ها در انتهای فایل ترتیب زمانی معتبر نیست. همچنین تاریخ شروع در 1-1 Remark!K22 برابر ۸ ژوئن ۲۰۲۴ است، در حالی که سرصفحه 103-1 Remark!N6 تاریخ ۷ ژوئن را نوشته است؛ مبنا باید با اسناد امضاشده تطبیق شود.' if fa else 'In 101-1 Remark, S5/N5 show report 107 dated 21 September 2024; sheets 102-1 and 103-1 show reports 117 and 118 dated 1 and 2 October, while sheets 104-1 to 112-1 carry reports 108–116. Thus late workbook sheet order is not chronological. The first-sheet spud note 1-1 Remark!K22 states 8 June 2024, whereas the 103-1 Remark!N6 header states 7 June; reconcile the basis with signed records.')
     p=d.add_paragraph(msg);style(p,fa);anchor._p.addprevious(p._p)
    if well=='122WO1':
     msg=('سلول 56-1 Remark!N5 به‌صورت تاریخ اکسل ۳ سپتامبر ۲۰۲۴ ذخیره شده، در حالی که یادداشت ترخیص K30 همان شیت، ۹ مارس ۲۰۲۴ را نوشته است. فاصله تقویمی درج‌شده در بازسازی از دو یادداشت صریح شروع و ترخیص محاسبه شده، نه از N5؛ تاریخ سرصفحه باید با نسخه امضاشده تطبیق شود.' if fa else '56-1 Remark!N5 is stored as the Excel date 3 September 2024, while the release note in K30 on the same sheet states 9 March 2024. The reconstructed calendar interval uses the explicit spud and release notes, not N5; reconcile the header date with the signed record.')
     p=d.add_paragraph(msg);style(p,fa);anchor._p.addprevious(p._p)
    if well=='122WO2':
     msg=('یادداشت ترخیص در 119-1 DDR Remark!S32 تاریخ ۲۲ فوریه ۲۰۲۵ را ثبت می‌کند، ولی 120-1!S29 و 121-1!S30 تاریخ ۲۳ فوریه را نوشته‌اند؛ زمان ۲۴:۰۰ و نسخه نهایی باید با گزارش امضاشده تطبیق شوند.' if fa else 'The release note in 119-1 DDR Remark!S32 states 22 February 2025, while 120-1!S29 and 121-1!S30 state 23 February; the 24:00 convention and final record require reconciliation against the signed report.')
     p=d.add_paragraph(msg);style(p,fa);anchor._p.addprevious(p._p)
   if well in ('85','127WO1') and not any('1-1 Remark!N5' in p.text for p in d.paragraphs):
    anchor=next(p for p in d.paragraphs if p.text in ('Selected daily-report observations','روایت منتخب از گزارش روزانه'))
    if well=='85':
     msg=('در 1-1 Remark!N5 تاریخ اکسل ۲ نوامبر ۲۰۲۶ ذخیره شده، اما K22 همان صفحه شروع را ۲۵ ژانویه ۲۰۲۴ ثبت کرده است؛ تاریخ سرصفحه باید با اصل امضاشده تطبیق شود.' if fa else 'In 1-1 Remark!N5 the stored Excel date is 2 November 2026, but K22 on the same page records spud on 25 January 2024; reconcile the header date with the signed record.')
    else:
     msg=('در 1-1 Remark!N5 تاریخ اکسل ۵ اوت ۲۰۲۴ ذخیره شده، اما K26 همان صفحه شروع مشروط را ۸ مه ۲۰۲۴ ثبت کرده است؛ تاریخ سرصفحه باید با اصل امضاشده تطبیق شود.' if fa else 'In 1-1 Remark!N5 the stored Excel date is 5 August 2024, but K26 on the same page records conditional spud on 8 May 2024; reconcile the header date with the signed record.')
    p=d.add_paragraph(msg);style(p,fa);anchor._p.addprevious(p._p)
   d.save(path)
  print(well,'mismatch',audit['mismatched'])
if __name__=='__main__':main()
