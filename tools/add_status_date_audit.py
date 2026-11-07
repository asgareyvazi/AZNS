#!/usr/bin/env python3
"""Add source-scoped header-date audit notes and correct volatile-date wording."""
import json
from pathlib import Path
from docx import Document
from replace_well_bit_evidence import style,text
ROOT=Path(__file__).resolve().parents[1]
A={x['file']:x for x in json.loads((ROOT/'evidence/status_report_date_audit.json').read_text())['workbooks']}
W={'10':'AZNS10Pack.xlsx','12':'AZNS12Pack.xlsx','75':'AZNS75Pack.xlsx','85':'AZNS85Pack.xlsx','116':'AZNS116Pack.xlsx','122WO1':'AZNS122WO1Pack.xlsx','122WO2':'AZNS122WO2Pack.xlsx','127WO1':'AZNS127WO1Pack.xlsx','127WO2':'AZNS127WOPack.xlsx','127ST2':'AZNS127ST2Pack.xlsx','166':'AZNS166Pack.xlsx','207':'AZNS207Pack.xlsx','40':'AZNS40Pack.xlsx'}
for well,source in W.items():
 b=A[source]
 for lang in ('EN','FA'):
  fa=lang=='FA';path=ROOT/f'{well}_Drilling_Engineering_Report_{lang}.docx';d=Document(path)
  if well=='85':
   for p in d.paragraphs:
    if '1-1 Remark!N5' in p.text and ('2026' in p.text or '۲۰۲۶' in p.text):
     msg=('سلول 1-1 Remark!N5 فرمول متغیر =TODAY()-1 دارد و مقدار ذخیره‌شده آن ۲ نوامبر ۲۰۲۶ است؛ این مقدار تاریخ تاریخی قابل اتکا نیست و ممکن است با محاسبه مجدد تغییر کند. K22 همان صفحه شروع را ۲۵ ژانویه ۲۰۲۴ نوشته است. در ۱۳ شیت، N5 همین فرمول متغیر را دارد؛ تاریخ هر گزارش باید با نسخه امضاشده تطبیق شود.' if fa else '1-1 Remark!N5 contains the volatile formula =TODAY()-1, with cached value 2 November 2026; this is not a reliable historical date and can change on recalculation. K22 on the same page states spud on 25 January 2024. Thirteen status sheets contain this volatile N5 formula; reconcile each report date with the signed original.')
     text(p,msg);style(p,fa)
  title='کنترل تاریخ سرصفحه گزارش‌ها' if fa else 'Status-page header-date audit'
  if not any(p.text==title for p in d.paragraphs):
   anchor=next(p for p in d.paragraphs if p.text in ('Selected daily-report observations','روایت منتخب از گزارش روزانه'))
   h=d.add_heading(title,level=2);h.alignment=1;anchor._p.addprevious(h._p)
   msg=(f"از {b['recognized_status_pages']} صفحه وضعیت، تاریخ {b['parseable_date_headers']} سرصفحه با قواعد محدود بازیابی شد و {b['unparsed_date_headers']} سرصفحه به‌دلیل قالب مبهم، داده مفقود یا فرمول متغیر بدون تفسیر ماند. در مقایسه صفحه‌های مجاورِ دارای تاریخ قابل خواندن به ترتیب پیشوند شیت، {b['adjacent_backwards']} تاریخ عقب‌گرد و {b['adjacent_same_date']} تاریخ تکراری دیده شد؛ این شمارش به‌تنهایی خطا، روز عملیاتی یا ترتیب واقعی را ثابت نمی‌کند. فرمول متغیر در {b['volatile_date_formulas']} صفحه دیده شد. حتی تاریخ قابل خواندن ممکن است در منبع یا قالب اکسل اشتباه باشد. تاریخ‌ها و نسخه‌های امضاشده باید تطبیق شوند. [evidence/status_report_date_audit.json]" if fa else f"Of {b['recognized_status_pages']} recognized status pages, {b['parseable_date_headers']} header dates were readable under narrow rules; {b['unparsed_date_headers']} remain uninterpreted because of ambiguous format, missing data or volatile formula. Among adjacent recognized worksheets with readable dates in prefix order, {b['adjacent_backwards']} go backwards and {b['adjacent_same_date']} repeat a date. These counts alone do not prove an error, operating days or true sequence. Volatile date formulas occur on {b['volatile_date_formulas']} pages. Even readable Excel dates can be wrong in the source or locale; reconcile against signed records. [evidence/status_report_date_audit.json]")
   p=d.add_paragraph(msg);style(p,fa);anchor._p.addprevious(p._p)
  d.save(path)
for typ in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
  fa=lang=='FA';path=ROOT/f'Field_{typ}_Report_{lang}.docx';d=Document(path);title='کنترل تاریخ‌ها و فرمول متغیر' if fa else 'Date-header and volatile-formula audit'
  if not any(p.text==title for p in d.paragraphs):
   h=d.add_heading(title,level=2);h.alignment=1
   msg=('از ۱۳۰۱ صفحه وضعیت در ۱۳ بسته، تاریخ ۱۰۴۲ سرصفحه با قواعد محدود خوانده شد؛ ۲۵۹ مورد شامل قالب مبهم یا فرمول متغیر بدون تفسیر ماندند. در چاه ۸۵، ۱۳ سلول N5 فرمول =TODAY()-1 دارند؛ تاریخ ذخیره‌شده فرمول، تاریخ تاریخی قابل اتکا نیست و می‌تواند با محاسبه مجدد تغییر کند. در چاه ۱۱۶ و چاه ۴۰ هر کدام یک عقب‌گرد در تاریخ‌های قابل‌خواندن صفحات مجاور به ترتیب شیت دیده شد. این مقایسه همه موارد ناهماهنگی را پوشش نمی‌دهد و تاریخ خوانده‌شده الزاماً صحیح نیست. [evidence/status_report_date_audit.json؛ AZNS85Pack.xlsx، 1-1 Remark!N5]' if fa else 'Of 1,301 status pages in 13 packs, 1,042 header dates were readable under narrow rules; 259 including ambiguous formats or volatile formulas remain uninterpreted. In well 85, thirteen N5 cells contain =TODAY()-1; the cached result is not a dependable historical date and can change on recalculation. Wells 116 and 40 each show a backward step between adjacent prefix-ordered readable headers. These comparisons do not identify every discrepancy, and a readable date is not necessarily correct. [evidence/status_report_date_audit.json; AZNS85Pack.xlsx, 1-1 Remark!N5]')
   p=d.add_paragraph(msg);style(p,fa)
  d.save(path)
print('Updated 30 reports; corrected volatile-date wording for 85')
