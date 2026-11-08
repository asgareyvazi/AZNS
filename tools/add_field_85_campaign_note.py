#!/usr/bin/env python3
"""Selected 85 campaign caveats for the field-level summaries."""
from pathlib import Path
from docx import Document
from replace_well_bit_evidence import style
for typ in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'Field_{typ}_Report_{lang}.docx');d=Document(path);title='بازبینی منتخب عملیات چاه ۸۵' if fa else 'Selected AZNS-85 campaign review'
  if any(p.text==title for p in d.paragraphs):continue
  h=d.add_heading(title,level=2);h.alignment=1
  msg=('گزارش‌های منتخب چاه ۸۵ خرابی قطعه پکر آزمون خشک و گم‌شدن یک drag block را ثبت می‌کنند؛ گزارش بعدی از بازیابی‌نشدن آن با آهنربا و شروع آسیاب‌کاری سخن می‌گوید. جمع‌های ۶۱، ۸۵ و ۹۲ ساعت در یادداشت‌های منبع به‌صورت ارقام تجمعی آمده‌اند و نباید با هم جمع یا به‌تنهایی مسئولیت قراردادی تلقی شوند. صفحه 39-1 Remark دو ران مشبک‌کاری را ثبت و بازه وسیع‌تری را تأییدشده توصیف می‌کند؛ نمی‌توان از آن مشبک‌شدن پیوسته کل بازه یا بهره‌دهی را نتیجه گرفت. آخرین صفحه شناسایی‌شده ترخیص برخی خدمات را ثبت می‌کند، نه زمان ترخیص دکل یا اجرای اسیدکاری بدون دکل برنامه‌ای. تاریخ ۱۳ سرصفحه فرمول متغیر دارد و برای گاه‌شمار باید با نسخه امضاشده تطبیق شود. [AZNS85Pack.xlsx، 18-1 Remark!K25:K30، 20-1 Remark!K30، 21-1 Remark!K23:K30، 23-1 Remark!K24:K44، 24-1 Remark!K25، 39-1 Remark!K24:K39، 53-1 Remark!K48:K50؛ evidence/85_campaign_review.json]' if fa else 'Selected AZNS-85 reports describe dry-test packer component failure and the loss of a drag block; later entries say a magnet did not recover it and milling began. The 61, 85 and 92-hour source notes are running figures and must not be added together or treated alone as contractual liability. 39-1 Remark records two perforating runs and describes a broader interval as approved; it does not prove continuous perforation across the whole span or productivity. The final recognized page records release of some services, not rig-release time or execution of planned rig-less acidizing. Thirteen header dates use volatile formulas and require signed-record reconciliation. [AZNS85Pack.xlsx, 18-1 Remark!K25:K30, 20-1 Remark!K30, 21-1 Remark!K23:K30, 23-1 Remark!K24:K44, 24-1 Remark!K25, 39-1 Remark!K24:K39, 53-1 Remark!K48:K50; evidence/85_campaign_review.json]')
  p=d.add_paragraph(msg);style(p,fa);d.save(path)
print('Updated four field reports')
