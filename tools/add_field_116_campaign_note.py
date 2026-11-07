#!/usr/bin/env python3
"""Add a bounded 116 source review to both field reports and both languages."""
from pathlib import Path
from docx import Document
from replace_well_bit_evidence import style
ROOT=Path(__file__).resolve().parents[1]
for kind in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
  fa=lang=='FA';path=ROOT/f'Field_{kind}_Report_{lang}.docx';d=Document(path);title='بازبینی منتخب عملیات چاه ۱۱۶' if fa else 'Selected AZNS-116 campaign review'
  if any(p.text==title for p in d.paragraphs):continue
  h=d.add_heading(title,level=2);h.alignment=1
  msg=('برنامه چاه ۱۱۶ در صفحه ۷، ۸۹ روز اصلی و ۴۵٫۲ روز احتیاطی را ذکر می‌کند؛ به دلیل اختلاف تاریخ شروع و مرز دامنه، انحراف مدت واقعی محاسبه نشده است. در گزارش‌های عملیاتی منتخب، شیت 40-1 Remark!K29:K30 بازیابی‌نشدن قطعه و باقی‌ماندن حدود ۱۲ سانتی‌متر تپر تپ شکسته را شرح می‌دهد؛ شیت 68-1 Remark!K26:K33 برداشت لاگ پیوند سیمان را ثبت می‌کند، نه تفسیر تأییدشده آن؛ شیت 91-1 Remark!K27:K30 مشبک‌کاری بازه ۳۹۷۹ تا ۳۹۷۵ متر را گزارش می‌کند، نه بهره‌دهی بازه. شیت 103-1 Remark!K24 ترخیص دکل را ۲ اکتبر ۲۰۲۴ ساعت ۰۶ ثبت می‌کند. این‌ها مشاهدات تاریخی و منتخبی از منبع‌اند، نه وضعیت کنونی موانع، مسئولیت قراردادی یا تکمیل تمام فعالیت‌ها. [AZNS-116 Workover Program Rev.04.pdf، صفحه ۷؛ AZNS116Pack.xlsx؛ evidence/116_campaign_review.json]' if fa else 'Programme p.7 states 89 main-plan days and 45.2 contingency days; an actual-duration variance is not calculated because the start date and scope boundaries conflict. Among selected operational reports, 40-1 Remark!K29:K30 describes no fish recovered and approximately 12 cm of broken taper tap left in the well; 68-1 Remark!K26:K33 records cement-bond log acquisition, not validated interpretation; 91-1 Remark!K27:K30 reports perforating 3,979–3,975 m, not interval productivity. 103-1 Remark!K24 records rig release on 2 October 2024 at 06:00. These are selected historical source accounts, not present barrier status, contractual liability or proof that every activity was completed. [AZNS-116 Workover Program Rev.04.pdf, p.7; AZNS116Pack.xlsx; evidence/116_campaign_review.json]')
  p=d.add_paragraph(msg);style(p,fa);d.save(path)
print('116 campaign note added to four field reports')
