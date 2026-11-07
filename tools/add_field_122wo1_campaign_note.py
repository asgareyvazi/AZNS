#!/usr/bin/env python3
"""Disseminate the selected 122 WO1 source findings without portfolio aggregation."""
from pathlib import Path
from docx import Document
from replace_well_bit_evidence import style
EN=['The 122 WO1 daily pack records rig spud on 14 January 2024 at 15:00 and rig release on 9 March 2024 at 18:00, 55 days and 3 hours apart. Programme p.7 prints 45 main-plan days plus 7 contingency-plan days. Neither difference establishes a contractual overrun without matched time boundaries, approved scope and service records. [AZNS122WO1Pack.xlsx, 1-1 Remark!K34 and 56-1 Remark!K30; AZNS-0122 Workover-1 Program Rev.07.pdf, p.7]', 'Report 29 gives conflicting fish-depth entries; report 34 states the fish tail was completely milled, not that the complete fish was independently recovered. Reports 41–42 refer to a deep liner-lap reference at 3,775.4 m and a new liner lap at 1,998 m; they are distinct references. Report 56 describes a coiled-tubing brine displacement, not confirmation of the separate rig-less acid wash on programme p.27. [AZNS122WO1Pack.xlsx, 29-1 Remark!K22:K29, 34-1 Remark!K24:K25, 41-1 Remark!K22, 42-1 Remark!K24, 56-1 Remark!K22:K27]']
FA=['بسته روزانه تعمیر اول چاه ۱۲۲ شروع عملیات دکل را در ۱۴ ژانویه ۲۰۲۴ ساعت ۱۵ و ترخیص را در ۹ مارس ۲۰۲۴ ساعت ۱۸ ثبت می‌کند؛ فاصله تقویمی ۵۵ روز و ۳ ساعت است. برنامه صفحه ۷، ۴۵ روز برنامه اصلی به‌علاوه ۷ روز برنامه احتیاطی را ذکر می‌کند. بدون تطبیق مرز زمان، دامنه مصوب و اسناد خدمات، این اختلاف اضافه‌زمان قراردادی را ثابت نمی‌کند. [AZNS122WO1Pack.xlsx، 1-1 Remark!K34 و 56-1 Remark!K30؛ AZNS-0122 Workover-1 Program Rev.07.pdf، صفحه ۷]', 'گزارش ۲۹ عمق‌های ناسازگار قطعه جامانده را ذکر می‌کند؛ گزارش ۳۴ آسیاب کامل انتهای قطعه را ادعا می‌کند، نه بازیابی مستقل کل قطعه. گزارش‌های ۴۱ و ۴۲ به مرجع هم‌پوشانی عمیق ۳۷۷۵٫۴ متر و هم‌پوشانی جدید ۱۹۹۸ متر اشاره می‌کنند؛ این دو مرجع متفاوت‌اند. گزارش ۵۶ جابه‌جایی آب‌نمک با لوله‌مغزی سیار را شرح می‌دهد، نه تأیید اسیدشویی بدون دکل صفحه ۲۷ برنامه. [AZNS122WO1Pack.xlsx، 29-1 Remark!K22:K29، 34-1 Remark!K24:K25، 41-1 Remark!K22، 42-1 Remark!K24، 56-1 Remark!K22:K27]']
for typ in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
  fa=lang=='FA';p=Path(f'Field_{typ}_Report_{lang}.docx');d=Document(p);title='شواهد منتخب تعمیر اول چاه ۱۲۲' if fa else 'Selected AZNS-122 workover 1 campaign evidence'
  if any(p.text==title for p in d.paragraphs):continue
  h=d.add_heading(title,level=2);h.alignment=1
  for v in FA if fa else EN:q=d.add_paragraph(v);style(q,fa)
  d.save(p)
