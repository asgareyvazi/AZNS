#!/usr/bin/env python3
"""Add material programme conflicts and ST1 scope to field summaries."""
from pathlib import Path
from docx import Document
from replace_well_bit_evidence import style
for typ in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'Field_{typ}_Report_{lang}.docx');d=Document(path)
  for p in list(d.paragraphs):
   if p.text.startswith(('Programme source discrepancies:', 'اختلاف‌های منابع برنامه:','ST1 pre-operation casing reference:', 'مرجع جداری پیش از انحراف اول:')):p._p.getparent().remove(p._p)
  anchor=next(p for p in d.paragraphs if p.text.startswith(('Additional ST1 phase references:','ارجاعات تکمیلی فاز انحراف اول:')))
  note=('اختلاف‌های منابع برنامه: برنامه چاه ۴۰ در صفحه فیزیکی ۱۳ عمق جداری ۲۰ اینچ را ۱۵۲٫۵ متر و در صفحه ۱۴۹ مقدار ۱۵۰ متر ذکر می‌کند. عبارت بازه فروریختگی ۳۸۸۹ تا ۳۹۱۱ متر نیز عیناً در برنامه چاه ۱۲ صفحه ۲۸ و برنامه تعمیر اول ۱۲۷ صفحه ۵۸ تکرار شده است. تعلق بازه و ارقام به وضعیت واقعی بدون سوابق تأییدشده روشن نیست.' if fa else 'Programme source discrepancies: physical p. 13 of the well 40 programme gives 152.5 m for the 20-inch casing, whereas p. 149 gives 150 m. Identical wording about a 3,889–3,911 m collapse interval also occurs on p. 28 of the well 12 programme and p. 58 of the 127 WO1 programme. Neither the interval attribution nor actual casing depth is resolved without approved well-specific records.')
  p=d.add_paragraph(note);style(p,fa);anchor._p.addnext(p._p)
  st=('مرجع جداری پیش از انحراف اول: برنامه انحراف اول ۱۲۷ در صفحه فیزیکی ۵۸، جداری‌های ۲۰، ۱۳⅜ و ۹⅝ اینچ را به‌ترتیب تا ۱۴۶، ۱۴۴۱ و ۳۹۰۹ متر و آستری ۷ اینچ را در بازه ۳۷۷۷ تا ۴۲۲۶ متر درج می‌کند؛ این جدول وضعیت نهایی پس از انحراف را تأیید نمی‌کند.' if fa else 'ST1 pre-operation casing reference: physical p. 58 of the 127 ST1 programme lists 20-, 13⅜- and 9⅝-inch casing to 146, 1,441 and 3,909 m, respectively, and a 7-inch liner from 3,777 to 4,226 m; it does not verify the final post-sidetrack configuration.')
  q=d.add_paragraph(st);style(q,fa);p._p.addnext(q._p);d.save(path)
