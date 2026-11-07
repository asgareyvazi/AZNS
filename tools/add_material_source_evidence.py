#!/usr/bin/env python3
"""Selected source-labelled mud materials in bilingual well reports."""
from pathlib import Path
import json,math
from docx import Document
from replace_well_bit_evidence import DATA,set_table,style,text

DAYS={'207':[2,72],'127ST2':[1,2],'40':[1,2],'127WO1':[2,3],'10':[1,3],'12':[6,9],'75':[3,4],'85':[2,7],'116':[2,13],'122WO1':[2,3],'122WO2':[54,55],'127WO2':[11,12],'166':[4,7]}
UNITS_FA={'25 kg':'۲۵ کیلوگرم','25 KG/SX':'۲۵ کیلوگرم / SX','1MT/BB':'۱ تن متریک / BB','1 MT/BB':'۱ تن متریک / BB','1 MT':'۱ تن متریک','55 lbs':'۵۵ پوند','MT':'تن متریک'}
FA={'Bentonite':'بنتونیت','Limestone':'سنگ آهک','Salt':'نمک','Caustic Soda':'سود سوزآور','Soda Ash':'کربنات سدیم','Soda Ash (25 kg)':'کربنات سدیم، بسته ۲۵ کیلوگرمی','Starch LV':'نشاسته کم‌گرانروی','Starch LV (Green)':'نشاسته کم‌گرانروی، سبز','Lime':'آهک','Xanthan Gum (XC Pol.)':'صمغ زانتان','Sodium Bicarbonate':'بی‌کربنات سدیم'}

def select(sheet):
 rows=[]
 for r in sheet['entries']:
  u=r['fields'].get('used')
  if not u or u['raw'] is None:continue
  try:v=float(u['raw'])
  except (ValueError,TypeError):continue
  if not math.isfinite(v) or v<=0:continue
  name=r['material']['raw'].strip()
  if name not in FA:continue
  if not r['unit'] or not r['unit']['raw']:continue
  rows.append(r)
  if len(rows)==3:break
 assert len(rows)>=2,(sheet['sheet'],len(rows))
 return rows

def main():
 sources={x['file']:x for x in json.loads(Path('evidence/mud_material_observations.json').read_text())};out=[]
 for well,w in DATA.items():
  src=sources[w['source']];samples=[]
  for day in DAYS[well]:
   s=next(s for s in src['sheets'] if s['day']==day);rows=select(s);samples.append({'day':day,'sheet':s['sheet'],'used_header':s['headers']['used']['cell'],'rows':rows})
  out.append({'well':well,'file':w['source'],'data_sheets':src['recognized_data_sheets'],'material_blocks':src['recognized_material_blocks'],'samples':samples})
  for lang in ('EN','FA'):
   fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');doc=Document(path)
   # Avoid double insertions on accidental re-execution.
   title='Selected mud-material source entries' if not fa else 'مقادیر منتخب مواد سیال در منبع'
   if any(p.text==title for p in doc.paragraphs):continue
   end=next(p for p in doc.paragraphs if p.text in ('Completion, Casing and Service-Module Findings','یافته‌های تکمیل، جداری و ماژول‌های خدماتی'))
   head=doc.add_heading(title,level=2);head.alignment=1;end._p.addprevious(head._p)
   intro=(f"جدول ۵-۲ چند ردیف مصرف ثبت‌شده را از دو شیت منتخب {w['source']} نشان می‌دهد. در {src['recognized_material_blocks']} شیت از {src['recognized_data_sheets']} شیت داده، سرستون‌های قابل تشخیص مواد یافت شدند؛ این تعداد نشان‌دهنده تکمیل یا صحت موجودی و مصرف نیست. نشانی سرستون «مصرف» و واحد هر ردیف در فایل شواهد ثبت شده است." if fa else f"Table 5-2 selects source-used entries from two sheets of {w['source']}. Recognizable materials headers were found in {src['recognized_material_blocks']} of {src['recognized_data_sheets']} data sheets; this count does not validate completeness or correctness of inventory or usage. Used-column header cells and row units are retained in the evidence file.")
   p=doc.add_paragraph(intro);style(p,fa);end._p.addprevious(p._p)
   table=doc.add_table(rows=1,cols=4);table.style='Table Grid';end._p.addprevious(table._tbl)
   rows=[['شیت منتخب','ماده و سلول','مصرف منبع و سلول','واحد منبع و سلول'] if fa else ['Selected sheet','Material and cell','Used value and cell','Source unit and cell']]
   for s in samples:
    for r in s['rows']:
     name=r['material']['raw'].strip();used=r['fields']['used'];unit=r['unit']
     # Raw unit text is preserved; the quantity is never converted or summed.
     rows.append([s['sheet'],f"{FA[name] if fa else name} [{r['material']['cell']}]",f"{used['raw']} [{used['cell']}]",f"{UNITS_FA[unit['raw']] if fa else unit['raw']} [{unit['cell']}]"])
   set_table(table,rows,[3.8,4.3,3.8,4.6],fa)
   p=doc.add_paragraph('جدول ۵-۲ ـ نمونه‌های منتخب ستون مصرف مواد سیال؛ واحد مطابق متن منبع' if fa else 'Table 5-2 — Selected mud-material used-column entries; units as printed');style(p,fa,True);end._p.addprevious(p._p)
   note=('ستون مصرف با ستون دریافت، دریافت تجمعی و موجودی تفاوت دارد. تکرار یک ماده یا مقدار در شیت‌ها لزوماً مصرف مستقل نیست؛ مقادیر خالی یا خطا صفر نیستند. اندازه بسته، معنی اختصار واحد، دوره ثبت، تطبیق مانده اول و آخر، دریافت و مصرف و سابقه اختلاط باید پیش از هر جمع یا تبدیل تأیید شوند. این نمونه‌ها به‌تنهایی مصرف درون‌چاهی، درمان هرزروی یا رابطه علّی با خواص گل را ثابت نمی‌کنند.' if fa else 'The used column is distinct from received, cumulative received and stock. Repeated material or values across sheets are not necessarily independent consumption; blanks and errors are not zero. Verify package size, unit abbreviation, reporting period, opening and closing balance, receipts, usage and mixing records before totals or conversion. These entries alone do not establish downhole use, treatment of losses or causality with fluid properties.')
   p=doc.add_paragraph(note);style(p,fa);end._p.addprevious(p._p)
   # Withdraw the remaining legacy POB calculation and unsupported claim of
   # comprehensive cell-level traceability rather than conceal with a warning.
   for p in doc.paragraphs:
    if p.text.startswith('Personnel on board averaged'):
     text(p,'Personnel-on-board averages and peaks in the legacy report are withdrawn pending reconciliation of daily role-level counts, totals, missing entries and personnel scope.')
    elif p.text.startswith('میانگین نفرات حاضر') or ('نفرات حاضر' in p.text and 'میانگین' in p.text and 'قرارداد' in p.text):
     text(p,'میانگین و بیشینه نفرات حاضرِ گزارش قبلی تا تطبیق تعداد روزانه به تفکیک نقش، جمع‌ها، مقادیر مفقود و دامنه نفرات کنار گذاشته شده‌اند.')
    elif p.text.startswith('Every quantitative statement in this report resolves'):
     text(p,'The table lists selected source anchors only. Other entries and interpretations still require independent cell/page checks, unit and scope review; the current report is not fully reconciled or reproducible without engineering interpretation.')
    elif p.text.startswith('هر گزاره کمّی این گزارش') or p.text.startswith('هر گزاره کمی این گزارش') or p.text.startswith('هر ادعای کمی این گزارش'):
     text(p,'جدول فقط ارجاعات منتخب منبع را نشان می‌دهد. سایر مقادیر و تفسیرها همچنان به کنترل مستقل سلول یا صفحه و بازبینی واحد و دامنه نیاز دارند؛ گزارش فعلی بدون تفسیر مهندسی به‌طور کامل تطبیق یا بازتولید نشده است.')
   doc.save(path)
 Path('evidence/mud_material_selected_reviewed.json').write_text(json.dumps(out,ensure_ascii=False,indent=2)+'\n')
 print('selected',len(out),'sheets',sum(len(x['samples']) for x in out))
if __name__=='__main__':main()
