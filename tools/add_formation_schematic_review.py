#!/usr/bin/env python3
"""Transcribe visually reviewed programme schematics; not verified as-built tops."""
from pathlib import Path
from hashlib import sha256
import json
from pypdf import PdfReader
from docx import Document
from replace_well_bit_evidence import style,set_table

N={'Asmari':'آسماری','Pabdeh':'پابده','Gurpi':'گورپی','Ilam':'ایلام','Sarvak':'سروک','Dariyan':'داریان','Gadvan':'گدوان','Gachsaran':'گچساران','Pabdeh/Jahrum':'پابده / جهرم','Pabdeh or Pd./Ja.':'پابده یا پابده / جهرم','Pabdeh Area-3':'پابده ناحیه ۳','Jahrum':'جهرم','U. Gurpi':'گورپی بالایی','Tarbur':'تاربور','L. Gurpi':'گورپی پایینی','Laffan':'لافان','Kazhdumi':'کژدمی','Fahliyan':'فهلیان','Gadvan-1':'گدوان ۱','Gadvan-5a':'گدوان ۵ الف','Gadvan-5c':'گدوان ۵ پ','Gadvan-8':'گدوان ۸','Sarvak-1':'سروک ۱'}
S={
'12':('AZNS-012 Workover Program Rev02.pdf',20,'embedded-image','before workover', [('Asmari','1267'),('Pabdeh','1460'),('Gurpi','1982'),('Ilam','2520'),('Sarvak','2592'),('Dariyan','3439'),('Gadvan','3597')]),
'75':('AZNS-075 Workover-1 Program Rev.05.pdf',21,'visual-pdf-text','before workover 1',[('Gachsaran','897.46'),('Asmari','1197.47'),('Pabdeh or Pd./Ja.','1382.48'),('Jahrum','1506.49'),('Pabdeh','1778.54'),('U. Gurpi','2023.55'),('Tarbur','2234.57'),('L. Gurpi','2334.58'),('Ilam','2472.58'),('Laffan','2547.59'),('Sarvak','2554.59'),('Kazhdumi','3196.65'),('Dariyan','3401.70'),('Gadvan-1','3557.76'),('Gadvan-5a','3660.82'),('Gadvan-5c','3670.82'),('Gadvan-8','3726.86')]),
'85':('AZNS-85 Workover Program Rev.02.pdf',19,'visual-pdf-text','before workover',[('Gachsaran','973'),('Asmari','1279/9'),('Pabdeh/Jahrum','1473/1'),('Jahrum','1567/6'),('Pabdeh Area-3','1864/9'),('Gurpi','2123'),('Ilam','2538'),('Sarvak','2619/4'),('Kazhdumi','3260'),('Dariyan','3463/7'),('Gadvan','3626/2'),('Fahliyan','3969/9')]),
'166':('AZNS-166 Workover-1 Program Rev.05.pdf',22,'visual-pdf-text','before workover 1',[('Gachsaran','1262'),('Asmari','1773'),('Pabdeh','2119'),('Gurpi','2353'),('Ilam','2616'),('Sarvak','2724'),('Kazhdumi','3847'),('Dariyan','4183')]),
'127WO1':('AZNS-127 Workover1 Program Rev.01.pdf',51,'visual-pdf-text','before workover',[('Gachsaran','1105'),('Pabdeh','1740.68'),('U. Gurpi','2182.49'),('Tarbur','2310.38'),('L. Gurpi','2390.33'),('Laffan','2618.34'),('Sarvak','2624.7'),('Kazhdumi','3273.82'),('Dariyan','3478.62'),('Gadvan','3650.57'),('Fahliyan','3937.73')]),
'127WO2':('AZNS-127 Workover #2 Program Rev.03.pdf',38,'visual-pdf-text','before workover 2',[('Gachsaran','1105'),('Asmari','1448.84'),('Pabdeh','1740.68'),('U. Gurpi','2182.49'),('Tarbur','2310.38'),('L. Gurpi','2390.33'),('Ilam','2516.7'),('Laffan','2618.34'),('Sarvak','2624.7'),('Kazhdumi','3273.82'),('Dariyan','3479.14'),('Gadvan','3667'),('Fahliyan','3938.87')]),
'127ST2':('AZNS -127 Side Track-2 Program Rev.01.pdf',25,'embedded-image','original well; MD(log)',[('Gachsaran','1084'),('Asmari','1448.84'),('Pabdeh','1740.68'),('U. Gurpi','2182.49'),('Tarbur','2310.38'),('L. Gurpi','2390.33'),('Ilam','2516.7'),('Laffan','2618.34'),('Sarvak-1','2624.7'),('Kazhdumi','3273.82'),('Dariyan','3478.62'),('Gadvan-1','3650.57'),('Fahliyan','3937.73')])}

NOTE_EN='The ST2 programme p.25 reproduces original-well tops with separate forecast, sample and log columns; the same source image occurs in the ST1 programme p.62. For Gachsaran these columns are 1,107, 1,105 and 1,084 m, respectively. The workover schematics show 1,105 m; this must not be silently identified as the log value. WO1 p.51 and WO2 p.38 also show different Dariyan (3,478.62 / 3,479.14 m), Gadvan (3,650.57 / 3,667 m) and Fahliyan (3,937.73 / 3,938.87 m) entries. These are source differences, not automatically geological movement or a validated sidetrack correlation.'
NOTE_FA='برنامه انحراف دوم در صفحه ۲۵ رأس‌های چاه اصلی را با ستون‌های جداگانه پیش‌بینی، نمونه و لاگ بازتولید می‌کند؛ همان تصویر در صفحه ۶۲ برنامه انحراف اول نیز آمده است. برای گچساران این سه ستون به‌ترتیب ۱۱۰۷، ۱۱۰۵ و ۱۰۸۴ مترند. شماتیک‌های تعمیر عدد ۱۱۰۵ متر را نشان می‌دهند؛ نباید آن را بدون توضیح مقدار لاگ دانست. همچنین شماتیک تعمیر اول صفحه ۵۱ و تعمیر دوم صفحه ۳۸ برای داریان (۳۴۷۸٫۶۲ / ۳۴۷۹٫۱۴ متر)، گدوان (۳۶۵۰٫۵۷ / ۳۶۶۷ متر) و فهلیان (۳۹۳۷٫۷۳ / ۳۹۳۸٫۸۷ متر) اعداد متفاوت دارند. این تفاوت منابع است، نه لزوماً جابه‌جایی زمین‌شناسی یا تطبیق تأییدشده شاخه انحرافی.'

def main():
 out=[]
 for well,(file,page,method,basis,items) in S.items():
  pdf=PdfReader(file);pg=pdf.pages[page-1];txt=pg.extract_text();image=max(pg.images,key=lambda i:len(i.data)) if method=='embedded-image' else None
  if method=='visual-pdf-text':
   for name,value in items:assert value in txt,(well,name,value)
  record={'well':well,'file':file,'physical_page':page,'method':method,'scope':basis,'depth_basis':'mMD as labelled; independently unverified', 'tops':[{'formation':n,'formation_fa':N[n],'source_depth_as_printed':v} for n,v in items]}
  if image:record['embedded_image_sha256']=sha256(image.data).hexdigest()
  out.append(record)
  for lang in ('EN','FA'):
   fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path);title='رأس‌های منتخب از شماتیک یا تصویر برنامه' if fa else 'Selected tops from the programme schematic or image'
   if any(p.text==title for p in d.paragraphs):continue
   anchor=next(p for p in d.paragraphs if p.text in ('Programme source summary and comparison limitations','خلاصه مستند برنامه و محدودیت مقایسه'))
   h=d.add_heading(title,level=2);h.alignment=1;anchor._p.addprevious(h._p)
   if well=='127ST2':en='These are the MD(log) values in the source-labelled actual tops of the original well, not proven tops of the ST2 branch.';pers='این‌ها مقادیر ستون عمق اندازه‌گیری‌شده لاگ برای رأس‌های موسوم به واقعی در چاه اصلی‌اند، نه رأس‌های اثبات‌شده شاخه انحراف دوم.'
   else:en='These values were transcribed from a schematic labelled before the workover, not independently confirmed current tops or as-built status.';pers='این مقادیر از شماتیک نام‌گذاری‌شده پیش از تعمیر نقل شده‌اند، نه رأس‌های فعلی یا وضعیت اجراشده‌ای که مستقل تأیید شده باشد.'
   p=d.add_paragraph((f'منبع: {file}، صفحه فیزیکی {page}. {pers} نام و عمق‌ها با بررسی بصری جفت شده‌اند؛ انتساب رخداد به سازند مستلزم تطبیق عمق رخداد، شاخه و مبنای عمق است.' if fa else f'Source: {file}, physical page {page}. {en} Names and depths were visually paired; event attribution requires reconciliation of the event depth, branch and datum.'));style(p,fa);anchor._p.addprevious(p._p)
   t=d.add_table(rows=1,cols=2);t.style='Table Grid';anchor._p.addprevious(t._tbl);set_table(t,[(['سازند یا عضو','عمق منبع، متر اندازه‌گیری‌شده'] if fa else ['Formation / member','Source depth, m MD'])]+[[N[n] if fa else n,v] for n,v in items],[8.0,8.4],fa)
   p=d.add_paragraph('جدول ـ رأس‌های منتخب سازند مطابق شماتیک یا تصویر منبع؛ مبنای چاه و مرحله در متن' if fa else 'Table — Selected formation tops as printed in the source schematic or image; see text for well and stage');style(p,fa,True);anchor._p.addprevious(p._p)
   if well=='85':
    p=d.add_paragraph('علامت ممیز در برخی اعداد به شکل «/» در منبع حفظ شده و به عدد اعشاری تبدیل نشده است؛ تأیید نسخه اصلی ضروری است.' if fa else 'Several source numbers contain a slash; it is retained as printed, not converted to a decimal without confirmation.');style(p,fa);anchor._p.addprevious(p._p)
   if well in ('127ST2','127WO1','127WO2'):
    p=d.add_paragraph(NOTE_FA if fa else NOTE_EN);style(p,fa);anchor._p.addprevious(p._p)
   d.save(path)
 # Link source images from ST1 and ST2 by hash; this proves identity of the reproduced image, not validity of the source numbers.
 im=max(PdfReader('AZNS -127 Side Track#01 Program Rev.03.pdf').pages[61].images,key=lambda x:len(x.data));hash1=sha256(im.data).hexdigest();assert hash1==next(x['embedded_image_sha256'] for x in out if x['well']=='127ST2')
 for fname in ('Field_Master_Drilling_Engineering_Report_EN.docx','Field_Master_Drilling_Engineering_Report_FA.docx','Field_Executive_Management_Report_EN.docx','Field_Executive_Management_Report_FA.docx'):
  d=Document(fname);fa='_FA' in fname;title='تفاوت مبنای رأس‌های چاه ۱۲۷' if fa else 'Well 127 formation-top source distinctions'
  if any(p.text==title for p in d.paragraphs):continue
  h=d.add_heading(title,level=2);h.alignment=1;p=d.add_paragraph(NOTE_FA if fa else NOTE_EN);style(p,fa);d.save(fname)
 Path('evidence/program_schematic_tops_reviewed.json').write_text(json.dumps({'notes':'Visually transcribed source-labelled historical schematic or original-well image, not independently validated actual sidetrack geology. Slash characters in 85 remain unchanged.','st1_same_image_sha256':hash1,'records':out},ensure_ascii=False,indent=2)+'\n')
 print('records',len(out),'tops',sum(len(x['tops']) for x in out))
if __name__=='__main__':main()
