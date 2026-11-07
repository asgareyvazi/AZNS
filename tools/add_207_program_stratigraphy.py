#!/usr/bin/env python3
"""Review 207 programme page 17 and compare to the labelled DDR snapshot."""
import json,sqlite3
from pathlib import Path
from docx import Document
from replace_well_bit_evidence import style,set_table,text

# Proof strings checked against physical PDF page text; row grouping was also
# inspected via pypdf layout. Values are programme estimates, not measured tops.
ROWS=[
 ('Aghajari','آغاجاری','Surface',0,'Aghajari Surface Surface 1021',71),
 ('Gachsaran','گچساران','1021',1021,'Gachsaran -1007 1021 347',72),
 ('Asmari','آسماری','1368',1368,'Asmari -1353 1368 215',73),
 ('Pabdeh / Jahrum','پابده / جهرم','1583',1583,'Pabdeh or Pd./Ja. -1568 1583 102',74),
 ('Jahrum','جهرم','1685',1685,'Jahrum -1671 1685 309',75),
 ('Pabdeh','پابده','1994',1994,'Pabdeh -1980 1994 184',76),
 ('Upper Gurpi','گورپی بالایی','2178',2178,'U. Gurpi -2164 2178 165',77),
 ('Tarbur','تاربور','2343',2343,'Tarbur -2328 2343 75',78),
 ('Lower Gurpi','گورپی پایینی','2418',2418,'L. Gurpi -2404 2418 142',79),
 ('Ilam','ایلام','2560',2560,'Ilam -2546 2560 69',80),
 ('Laffan','لافان','2629',2629,'Laffan -2614 2629 6',81),
 ('Sarvak','سروک','2635',2635,'Sarvak -2621 2635 635',82),
 ('Kazhdumi','کژدمی','3270',3270,'Kazhdumi -3256 3270 194',83),
 ('Dariyan','داریان','3464',3464,'Dariyan -3450 3464 158',84),
 ('Gadvan','گدوان','3622',3622,'Gadvan -3607 3622 293',85),
 ('Fahliyan','فهلیان','3915',3915,'Fahliyan -3900 3915 40',86),
]

def main():
 db=sqlite3.connect('file:/tmp/azns-evidence.sqlite?mode=ro',uri=True);filename='AZNS-207 Program.pdf'
 pdf=' '.join(db.execute('select text from pages join files on files.id=pages.file_id where name=? and page=17',(filename,)).fetchone()[0].split())
 snap=next(x for x in json.loads(Path('evidence/formation_table_snapshots.json').read_text()) if x['file']=='AZNS207Pack.xlsx');assert snap['sheet']=='155-2 DDR Data'
 byrow={int(r['cells'][0]['cell'][1:]):r for r in snap['tables'][0]['rows']};items=[]
 for name,fa,label,md,proof,row in ROWS:
  assert proof in pdf,proof;r=byrow[row];raw=next(c for c in r['cells'] if c['cell']==f'P{row}');val=None if raw['value'].strip().lower()=='surface' else float(raw['value'])
  items.append({'name':name,'name_fa':fa,'programme_top_label':label,'programme_md_for_plot':md,'programme_source_text':proof,'ddr_name':r['formation'],'ddr_cell':raw['cell'],'ddr_value':raw['value'],'numeric_difference_ddr_minus_programme_m':None if val is None else val-md})
 output={'programme_file':filename,'physical_page':17,'programme_basis':'Prognosed mMD; RTE 14.45 m, see physical p.19. Values subject to revision during work.','programme_limit_md_m':4180,'pack_file':'AZNS207Pack.xlsx','pack_sheet':'155-2 DDR Data','ddr_basis':'Field labelled m MD; actual versus prognosed status not independently established.','rows':items,'limitations':'Arithmetic differences do not establish geological deviation, actual stratigraphic tops, event attribution or datum consistency. Pabdeh interval 1994–2178 m has 184 m in the table and a 165 m thickness stated in page 14 text.'}
 Path('evidence/program_207_stratigraphy_reviewed.json').write_text(json.dumps(output,indent=2,ensure_ascii=False)+'\n')
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'207_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
  title='رأس‌های پیش‌بینی‌شده برنامه و جدول روزانه چاه ۲۰۷' if fa else 'Programme-prognosed tops and 207 daily-table comparison'
  if any(p.text==title for p in d.paragraphs):continue
  anchor=next(p for p in d.paragraphs if p.text in ('Programme source summary and comparison limitations','خلاصه مستند برنامه و محدودیت مقایسه'))
  h=d.add_heading(title,level=2);h.alignment=1;anchor._p.addprevious(h._p)
  intro=('صفحه فیزیکی ۱۷ برنامه AZNS-207 Program.pdf رأس‌های پیش‌بینی‌شده را با مرجع ارتفاع میز دوار ۱۴٫۴۵ متر و ستون عمق اندازه‌گیری‌شده درج می‌کند؛ صفحه ۱۹ تصریح می‌کند که عمق‌ها ممکن است در حین عملیات اصلاح شوند. جدول زیر این مقادیر را با AZNS207Pack.xlsx، شیت 155-2 DDR Data، ستون P مقایسه می‌کند. جدول روزانه به‌تنهایی واقعی یا پیش‌بینی بودن رأس‌ها را اثبات نمی‌کند.' if fa else 'Physical page 17 of AZNS-207 Program.pdf gives prognosed tops in measured depth, with RTE 14.45 m; page 19 says formation depths are subject to adjustment during work. The table compares these values with the m MD column P of AZNS207Pack.xlsx, sheet 155-2 DDR Data. The DDR table alone does not establish whether the entries are actual or prognosed.')
  p=d.add_paragraph(intro);style(p,fa);anchor._p.addprevious(p._p)
  tab=d.add_table(rows=1,cols=4);tab.style='Table Grid';anchor._p.addprevious(tab._tbl)
  rows=[['سازند / عضو','برنامه، متر','ثبت روزانه و سلول','اختلاف حسابی، متر'] if fa else ['Formation / member','Programme, m MD','DDR value and cell','Arithmetic difference, m']]
  for x in items:
   delta=x['numeric_difference_ddr_minus_programme_m'];rows.append([x['name_fa'] if fa else x['name'],('سطح' if fa else 'Surface') if x['programme_top_label']=='Surface' else x['programme_top_label'],(('سطح' if fa else 'Surface') if x['ddr_value']=='Surface' else x['ddr_value'])+f" [{x['ddr_cell']}]",'—' if delta is None else f'{delta:+g}'])
  set_table(tab,rows,[4.3,3,4.2,5],fa)
  p=d.add_paragraph('جدول پیوست ـ رأس‌های پیش‌بینی‌شده برنامه در برابر مقادیر جدول روزانه؛ اختلاف به معنی انحراف واقعی نیست' if fa else 'Appendix table — Programme prognosed tops versus DDR entries; difference is not proven geological deviation');style(p,fa,True);anchor._p.addprevious(p._p)
  notes=[
   ('For example, the programme Fahliyan top is 3,915 m and DDR P86 is 3,968 m: the 53 m difference is arithmetic only. Confirm depth datum, table identity, actual versus forecast basis, approved geological interpretation and event depth before formation attribution.', 'برای نمونه رأس فهلیان در برنامه ۳۹۱۵ متر و مقدار P86 در جدول روزانه ۳۹۶۸ متر است؛ اختلاف ۵۳ متر صرفاً حسابی است. پیش از انتساب به سازند، مبنای عمق، هویت جدول، واقعی یا پیش‌بینی بودن، تفسیر زمین‌شناسی مصوب و عمق رویداد تأیید شوند.'),
   ('The programme describes Pabdeh as argillaceous limestone and marl, with chert and cherty limestone at its base; page 14 flags the latter for PDC-bit parameter review. This is a predicted geological description and not proof that a particular bit was damaged or an event occurred there. [physical p.14]', 'برنامه پابده را شامل آهک رسی و مارن و در بخش پایینی دارای چرت و آهک چرت‌دار توصیف می‌کند و در صفحه ۱۴ توجه به پارامترهای مته پلی‌کریستال را مطرح می‌سازد. این توصیف زمین‌شناسی پیش‌بینی‌شده است و خرابی مته یا وقوع رویدادی خاص در آن بازه را ثابت نمی‌کند. [صفحه فیزیکی ۱۴]'),
   ('The programme describes the Gurpi section as argillaceous limestone, Laffan as a thin shale, Sarvak as carbonate and Kazhdumi as shale with interbeds; page 15 supplies the lithological descriptions. Their different properties warrant review of actual samples, logs, tool response and pressure records, not an invented operating limit. [physical p.15]', 'برنامه گورپی را آهک رسی، لافان را لایه نازک شیل، سروک را واحد کربناته و کژدمی را شیل همراه میان‌لایه‌ها توصیف می‌کند. تفاوت این ویژگی‌ها بررسی نمونه، لاگ، پاسخ ابزار و فشار واقعی را ایجاب می‌کند، نه تعیین حد عملیاتی ساختگی. [صفحه فیزیکی ۱۵]'),
   ('The page 14 description calls Pabdeh 1,994–2,178 m and reports 165 m thickness, while page 17 lists 184 m, matching the arithmetic endpoint difference. The contradiction remains a source-review item; do not silently alter either value. [physical pp.14, 17]', 'شرح صفحه ۱۴ بازه پابده را ۱۹۹۴ تا ۲۱۷۸ متر و ضخامت را ۱۶۵ متر می‌گوید؛ جدول صفحه ۱۷ ضخامت ۱۸۴ متر را درج می‌کند که با اختلاف حسابی ابتدا و انتها برابر است. این تناقض باید بررسی شود و هیچ مقدار منبع بی‌صدا اصلاح نشده است. [صفحات فیزیکی ۱۴ و ۱۷]')]
  for en,fr in notes:p=d.add_paragraph(fr if fa else en);style(p,fa);anchor._p.addprevious(p._p)
  d.save(path)
 print('207 programme tops reviewed:',len(items))
if __name__=='__main__':main()
