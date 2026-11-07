#!/usr/bin/env python3
"""Add selected directly transcribed labelled formation tops from programme pages.
Avoid deriving mappings from PDF tables where extraction loses spatial alignment.
"""
from pathlib import Path
import json,sqlite3
from docx import Document
from replace_well_bit_evidence import style,set_table

NAMES={'Gachsaran':'گچساران','Asmari':'آسماری','Pabdeh':'پابده','Pabdeh/Jahrum':'پابده / جهرم','Jahrum':'جهرم','U. Gurpi':'گورپی بالایی','Tarbur':'تاربور','L. Gurpi':'گورپی پایینی','Gurpi':'گورپی','Ilam':'ایلام','Laffan':'لافان','Sarvak':'سروک','Kazhdumi':'کژدمی','Dariyan':'داریان','Gadvan':'گدوان','U. Gadvan':'گدوان بالایی','Khalij Mbr':'عضو خلیج','L. Gadvan':'گدوان پایینی','Fahliyan':'فهلیان','Sar-1':'واحد سروک ۱','Kaz-Top':'رأس کژدمی','Gad-1':'واحد گدوان ۱'}
S={
'10':('AZNS-010 Workover Program Rev.06.pdf',13,'actual_label','RTE 13.45 m',[
('Gachsaran','1151.5','Gachsaran -1138.0 1151.5 350.5'),('Asmari','1502.1','Asmari -1488.5 1502.1 282.7'),('Pabdeh','1784.8','Pabdeh -1771.2 1784.8 415.0'),('U. Gurpi','2199.8','U-Gurpi -2186.2 2199.8 121.8'),('Tarbur','2321.6','Tarbur -2308.0 2321.6 85.6'),('L. Gurpi','2407.2','L-Gurpi -2393.6 2407.2 125.1'),('Ilam','2532.3','Ilam -2518.7 2532.3 107.9'),('Laffan','2640.2','Laffan -2626.6 2640.2 6.1'),('Sarvak','2646.3','Sarvak -2632.7 2646.3 652.0'),('Kazhdumi','3298.3','Kazhdumi -3284.7 3298.3 198.8'),('Dariyan','3497.1','Dariyan -3483.5 3497.1 197.1'),('Gadvan','3694.2','Gadvan -3680.6 3694.2 293.6'),('Fahliyan','3987.8','Fahliyan -3974.2 3987.8 37.6')]),
'116':('AZNS-116 Workover Program Rev.04.pdf',40,'log_label','RTE 14.45 m',[
('Gachsaran','902','Gachsaran 893 892 902'),('Asmari','1201','Asmari 1204 1182 1201'),('Pabdeh/Jahrum','1377','Pabdeh/Jahrum 1361 1375 1377'),('Jahrum','1506','Jahrum 1493 1502 1506'),('Pabdeh','1714','Pabdeh 1724 1744 1714'),('U. Gurpi','2027','U. Gurpi 2012 2026 2027'),('Tarbur','2232','Tarbur 2219 2230 2232'),('L. Gurpi','2339','L. Gurpi 2326 2336 2339'),('Ilam','2482','Ilam 2462 2480 2482'),('Laffan','2558','Laffan 2536 2555 2558'),('Sarvak','2566','Sarvak 2543 2564 2566'),('Kazhdumi','3202','Kazhdumi 3185 3198 3202'),('Dariyan','3401','Dariyan 3390 3396 3401'),('U. Gadvan','3557','U. Gadvan 3548 3552 3557'),('Khalij Mbr','3628','Khalij Mbr - 3622 3628'),('L. Gadvan','3635','L. Gadvan - 3630 3635'),('Fahliyan','3900','Fahliyan 3906 3896 3900')]),
'122WO1':('AZNS-0122 Workover-1 Program Rev.07.pdf',19,'pre_wo1','not specified in figure',[]),
'122WO2':('AZNS-0122 Workover-2 Program Rev.00.pdf',20,'pre_wo1','not specified in figure',[]),
'40':('AZNS-040 RE-Entry Well Program Rev.03.pdf',39,'expected','RTE 13.45 m',[
('U. Gurpi','2197.9','U-Gurpi -2184.4 2197.9 124.1'),('Tarbur','2322.0','Tarbur -2308.5 2322.0 87.6'),('L. Gurpi','2409.6','L-Gurpi -2396.1 2409.6 125.4'),('Sar-1','2645.4','Sar-1 -2631.9 2645.4 19.6'),('Kaz-Top','3283.5','Kaz-Top -3270.0 3283.5 147.3'),('Gad-1','3682.7','Gad-1 -3669.3 3682.7 9.2'),('Fahliyan','3985.79','Fahliyan -3972.34 3985.79 37.0')])}
common=[('Gachsaran','1115'),('Asmari','1461'),('Pabdeh','1728.3'),('Gurpi','2196.5'),('Ilam','2535.6'),('Sarvak','2632.7'),('Kazhdumi','3277.7'),('Dariyan','3481'),('Gadvan','3647.5'),('Fahliyan','3937.6')]
for key in ('122WO1','122WO2'):S[key][4].extend((name,md,f'{name} @ {md}') for name,md in common)

def main():
 c=sqlite3.connect('file:/tmp/azns-evidence.sqlite?mode=ro',uri=True);out=[]
 for well,(file,page,basis,datum,entries) in S.items():
  raw=' '.join(c.execute('select text from pages join files on files.id=pages.file_id where name=? and page=?',(file,page)).fetchone()[0].split())
  # PDF line extraction changes spacing around @, but text itself must match.
  normalized=raw.replace('  ',' ')
  for name,md,proof in entries:
   assert proof in normalized,(well,proof)
  out.append({'well':well,'file':file,'physical_page':page,'source_basis':basis,'datum':datum,'tops':[{'name':name,'name_fa':NAMES[name],'md_m':md,'source_text':proof} for name,md,proof in entries]})
 Path('evidence/program_formation_tops_reviewed.json').write_text(json.dumps(out,ensure_ascii=False,indent=2)+'\n')
 for w in out:
  for lang in ('EN','FA'):
   fa=lang=='FA';path=Path(f"{w['well']}_Drilling_Engineering_Report_{lang}.docx");doc=Document(path)
   title='رأس‌های منتخب سازند از صفحه برنامه' if fa else 'Selected formation tops from programme page'
   if any(p.text==title for p in doc.paragraphs):continue
   anchor=next(p for p in doc.paragraphs if p.text in ('Programme source summary and comparison limitations','خلاصه مستند برنامه و محدودیت مقایسه'))
   h=doc.add_heading(title,level=2);h.alignment=1;anchor._p.addprevious(h._p)
   detail={'actual_label':('The programme labels the table actual formation tops; these are historical source values, not independent verification.','برنامه جدول را رأس‌های واقعی سازند نامیده است؛ این‌ها مقادیر تاریخی منبع‌اند، نه تأیید مستقل.'),'log_label':('The programme labels the table actual; values below are the MD (log) column, not the forecast or sample column.','برنامه جدول را واقعی نامیده است؛ اعداد زیر ستون عمق اندازه‌گیری‌شده لاگ‌اند، نه ستون پیش‌بینی یا نمونه.'),'pre_wo1':('These values are transcribed from the schematic explicitly labelled before workover #1, including in the WO2 programme; they do not establish the present configuration.','این مقادیر از شماتیکی که صراحتاً پیش از تعمیر شماره ۱ نام‌گذاری شده، حتی در برنامه تعمیر شماره ۲، نقل شده‌اند و وضعیت کنونی را ثابت نمی‌کنند.'),'expected':('The programme labels these expected tops; only clearly paired names and depths are included.','برنامه این رأس‌ها را پیش‌بینی‌شده نامیده است؛ فقط نام‌ها و عمق‌هایی که جفت آنها روشن است درج شده‌اند.')}
   en,fr=detail[w['source_basis']]
   p=doc.add_paragraph((f"منبع: {w['file']}، صفحه فیزیکی {w['physical_page']}. {fr} مرجع عمق: {w['datum']}. انتساب رخداد به سازند بدون تطبیق عمق رخداد و مرجع عمق معتبر نیست." if fa else f"Source: {w['file']}, physical page {w['physical_page']}. {en} Depth reference: {w['datum']}. Event-to-formation attribution requires a reconciled event depth and datum."));style(p,fa);anchor._p.addprevious(p._p)
   table=doc.add_table(rows=1,cols=2);table.style='Table Grid';anchor._p.addprevious(table._tbl)
   rows=[(['سازند یا عضو','رأس در منبع، متر اندازه‌گیری‌شده'] if fa else ['Formation / member','Source top, m MD'])]+[[top['name_fa'] if fa else top['name'],top['md_m']] for top in w['tops']]
   set_table(table,rows,[8.0,8.4],fa);doc.save(path)
  print(w['well'],len(w['tops']))
if __name__=='__main__':main()
