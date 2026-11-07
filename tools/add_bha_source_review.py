#!/usr/bin/env python3
"""Insert selected, reviewed component observations without claiming BHA runs."""
from pathlib import Path
import json
from docx import Document
from replace_well_bit_evidence import DATA,set_table,style

# Day and four independently checked translations in the exact source row order.
SELECT={
 '207':(72,59,['مته پلی‌کریستال','رابط مته','دو شاخه طوق حفاری','پایدارکننده']),
 '127ST2':(48,65,['مته سه‌مخروطی','رابط مته','هفت دسته طوق حفاری','دو دسته لوله حفاری سنگین']),
 '40':(2,91,['مته چرخشی','رابط مته با شیر شناور','طوق حفاری','تبدیل اتصال']),
 '127WO1':(52,62,['آسیاب فلزی ضایعات','گیرنده ضایعات','گیرنده ضایعات','رابط دو سر نر']),
 '10':(2,32,['مته چرخشی','رابط مته','طوق حفاری','لوله حفاری سنگین']),
 '12':(37,33,['آسیاب فلزی ضایعات','رابط مته','۱۷ شاخه طوق حفاری','تبدیل اتصال']),
 '75':(80,42,['مته سه‌مخروطی','رابط مته','لوله حفاری سنگین','لوله حفاری']),
 '85':(5,28,['مته','رابط مته','طوق حفاری','لوله حفاری سنگین']),
 '116':(10,50,['آسیاب','گیرنده ضایعات','رابط دو سر نر','گیرنده ضایعات']),
 '122WO1':(12,28,['مته چرخشی شش اینچ','خراشنده هفت اینچ','رابط مته با شیر شناور','۹ شاخه طوق حفاری']),
 '122WO2':(54,32,['مته پلی‌کریستال','موتور درون‌چاهی','رابط شناور','محفظه ابزار جهت‌یابی']),
 '127WO2':(34,50,['اورشات ماهیگیری','یک شاخه طوق حفاری','جار ماهیگیری','یک دسته طوق حفاری']),
 '166':(9,33,['آسیاب فلزی ضایعات','رابط مته','خراشنده','طوق حفاری']),
}

def main():
 allraw={x['file']:x for x in json.loads(Path('evidence/bha_observations.json').read_text())};review=[]
 for well,w in DATA.items():
  day,page,fa_names=SELECT[well];raw=allraw[w['source']];s=next(x for x in raw['sheets'] if x['day']==day);b=s['blocks'][0];components=b['components'][:4];assert len(components)==len(fa_names)
  review.append({'well':well,'source_file':w['source'],'programme_page':page,'note':'Selected programme BHA page is a navigation anchor, not a verified comparison with the daily configuration. PDF 207 p.59 is descriptive; p.61–66 contain image-only figures.','sheet':s['sheet'],'identifier':b['identifier'],'block_description':b['block_description'],'components':components,'persian_names':fa_names,'sheet_coverage':raw['sheets_with_nonempty_bha_blocks'],'data_sheets_considered':raw['data_sheets_considered']})
  for lang in ('EN','FA'):
   fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
   old=next((p for p in d.paragraphs if p.text in ('Selected daily BHA components and hydraulic review','اجزای منتخب رشته ته‌چاهی و بازبینی هیدرولیکی')),None)
   if old:raise ValueError(f'Already inserted: {path}')
   anchor=next(p for p in d.paragraphs if p.text in ('Source-coded non-planned time and partial coverage','زمان غیربرنامه‌ای بر اساس کد منبع و پوشش جزئی'))
   head=d.add_heading('اجزای منتخب رشته ته‌چاهی و بازبینی هیدرولیکی' if fa else 'Selected daily BHA components and hydraulic review',level=2);head.alignment=1;anchor._p.addprevious(head._p)
   intro=(f"از {raw['data_sheets_considered']} شیت داده، در {raw['sheets_with_nonempty_bha_blocks']} شیت دست‌کم یک بلوک غیرخالی با قالب قابل تشخیص بازیابی شد. جدول زیر چهار ردیف از یک بلوک در {w['source']}، شیت {s['sheet']} است؛ این گزینش تمام رشته، تمام بازران‌ها یا وضعیت فعلی چاه نیست. شناسه تکراری یا ستون‌های هم‌زمان به‌تنهایی تعداد بازران را ثابت نمی‌کنند." if fa else f"At least one nonempty recognized-format block was recovered in {raw['sheets_with_nonempty_bha_blocks']} of {raw['data_sheets_considered']} data sheets. The table selects four rows from one block in {w['source']}, sheet {s['sheet']}; it is not the entire string, all runs, or the current well condition. A repeated identifier or parallel columns do not establish the number of runs.")
   p=d.add_paragraph(intro);style(p,fa);anchor._p.addprevious(p._p)
   table=d.add_table(rows=1,cols=3);table.style='Table Grid';anchor._p.addprevious(table._tbl)
   rows=[['جزء منتخب و سلول','قطر خارجی ثبت‌شده و سلول','کاربرد داده'] if fa else ['Selected component and cell','Recorded outside diameter and cell','Evidence scope']]
   for c,n in zip(components,fa_names):
    item=c['item'];od=c['od'];os=(f"{od['raw']} [{od['cell']}]" if od else ('ثبت نشده' if fa else 'Not recorded'))
    rows.append([f"{n if fa else item['raw']} [{item['cell']}]",os,'مشاهده منبع' if fa else 'Source observation'])
   set_table(table,rows,[7.3,4.5,3.8],fa)
   p=d.add_paragraph('جدول ۶-۲ ـ اجزای منتخب یک بلوک منبع، با مشخصات ثبت‌شده' if fa else 'Table 6-2 — Selected source block components and recorded dimensions');style(p,fa,True);anchor._p.addprevious(p._p)
   detail=(f"صفحه فیزیکی {page} برنامه چاه، محل رجوع به بحث رشته ته‌چاهی برنامه‌ای است؛ یکسانی فاز یا انطباق آن با این نمونه روزانه اثبات نشده است. ابعاد درج‌شده باید با نام ابزار، تعداد، طول، قطر داخلی، محدودیت اتصالات، ترتیب، سابقه بازران و برگه سازنده تطبیق شوند. " if fa else f"Physical programme page {page} is a navigation anchor for the planned BHA discussion; matching phase or agreement with this daily example is not established. Reconcile stated dimensions with the component name, count, length, inside diameter, connection limits, sequence, run record and manufacturer data. ")
   if well=='122WO1':detail+=('در این نمونه، نام خراشنده در سلول I43 برابر ۷ اینچ و مقدار ستون قطر خارجی K43 برابر ۵⅜ اینچ است؛ این اختلاف بدون مدرک حل نشده است. ' if fa else 'Here the scraper description in I43 says 7 inches while the OD column K43 says 5⅜ inches; the difference is unresolved. ')
   if well=='207':detail+=('صفحات ۶۱ تا ۶۶ برنامه حاوی نمودار تصویری‌اند و متن کامل آن‌ها از استخراج متنی قابل احراز نیست. ' if fa else 'Programme pages 61–66 contain graphical figures whose complete contents are not verified by text extraction. ')
   detail+=('برای بازبینی هیدرولیکی باید قطر داخلی و بیرونی هر جزء و فضای حلقوی، دبی و خواص سیال هم‌زمان، مشخصات نازل و تلفات سطحی و کالیبراسیون فشار معلوم باشند. بدون این ورودی‌ها افت فشار، سرعت حلقوی، تمیزکاری چاه یا محدوده عملیاتی قابل اتکا محاسبه نمی‌شود؛ داده‌های W3 به‌تنهایی طراحی نیستند.' if fa else 'Hydraulic review requires verified inside and outside diameters and annular geometry for each section, simultaneous flow and fluid properties, nozzle and surface-loss information and pressure calibration. Without these inputs, pressure loss, annular velocity, hole cleaning and a reliable operating window cannot be calculated; W3 alone is not a design.')
   p=d.add_paragraph(detail);style(p,fa);anchor._p.addprevious(p._p);d.save(path)
 Path('evidence/bha_selected_reviewed.json').write_text(json.dumps(review,indent=2,ensure_ascii=False)+'\n')
if __name__=='__main__':main()
