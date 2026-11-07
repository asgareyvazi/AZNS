#!/usr/bin/env python3
"""Add strictly source-cited CBS service samples; never sum across scopes."""
import json
from pathlib import Path
from openpyxl import load_workbook
from docx import Document
from replace_well_bit_evidence import style

CT={'12':('CBS AZNS-12.xlsx','Coil Tubing Service-12'),'75':('CBS AZNS-075.xlsx','Coil Tubing Service-075'),'116':('CBS AZNS-127WO1 and 116.xlsx','Coil Tubing Service-116'),'122WO2':('CBS AZNS-122WO2.xlsx','Coil Tubing Service-122'),'127WO2':('CBS AZNS-127WO2.xlsx','Coil Tubing Service-127#2'),'166':('CBS AZNS-166.xlsx','Coil Tubing Service-166'),'207':('CBS AZNS-207.xlsx','Coil Tubing Service-F-20')}

def sample(file,sheet,addresses):
 out={'file':file,'sheet':sheet,'cells':{}}
 for data in (False,True):
  wb=load_workbook(file,read_only=True,data_only=data);ws=wb[sheet]
  for addr in addresses:
   v=ws[addr].value
   out['cells'].setdefault(addr,{})['cached' if data else 'formula_or_value']=v
  wb.close()
 return out

def insert(well,en,fa):
 for lang,body in [('EN',en),('FA',fa)]:
  f=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');doc=Document(f);title='Selected CBS service-sheet checks' if lang=='EN' else 'کنترل‌های منتخب شیت خدمات در ساختار شکست هزینه'
  if any(p.text==title for p in doc.paragraphs):continue
  end=next(p for p in doc.paragraphs if p.text in ('Operational observations and engineering decision gates','بازبینی پارامترهای عملیاتی و دروازه‌های تصمیم'))
  h=doc.add_heading(title,level=2);h.alignment=1;end._p.addprevious(h._p)
  for t in body:
   p=doc.add_paragraph(t);style(p,lang=='FA');end._p.addprevious(p._p)
  doc.save(f)

def main():
 out=[]
 for well,(file,sheet) in CT.items():
  rec=sample(file,sheet,['A1','B3','C3','D3','E3','F3','G3','H3']);rec['well']=well;out.append(rec)
  c=rec['cells'];v={k:x['cached'] for k,x in c.items()};assert v['B3']=='Coiled Tubing Unit (1 1/2")';assert v['H3']==v['D3']*v['F3']+v['E3']*v['G3'];assert v['F3']==v['G3']==4800
  en=[f"The coiled-tubing service sheet in {file}, {sheet}!B3:H3, lists one 1½-inch coiled-tubing unit, {v['D3']} standby days and {v['E3']} operating days, with both day rates at USD 4,800 and a cached line amount of USD {v['H3']:,.0f}. The source formula multiplies each day count by its corresponding rate. This is a selected estimated service line, not a verified operational record or invoice.", 'This sample is not an additional cost to be added to the main breakdown or ceiling: the roll-up, contingencies, units, contract scope and service execution must first be traced and reconciled. A repeated template line across workbooks is not independent evidence of repeated deployment.']
  fa=[f"شیت خدمات لوله‌مغزی سیار در {file}، {sheet}!B3:H3 یک دستگاه ۱½ اینچ، {v['D3']} روز آماده‌به‌کار و {v['E3']} روز عملیاتی با نرخ روزانه هر دو ۴۸۰۰ دلار و مبلغ ذخیره‌شده ردیف {v['H3']:,.0f} دلار را نشان می‌دهد. فرمول منبع هر تعداد روز را در نرخ متناظر ضرب می‌کند. این یک ردیف منتخب برآورد خدمات است، نه سند تأییدشده عملیات یا صورت‌حساب.", 'این نمونه هزینه اضافی قابل جمع با ساختار اصلی یا سقف نیست: ابتدا مسیر تجمیع، مبالغ احتیاطی، واحدها، دامنه قرارداد و اجرای خدمت باید ردیابی و تطبیق شوند. تکرار ردیف قالب در فایل‌های مختلف، شاهد مستقل استقرار مکرر تجهیزات نیست.']
  if well=='207':
   extra=sample(file,'Well testing and production-127',['A1','H3']);extra['well']=well;out.append(extra);assert 'F-20' in extra['cells']['A1']['cached']
   en.append('The tab “Well testing and production-127” identifies Well F-20 in A1 despite the 127 suffix; workbook-to-well identity and service allocation require confirmation. The tab name alone is not an assignment.')
   fa.append('عنوان سلول A1 در شیت «Well testing and production-127» با وجود پسوند ۱۲۷، چاه F-20 را معرفی می‌کند؛ هویت فایل و تخصیص خدمت باید تأیید شوند. نام شیت به‌تنهایی مبنای انتساب نیست.')
  insert(well,en,fa)
 rec=sample('CBS AZNS-40.xlsx','Cost breakdown Rigless Stim.-40',['A1']+[f'{c}{r}' for r in range(3,16) for c in 'BC']+['C16']);rec['well']='40';out.append(rec)
 vals=[rec['cells'][f'C{r}']['cached'] for r in range(3,16)];assert round(sum(vals),2)==rec['cells']['C16']['cached']==749473.64
 insert('40',['The rig-less stimulation bidder breakdown is separately labelled for AZN-40: CBS AZNS-40.xlsx, Cost breakdown Rigless Stim.-40!A1, B3:C16. Its 13 source lines total USD 749,473.64 in C16; the independently checked sum of C3:C15 matches that amount. Examples are materials USD 159,182.50 (C4), acidizing and nitrogen services USD 117,529 (C5), coiled tubing USD 44,900 (C6), and well testing and production USD 83,282.84 (C8).','This is a bidder-scope breakdown, not an invoice or confirmation of rig-less treatment. Do not add it to the main-sheet total until the contract boundaries, any included services and roll-up are verified; the arithmetic match alone does not validate inputs or rates.'],['ساختار شکست تحریک بدون دکل برای پیشنهاددهندگان به‌طور جداگانه برای چاه ۴۰ درج شده است: CBS AZNS-40.xlsx، Cost breakdown Rigless Stim.-40!A1 و B3:C16. جمع ۱۳ ردیف منبع در C16 برابر ۷۴۹٬۴۷۳٫۶۴ دلار است؛ جمع مستقل C3:C15 نیز با آن مطابقت دارد. نمونه‌ها شامل مواد ۱۵۹٬۱۸۲٫۵۰ دلار (C4)، خدمات اسیدکاری و نیتروژن ۱۱۷٬۵۲۹ دلار (C5)، لوله‌مغزی سیار ۴۴٬۹۰۰ دلار (C6) و آزمایش و تولید ۸۳٬۲۸۲٫۸۴ دلار (C8) هستند.','این ساختار دامنه پیشنهاددهنده است، نه صورت‌حساب یا تأیید اجرای تحریک بدون دکل. تا بررسی مرزهای قرارداد، خدمات احتمالی منظورشده و مسیر تجمیع، آن را با جمع شیت اصلی اضافه نکنید؛ تطابق حسابی به‌تنهایی ورودی‌ها یا نرخ‌ها را تأیید نمی‌کند.'])
 rec=sample('CBS AZNS-127 WO1.xlsx','AZNS-127 WO (Side Tracking) ',['B2','R2','R23','R24','R83']);rec['well']='127WO1';out.append(rec);assert rec['cells']['R23']['cached']=='#REF!' and '#REF!' in rec['cells']['R23']['formula_or_value']
 insert('127WO1',['In CBS AZNS-127 WO1.xlsx, AZNS-127 WO (Side Tracking) !B2 and R2, the source labels a workover/side-tracking schedule and a cumulative duration with contingency. The formula at R23 contains a broken reference and its cached value is #REF!; subsequent cached errors occur in the same column. This schedule must not be used as a validated duration until repaired and reconciled with the separate programme.','The sheet name and workbook do not by themselves prove that the proposed side track was executed or establish a mapping to the distinct ST1 or ST2 programmes.'],['در CBS AZNS-127 WO1.xlsx، سلول‌های AZNS-127 WO (Side Tracking) !B2 و R2 برنامه تعمیر/انحراف و مدت تجمعی با احتیاط را معرفی می‌کنند. فرمول R23 ارجاع شکسته دارد و مقدار ذخیره‌شده آن #REF! است؛ خطاهای ذخیره‌شده بعدی نیز در همان ستون دیده می‌شوند. تا ترمیم و تطبیق با برنامه مستقل، این شیت نباید مدت معتبر تلقی شود.','نام شیت و فایل به‌تنهایی اجرای انحراف پیشنهادی یا ارتباط آن با برنامه‌های متمایز ST1 یا ST2 را اثبات نمی‌کند.'])
 Path('evidence/cbs_service_spot_checks.json').write_text(json.dumps(out,ensure_ascii=False,indent=2)+'\n');print('spot-check records',len(out))
if __name__=='__main__':main()
