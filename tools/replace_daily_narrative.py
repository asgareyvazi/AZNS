#!/usr/bin/env python3
"""Replace unsupported daily causal narrative with documented examples and status log."""
from pathlib import Path
from io import BytesIO
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from docx import Document
from docx.oxml.ns import qn
from replace_well_bit_evidence import DATA,style
from apply_program_and_safety_revision import TRACKS,ACTIVITY

# Manually translated selected remarks. Match anchors prevent applying the
# translation to a different sheet/cell or to a different source edition.
EXAMPLES={
'207':[(17,'S27','Severe FM losses','The report notes 40–50 barrels per hour of formation losses, stopping drilling, pulling one stand and pumping a 50-barrel loss-control pill.','گزارش از هرزروی سازندی ۴۰ تا ۵۰ بشکه در ساعت، توقف حفاری، بیرون کشیدن یک پایه و پمپاژ ۵۰ بشکه گلوله کنترل هرزروی خبر می‌دهد.'),(87,'S25','3017.5m','The report describes resumed drilling of the 12¼-inch hole from 3,006 to 3,017.5 m, with 2–4 barrels per hour of losses.','گزارش ادامه حفاری حفره ۱۲¼ اینچ از ۳۰۰۶ تا ۳۰۱۷٫۵ متر را با هرزروی ۲ تا ۴ بشکه در ساعت توصیف می‌کند.')],
'127ST2':[(7,'S27','Cased-Hole Sidetrack BHA','The record describes assembling a cased-hole sidetrack assembly, including a hydraulic whipstock system and an 8½-inch mill.','ثبت روزانه مونتاژ مجموعه انحراف در جداری شامل سیستم ویپ‌استاک هیدرولیکی و آسیاب ۸½ اینچ را توصیف می‌کند.'),(33,'S34','MWD signal & DHM Diff. Press. loss','The report states nine hours of NPT associated with loss of MWD signal and downhole-motor differential pressure; the record’s party attribution is not a contractual determination.','گزارش ۹ ساعت زمان غیرمولد مرتبط با فقدان سیگنال اندازه‌گیری حین حفاری و اختلاف فشار موتور درون‌چاهی را ذکر می‌کند؛ انتساب طرف در گزارش به‌معنی تعیین مسئولیت قراردادی نیست.')],
'40':[(6,'K24','Recover all components of fish','The report records pulling the fishing assembly from 1,547.2 m and recovering all fish components, with 1–2 barrels per hour of losses.','گزارش بیرون کشیدن مجموعه ماهیگیری از ۱۵۴۷٫۲ متر و بازیابی تمامی اجزای قطعه جامانده را همراه با هرزروی ۱ تا ۲ بشکه در ساعت ثبت می‌کند.'),(147,'K26','6000 Psi; OK','The report says surface leakages were worked on and a 6,000 psi pressure test was reported as satisfactory; the cell does not identify a complete test programme.','گزارش رسیدگی به نشتی‌ها و نتیجه رضایت‌بخش آزمون فشار ۶۰۰۰ پوند بر اینچ مربع را ذکر می‌کند؛ این سلول برنامه کامل آزمون را مشخص نمی‌کند.')],
'127WO1':[(39,'K26','Retreive 21.1m','The report describes pulling the overshot and retrieving 21.1 m of 4½-inch tubing; the individual cell continues into subsequent text.','گزارش بیرون کشیدن اورشات و بازیابی ۲۱٫۱ متر لوله مغزی ۴½ اینچ را شرح می‌دهد؛ شرح این سلول در متن بعدی ادامه دارد.'),(130,'K22','pump truck surf. Lines','A later report records work to rectify leakage in surface lines of the pump truck; the original contains an NPT party label that is not adopted here as contractual liability.','گزارش بعدی رسیدگی به نشتی خطوط سطحی کامیون پمپاژ را ثبت می‌کند؛ برچسب طرف مربوط به زمان غیرمولد در منبع، در اینجا مسئولیت قراردادی تلقی نشده است.')],
'10':[(5,'K24','Mud Pump#1','The report records work on mud pump number 1; the selected cell does not identify a failure mechanism.','گزارش کار روی پمپ گل شماره ۱ را ثبت می‌کند؛ این سلول سازوکار خرابی را مشخص نمی‌کند.'),(65,'K30','Pressure Test Top Ram','The report records a two-step top-ram pressure test, 500 and 7,500 psi for 15 minutes, with an OK result as written.','گزارش آزمون فشار رام بالایی را در دو مرحله ۵۰۰ و ۷۵۰۰ پوند بر اینچ مربع به مدت ۱۵ دقیقه با نتیجه درج‌شده «مطلوب» ثبت می‌کند.')],
'12':[(16,'S24','Fishing BHA#02','The report records running fishing assembly number 2 from 3,290 to 3,305 m and recording upward and downward weights with and without pumping.','گزارش راندن مجموعه ماهیگیری شماره ۲ از ۳۲۹۰ تا ۳۳۰۵ متر و ثبت وزن بالا و پایین با و بدون پمپاژ را بیان می‌کند.'),(48,'S25','Magneset plug','The report describes drilling a soft-to-medium Magneset plug from 3,180 to 3,270 m, then washing down to 3,285 m.','گزارش حفاری پلاگ مگنست نرم تا متوسط از ۳۱۸۰ تا ۳۲۷۰ متر و سپس شست‌وشو تا ۳۲۸۵ متر را شرح می‌دهد.')],
'75':[(5,'S25','BOP test several Time No Succ','The report describes repeated unsuccessful BOP test attempts while work continued on BOP leakages.','گزارش تلاش‌های ناموفق مکرر برای آزمون فوران‌گیر و ادامه کار برای رفع نشتی‌های آن را شرح می‌دهد.'),(5,'S26','Wait on OEOC to Complete BOP Test','The next entry records waiting for completion of the BOP test while leakage repairs continued to midnight; attribution requires contract review.','ثبت بعدی انتظار برای تکمیل آزمون فوران‌گیر و ادامه رفع نشتی تا نیمه‌شب را ذکر می‌کند؛ انتساب مسئولیت نیازمند بررسی قرارداد است.')],
'85':[(5,'K32','No success','The report describes unsuccessful attempts to free a pressure tester using overpull, followed by picking up a fishing jar.','گزارش تلاش ناموفق برای آزاد کردن تستر فشار با کشش اضافی و سپس برداشتن جار ماهیگیری را شرح می‌دهد.'),(42,'K34','4033m','The report records running the 5⅞-inch clean-out assembly from 3,670 to 4,033 m.','گزارش راندن مجموعه تمیزکاری ۵⅞ اینچ از ۳۶۷۰ تا ۴۰۳۳ متر را ثبت می‌کند.')],
'116':[(2,'K22','Ring Gasket BX-152','The report records continuing work on leakage at the BX-152 ring gasket and says the gasket was changed with an OK result.','گزارش ادامه رسیدگی به نشتی در واشر حلقوی BX-152 و تعویض آن با نتیجه درج‌شده «مطلوب» را ثبت می‌کند.'),(98,'K22','3947 to 3969m','The report describes continuing milling and smoothing in the 3,947–3,969 m interval, with no loss or gain reported.','گزارش ادامه آسیاب‌کاری و هموارسازی بازه ۳۹۴۷ تا ۳۹۶۹ متر را با ثبت نبود افت یا افزایش سیال شرح می‌دهد.')],
'122WO1':[(8,'K22','Obs. Leakage f/ string','During an attempted pressure test, the report notes leakage from the tubing-running tools and records backing off those tools.','گزارش هنگام تلاش برای آزمون فشار، نشتی از ابزار راندن مغزی و بازکردن آن ابزار را ثبت می‌کند.'),(53,'K26','7000 psi for 15 min','The report records a 7,000 psi, 15-minute test of the completion string and associated equipment, with an OK result as written.','گزارش آزمون ۷۰۰۰ پوند بر اینچ مربع به مدت ۱۵ دقیقه برای رشته تکمیل و تجهیزات مرتبط را با نتیجه درج‌شده «مطلوب» ثبت می‌کند.')],
'122WO2':[(9,'S26','3721 to 3721.8m','The report records starting to mill the packer from 3,721 to 3,721.8 m; recorded parameters are historical, not future setpoints.','گزارش آغاز آسیاب‌کاری پکر از ۳۷۲۱ تا ۳۷۲۱٫۸ متر را ثبت می‌کند؛ پارامترهای درج‌شده تاریخی‌اند و نقطه تنظیم آینده نیستند.'),(118,'S31','leakage BOP','The report records work to resolve a BOP leakage; the selected cell does not establish its cause.','گزارش رسیدگی به نشتی فوران‌گیر را ثبت می‌کند؛ سلول انتخاب‌شده علت آن را اثبات نمی‌کند.')],
'127WO2':[(18,'S33','Pressure Test TWCV','The report records setting and pressure-testing the tubing-hanger valve to 7,000 psi, with an OK result as written.','گزارش نصب و آزمون فشار شیر آویز مغزی تا ۷۰۰۰ پوند بر اینچ مربع را با نتیجه درج‌شده «مطلوب» ثبت می‌کند.'),(44,'S35','44m of Fish was recovered','The report says 44 m of the fish was recovered while dismantling the fishing assembly; it does not by itself establish complete fish recovery.','گزارش بازیابی ۴۴ متر از قطعه جامانده هنگام بازکردن مجموعه ماهیگیری را بیان می‌کند؛ این سلول به‌تنهایی بازیابی کامل قطعه را ثابت نمی‌کند.')],
'166':[(5,'S25','BOP Stack While Work on fixing BOP leakages','The report describes waiting for continuation of the BOP test while BOP leakages were being repaired.','گزارش انتظار برای ادامه آزمون فوران‌گیر هم‌زمان با رفع نشتی‌های آن را شرح می‌دهد.'),(6,'S26','Leakage TDS Kelly Hose','The report records an unsuccessful attempt to circulate mud through the string after leakage was observed at the top-drive kelly hose.','گزارش تلاش ناموفق برای گردش گل از رشته پس از مشاهده نشتی شیلنگ کلیِ تاپ‌درایو را ثبت می‌کند.')]
}

def chart(well,wellinfo,activity):
 days=activity['days'];complete=[x for x in days if x['activity']['complete']];last=max(TRACKS[wellinfo['source']]['paired_sheet_candidates'],max((d['day'] for d in days),default=0));fig,ax=plt.subplots(figsize=(13,6),dpi=300);fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white')
 colors=[('#28688d','planned_h','Reported planned'),('#d0a153','npt_h','N / NPT'),('#c65d42','unplanned_h','UP / unplanned'),('#8a65a5','mixed_h','NPT-UNPLN')]
 for d in complete:
  bottom=0
  for color,key,label in colors:
   val=d['activity'][key];ax.bar(d['day'],val,width=.9,bottom=bottom,color=color,linewidth=0);bottom+=val
 missing=sorted(set(range(1,last+1))-{d['day'] for d in complete})
 if missing:ax.scatter(missing,[24.8]*len(missing),marker='|',s=30,c='#a23e32',linewidths=.8)
 ax.set_xlim(.2,last+1);step=20 if last>100 else 10;ax.set_xticks([1]+list(range(step,last+1,step)));ax.set_ylim(0,26.3);ax.set_yticks([0,6,12,18,24]);ax.set_ylabel('Source-classified hours',fontsize=10);ax.set_xlabel('Report day number, not verified elapsed rig duration',fontsize=10);ax.grid(axis='y',alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right']].set_visible(False)
 fig.suptitle(f'AZNS-{well} | Daily source status codes',x=.05,y=.98,ha='left',fontsize=15,fontweight='bold',color='#223952')
 fig.text(.05,.916,f"{len(complete)} complete classified days in day-number span 1–{last}; missing or incomplete marked above 24h.",fontsize=9,color='#9b431c')
 fig.legend(handles=[*(Patch(color=c,label=l) for c,k,l in colors),Patch(color='#a23e32',label='Incomplete / missing')],loc='upper center',bbox_to_anchor=(.5,.89),ncol=5,fontsize=8,frameon=False)
 fig.text(.05,.04,'Status codes are not independently validated causes, responsibility or cost. Sources: activity_observations.json, daily Remark cells.',fontsize=8.5,color='#9b431c')
 fig.subplots_adjust(left=.08,right=.97,top=.79,bottom=.18);fig.canvas.draw();ren=fig.canvas.get_renderer()
 for t in [*fig.texts,*ax.get_xticklabels(),*ax.get_yticklabels()]:
  box=t.get_window_extent(ren)
  if box.x0<0 or box.x1>fig.bbox.x1 or box.y0<0 or box.y1>fig.bbox.y1:raise ValueError((well,t.get_text()))
 b=BytesIO();fig.savefig(b,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return b.getvalue()

def examples(well,activity):
 result=[]
 for day,cell,anchor,en,fa in EXAMPLES[well]:
  d=next(d for d in activity['days'] if d['day']==day);x=next(x for x in d['detail_cells'] if x['cell']==cell);assert anchor in x['raw'],(well,day,cell)
  result.append({'day':day,'sheet':x['sheet'],'cell':cell,'raw':x['raw'],'english_summary':en,'persian_summary':fa})
 return result

def main():
 output=[]
 for well,w in DATA.items():
  a=ACTIVITY[w['source']];samples=examples(well,a);output.append({'well':well,'file':w['source'],'selected_examples':samples});png=chart(well,w,a);complete=sum(d['activity']['complete'] for d in a['days'])
  for lang in ('EN','FA'):
   fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
   begin=next(p for p in d.paragraphs if p.text in ('Chronological Operational Narrative','روایت زمانی عملیات','روایت منتخب از گزارش روزانه','Selected daily-report observations'))
   end=next(p for p in d.paragraphs if p.text.startswith(('Performance and Parameter Analysis','تحلیل عملکرد و پارامترها')))
   for r in begin.runs:r.text=''
   begin.add_run('روایت منتخب از گزارش روزانه' if fa else 'Selected daily-report observations');begin.alignment=1
   caption=next(p for p in d.paragraphs if p.text.startswith(('Figure W2','شکل W2')));drawing=caption._p.getprevious();bl=drawing.xpath('.//a:blip')[0];rid=bl.get(qn('r:embed'));d.part.related_parts[rid]._blob=png
   shape=next(s for s in d.inline_shapes if any(x.get(qn('r:embed'))==rid for x in s._inline.xpath('.//a:blip')));shape.height=int(shape.width*6/13)
   # Remove only the replaced section, keeping the drawing and caption.
   cursor=begin._p.getnext()
   while cursor is not end._p:
    nxt=cursor.getnext()
    if nxt is None:raise ValueError((well,lang,'missing end'))
    if cursor is not drawing and cursor is not caption._p:cursor.getparent().remove(cursor)
    cursor=nxt
   para=d.add_paragraph((f"از {a['remark_days']} صفحه وضعیت شناسایی‌شده، {complete} گزارش مدت‌های قابل طبقه‌بندی با جمع نزدیک ۲۴ ساعت دارند. برچسب‌های برنامه‌ای، غیرمولد و غیربرنامه‌ای عین طبقه‌بندی منبع‌اند و به‌تنهایی علت، مسئولیت یا مدت کل عملیات را ثابت نمی‌کنند. نمونه‌های زیر فقط گزیده‌اند؛ سلول‌های کامل در پرونده evidence/activity_observations.json نگهداری می‌شوند." if fa else f"Of {a['remark_days']} recognized status pages, {complete} contain classified durations totalling approximately 24 hours. Planned, NPT and unplanned labels are source classifications; they do not independently establish cause, liability or total operation duration. The following examples are selective; complete extracted detail cells are retained in evidence/activity_observations.json."));style(para,fa);drawing.addprevious(para._p)
   for s in samples:
    msg=(f"گزارش {s['day']} ـ {s['persian_summary']} [{w['source']}، {s['sheet']}!{s['cell']}]" if fa else f"Report {s['day']} — {s['english_summary']} [{w['source']}, {s['sheet']}!{s['cell']}]")
    p=d.add_paragraph(msg);style(p,fa);drawing.addprevious(p._p)
   p=d.add_paragraph('این نقل‌ها شرح منبع‌اند، نه تأیید مستقل آزمون، علت خرابی، سلامت فعلی چاه یا دستور عملیاتی. یک شرح ممکن است به سلول بعدی ادامه یابد؛ پیش از تصمیم‌گیری باید ردیف کامل و زمان آن تطبیق شود.' if fa else 'These are source statements, not independent verification of a test, failure mechanism, current well condition or an operating instruction. A description may continue into the next cell; reconcile the complete row and time before making decisions.');style(p,fa);drawing.addprevious(p._p)
   # Keep captions in their original centered style.
   if caption.runs:
    caption.runs[0].text='شکل W2 ـ ترکیب روزانه کدهای منبع در گزارش‌های کامل، با پوشش جزئی' if fa else 'Figure W2 — Daily source-code composition in complete reports, with partial coverage'
    for r in caption.runs[1:]:r.text=''
   d.save(path)
  print(well,complete,'selected',len(samples),flush=True)
 Path('evidence/daily_observation_examples.json').write_text(json.dumps(output,ensure_ascii=False,indent=2)+'\n')
if __name__=='__main__':main()
