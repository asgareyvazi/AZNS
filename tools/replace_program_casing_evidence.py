#!/usr/bin/env python3
"""Replace legacy casing tables/W6 with selected PDF-stated intervals."""
import json,textwrap
from io import BytesIO
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.oxml.ns import qn
from replace_well_bit_evidence import set_table,style,text

DATA=json.loads(Path('evidence/program_casing_intervals_reviewed.json').read_text())

def chart(w):
 fig,ax=plt.subplots(figsize=(11,6),dpi=300);fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white');data=w['intervals'];last=max(i['to_m'] for i in data)
 for n,x in enumerate(data):
  ax.vlines(n,x['from_m'],x['to_m'],color='#28688d',lw=7,zorder=2);ax.scatter([n],[x['from_m']],color='#208365',s=22,zorder=3);ax.scatter([n],[x['to_m']],color='#a94331',marker='_',s=130,linewidths=2,zorder=3)
 ax.set_xticks(range(len(data)),['\n'.join(textwrap.wrap(x['label_en'],17,break_long_words=False,break_on_hyphens=False))+f"\n{x['from_m']:g}–{x['to_m']:g} m" for x in data],fontsize=8,rotation=35,ha='right',rotation_mode='anchor');ax.tick_params(axis='x',length=0,pad=7);ax.set_xlim(-.7,len(data)-.3);ax.set_ylim(last*1.08,-last*.045);ax.set_ylabel('Programme-stated depth (m); datum to verify');ax.grid(axis='y',alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right']].set_visible(False)
 ax.plot([],[],marker='o',linestyle='None',color='#208365',label='Interval start');ax.plot([],[],marker='_',linestyle='None',markersize=12,color='#a94331',label='Interval end');fig.legend(loc='upper right',bbox_to_anchor=(.97,.89),ncol=2,fontsize=8,frameon=False)
 fig.suptitle(f"AZNS-{w['well']} | Programme-stated casing intervals",x=.05,y=.98,ha='left',fontsize=15,fontweight='bold',color='#223952')
 fig.text(.05,.92,f"Source: {w['file']}, physical p. {w['physical_page']}; programme table, not verified as-built.",fontsize=8.4,color='#9b431c')
 fig.text(.05,.04,'Segments can denote changes in specification, not separate shoes. No cement, tubing, barriers or nesting inferred.',fontsize=8.4,color='#9b431c')
 fig.subplots_adjust(left=.10,right=.97,top=.82,bottom=.31);fig.canvas.draw();renderer=fig.canvas.get_renderer()
 for t in [*ax.get_xticklabels(),*fig.texts]:
  b=t.get_window_extent(renderer)
  if b.x0<0 or b.x1>fig.bbox.x1 or b.y0<0 or b.y1>fig.bbox.y1:raise ValueError((w['well'],t.get_text()))
 out=BytesIO();fig.savefig(out,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return out.getvalue()

def main():
 for w in DATA:
  well=w['well']
  if well=='127ST1':continue # No separate ST1 daily-report edition.
  png=chart(w)
  for lang in ('EN','FA'):
   fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
   old=next((t for t in d.tables if t.rows and t.rows[0].cells[0].text in ('Casing / liner','جداری / آستری','Programme element','جزء در برنامه')),None)
   assert old is not None,(well,lang)
   t=d.add_table(rows=1,cols=4);t.style='Table Grid';old._tbl.addprevious(t._tbl);old._tbl.getparent().remove(old._tbl)
   row=[('جزء در برنامه','ابتدای بازه، متر','انتهای بازه، متر','مشخصات منبع') if fa else ('Programme element','Interval start, m','Interval end, m','Source specification')]
   row.extend([(i['label_fa'] if fa else i['label_en'],f"{i['from_m']:g}",f"{i['to_m']:g}",i['specification']) for i in w['intervals']]);set_table(t,row,[4.2,2.4,2.4,7.0],fa)
   cap=next(p for p in d.paragraphs if p.text.startswith(('Table 6-1','جدول ۶-۱')));text(cap,(f"جدول ۶-۱ ـ بازه‌های جداری و آستری در برنامه، {w['file']}، صفحه فیزیکی {w['physical_page']}" if fa else f"Table 6-1 — Casing and liner intervals in {w['file']}, physical p. {w['physical_page']}"));cap.alignment=1
   # Re-run safe source note
   for p in list(d.paragraphs):
    if p.text.startswith(('Programme casing scope:', 'محدوده داده جداری برنامه:')):p._p.getparent().remove(p._p)
   scope=('محدوده داده جداری برنامه: این ارقام برداشت از جدول برنامه و مربوط به وضعیت تاریخی یا پیش از عملیات هستند، نه نقشه نهایی تأییدشده. ابتدا و انتهای بخش‌های مشخصات لزوماً کفشک یا آویز نیستند؛ مرجع عمق، آرایش واقعی، تالی و سلامت موانع باید با سوابق مصوب تطبیق شود.' if fa else 'Programme casing scope: these values transcribe the programme table and describe historical or pre-operation conditions, not a verified final wellbore. Endpoints of specification segments are not necessarily shoes or hangers; datum, actual configuration, tally and barrier integrity require approved records.')
   if well=='207':scope=('محدوده داده جداری برنامه: این ارقام طراحی برنامه‌ریزی‌شده‌اند، نه اجرای تأییدشده؛ عمق، آرایش و سلامت موانع باید با تالی و سوابق نهایی تطبیق شود.' if fa else 'Programme casing scope: these are planned design intervals, not verified execution. Depths, configuration and barrier integrity require reconciliation with the final tally and records.')
   if well=='40':scope=('محدوده داده جداری برنامه: برنامه دو جداری نخست را قبلاً نصب‌شده می‌نامد و دو ردیف بعدی طراحی‌اند؛ هیچ‌یک به‌تنهایی وضعیت نهایی تأییدشده نیست. صفحه ۱۳ همان برنامه عمق جداری ۲۰ اینچ را ۱۵۲٫۵ متر و صفحه ۱۴۹ مقدار ۱۵۰ متر ذکر می‌کند؛ اختلاف نیازمند تطبیق است.' if fa else 'Programme casing scope: the source labels the first two casings already run; the later two rows are design, not verified final configuration. Page 13 states 152.5 m for the 20-inch casing while page 149 states 150 m; reconcile the discrepancy.')
   if well=='127ST2':scope+=(' منبع این جدول را صراحتاً پیش از انحراف دوم نام‌گذاری می‌کند.' if fa else ' The source explicitly labels this table before sidetrack 2.')
   if well=='12':scope+=(' رأس آستری ۷ اینچ در برنامه ۳۲۸۵ متر است؛ عدد ۳۲۸۶٫۵ جدول قبلی با این منبع یکسان نبود و باید جداگانه تطبیق شود.' if fa else ' The programme gives the 7-inch liner top as 3,285 m; the legacy table stated 3,286.5 m, requiring separate reconciliation.')
   if well=='127WO1':scope+=(' عمق جداری ۲۰ اینچ در برنامه ۱۴۶ متر است؛ عدد ۱۵۵ متر جدول قبلی با این منبع یکسان نبود و باید جداگانه تطبیق شود.' if fa else ' The programme gives the 20-inch casing depth as 146 m; the legacy table stated 155 m, requiring separate reconciliation.')
   p=d.add_paragraph(scope);style(p,fa);cap._p.addnext(p._p)
   cap6=next(p for p in d.paragraphs if p.text.startswith(('Figure W6','شکل W6')));drawing=cap6._p.getprevious();bl=drawing.xpath('.//a:blip')[0];rid=bl.get(qn('r:embed'));d.part.related_parts[rid]._blob=png
   sh=next(s for s in d.inline_shapes if any(x.get(qn('r:embed'))==rid for x in s._inline.xpath('.//a:blip')));sh.height=int(sh.width*6/11)
   text(cap6,'شکل W6 ـ بازه‌های جداری درج‌شده در برنامه؛ شماتیک تأییدشده نیست' if fa else 'Figure W6 — Programme-stated casing intervals; not a verified as-built schematic');cap6.alignment=1
   # W7 no longer shows a trajectory; withdraw the unsupported old numeric claim.
   for p in d.paragraphs:
    if 'trajectory panel below' in p.text.lower():text(p,'The legacy directional survey statistics and trajectory have not been independently reconciled; W7 now shows bit/tool evidence, not a trajectory.');style(p,fa)
    if ('نمودار مسیر' in p.text or 'پنل مسیر زیر' in p.text) and ('ایستگاه' in p.text or 'انحراف' in p.text):text(p,'آمار پیمایش و مسیر قبلی مستقلاً تطبیق نشده است؛ شکل W7 اکنون شواهد مته و ابزار را نشان می‌دهد، نه مسیر چاه.');style(p,fa)
   d.save(path)
  print(well,len(w['intervals']),flush=True)
if __name__=='__main__':main()
