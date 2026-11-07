#!/usr/bin/env python3
"""Correct the M7 axis and annotate field reports with header-vs-sheet audit."""
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from replace_well_bit_evidence import style,text
import build_field_source_figures as b

meta=b.load('field_figure_observations');rows=meta['figures']['M7']
fig=b.canvas('Selected operations | Report-end depth observations','Four operations with varying recorded depth; dots are daily samples, not a field learning curve.','Worksheet prefix is not necessarily report number or elapsed time; report-end MD is not event depth.')
axs=fig.subplots(2,2)
for ax,row in zip(axs.flat,rows):
 pts=row['points'];ax.scatter([x['day'] for x in pts],[x['md'] for x in pts],s=11,c=b.BLUE,alpha=.7);ax.set_title(f"AZNS-{row['well']}  |  {len(pts)}/{row['recognized_data_sheets']} data pages",loc='left',fontsize=10,fontweight='bold',color=b.INK);ax.set_xlabel('Worksheet prefix');ax.set_ylabel('Report-end MD (m)');ax.invert_yaxis();ax.grid(alpha=.2);ax.spines[['top','right']].set_visible(False)
fig.subplots_adjust(left=.10,right=.96,top=.84,bottom=.13,hspace=.50,wspace=.30);b.save(fig,7);png=(b.OUT/'M7.png').read_bytes()
for typ in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'Field_{typ}_Report_{lang}.docx');d=Document(path)
  if typ=='Master_Drilling_Engineering':
   pars=d.paragraphs;i=next(i for i,p in enumerate(pars) if p.text.startswith(('Figure M7','شکل M7')));rid=pars[i-1]._p.xpath('.//a:blip')[0].get(qn('r:embed'));d.part.related_parts[rid]._blob=png
   for p in pars:
    if p.text.startswith('Four operations with changing source depth'):
     text(p,'Four operations are shown as individual report-end depth samples by numeric worksheet prefix. The prefix is not necessarily the report number or chronology; points are not continuous trajectories, drilling hours, event depths or evidence of a learning curve.');style(p,fa)
    elif p.text.startswith('چهار عملیات دارای تغییر عمق'):
     text(p,'چهار عملیات به‌صورت نمونه‌های مستقل عمق پایان گزارش برحسب پیشوند عددی شیت نمایش داده شده‌اند. پیشوند الزاماً شماره گزارش یا توالی زمانی نیست؛ نقاط مسیر پیوسته، ساعت حفاری، عمق رخداد یا شاهد منحنی یادگیری نیستند.');style(p,fa)
  title='تفاوت شماره شیت با شماره گزارش' if fa else 'Worksheet prefix versus report number'
  if not any(p.text==title for p in d.paragraphs):
   h=d.add_heading(title,level=2);h.alignment=1
   msg=('بازبینی سرصفحه ۱۳۰۱ صفحه وضعیت نشان داد در ۴۴۷ صفحه، شماره گزارش عددی با پیشوند نام شیت تفاوت دارد و یک سرصفحه عددی قابل خواندن نیست. برای نمونه، در بسته چاه ۱۱۶ شیت‌های 102-1 و 103-1 به‌ترتیب گزارش‌های ۱۱۷ و ۱۱۸ هستند، اما شیت‌های 104-1 تا 112-1 گزارش‌های ۱۰۸ تا ۱۱۶ را دارند. همچنین یادداشت شروع، ۸ ژوئن ۲۰۲۴ و سرصفحه آخر ۷ ژوئن را نشان می‌دهند. ترتیب ظاهری شیت گاه‌شمار یا مدت عملیات نیست؛ داده‌های سرصفحه نیز باید با نسخه امضاشده تطبیق شوند. [AZNS116Pack.xlsx، 1-1 Remark!K22، شیت‌های 101-1 تا 112-1 سلول S5 و 103-1 Remark!N6؛ evidence/sheet_report_number_audit.json]' if fa else 'A header review of 1,301 status pages found 447 where the numeric report header differs from the worksheet prefix, plus one unreadable numeric header. In well 116, sheets 102-1 and 103-1 carry reports 117 and 118, whereas sheets 104-1 to 112-1 carry reports 108–116. Its start note states 8 June 2024 while the final header states 7 June. Sheet order is not chronology or operating duration; headers themselves must be reconciled to signed records. [AZNS116Pack.xlsx, 1-1 Remark!K22, sheets 101-1 through 112-1 cell S5 and 103-1 Remark!N6; evidence/sheet_report_number_audit.json]')
   p=d.add_paragraph(msg);style(p,fa)
  d.save(path)
print('M7 updated, field notes added')
