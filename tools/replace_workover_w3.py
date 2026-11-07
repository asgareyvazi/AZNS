#!/usr/bin/env python3
"""Daily workover observations, with explicit coverage and unknown handling.
Activity statuses are source classifications, not independently validated NPT.
"""
from pathlib import Path
from io import BytesIO
import json,math,re
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from matplotlib.ticker import MaxNLocator
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from replace_drilling_w3 import number

files={'10':'AZNS10Pack.xlsx','12':'AZNS12Pack.xlsx','75':'AZNS75Pack.xlsx','85':'AZNS85Pack.xlsx','116':'AZNS116Pack.xlsx','122WO1':'AZNS122WO1Pack.xlsx','122WO2':'AZNS122WO2Pack.xlsx','127WO2':'AZNS127WOPack.xlsx','166':'AZNS166Pack.xlsx'}
tracks={x['file']:x for x in json.loads(Path('evidence/drilling_track_observations.json').read_text())}

def series(days,key,allow=None):
 out=[]
 for d in days:
    x=number(d['fields'],key)
    if x is None or not math.isfinite(x) or x<0:continue
    if allow and not allow(x):continue
    out.append((x,d['day']))
 return out

def dots(ax,points,color,marker='o'):
 if points:ax.scatter([x for x,y in points],[y for x,y in points],s=15,color=color,alpha=.7,marker=marker,edgecolors='none')

def chart(well,w):
 days=w['days'];assert days
 total=w['paired_sheet_candidates'];maxday=max(total,max(d['day'] for d in days));complete=[d for d in days if d['activity']['complete']]
 md=series(days,'md_end',lambda v:v<=6000);mw=series(days,'mud_weight',lambda v:30<=v<=200)
 pmin=series(days,'spp_a_min',lambda v:v<=15000);pmax=series(days,'spp_a_max',lambda v:v<=15000);loss=series(days,'downhole_loss',lambda v:v<=50000)
 tracks=[];omitted=[]
 for name,have in [('MD',bool(md)),('MW',bool(mw)),('pressure / loss',bool(pmin or pmax or loss)),('status',bool(complete)),('field counts',True)]:
  (tracks if have else omitted).append(name)
 fig,plots=plt.subplots(1,len(tracks),figsize=(13,9),dpi=300,sharey=True,gridspec_kw={'wspace':.28});fig.patch.set_facecolor('#fff');axs=dict(zip(tracks,plots))
 for i,ax in enumerate(plots):
  ax.set_ylim(maxday+3,0);ax.grid(axis='y',alpha=.18);ax.spines[['top','right']].set_visible(False);ax.tick_params(axis='x',labelsize=8);ax.tick_params(axis='y',labelsize=8)
  if i:ax.tick_params(axis='y',labelleft=False,length=0)
 plots[0].set_ylabel('Worksheet prefix (not report number)',fontsize=10)
 if md:
  ax=axs['MD'];dots(ax,md,'#28688d');ax.set_xlim(0,max(v for v,_ in md)*1.06 or 1);ax.set_xlabel('Report-end MD (m)',fontsize=9)
 if mw:
  ax=axs['MW'];dots(ax,mw,'#28688d');ax.xaxis.set_major_locator(MaxNLocator(nbins=5));ppg=ax.secondary_xaxis('top',functions=(lambda v:v/7.48052,lambda v:v*7.48052));ppg.xaxis.set_major_locator(MaxNLocator(nbins=5));ppg.set_xlabel('Density (ppg)',fontsize=9,labelpad=5);ppg.tick_params(labelsize=8);ax.set_xlabel('Fluid density (pcf)',fontsize=9)
 if pmin or pmax or loss:
  ax=axs['pressure / loss']
  if pmin or pmax:
   dots(ax,pmin,'#28688d');dots(ax,pmax,'#80b0ca');ax.set_xlim(left=0);ax.set_xlabel('Pressure (psi)',fontsize=9)
  else:
   ax.set_xticks([]);ax.set_xlabel('Pressure not recovered',fontsize=9,color='#9b431c')
  if loss:
   x=ax.twiny();dots(x,loss,'#c35c40');largest=max(z[0] for z in loss)
   if largest==0:x.set_xlim(-.5,.5);x.set_xticks([0])
   else:x.set_xlim(-largest*.03,largest*1.08)
   x.tick_params(axis='x',labelsize=8,colors='#a23e32');x.set_xlabel('Loss (bbl)',fontsize=9,color='#a23e32',labelpad=5)
 if complete:
  ax=axs['status']
  for d in complete:
   a=d['activity'];left=0
   for key,color in [('planned_h','#28688d'),('npt_h','#d0a153'),('unplanned_h','#c65d42'),('mixed_h','#8a65a5')]:
    ax.barh(d['day'],a[key],left=left,height=.8,color=color);left+=a[key]
  missing=sorted(set(range(1,maxday+1))-{d['day'] for d in complete})
  if missing:ax.scatter([24.5]*len(missing),missing,marker='|',s=14,color='#a9a9a9',linewidths=.5)
  ax.set_xlim(0,25);ax.set_xticks([0,12,24]);ax.set_xlabel('Daily hours',fontsize=9)
 ax=axs['field counts'];fields=('md_end','mud_weight','spp_a_min','spp_a_max','downhole_loss','pv','yp')
 for d in days:
  count=sum(number(d['fields'],key) is not None for key in fields);ax.scatter(count,d['day'],s=15,color='#4b7b74',alpha=.7,edgecolors='none')
 ax.set_xlim(-.5,len(fields)+.5);ax.set_xticks(range(0,len(fields)+1,2));ax.set_xlabel('Recognized numeric fields\n(count, not completeness)',fontsize=9)
 fig.suptitle(f'AZNS-{well} | Workover daily reporting log',x=.05,y=.985,ha='left',fontweight='bold',fontsize=15,color='#223952')
 fig.text(.05,.943,f"{len(days)} interpreted data sheets / {total} sheet pairs; {len(complete)} complete 24-hour status records.",fontsize=9,color='#9b431c')
 fig.text(.05,.918,('Omitted tracks: '+', '.join(omitted)+'. ' if omitted else '')+'Blue=pressure; orange=loss. Missing values are not zero.',fontsize=8.5,color='#263a4b')
 if complete:fig.legend(handles=[Patch(color='#28688d',label='Reported planned'),Patch(color='#d0a153',label='N / NPT'),Patch(color='#c65d42',label='UP / unplanned'),Patch(color='#8a65a5',label='NPT-UNPLN'),Patch(color='#aaa',label='Missing')],loc='upper right',bbox_to_anchor=(.98,.90),ncol=5,fontsize=8,frameon=False)
 fig.text(.05,.035,'Field counts are not data completeness; source NPT and unplanned codes remain distinct. Source: evidence/drilling_track_observations.json.',fontsize=8.5,color='#9b431c')
 fig.subplots_adjust(left=.07,right=.98,top=.84,bottom=.12)
 fig.canvas.draw();renderer=fig.canvas.get_renderer()
 for t in fig.findobj(matplotlib.text.Text):
  if not t.get_visible() or not t.get_text() or t.get_clip_on():continue
  box=t.get_window_extent(renderer)
  if box.width and box.height and (box.x0 < -2 or box.y0 < -2 or box.x1 > fig.bbox.x1+2 or box.y1 > fig.bbox.y1+2):raise ValueError((well,'clipped',t.get_text()))
 output=BytesIO();fig.savefig(output,format='png',dpi=300);plt.close(fig);return output.getvalue(),len(complete)

def main():
 for well,file in files.items():
    w=tracks[file];png,complete=chart(well,w)
    for lang in ('EN','FA'):
        fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');doc=Document(path)
        i=next(i for i,p in enumerate(doc.paragraphs) if p.text.startswith(('Figure W3','شکل W3')))
        bl=doc.paragraphs[i-1]._element.xpath('.//a:blip')[0];rid=bl.get(qn('r:embed'));doc.part.related_parts[rid]._blob=png
        sh=next(s for s in doc.inline_shapes if any(x.get(qn('r:embed'))==rid for x in s._inline.xpath('.//a:blip')));sh.height=int(sh.width*9/13)
        cap=doc.paragraphs[i];cap.runs[0].text=(f'شکل W3 ـ داده روزانه عملیات تعمیری برحسب شماره شیت، چاه {well}' if fa else f'Figure W3 — Workover daily source observations by worksheet prefix, AZNS-{well}')
        for r in cap.runs[1:]:r.text=''
        for old in list(doc.paragraphs):
            if old.text.startswith(('نمودار روزانه تعمیری فقط','The workover log uses only')):old._element.getparent().remove(old._element)
        note=doc.add_paragraph((f'نمودار روزانه تعمیری فقط از {len(w["days"])} شیت داده قابل تشخیص در میان {w["paired_sheet_candidates"]} جفت شیت و {complete} گزارش با جمع زمان طبقه‌بندی‌شده نزدیک به ۲۴ ساعت استفاده می‌کند. مقادیر نامعلوم صفر فرض نشده‌اند. «برنامه‌ای» به معنی مولد بودن قطعی نیست؛ تعداد فیلد عددی معیار کامل‌بودن داده نیست؛ کدهای NPT و غیربرنامه‌ای جدا حفظ شده‌اند و این جمع‌ها NPT کل یا پیشرفت تأییدشده نیستند. عمق پایان گزارش الزاماً عمق ابزار یا حادثه نیست.' if fa else f'The workover log uses only {len(w["days"])} recognized data sheets of {w["paired_sheet_candidates"]} sheet pairs and {complete} reports with classified durations totaling approximately 24 hours. Unknown values are not treated as zero. Planned does not necessarily mean productive; field counts do not measure data completeness; NPT and unplanned remain distinct and these are not verified total NPT or work progress. Report-end depth is not necessarily tool or incident depth.'))
        cap._p.addnext(note._p)
        if fa:
            pp=note._p.get_or_add_pPr();pp.append(OxmlElement('w:bidi'))
        for run in note.runs:
            run.font.name='Times New Roman';rp=run._element.get_or_add_rPr();ft=rp.find(qn('w:rFonts'));ft.set(qn('w:cs'),'B Nazanin' if fa else 'Times New Roman')
            if fa:rp.append(OxmlElement('w:rtl'))
        doc.save(path)
    print(well,'sheets',len(w['days']),'/',w['paired_sheet_candidates'],'complete',complete,flush=True)
if __name__=='__main__':main()
