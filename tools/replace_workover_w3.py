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
 fig,axs=plt.subplots(1,5,figsize=(13,9),dpi=300,sharey=True,gridspec_kw={'wspace':.28});fig.patch.set_facecolor('#fff')
 for ax in axs:
    ax.set_ylim(maxday+3,0);ax.grid(axis='y',alpha=.18);ax.spines[['top','right']].set_visible(False);ax.tick_params(axis='x',labelsize=8);ax.tick_params(axis='y',labelsize=8)
 for ax in axs[1:]:ax.tick_params(axis='y',labelleft=False,length=0)
 axs[0].set_ylabel('Report day (not verified rig days)',fontsize=10)
 def empty(ax,message):
    ax.set_xlim(0,1);ax.set_xticks([]);ax.text(.5,.5,message,transform=ax.transAxes,ha='center',va='center',fontsize=8,color='#9b431c',wrap=True)
 md=series(days,'md_end',lambda v:v<=6000)
 if md:dots(axs[0],md,'#28688d');axs[0].set_xlim(left=0)
 else:empty(axs[0],'No numeric\nend-depth data')
 axs[0].set_xlabel('Report-end MD (m)',fontsize=9)
 mw=series(days,'mud_weight',lambda v:30<=v<=200)
 if mw:
    dots(axs[1],mw,'#28688d');ppg=axs[1].secondary_xaxis('top',functions=(lambda v:v/7.48052,lambda v:v*7.48052));ppg.set_xlabel('Density (ppg)',fontsize=9,labelpad=5);ppg.tick_params(labelsize=8)
 else:empty(axs[1],'No usable\ndensity data')
 axs[1].set_xlabel('Fluid density (pcf)',fontsize=9)
 pmin=series(days,'spp_a_min',lambda v:v<=15000);pmax=series(days,'spp_a_max',lambda v:v<=15000);loss=series(days,'downhole_loss',lambda v:v<=50000)
 if pmin or pmax:
    dots(axs[2],pmin,'#28688d');dots(axs[2],pmax,'#80b0ca');axs[2].set_xlim(left=0)
 elif not loss:empty(axs[2],'No numeric\npressure or loss')
 else:axs[2].set_xticks([]);axs[2].text(.5,.5,'No numeric\npressure',transform=axs[2].transAxes,ha='center',fontsize=8,color='#9b431c')
 if loss:
    a=axs[2].twiny();dots(a,loss,'#c35c40');a.set_xlim(left=0);a.tick_params(axis='x',labelsize=8,colors='#a23e32');a.set_xlabel('Loss (bbl)',fontsize=9,color='#a23e32',labelpad=5)
 axs[2].set_xlabel('Pressure (psi)',fontsize=9)
 if complete:
    for d in complete:
        a=d['activity'];left=0
        for key,color in [('planned_h','#28688d'),('npt_h','#d0a153'),('unplanned_h','#c65d42'),('mixed_h','#8a65a5')]:
            axs[3].barh(d['day'],a[key],left=left,height=.8,color=color);left+=a[key]
    missing=sorted(set(range(1,maxday+1))-{d['day'] for d in complete})
    if missing:axs[3].scatter([24.5]*len(missing),missing,marker='|',s=14,color='#a9a9a9',linewidths=.5)
    axs[3].set_xlim(0,25);axs[3].set_xticks([0,12,24]);
 else:empty(axs[3],'No complete classified\n24-hour records')
 axs[3].set_xlabel('Daily hours',fontsize=9)
 for key,color in [('npt_h','#d0a153'),('unplanned_h','#c65d42'),('mixed_h','#8a65a5')]:
    running=0;points=[]
    for d in sorted(complete,key=lambda x:x['day']):
        running+=d['activity'][key];points.append((running,d['day']))
    if points:dots(axs[4],points,color)
 if complete:axs[4].set_xlim(left=0)
 else:empty(axs[4],'No complete\nclassified records')
 axs[4].set_xlabel('Cumulative source-code\nhours (partial)',fontsize=9)
 fig.suptitle(f'AZNS-{well} | Workover daily reporting log',x=.05,y=.985,ha='left',fontweight='bold',fontsize=15,color='#223952')
 fig.text(.05,.943,f"{len(days)} interpreted data sheets / {total} sheet pairs; {len(complete)} complete 24-hour status records.",fontsize=9,color='#9b431c')
 fig.text(.05,.918,'Pressure blue=lower/single, pale blue=upper; loss orange. Missing/ambiguous values are not plotted.',fontsize=8.5,color='#263a4b')
 fig.legend(handles=[Patch(color='#28688d',label='Reported planned'),Patch(color='#d0a153',label='N / NPT'),Patch(color='#c65d42',label='UP / unplanned'),Patch(color='#8a65a5',label='NPT-UNPLN'),Patch(color='#aaa',label='Missing')],loc='upper right',bbox_to_anchor=(.98,.90),ncol=5,fontsize=8,frameon=False)
 fig.text(.05,.035,'Cumulative series use complete classified days only; source NPT and unplanned codes remain distinct. Source: evidence/drilling_track_observations.json.',fontsize=8.5,color='#9b431c')
 fig.subplots_adjust(left=.07,right=.98,top=.84,bottom=.12)
 output=BytesIO();fig.savefig(output,format='png',dpi=300);plt.close(fig);return output.getvalue(),len(complete)

def main():
 for well,file in files.items():
    w=tracks[file];png,complete=chart(well,w)
    for lang in ('EN','FA'):
        fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');doc=Document(path)
        i=next(i for i,p in enumerate(doc.paragraphs) if p.text.startswith(('Figure W3','شکل W3')))
        bl=doc.paragraphs[i-1]._element.xpath('.//a:blip')[0];rid=bl.get(qn('r:embed'));doc.part.related_parts[rid]._blob=png
        sh=next(s for s in doc.inline_shapes if any(x.get(qn('r:embed'))==rid for x in s._inline.xpath('.//a:blip')));sh.height=int(sh.width*9/13)
        cap=doc.paragraphs[i];cap.runs[0].text=(f'شکل W3 ـ داده روزانه عملیات تعمیری برحسب شماره گزارش، چاه {well}' if fa else f'Figure W3 — Workover daily source observations by report day, AZNS-{well}')
        for r in cap.runs[1:]:r.text=''
        for old in list(doc.paragraphs):
            if old.text.startswith(('نمودار روزانه تعمیری فقط','The workover log uses only')):old._element.getparent().remove(old._element)
        note=doc.add_paragraph((f'نمودار روزانه تعمیری فقط از {len(w["days"])} شیت داده قابل تشخیص در میان {w["paired_sheet_candidates"]} جفت شیت و {complete} گزارش با جمع زمان طبقه‌بندی‌شده نزدیک به ۲۴ ساعت استفاده می‌کند. مقادیر نامعلوم صفر فرض نشده‌اند. «برنامه‌ای» به معنی مولد بودن قطعی نیست؛ جمع تجمعی هر کد فقط شامل روزهای قابل طبقه‌بندی است؛ کدهای NPT و غیربرنامه‌ای جدا حفظ شده‌اند و این جمع‌ها NPT کل یا پیشرفت تأییدشده نیستند. عمق پایان گزارش الزاماً عمق ابزار یا حادثه نیست.' if fa else f'The workover log uses only {len(w["days"])} recognized data sheets of {w["paired_sheet_candidates"]} sheet pairs and {complete} reports with classified durations totaling approximately 24 hours. Unknown values are not treated as zero. Planned does not necessarily mean productive; cumulative code totals include only complete classified days; NPT and unplanned remain distinct and these are not verified total NPT or work progress. Report-end depth is not necessarily tool or incident depth.'))
        cap._p.addnext(note._p)
        if fa:
            pp=note._p.get_or_add_pPr();pp.append(OxmlElement('w:bidi'))
        for run in note.runs:
            run.font.name='Times New Roman';rp=run._element.get_or_add_rPr();ft=rp.find(qn('w:rFonts'));ft.set(qn('w:cs'),'B Nazanin' if fa else 'Times New Roman')
            if fa:rp.append(OxmlElement('w:rtl'))
        doc.save(path)
    print(well,'sheets',len(w['days']),'/',w['paired_sheet_candidates'],'complete',complete,flush=True)
if __name__=='__main__':main()
