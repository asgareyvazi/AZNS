#!/usr/bin/env python3
"""Source-labelled six-track exploratory drilling logs for 207 and 127ST2.
Tracks are daily samples at report end MD, not continuous downhole measurements.
No target operating envelope or formation/event attribution is inferred.
"""
from pathlib import Path
from io import BytesIO
import json,math,re
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
from docx import Document
from docx.oxml.ns import qn
from replace_w6_intervals import rows as casing_rows

tracks={x['file']:x for x in json.loads(Path('evidence/drilling_track_observations.json').read_text())}
formations={x['file']:x for x in json.loads(Path('evidence/formation_table_snapshots.json').read_text())}
files={'207':'AZNS207Pack.xlsx','127ST2':'AZNS127ST2Pack.xlsx','40':'AZNS40Pack.xlsx','127WO1':'AZNS127WO1Pack.xlsx'}

def number(f,key):
    x=f.get(key)
    if not x:return None
    if x['value'] is not None:return x['value'] if math.isfinite(x['value']) else None
    # Explicitly unit-suffixed MD is accepted only when the complete entry is a scalar in metres.
    if key=='md_end':
        m=re.fullmatch(r'\s*(\d+(?:\.\d+)?)\s*m\s*',x['raw'],re.I)
        if m:return float(m.group(1))
    return None

def form_tops(snapshot):
    out=[]
    for t in snapshot['tables']:
        for r in t['rows']:
            md=next((c for c in r['cells'] if c['header'].strip().lower() in ('m md','act. md (m.)')),None)
            if md is None:continue
            try:depth=float(md['value'])
            except ValueError:
                if md['value'].strip().lower()=='surface':depth=0.0
                else:continue
            if depth>=0:out.append((r['formation'],depth))
    return sorted(out,key=lambda x:x[1])

def get_data(w):
    records=[]
    for day in w['days']:
        f=day['fields'];md=number(f,'md_end')
        if md is None or not (0<=md<=6000):continue
        vals={'md':md,'day':day['day']}
        for k in ('daily_rop','wob_a_min','wob_a_max','rpm_a_min','rpm_a_max','flow_a_min','flow_a_max','spp_a_min','spp_a_max','mud_weight','pv','yp','downhole_loss'):
            v=number(f,k)
            if v is not None and v<0:v=None
            # Exclude a conspicuous 7500 gpm value from the plot pending source review;
            # the original remains in the evidence JSON with exact provenance.
            if k.startswith('flow_') and v is not None and v>2000:v=None
            if k=='mud_weight' and v is not None and not 30<=v<=200:v=None
            vals[k]=v
        records.append(vals)
    return records

def plot_pairs(ax,records,keys,colors,labels):
    for key,color,label in zip(keys,colors,labels):
        pairs=[(r[key],r['md']) for r in records if r[key] is not None]
        if pairs:
            ax.scatter([x for x,y in pairs],[y for x,y in pairs],s=14,color=color,alpha=.68,label=label,edgecolors='none')
    # Blank fields remain missing, never drawn as zero.

def plot(well,file):
    rec=get_data(tracks[file]);tops=form_tops(formations[file]);assert rec and tops
    maxdepth=max([r['md'] for r in rec]+[d for n,d in tops]);maxdepth=math.ceil((maxdepth+120)/500)*500
    mindepth=max(0,math.floor((min([r['md'] for r in rec]+[d for n,d in tops])-100)/500)*500)
    fig=plt.figure(figsize=(14.5,10),dpi=300,facecolor='white');gs=GridSpec(1,7,figure=fig,width_ratios=[1.38,1,1,1,1.15,1.1,1.25],wspace=.25)
    ax0=fig.add_subplot(gs[0]);axs=[fig.add_subplot(gs[i],sharey=ax0) for i in range(1,7)]
    palette=['#e5eef7','#f5ebdc','#e9f3e8','#f2e5e6','#edf0da','#e4e7f4']
    allaxes=[ax0]+axs
    # Bands transcribe the selected snapshot's labeled tops only; no measured contact asserted.
    for idx,(name,start) in enumerate(tops):
        stop=tops[idx+1][1] if idx+1<len(tops) else start
        if stop<=start:continue
        for a in allaxes:a.axhspan(start,stop,color=palette[idx%len(palette)],alpha=.75,zorder=0)
    # Formation labels are intentionally spaced; leader marks point to the band midpoint.
    centers=[(start+tops[i+1][1])/2 if i+1<len(tops) else start for i,(name,start) in enumerate(tops)]
    span=maxdepth-mindepth;sep=min(span/(max(len(centers),1)+3)*.85,span*.045)
    positions=[]
    for x in centers:positions.append(max(x,positions[-1]+sep if positions else mindepth+span*.025))
    if positions and positions[-1]>maxdepth-span*.02:
        positions[-1]=maxdepth-span*.02
        for i in range(len(positions)-2,-1,-1):positions[i]=min(positions[i],positions[i+1]-sep)
    ax0.set_xlim(0,1);ax0.set_xticks([]);ax0.set_ylabel('End-of-report measured depth (m)',fontsize=10)
    for (name,_),center,pos in zip(tops,centers,positions):
        ax0.annotate(name,xy=(.98,center),xytext=(.02,pos),textcoords='data',va='center',fontsize=8.7,color='#263a4b',arrowprops=dict(arrowstyle='-',color='#738797',lw=.5))
    ax0.set_ylim(maxdepth,mindepth);ax0.grid(axis='y',alpha=.15)
    for a in axs:
        a.grid(axis='y',alpha=.2);a.spines[['top','right']].set_visible(False);a.tick_params(axis='y',labelleft=False,length=0);a.tick_params(axis='x',labelsize=8)
    plot_pairs(axs[0],rec,['daily_rop'],['#206889'],['ROP'])
    for k in range(2,len(rec)):
        tri=rec[k-2:k+1]
        if all(x['daily_rop'] is not None for x in tri) and tri[2]['day']-tri[0]['day']==2:
            axs[0].scatter(sum(x['daily_rop'] for x in tri)/3,tri[-1]['md'],s=20,marker='^',color='#ca9639',alpha=.8,edgecolors='none')
    plot_pairs(axs[1],rec,['wob_a_min','wob_a_max'],['#206889','#c45d3d'],['min','max'])
    plot_pairs(axs[2],rec,['rpm_a_min','rpm_a_max'],['#206889','#c45d3d'],['min','max'])
    plot_pairs(axs[3],rec,['flow_a_min','flow_a_max'],['#206889','#53a1c2'],['flow min','flow max'])
    pressure=axs[3].twiny();plot_pairs(pressure,rec,['spp_a_min','spp_a_max'],['#b14334','#e58d80'],['SPP min','SPP max']);pressure.tick_params(axis='x',labelsize=7,colors='#a23e32');pressure.set_xlabel('SPP (psi)',fontsize=9,color='#a23e32',labelpad=6)
    plot_pairs(axs[4],rec,['mud_weight'],['#375b7d'],['weight'])
    ppg=axs[4].secondary_xaxis('top',functions=(lambda v:v/7.48052,lambda v:v*7.48052));ppg.set_xlabel('MW (ppg)',fontsize=9,labelpad=6);ppg.tick_params(labelsize=7)
    plot_pairs(axs[5],rec,['pv','yp'],['#206889','#8d5caf'],['PV','YP'])
    loss=None
    if any(r['downhole_loss'] is not None for r in rec):
        loss=axs[5].twiny();plot_pairs(loss,rec,['downhole_loss'],['#c35c40'],['loss']);loss.set_xlabel('Loss (bbl)',fontsize=9,color='#a23e32',labelpad=6);loss.tick_params(axis='x',labelsize=7,colors='#a23e32')
    else:
        axs[5].text(.04,.98,'No numeric loss data',transform=axs[5].transAxes,va='top',fontsize=7,color='#a23e32')
    labels=['ROP\n(m/h)','WOB\n(klb)','RPM (source)','Flow\n(gpm)','MW\n(pcf)','PV (cP) / YP\n(lb/100 ft²)']
    for a,l in zip(axs,labels):a.set_xlabel(l,fontsize=9,labelpad=7)
    for a in [axs[0],axs[1],axs[2],axs[3],pressure,axs[5],loss]:
        if a is not None:a.set_xlim(left=0)
    # Legacy shoe markers are omitted until a common datum and operation
    # stage have been reconciled with the source casing records.
    fig.text(.055,.916,'WOB/RPM: blue=lower or single, orange=upper. Flow: blue=lower or single, light blue=upper; SPP: red/pink.',fontsize=8.3,color='#263a4b')
    fig.text(.055,.893,'ROP: blue=daily, gold triangle=mean of 3 consecutive reports. PV=blue, YP=purple, loss=orange; casing shoes omitted pending reconciliation.',fontsize=8.3,color='#263a4b')
    fig.suptitle(f'AZNS-{well} | Daily drilling parameters at report-end MD',fontsize=17,fontweight='bold',x=.055,y=.985,ha='left',color='#223952')
    fig.text(.055,.944,f"{len(rec)} plotted reports / {tracks[file]['paired_sheet_candidates']} sheet pairs; displayed depth range covers observations and tops; no verified formation attribution.",fontsize=9,color='#9b431c')
    fig.text(.055,.035,'DDR sources and rejected/non-numeric values: evidence/drilling_track_observations.json. Bands stop at the last listed top; no lower contact assumed.',fontsize=8.5,color='#9b431c')
    fig.subplots_adjust(left=.08,right=.98,top=.84,bottom=.115)
    fig.canvas.draw();out=BytesIO();fig.savefig(out,format='png',dpi=300);plt.close(fig);return out.getvalue()

def main():
    for well,file in files.items():
        image=plot(well,file)
        for lang in ('EN','FA'):
            path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
            i=next(i for i,p in enumerate(d.paragraphs) if p.text.startswith(('Figure W3','شکل W3')))
            bl=d.paragraphs[i-1]._element.xpath('.//a:blip')[0];rid=bl.get(qn('r:embed'));d.part.related_parts[rid]._blob=image
            sh=next(s for s in d.inline_shapes if any(x.get(qn('r:embed'))==rid for x in s._inline.xpath('.//a:blip')));sh.height=int(sh.width*10/14.5)
            cap=d.paragraphs[i];cap.runs[0].text=(f'شکل W3 ـ پارامترهای روزانه در عمق پایان گزارش، چاه {well}؛ انتساب سازند تأیید نشده است' if lang=='FA' else f'Figure W3 — Daily drilling parameters at report-end depth, AZNS-{well}; formation attribution unverified')
            for r in cap.runs[1:]:r.text=''
            for old in list(d.paragraphs):
                if old.text.startswith(('مقادیر نقاط از داده روزانه','Points combine daily entries with report-end depth')):old._element.getparent().remove(old._element)
            note=d.add_paragraph(('مقادیر نقاط از داده روزانه و عمق پایان گزارش گرفته شده‌اند؛ عمق حادثه یا ابزار و توالی پیوسته را نشان نمی‌دهند. بازه‌های کمینه/بیشینه از ستون نخست یا بازه صریح منبع‌اند، نه پنجره بهینه؛ RPM در قالب قدیمی ممکن است شامل دور موتور باشد. مثلث‌های ROP میانگین سه گزارش متوالی هستند. رأس سازند و کفشک‌ها به تطبیق نیاز دارند. مقدار مشکوک ۷۵۰۰ gpm چاه 207 از نمودار حذف شده اما در داده منبع محفوظ است.' if lang=='FA' else 'Points combine daily entries with report-end depth; they do not establish event/tool depth or a continuous log. Points use the first bit column or explicit source range, not an optimum envelope; scalar entries are not necessarily minima, and legacy RPM may include motor RPM; ROP triangles average three consecutive reports. Formation tops and shoe markers require reconciliation. The suspect 7,500 gpm entry in well 207 is excluded from the plot but retained in the source evidence.'))
            cap._p.addnext(note._p)
            for run in note.runs:
                run.font.name='Times New Roman';pr=run._element.get_or_add_rPr();rfonts=pr.find(qn('w:rFonts'));rfonts.set(qn('w:cs'),'B Nazanin' if lang=='FA' else 'Times New Roman')
                if lang=='FA':pr.append(__import__('docx').oxml.OxmlElement('w:rtl'))
            d.save(path)
        print(well,len(get_data(tracks[file])),'plotted days')
if __name__=='__main__':main()
