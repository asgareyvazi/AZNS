#!/usr/bin/env python3
"""Withdraw false daily cumulative cost lines and show legacy category estimates.

Charts are a provisional transcription of each report's EN table 8-2, not a
fresh CBS reconstruction. Run after the interim safety revision. Requires
python-docx and matplotlib; outputs 300 dpi English charts in both languages.
"""
from pathlib import Path
from io import BytesIO
import re, textwrap
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document

root=Path('.')

def chart(well, rows):
    rows=sorted(rows,key=lambda r:r[1],reverse=True)
    fig,ax=plt.subplots(figsize=(11,5.7),dpi=300)
    fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white')
    labels=['\n'.join(textwrap.wrap(k,width=32,break_long_words=False,break_on_hyphens=False)) for k,v in rows]
    vals=[v/1e6 for k,v in rows]; y=list(range(len(vals)))
    ax.barh(y,vals,color='#28688d',height=.25 if len(vals)==1 else .68)
    if len(vals)==1: ax.set_ylim(.65,-.65)
    ax.set_yticks(y,labels,fontsize=9);ax.invert_yaxis();ax.set_xlabel('Reported category estimate (USD million)',fontsize=10)
    xmax=max(vals) if vals else 1;ax.set_xlim(0,max(xmax*1.30,.1));ax.grid(axis='x',alpha=.22);ax.set_axisbelow(True)
    for i,v in enumerate(vals): ax.text(v+xmax*.02,i,f'{v:,.3f}',va='center',fontsize=9,color='#203047')
    ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0)
    fig.suptitle(f'AZNS-{well} | Reported cost categories',fontweight='bold',fontsize=15,x=.06,ha='left',y=.98)
    fig.text(.06,.915,'Up to eight categories from legacy table 8-2; CBS amounts and scope are not yet reconciled.',fontsize=9,color='#9b431c')
    fig.text(.06,.035,'No daily cost series shown. Category estimates are not a validated actual expenditure timeline.',fontsize=9,color='#9b431c')
    fig.subplots_adjust(left=.32,right=.94,top=.83,bottom=.17)
    out=BytesIO();fig.savefig(out,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return out.getvalue()

def set_text(p,text):
    if p.runs:
        p.runs[0].text=text
        for r in p.runs[1:]: r.text=''
    else: p.add_run(text)

def data(d):
    for t in d.tables:
        if not t.rows or not t.rows[0].cells: continue
        if t.rows[0].cells[0].text.strip()=='CBS cost category':
            out=[]
            for row in t.rows[1:]:
                cells=row.cells
                try: value=float(cells[1].text.strip().replace(',','').replace('$',''))
                except ValueError: continue
                if value>0: out.append((cells[0].text.strip(),value))
            if out: return out
    raise ValueError('Category table not found')

files=sorted(p for p in root.glob('*_Drilling_Engineering_Report_EN.docx') if p.name[0].isdigit())
assert len(files)==13,len(files)
for en in files:
    well=en.name.split('_')[0];d=Document(en);rows=data(d);png=chart(well,rows)
    for language in ('EN','FA'):
        path=root/(well+'_Drilling_Engineering_Report_'+language+'.docx');doc=Document(path);fa=language=='FA'
        captions=[i for i,p in enumerate(doc.paragraphs) if p.text.startswith(('Figure W5', 'شکل W5'))]
        if len(captions)!=1: raise ValueError((path,captions))
        i=captions[0]; caption=doc.paragraphs[i];previous=doc.paragraphs[i-1]
        blips=previous._element.xpath('.//a:blip')
        if len(blips)!=1: raise ValueError(f'{path}: chart W5 image not found before caption')
        rid=blips[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        doc.part.related_parts[rid]._blob=png
        set_text(caption, ('شکل W5 ـ برآورد اولیه طبقات هزینه مندرج در جدول ۸-۲، چاه AZNS-'+well if fa else 'Figure W5 — Provisional reported cost categories (table 8-2), AZNS-'+well))
        for p in doc.paragraphs:
            if p.text.startswith('The financial position of the well is reconstructed from two sources:'):
                set_text(p,'The legacy report combines a rig day-rate estimate with CBS service categories. The category estimates shown in table 8-2 and Figure W5 remain provisional pending source and scope reconciliation. The former straight-line cumulative-cost curve was removed: no actual daily expenditure chronology has been established. A credible timeline requires dated invoices, service tickets or a clearly identified, justified allocation model.')
            elif p.text.startswith('وضعیت مالی چاه از دو منبع بازسازی شده است:'):
                set_text(p,'در برآورد قبلی، نرخ روزانه دکل و طبقات خدمات CBS با یکدیگر ترکیب شده‌اند. ارقام جدول ۸-۲ و شکل W5 تا تطبیق منبع و دامنه هزینه‌ها صرفاً برآورد اولیه‌اند. منحنی خطی هزینه تجمعی حذف شد، زیرا توالی واقعی مخارج روزانه احراز نشده است. برای بازسازی معتبر، صورت‌حساب‌ها یا برگه‌های خدمات تاریخ‌دار، یا یک مدل تخصیص با فرض‌های شفاف و قابل دفاع لازم است.')
            elif p.text=='CBS Financial Analysis and Reconstructed Cumulative Cost': set_text(p,'CBS Financial Analysis and Provisional Cost Categories')
            elif p.text=='تحلیل مالی CBS و بازسازی هزینه تجمعی': set_text(p,'تحلیل مالی CBS و برآورد اولیه طبقات هزینه')
        doc.save(path)
    print(well,len(rows),'categories',flush=True)
