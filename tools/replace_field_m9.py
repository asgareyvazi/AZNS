#!/usr/bin/env python3
"""Redraw M9 as a labelled provisional schedule/cost scatter.
Missing budgets are excluded, not treated as zero. Requires adjustText.
"""
from io import BytesIO
import re
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from adjustText import adjust_text
from docx import Document
from docx.oxml.ns import qn

master=Document('Field_Master_Drilling_Engineering_Report_EN.docx')
table=next(t for t in master.tables if t.rows and t.rows[0].cells[0].text=='Operation' and len(t.rows[0].cells)==8)
pts=[];missing=[]
for row in table.rows[1:]:
    c=row.cells;name=c[0].text.replace('AZNS-','').replace(' (F-20)','')
    if name=='Portfolio':continue
    try:actual=float(c[2].text);planned=float(c[3].text);spend=float(c[5].text);variance=float(c[7].text.replace('%',''))
    except ValueError:missing.append(name);continue
    assert planned>0;pts.append((name,actual/planned,variance,spend))
assert len(pts)==11 and set(missing)=={'122 WO1','85'}
np.random.seed(42)
fig,ax=plt.subplots(figsize=(11,6.8),dpi=300);fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white')
ax.axhline(0,color='#8494a2',lw=.9);ax.axvline(1,color='#8494a2',lw=.9)
x=[v[1] for v in pts];y=[v[2] for v in pts];sizes=[80+v[3]*30 for v in pts]
ax.scatter(x,y,s=sizes,color='#4b88a7',edgecolor='#23546b',alpha=.6,zorder=2)
ax.set_xlim(.38,2.85);ax.set_ylim(-82,78);ax.grid(alpha=.18);ax.set_axisbelow(True)
ax.set_xlabel('Reported actual days / planned days');ax.set_ylabel('Legacy budget variance (%)')
texts=[ax.text(px,py,name,color='#1e2d3d',fontsize=9,fontweight='bold',zorder=3) for name,px,py,_ in pts]
adjust_text(texts,x=x,y=y,ax=ax,iter_lim=1000,force_text=(.55,.75),force_static=(.35,.55),expand=(1.18,1.35),arrowprops=dict(arrowstyle='-',color='#536778',lw=.7))
fig.suptitle('AZNS | Schedule and cost comparison',x=.06,y=.98,ha='left',fontsize=15,fontweight='bold')
fig.text(.06,.935,'Provisional legacy values; bubble area scales with reported spend, not risk severity.',fontsize=9,color='#9b431c')
fig.text(.06,.03,'Excluded: 122 WO1 and 85 (budget not reconciled). Ratios and variance are not independently validated.',fontsize=8.8,color='#9b431c')
fig.subplots_adjust(left=.11,right=.96,top=.86,bottom=.14);fig.canvas.draw()
renderer=fig.canvas.get_renderer();boxes=[t.get_window_extent(renderer).expanded(1.01,1.01) for t in texts]
for i,b in enumerate(boxes):
 if b.x0<fig.bbox.x0 or b.y0<fig.bbox.y0 or b.x1>fig.bbox.x1 or b.y1>fig.bbox.y1:raise ValueError('Clipped label')
 if any(b.overlaps(other) for other in boxes[:i]):raise ValueError('Overlapping labels')
out=BytesIO();fig.savefig(out,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);png=out.getvalue()
for kind in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
    f=f'Field_{kind}_Report_{lang}.docx';d=Document(f);i=next(i for i,p in enumerate(d.paragraphs) if re.match(r'^(?:Figure|شکل) M9\b',p.text))
    b=d.paragraphs[i-1]._element.xpath('.//a:blip')[0];rid=b.get(qn('r:embed'));d.part.related_parts[rid]._blob=png
    sh=next(s for s in d.inline_shapes if any(x.get(qn('r:embed'))==rid for x in s._inline.xpath('.//a:blip')));sh.height=int(sh.width*6.8/11)
    p=d.paragraphs[i];p.runs[0].text='شکل M9 ـ مقایسه موقت زمان و هزینه در گزارش قبلی' if lang=='FA' else 'Figure M9 — Provisional schedule and cost comparison'
    for r in p.runs[1:]:r.text=''
    d.save(f);print(f)
