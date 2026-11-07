#!/usr/bin/env python3
"""Historical component-depth chart; it is deliberately not a barrier schematic."""
import json
from pathlib import Path
from io import BytesIO
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Cm
from replace_well_bit_evidence import style

j=json.loads(Path('evidence/122wo1_campaign_review.json').read_text());v=j['completion_depths'];assert len(v)==9
fig,ax=plt.subplots(figsize=(10.5,6.0),dpi=300);fig.patch.set_facecolor('#f8fafc');ax.set_facecolor('white')
y=list(range(len(v)));depth=[float(r['source_depth_text']) for r in v]
ax.hlines(y,0,depth,color='#b1c9d8',lw=3,zorder=1);ax.scatter(depth,y,s=55,color='#216688',edgecolor='white',linewidth=.7,zorder=3)
ax.set_yticks(y,[r['english'] for r in v],fontsize=9);ax.invert_yaxis();ax.set_xlim(0,4400);ax.set_xlabel('Reported depth (m; datum requires reconciliation)',fontsize=10);ax.grid(axis='x',alpha=.2);ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0)
for i,r in enumerate(v):ax.text(float(r['source_depth_text'])+48,i,r['source_depth_text']+' m',fontsize=8.5,va='center',color='#244255')
fig.suptitle('AZNS-122 WO1 | Completion components after packer setting',x=.06,y=.97,ha='left',fontsize=15,fontweight='bold',color='#223952')
fig.text(.06,.91,'Selected historical depths; not an as-built, current barrier or wellbore schematic.',fontsize=9,color='#994a22')
fig.text(.06,.04,'Source: AZNS122WO1Pack.xlsx, 54-1 Remark!K39:K40. Depth labels and component names as reported.',fontsize=8.4,color='#994a22')
fig.subplots_adjust(left=.25,right=.96,top=.82,bottom=.16)
b=BytesIO();fig.savefig(b,format='png',dpi=300);plt.close(fig);png=b.getvalue();Path('evidence/well_figures/122WO1_completion_depths.png').write_bytes(png)
for lang in ('EN','FA'):
 fa=lang=='FA';f=Path(f'122WO1_Drilling_Engineering_Report_{lang}.docx');d=Document(f);caption='شکل ـ عمق‌های منتخب اجزای تکمیل پس از نصب پکر؛ نه شماتیک مانع' if fa else 'Figure — Selected completion-component depths after setting the packer; not a barrier schematic'
 if any(p.text==caption for p in d.paragraphs):continue
 prev=next(p for p in d.paragraphs if p.text.startswith(('Table — Selected completion depths after packer setting','جدول ـ عمق‌های منتخب تکمیل پس از نصب پکر')))
 p=d.add_paragraph();p.alignment=1;p.add_run().add_picture(BytesIO(png),width=Cm(16));prev._p.addnext(p._p)
 c=d.add_paragraph(caption);style(c,fa,True);p._p.addnext(c._p);d.save(f)
print('122WO1 completion depth chart',len(png))
