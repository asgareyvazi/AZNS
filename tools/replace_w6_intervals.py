#!/usr/bin/env python3
"""Replace W6 with a source-table interval plot, not a wellbore schematic.
This avoids empty panels and invented casing nesting. Legacy table values remain
unverified; malformed/ambiguous intervals are excluded, not silently repaired.
"""
from pathlib import Path
from io import BytesIO
import re,textwrap
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document


def depth(s):
    s=s.strip()
    if s.lower()=='surface':return 0.0
    if not re.fullmatch(r'\d+(?:\.\d+)?',s):return None
    return float(s)

def rows(table):
    out=[];excluded=[]
    for r in table.rows[1:]:
        c=r.cells
        names=[s.strip() for s in c[0].text.splitlines() if s.strip()]
        if not names:continue
        if any(re.search(r'rigging|vertical drilling',x,re.I) for x in names):
            excluded.extend(names);continue
        tops=[s.strip() for s in c[4].text.splitlines() if s.strip()]
        shoes=[s.strip() for s in c[5].text.splitlines() if s.strip()]
        if len(names)!=len(tops) or len(names)!=len(shoes):
            excluded.extend(names);continue
        for name,ts,ss in zip(names,tops,shoes):
            a,b=depth(ts),depth(ss)
            if not re.match(r'^(?:New\s+)?\d+(?:-\d+/\d+)?\s*(?:"|\'\')',name,re.I) or re.search(r'\b(?:pkr|packer)\b',name,re.I) or a is None or b is None or a>=b or b>10000:
                excluded.append(name);continue
            out.append((name,a,b))
    return out,excluded

def chart(well,intervals,excluded):
    fig,ax=plt.subplots(figsize=(11,5.7),dpi=300);fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white')
    if intervals:
        ymax=max(x[2] for x in intervals);n=len(intervals)
        for i,(name,top,shoe) in enumerate(intervals):
            ax.vlines(i,top,shoe,color='#28688d',linewidth=8,zorder=2)
            ax.scatter([i],[top],marker='o',s=22,color='#216c58',zorder=3)
            ax.scatter([i],[shoe],marker='_',s=130,linewidth=2,color='#a84531',zorder=3)
        labels=['\n'.join(textwrap.wrap(name,16,break_long_words=False,break_on_hyphens=False)) for name,_,_ in intervals]
        ax.set_xticks(range(n),labels,fontsize=7.5,rotation=35,ha='right',rotation_mode='anchor');ax.tick_params(axis='x',length=0,pad=7)
        ax.set_xlim(-.7,n-.3);ax.set_ylim(ymax*1.07,-ymax*.04);ax.set_ylabel('Reported depth (m); datum unverified')
        ax.grid(axis='y',alpha=.2);ax.set_axisbelow(True)
        ax.spines[['top','right']].set_visible(False)
        ax.plot([],[],marker='o',linestyle='None',color='#216c58',label='Reported top')
        ax.plot([],[],marker='_',linestyle='None',markersize=12,color='#a84531',label='Reported shoe')
        handles,legend_labels=ax.get_legend_handles_labels()
        fig.legend(handles,legend_labels,loc='upper right',bbox_to_anchor=(.97,.89),ncol=2,fontsize=8,frameon=False)
    else:
        ax.axis('off');ax.text(.5,.5,'No unambiguous top / shoe pairs in the legacy table.\nConsult the source completion records.',ha='center',va='center',transform=ax.transAxes,fontsize=12)
    fig.suptitle(f'AZNS-{well} | Reported casing intervals',x=.05,y=.98,ha='left',fontsize=15,fontweight='bold')
    fig.text(.05,.92,'Source: legacy report table 6-1; not a validated as-built wellbore or completion schematic.',fontsize=9,color='#9b431c')
    fig.text(.05,.028,f'Surface plotted as zero; {len(excluded)} ambiguous or non-casing table item(s) excluded. No tubing, geometry or cement inferred.',fontsize=8.6,color='#9b431c')
    fig.subplots_adjust(left=.10,right=.97,top=.82,bottom=.33);fig.canvas.draw();renderer=fig.canvas.get_renderer()
    for artist in ax.get_xticklabels():
        b=artist.get_window_extent(renderer)
        if b.x0<0 or b.x1>fig.bbox.x1 or b.y0<0 or b.y1>fig.bbox.y1:raise ValueError(f'{well}: clipped {artist.get_text()}')
    b=BytesIO();fig.savefig(b,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return b.getvalue()

def main():
    files=sorted(p for p in Path('.').glob('*_Drilling_Engineering_Report_EN.docx') if p.name[0].isdigit());assert len(files)==13
    for p in files:
        well=p.name.split('_')[0];d=Document(p);t=next(t for t in d.tables if t.rows and t.rows[0].cells[0].text=='Casing / liner')
        data,excluded=rows(t);image=chart(well,data,excluded)
        for lang in ('EN','FA'):
            f=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');doc=Document(f)
            cap=next((i for i,p in enumerate(doc.paragraphs) if p.text.startswith(('Figure W6','شکل W6'))),None)
            if cap is None:raise ValueError(f'No W6: {f}')
            b=doc.paragraphs[cap-1]._element.xpath('.//a:blip')[0];rid=b.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            doc.part.related_parts[rid]._blob=image
            shape=next(s for s in doc.inline_shapes if any(x.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')==rid for x in s._inline.xpath('.//a:blip')))
            shape.height=int(shape.width*5.7/11)
            caption=doc.paragraphs[cap];caption.runs[0].text=(f'شکل W6 ـ بازه‌های گزارش‌شده جداری در جدول ۶-۱، چاه AZNS-{well} (شماتیک تأییدشده نیست)' if lang=='FA' else f'Figure W6 — Reported casing intervals from table 6-1, AZNS-{well} (not an as-built schematic)')
            for r in caption.runs[1:]:r.text=''
            # These rows are phase labels accidentally included in the casing table.
            if well in ('10','85'):
                table=next(t for t in doc.tables if t.rows and t.rows[0].cells[0].text in ('Casing / liner','جداری / آستری'))
                for row in list(table.rows)[1:]:
                    if re.search(r'rigging|vertical drilling',row.cells[0].text,re.I):table._tbl.remove(row._tr)
            doc.save(f)
        print(well,len(data),'intervals;',len(excluded),'excluded',flush=True)
if __name__=='__main__':main()
