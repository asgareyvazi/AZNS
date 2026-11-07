#!/usr/bin/env python3
"""Replace backwards W4 Pareto charts using the report's displayed cause table.
This is a presentation fix, not a validation of the original NPT classifications.
"""
from io import BytesIO
from pathlib import Path
import textwrap
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document


def get_rows(doc):
    table=next(t for t in doc.tables if t.rows and t.rows[0].cells[0].text=='Technical root cause')
    result=[]
    for row in table.rows[1:]:
        name=row.cells[0].text.strip()
        value=float(row.cells[1].text.strip().replace(',',''))
        if value>0: result.append((name,value))
    if not result: raise ValueError('No positive cause entries')
    return sorted(result,key=lambda x:-x[1])

def pareto(rows):
    total=sum(v for _,v in rows); acc=0;out=[]
    for name,value in rows:
        acc+=value;out.append((name,value,100*acc/total))
    return out

def make_png(well, rows):
    items=pareto(rows);n=len(items)
    fig,ax=plt.subplots(figsize=(11,5.7),dpi=300)
    fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white')
    labels=['\n'.join(textwrap.wrap(name,31,break_long_words=False,break_on_hyphens=False)) for name,_,_ in items]
    values=[v for _,v,_ in items];ys=list(range(n));maxval=max(values)
    ax.barh(ys,values,height=.25 if n==1 else .65,color='#28688d')
    ax.set_yticks(ys,labels,fontsize=9);ax.invert_yaxis()
    if n==1:ax.set_ylim(.65,-.65)
    ax.set_xlim(0,maxval*1.62);ax.set_xlabel('Reported NPT hours',fontsize=10)
    ax.grid(axis='x',alpha=.2);ax.set_axisbelow(True)
    ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0)
    for i,(_,v,pct) in enumerate(items):
        ax.text(v+maxval*.025,i,f'{v:,.0f} h  |  {pct:.1f}%',va='center',fontsize=9,color='#24354a')
    fig.suptitle(f'AZNS-{well} | NPT cause Pareto',x=.05,y=.98,ha='left',fontsize=15,fontweight='bold')
    fig.text(.05,.92,'Cumulative percentage runs from the largest listed cause downward.',fontsize=9,color='#9b431c')
    fig.text(.05,.035,'Basis: displayed, rounded report table 7-1 only; omitted causes and original classifications need review.',fontsize=8.8,color='#9b431c')
    fig.subplots_adjust(left=.38,right=.97,top=.84,bottom=.17)
    # Guard against clipping of any axis label or value at the figure boundary.
    fig.canvas.draw(); renderer=fig.canvas.get_renderer(); bounds=fig.bbox
    for artist in [*ax.get_yticklabels(),*ax.texts]:
        b=artist.get_window_extent(renderer)
        if b.x0 < bounds.x0 or b.x1 > bounds.x1 or b.y0 < bounds.y0 or b.y1 > bounds.y1:
            raise ValueError(f'Clipped text for {well}: {artist.get_text()}')
    out=BytesIO();fig.savefig(out,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return out.getvalue()

def set_text(p,text):
    p.runs[0].text=text
    for r in p.runs[1:]:r.text=''

def main():
    files=sorted(p for p in Path('.').glob('*_Drilling_Engineering_Report_EN.docx') if p.name[0].isdigit())
    assert len(files)==13
    for path in files:
        well=path.name.split('_')[0];rows=get_rows(Document(path));image=make_png(well,rows)
        for language in ('EN','FA'):
            target=Path(f'{well}_Drilling_Engineering_Report_{language}.docx');doc=Document(target)
            captions=[i for i,p in enumerate(doc.paragraphs) if p.text.startswith(('Figure W4','شکل W4'))]
            if len(captions)!=1:raise ValueError((str(target),captions))
            i=captions[0];blips=doc.paragraphs[i-1]._element.xpath('.//a:blip')
            if len(blips)!=1:raise ValueError('W4 image not found in '+str(target))
            rid=blips[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            doc.part.related_parts[rid]._blob=image
            shape=next(s for s in doc.inline_shapes if any(b.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')==rid for b in s._inline.xpath('.//a:blip')))
            shape.height=int(shape.width*5.7/11)
            set_text(doc.paragraphs[i],f'شکل W4 ـ پارتوی علل ثبت‌شده در جدول ۷-۱، چاه AZNS-{well}' if language=='FA' else f'Figure W4 — Pareto of causes listed in table 7-1, AZNS-{well}')
            doc.save(target)
        print(well,len(rows),'categories',flush=True)
if __name__=='__main__':main()
