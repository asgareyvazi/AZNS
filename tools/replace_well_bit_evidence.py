#!/usr/bin/env python3
"""Replace legacy bit-run tables/W7 with source-cited preliminary observations."""
from pathlib import Path
from io import BytesIO
from collections import OrderedDict
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm,Pt

DATA={w['well']:w for w in json.loads(Path('evidence/well_bit_observations.json').read_text())}

def chart(w):
    selected=sorted([g for g in w['groups'] if g['maximum_cumulative_m'] and g['maximum_cumulative_m']['value']>0],key=lambda g:g['maximum_cumulative_m']['value'],reverse=True)[:10]
    fig,ax=plt.subplots(figsize=(11,6),dpi=300);fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white')
    if selected:
        labels=[f"Bit {g['bit_no']} / rerun {g['rerun'] if g['rerun'] is not None else 'unknown'}" for g in selected]
        vals=[g['maximum_cumulative_m']['value'] for g in selected];ax.barh(range(len(vals)),vals,color='#28688d',height=.65);ax.set_yticks(range(len(vals)),labels,fontsize=9);ax.invert_yaxis();mx=max(vals);ax.set_xlim(0,mx*1.25)
        for i,x in enumerate(vals):ax.text(x+mx*.02,i,f'{x:,.1f}',va='center',fontsize=9)
        ax.set_xlabel('Maximum recorded cumulative counter (m)',fontsize=10)
        subtitle='Up to 10 largest observed counters; counters are not independent metres drilled per rerun.'
    else:
        labels=['Recognized data sheets','Numbered bit-column entries','Unassigned tool-column entries']
        vals=[w['recognized_data_sheets'],len(w['observations']),len(w['unassigned_observations'])]
        ax.barh(range(3),vals,color=['#28688d','#6d9db6','#d0a153'],height=.45);ax.set_yticks(range(3),labels,fontsize=9);ax.invert_yaxis();mx=max(vals,default=1);ax.set_xlim(0,max(mx*1.28,1))
        for i,x in enumerate(vals):ax.text(x+max(mx*.02,.1),i,str(x),va='center',fontsize=9)
        ax.set_xlabel('Sheets or column entries, as labelled',fontsize=10)
        subtitle='No positive numerical cumulative-metre counter for an identified bit was recovered in these sheets.'
    ax.grid(axis='x',alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0)
    fig.suptitle(f"AZNS-{w['well']} | Bit and tool evidence",x=.05,y=.98,ha='left',fontsize=15,fontweight='bold',color='#223952')
    fig.text(.05,.91,subtitle,fontsize=8.8,color='#9b431c')
    fig.text(.05,.055,f"Source: {w['source']} | {w['recognized_data_sheets']} recognized data sheets / {w['candidate_pairs']} candidate pairs. Missing is not zero.",fontsize=8.5,color='#9b431c')
    fig.subplots_adjust(left=.31,right=.95,top=.82,bottom=.18)
    fig.canvas.draw();renderer=fig.canvas.get_renderer()
    for t in [*ax.get_yticklabels(),*ax.texts,*fig.texts]:
        box=t.get_window_extent(renderer)
        if box.x0<0 or box.x1>fig.bbox.x1 or box.y0<0 or box.y1>fig.bbox.y1:raise ValueError(('clipped text',w['well'],t.get_text()))
    out=BytesIO();fig.savefig(out,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return out.getvalue(),bool(selected)

def style(p,fa,center=False):
    if center:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    if fa:
        pr=p._p.get_or_add_pPr()
        if pr.find(qn('w:bidi')) is None:
            b=OxmlElement('w:bidi');before=next((x for x in pr if x.tag in [qn('w:'+n) for n in ('jc','outlineLvl','rPr','sectPr','pPrChange')]),None)
            if before is None:pr.append(b)
            else:before.addprevious(b)
    for r in p.runs:
        r.font.name='Times New Roman';r.font.size=Pt(8)
        rp=r._element.get_or_add_rPr();fonts=rp.find(qn('w:rFonts'))
        if fonts is None:fonts=OxmlElement('w:rFonts');rp.insert(0,fonts)
        fonts.set(qn('w:cs'),'B Nazanin' if fa else 'Times New Roman')
        if fa and rp.find(qn('w:rtl')) is None:rp.append(OxmlElement('w:rtl'))

def text(p,s):
    if p.runs:
        p.runs[0].text=s
        for r in p.runs[1:]:r.text=''
    else:p.add_run(s)

def digits(v):return str(v).translate(str.maketrans('0123456789','۰۱۲۳۴۵۶۷۸۹'))

def counter(g,key):
    x=g.get('maximum_'+key)
    return f"{x['value']:g} [{x['sheet']}!{x['cell']}]" if x else '—'

def metadata(g,fa):
    values=g['type'][:2]+g['size'][:2]
    raw='; '.join(values) if values else ('ثبت نشده' if fa else 'Not recorded')
    if g['conflicts']:raw+=('؛ اختلاف داده' if fa else '; CONFLICT')
    return raw

def set_table(table,rows,widths,fa):
    while len(table.rows)>1:table._tbl.remove(table.rows[-1]._tr)
    table.autofit=False
    for c,w in zip(table.columns,widths):c.width=Cm(w)
    trpr=table.rows[0]._tr.get_or_add_trPr()
    if trpr.find(qn('w:tblHeader')) is None:trpr.append(OxmlElement('w:tblHeader'))
    for i,row in enumerate(rows):
        cells=table.rows[0].cells if i==0 else table.add_row().cells
        for c,value,width in zip(cells,row,widths):
            c.width=Cm(width);c.text=value
            for p in c.paragraphs:style(p,fa,i==0)
    if fa and table._tbl.tblPr.find(qn('w:bidiVisual')) is None:
        pr=table._tbl.tblPr;b=OxmlElement('w:bidiVisual');tw=pr.find(qn('w:tblW'))
        if tw is None:pr.insert(0,b)
        else:tw.addprevious(b)

def unassigned_groups(w):
    groups=OrderedDict()
    for o in w['unassigned_observations']:
        f=o['fields'];k=tuple(f.get(x,{}).get('value','').strip() or '—' for x in ('bit_no','rerun','size','type'))
        if k not in groups:groups[k]=[0,o]
        groups[k][0]+=1
    return groups

def source_refs(entry):
    f=entry['fields'];selected=[f[k]['cell'] for k in ('bit_no','rerun','size','type') if k in f]
    return entry['sheet']+'!'+','.join(selected) if selected else entry['sheet']

def main():
 for well,w in DATA.items():
    png,has_counter=chart(w)
    for lang in ('EN','FA'):
        fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
        table=next((t for t in d.tables if t.rows and t.rows[0].cells[0].text in ('Bit','مته','شماره') and len(t.columns)==6),None)
        cap=next((p for p in d.paragraphs if p.text.startswith(('Table 4-2','جدول ۴-۲'))),None)
        if table is None:
            if cap is not None:raise ValueError((path,'caption exists but table missing'))
            # Some earlier reports omitted a bit table altogether. Place a
            # new caption and table before the existing W6 drawing.
            i=next(i for i,p in enumerate(d.paragraphs) if p.text.startswith(('Figure W6','شکل W6')))
            cap=d.add_paragraph('جدول ۴-۲' if fa else 'Table 4-2');style(cap,fa,True)
            d.paragraphs[i-1]._p.addprevious(cap._p)
            table=d.add_table(rows=1,cols=6);table.style='Table Grid';cap._p.addnext(table._tbl)
        if cap is None:raise ValueError((path,'bit caption missing'))
        header=['شماره','بازران','نوع / اندازه','بیشینه متراژ تجمعی / منبع','بیشینه ساعت تجمعی / منبع','سایش در ترتیب شیت'] if fa else ['Bit','Rerun','Type / size','Max cumulative m / source','Max cumulative h / source','Dull in sheet order']
        rows=[header]
        for g in w['groups']:
            dull=g['last_nonblank_dull'];dulltext=f"{dull['value']} [{dull['sheet']}!{dull['cell']}]" if dull else '—'
            bit=digits(g['bit_no']) if fa else str(g['bit_no']);rr=('نامشخص' if fa else 'Unknown') if g['rerun'] is None else (digits(g['rerun']) if fa else str(g['rerun']))
            rows.append([bit,rr,metadata(g,fa),counter(g,'cumulative_m'),counter(g,'cumulative_h'),dulltext])
        if len(rows)==1:rows.append(['—','—','شناسه عددی مته شناسایی نشد' if fa else 'No numeric bit ID recognized','—','—','—'])
        set_table(table,rows,[1.25,1.4,3.25,3.7,3.7,3.6],fa)
        text(cap,'جدول ۴-۲ ـ مشاهدات مته با شناسه قابل تشخیص و منبع شمارنده.' if fa else 'Table 4-2 — Identified bit observations and cumulative-counter sources.')
        for old in list(d.paragraphs):
            if old.text.startswith(('All rows come from AZNS207Pack.xlsx.','منبع تمام ردیف‌ها AZNS207Pack.xlsx است.','Bit evidence is preliminary:','شواهد مته مقدماتی است:')):old._element.getparent().remove(old._element)
        note=d.add_paragraph((f"شواهد مته مقدماتی است: منبع {w['source']} است؛ {w['recognized_data_sheets']} شیت داده از {w['candidate_pairs']} جفت شیت نامزد بررسی شد. {len(w['unassigned_observations'])} ستون دارای داده ابزار بدون شناسه عددی روشن باقی مانده است. بیشینه شمارنده تجمعی لزوماً مقدار افزایشی یک بازران نیست؛ شناسه مته و بازران تکراری، اختلاف مشخصات و سایش باید با دفاتر مته تطبیق شوند. ترتیب شیت‌ها زمان وقوع را ثابت نمی‌کند؛ سایش نمایش‌داده‌شده الزاماً نهایی نیست. مقدار نامعلوم صفر نیست." if fa else f"Bit evidence is preliminary: source {w['source']}; {w['recognized_data_sheets']} data sheets were recognized from {w['candidate_pairs']} candidate sheet pairs. {len(w['unassigned_observations'])} tool-column entries lack an unambiguous numeric bit ID. A cumulative-counter maximum is not necessarily incremental run performance; bit-number reuse, rerun, metadata conflicts and dull reports require reconciliation. Workbook order does not prove chronology; the displayed dull observation is not necessarily final. Missing values are not zero."))
        style(note,fa);cap._p.addnext(note._p)
        # Delete previous generated supplemental table/heading, if this script is rerun.
        for t in list(d.tables):
            if t.rows and t.rows[0].cells[0].text in ('Unassigned raw ID','شناسه خام نامشخص'):t._tbl.getparent().remove(t._tbl)
        for p in list(d.paragraphs):
            if p.text.startswith(('Unassigned tool observations —','مشاهدات ابزار بدون شناسه قطعی ـ')):p._p.getparent().remove(p._p)
        groups=unassigned_groups(w)
        if groups:
            title=d.add_paragraph('مشاهدات ابزار بدون شناسه قطعی ـ گروه‌بندی فقط برای نمایش متن یکسان، نه شمارش بازران.' if fa else 'Unassigned tool observations — grouped only by identical text, not verified runs.');style(title,fa,True)
            table._tbl.addnext(title._p)
            detail=d.add_table(rows=1,cols=4);detail.style='Table Grid';title._p.addnext(detail._tbl)
            dat=[(['شناسه خام نامشخص','اندازه / نوع ثبت‌شده','تعداد مشاهده','نخستین منبع'] if fa else ['Unassigned raw ID','Reported size / type','Entries','First source'])]
            for (raw,rr,size,kind),(count,entry) in groups.items():
                label=raw+((' / '+rr) if rr!='—' else '')
                dat.append([label,size+' / '+kind,digits(count) if fa else str(count),source_refs(entry)])
            set_table(detail,dat,[3.2,5.1,2.1,6.6],fa)
        i=next(i for i,p in enumerate(d.paragraphs) if p.text.startswith(('Figure W7','شکل W7')))
        bl=d.paragraphs[i-1]._element.xpath('.//a:blip')[0];rid=bl.get(qn('r:embed'));d.part.related_parts[rid]._blob=png
        shape=next(s for s in d.inline_shapes if any(b.get(qn('r:embed'))==rid for b in s._inline.xpath('.//a:blip')));shape.height=int(shape.width*6/11)
        caption=('شکل W7 ـ بیشینه شمارنده‌های تجمعی منبع؛ نه عملکرد افزایشی بازران' if fa else 'Figure W7 — Maximum source cumulative counters, not incremental run performance') if has_counter else ('شکل W7 ـ پوشش شناسایی داده مته و ابزار، نه شمار بازران‌ها' if fa else 'Figure W7 — Bit/tool evidence coverage, not a run count')
        text(d.paragraphs[i],caption)
        for p in d.paragraphs:
            if p.text.startswith(('The bonus panel below collects','پنل تکمیلی زیر ماژول‌های خاص')):
                text(p,'شکل بعدی فقط شواهد قابل تشخیص مته و ابزار را نشان می‌دهد؛ جدول جداری و داده هزینه در بخش‌های مربوط ارائه شده‌اند. این شکل تعداد بازران تأییدشده یا لاگ پیوسته نیست.' if fa else 'The next figure displays only recognized bit and tool evidence; casing and cost evidence are presented in their respective sections. This is not a verified run count or continuous log.')
        for t in d.tables:
            if t.rows:
                for c in t.rows[0].cells:
                    for p in c.paragraphs:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        d.save(path)
    print(well,'identified',len(w['groups']),'unassigned text groups',len(unassigned_groups(w)),'chart counters',has_counter,flush=True)
if __name__=='__main__':main()
