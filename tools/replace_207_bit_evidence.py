#!/usr/bin/env python3
"""Replace 207's erroneous bit table/plot with sourced cumulative observations.
Does not claim that maxima are incremental per-run performance.
"""
from io import BytesIO
from pathlib import Path
import json,textwrap
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm,Pt

source=json.loads(Path('evidence/207_bit_observations.json').read_text())
groups=source['groups'];assert len(groups)>0 and not any('Nozzle' in str(g['bit_no']) for g in groups)

def chart():
    eligible=[g for g in groups if g['maximum_cumulative_m'] is not None]
    selected=sorted(eligible,key=lambda g:-g['maximum_cumulative_m']['value'])[:10]
    fig,ax=plt.subplots(figsize=(11,5.7),dpi=300);fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white')
    labs=[f"Bit {g['bit_no']}, rerun {g['rerun'] if g['rerun'] is not None else 'unknown'}" for g in selected]
    vals=[g['maximum_cumulative_m']['value'] for g in selected]
    y=list(range(len(vals)));ax.barh(y,vals,color='#28688d',height=.66);ax.set_yticks(y,labs,fontsize=9);ax.invert_yaxis()
    ax.set_xlabel('Maximum reported cumulative counter (m)',fontsize=10);ax.set_xlim(0,max(vals)*1.3)
    ax.grid(axis='x',alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0)
    for i,v in enumerate(vals):ax.text(v+max(vals)*.02,i,f'{v:,.1f}',va='center',fontsize=9)
    fig.suptitle('AZNS-207 | Reported bit counters',x=.05,y=.98,ha='left',fontsize=15,fontweight='bold')
    fig.text(.05,.92,'Ten largest observed cumulative values; not metres drilled in each rerun.',fontsize=9,color='#9b431c')
    fig.text(.05,.035,'Source: AZNS207Pack.xlsx daily data; full observations and cell references in the report and evidence index.',fontsize=8.5,color='#9b431c')
    fig.subplots_adjust(left=.27,right=.96,top=.84,bottom=.17);fig.canvas.draw();renderer=fig.canvas.get_renderer()
    for a in [*ax.get_yticklabels(),*ax.texts]:
        b=a.get_window_extent(renderer)
        if b.x0<0 or b.x1>fig.bbox.x1 or b.y0<0 or b.y1>fig.bbox.y1:raise ValueError('clipped '+a.get_text())
    output=BytesIO();fig.savefig(output,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return output.getvalue()

def setup(p,fa,center=False):
    if center:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    if fa:
        pp=p._p.get_or_add_pPr()
        if pp.find(qn('w:bidi')) is None:
            b=OxmlElement('w:bidi');before=next((x for x in pp if x.tag in [qn('w:'+s) for s in ('jc','outlineLvl','rPr','sectPr','pPrChange')]),None)
            if before is not None:before.addprevious(b)
            else:pp.append(b)
    for r in p.runs:
        r.font.name='Times New Roman';r.font.size=Pt(8)
        rp=r._element.get_or_add_rPr();fonts=rp.find(qn('w:rFonts'))
        if fonts is None:fonts=OxmlElement('w:rFonts');rp.insert(0,fonts)
        fonts.set(qn('w:cs'),'B Nazanin' if fa else 'Times New Roman')
        if fa and rp.find(qn('w:rtl')) is None:rp.append(OxmlElement('w:rtl'))

def fa_num(s):return str(s).translate(str.maketrans('0123456789','۰۱۲۳۴۵۶۷۸۹'))

def counter(g,key):
    x=g.get('maximum_'+key)
    return f"{x['value']:g} [{x['sheet']}!{x['cell']}]" if x else '—'

png=chart()
for lang in ('EN','FA'):
    fa=lang=='FA';path=f'207_Drilling_Engineering_Report_{lang}.docx';d=Document(path)
    table=next(t for t in d.tables if t.rows and t.rows[0].cells[0].text in ('Bit','مته'))
    for r in list(table.rows)[1:]:table._tbl.remove(r._tr)
    headings=['شماره','بازران','نوع / اندازه','بیشینه شمارنده متراژ / منبع','بیشینه شمارنده ساعت / منبع','آخرین سایش ثبت‌شده'] if fa else ['Bit','Rerun','Type / size','Max cumulative m / source','Max cumulative h / source','Last reported dull']
    for c,text in zip(table.rows[0].cells,headings):c.text=text
    trpr=table.rows[0]._tr.get_or_add_trPr()
    if trpr.find(qn('w:tblHeader')) is None:trpr.append(OxmlElement('w:tblHeader'))
    for g in groups:
        cells=table.add_row().cells
        info='; '.join(g['type']+g['size']) or ('ثبت نشده' if fa else 'Not recorded')
        if g['conflicts']:info+=(' — اختلاف داده' if fa else ' — CONFLICT')
        dull=g['last_nonblank_dull'];dulltext=f"{dull['value']} [{dull['sheet']}!{dull['cell']}]" if dull else '—'
        bit=fa_num(g['bit_no']) if fa else str(g['bit_no'])
        rr=('نامشخص' if fa else 'Unknown') if g['rerun'] is None else (fa_num(g['rerun']) if fa else str(g['rerun']))
        for c,text in zip(cells,[bit,rr,info,counter(g,'cumulative_m'),counter(g,'cumulative_h'),dulltext]):c.text=text
    widths=[1.3,1.4,3.4,3.6,3.6,3.7]
    for col,w in zip(table.columns,widths):col.width=Cm(w)
    for i,row in enumerate(table.rows):
        for c,w in zip(row.cells,widths):
            c.width=Cm(w)
            for p in c.paragraphs:setup(p,fa,i==0)
    cap=next(p for p in d.paragraphs if p.text.startswith(('Table 4-2','جدول ۴-۲')))
    cap.runs[0].text='جدول ۴-۲ ـ مشاهدات شمارنده مته به تفکیک شماره و بازران، با ارجاع منبع.' if fa else 'Table 4-2 — Bit and rerun cumulative-counter observations with source references.'
    for r in cap.runs[1:]:r.text=''
    note=d.add_paragraph('منبع تمام ردیف‌ها AZNS207Pack.xlsx است. شمارنده‌های تجمعی ممکن است بازران‌های قبلی را شامل شوند؛ بیشینه ثبت‌شده، متراژ یا ساعت افزایشی هر بازران نیست. بازران خالی نامشخص باقی مانده و به صفر تبدیل نشده است. مقادیر دارای اختلاف مشخصات و آخرین وضعیت سایش باید با سوابق مته تطبیق شوند. شماره‌های گزارش و سلول‌ها در کروشه‌اند؛ مقدار صفر به‌تنهایی اثبات عدم حفاری نیست.' if fa else 'All rows come from AZNS207Pack.xlsx. Cumulative counters may include earlier reruns; a maximum recorded counter is not incremental metres or hours for a rerun. A missing rerun remains unknown rather than being converted to zero. Metadata conflicts and last reported dull grades require reconciliation to bit records. Sheet and cell references are in brackets; a zero alone does not prove no drilling.')
    setup(note,fa);cap._p.addnext(note._p)
    i=next(i for i,p in enumerate(d.paragraphs) if p.text.startswith(('Figure W7','شکل W7')))
    bl=d.paragraphs[i-1]._element.xpath('.//a:blip')[0];rid=bl.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
    d.part.related_parts[rid]._blob=png
    shape=next(s for s in d.inline_shapes if any(b.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')==rid for b in s._inline.xpath('.//a:blip')))
    shape.height=int(shape.width*5.7/11)
    caption=d.paragraphs[i];caption.runs[0].text='شکل W7 ـ بیشینه شمارنده‌های تجمعی ثبت‌شده مته، چاه 207' if fa else 'Figure W7 — Maximum reported cumulative bit counters, AZNS-207'
    for r in caption.runs[1:]:r.text=''
    d.save(path)
print('Updated 207 bit table and W7 in EN/FA')
