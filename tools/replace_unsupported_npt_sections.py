#!/usr/bin/env python3
"""Remove unsupported mechanism and contractual-attribution assertions.

Replacement uses only activity-code hours from complete daily records and
makes no causal, contractual, or financial inference.
"""
from pathlib import Path
from io import BytesIO
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.oxml.ns import qn
from replace_well_bit_evidence import DATA,set_table,text,style

TRACKS={x['file']:x for x in json.loads(Path('evidence/drilling_track_observations.json').read_text())}
CODES=[('npt','npt_h','N / NPT','کد N / NPT'),('unplanned','unplanned_h','UP / unplanned','کد UP / غیربرنامه‌ای'),('mixed','mixed_h','NPT-UNPLN','کد ترکیبی NPT-UNPLN')]

def values(track):
    complete=[x for x in track['days'] if x['activity']['complete']]
    result=[]
    for classification,key,en,fa in CODES:
        total=sum(x['activity'][key] for x in complete)
        examples=[]
        for day in complete:
            for segment in day['activity']['segments']:
                if segment['classification']==classification and segment['duration']['value'] is not None:
                    examples.append(f"{day['remark_sheet']}!{segment['status_cell']},{segment['duration']['cell']}")
        result.append({'total':total,'en':en,'fa':fa,'examples':examples[:2]})
    return complete,result

def chart(w,track,complete,rows):
    rows=sorted(rows,key=lambda x:x['total'],reverse=True);fig,ax=plt.subplots(figsize=(11,6),dpi=300);fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white')
    vals=[r['total'] for r in rows];x=range(3);bars=ax.bar(x,vals,width=.48,color=['#28688d','#c65d42','#b69a50']);top=max(vals) or 1
    ax.set_ylim(0,top*1.22);ax.set_xticks(list(x),[r['en'] for r in rows],fontsize=10);ax.set_ylabel('Hours in selected complete days',fontsize=10);ax.grid(axis='y',alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right']].set_visible(False)
    for rect,val in zip(bars,vals):ax.text(rect.get_x()+rect.get_width()/2,rect.get_height()+top*.025,f'{val:,.1f}',ha='center',va='bottom',fontsize=10)
    total=sum(vals)
    if total:
        second=ax.twinx();acc=0;ys=[]
        for v in vals:acc+=v;ys.append(100*acc/total)
        second.plot(list(x),ys,'o-',color='#8f3441',linewidth=2,markersize=5);second.set_ylim(0,110);second.set_ylabel('Cumulative share of displayed codes (%)',color='#8f3441',fontsize=9);second.tick_params(axis='y',colors='#8f3441');second.spines[['top','left']].set_visible(False)
        for i,y in enumerate(ys):second.annotate(f'{y:.1f}%',(i,y),xytext=(34,6),textcoords='offset points',ha='left',color='#8f3441',fontsize=8)
    fig.suptitle(f"AZNS-{w['well']} | Source status-code distribution",x=.05,y=.98,ha='left',fontsize=15,fontweight='bold',color='#223952')
    fig.text(.05,.915,'Partial coverage; the code is not a technical root cause, contractual responsibility or a financial cost.',fontsize=9,color='#9b431c')
    fig.text(.05,.04,f"Source: {w['source']} | {len(complete)} complete 24-hour records / {track['paired_sheet_candidates']} candidate pairs. Missing days are not zero.",fontsize=8.5,color='#9b431c')
    fig.subplots_adjust(left=.10,right=.89,top=.82,bottom=.18);fig.canvas.draw();renderer=fig.canvas.get_renderer()
    for a in fig.axes:
        for t in [*a.get_xticklabels(),*a.get_yticklabels(),*a.texts]:
            box=t.get_window_extent(renderer)
            if box.x0 < -1 or box.x1>fig.bbox.x1+1 or box.y0 < -1 or box.y1>fig.bbox.y1+1:raise ValueError((w['well'],t.get_text()))
    out=BytesIO();fig.savefig(out,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return out.getvalue()

def main():
 for well,w in DATA.items():
    track=TRACKS[w['source']];complete,rows=values(track);png=chart(w,track,complete,rows)
    for lang in ('EN','FA'):
        fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
        h=next(p for p in d.paragraphs if p.text.startswith(('Root-Cause Analysis and Three-Dimensional NPT Attribution','ریشه‌یابی و تفکیک سه‌بُعدی زمان غیرمولد','Source-coded non-planned time','زمان غیربرنامه‌ای بر اساس کد منبع')))
        text(h,'زمان غیربرنامه‌ای بر اساس کد منبع و پوشش جزئی' if fa else 'Source-coded non-planned time and partial coverage');h.alignment=1
        figcap=next(p for p in d.paragraphs if p.text.startswith(('Figure W4','شکل W4')))
        drawing=figcap._p.getprevious()
        if not drawing.xpath('.//a:blip'):raise ValueError((well,lang,'W4 drawing missing'))
        cursor=h._p.getnext()
        while cursor is not drawing:
            nxt=cursor.getnext()
            if nxt is None:raise ValueError((well,lang,'W4 precedes heading'))
            cursor.getparent().remove(cursor);cursor=nxt
        p=d.add_paragraph((f"از {track['paired_sheet_candidates']} جفت شیت نامزد، فقط {len(complete)} گزارش روزانه با جمع مدت‌های طبقه‌بندی‌شده نزدیک ۲۴ ساعت در محاسبه وارد شده است. مقادیر زیر ساعت کدهای منبع‌اند، نه کل قطعی زمان غیرمولد عملیات؛ روزهای ناقص یا تشخیص‌نیافته صفر فرض نشده‌اند. جزئیات ردیف‌ها و سلول‌ها در evidence/drilling_track_observations.json ثبت است." if fa else f"Only {len(complete)} daily reports with classified durations totalling approximately 24 hours were included out of {track['paired_sheet_candidates']} candidate sheet pairs. These are source-code hours, not a verified total for the operation; incomplete or unrecognized days are not zero. Row-level cells are retained in evidence/drilling_track_observations.json."));style(p,fa);drawing.addprevious(p._p)
        p=d.add_paragraph(('جدول ۷-۱ ـ ساعت کدهای منبع در گزارش‌های کامل؛ ارجاعات صرفاً نمونه‌اند.' if fa else 'Table 7-1 — Source-code hours in complete reports; references are examples.'));style(p,fa,True);drawing.addprevious(p._p)
        t=d.add_table(rows=1,cols=3);t.style='Table Grid';drawing.addprevious(t._tbl)
        values_table=[(['کد منبع','ساعت در پوشش موجود','نمونه شیت و سلول وضعیت و مدت'] if fa else ['Source code','Hours in coverage','Example status and duration cells'])]
        for row in rows:values_table.append([row['fa'] if fa else row['en'],f"{row['total']:,.2f}",' ; '.join(row['examples']) if row['examples'] else ('بدون مورد در روزهای کامل' if fa else 'None in included complete days')])
        set_table(t,values_table,[4.3,3.0,8.8],fa)
        p=d.add_paragraph('برچسب‌های فعالیت به‌تنهایی علت ریشه‌ای یا مسئول قراردادی را اثبات نمی‌کنند. تحلیل سازوکار، انتساب پیمانکار و ریسک مالی قبلی تا بازبینی متن واقعی گزارش روزانه، قرارداد و شواهد رویداد کنار گذاشته شده‌اند؛ جمع‌زدن کدهای متفاوت به‌عنوان یک تعریف NPT مجاز نیست.' if fa else 'Activity codes alone do not establish root cause or contractual responsibility. Legacy mechanism claims, contractor attribution and financial exposure are withdrawn pending review of actual daily narrative, contracts and event evidence. Different source codes must not be presented as a single NPT definition.')
        style(p,fa);drawing.addprevious(p._p)
        bl=drawing.xpath('.//a:blip')[0];rid=bl.get(qn('r:embed'));d.part.related_parts[rid]._blob=png
        shape=next(s for s in d.inline_shapes if any(b.get(qn('r:embed'))==rid for b in s._inline.xpath('.//a:blip')));shape.height=int(shape.width*6/11)
        text(figcap,'شکل W4 ـ پارتوی کدهای منبع در روزهای کامل؛ علت فنی نیست' if fa else 'Figure W4 — Pareto of source status codes in complete days, not technical causes')
        for table in d.tables:
            if table.rows:
                for c in table.rows[0].cells:
                    for p in c.paragraphs:p.alignment=1
        d.save(path)
    print(well,len(complete),track['paired_sheet_candidates'],[round(r['total'],2) for r in rows],flush=True)
if __name__=='__main__':main()
