#!/usr/bin/env python3
"""Replace unsourced legacy cost reconstructions with directly cited CBS values.

All figures remain estimates and are not actual invoices. Run on the interim
reports once; source evidence is created by extract_cbs_evidence.py.
"""
from pathlib import Path
from io import BytesIO
import json, re, math, textwrap
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Cm,Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

mapping={
'10':('CBS AZNS-10.xlsx','Cost Breakdown-10 w. Extra Wor','F191'),
'12':('CBS AZNS-12.xlsx','Cost Breakdown-12','F197'),
'75':('CBS AZNS-075.xlsx','Cost Breakdown-075','F197'),
'85':('CBS AZNS-85.xlsx',None,None),
'116':('CBS AZNS-127WO1 and 116.xlsx','Cost Breakdown-116','F197'),
'122WO1':('CBS AZNS-122WO1.xlsx','Cost Breakdown-122','G190'),
'122WO2':('CBS AZNS-122WO2.xlsx','Cost Breakdown-122','F189'),
'127ST2':('CBS AZNS-127 ST 2.xlsx','Cost Breakdown-127','F197'),
'127WO1':('CBS AZNS-127 WO1.xlsx','Cost Breakdown-127','F199'),
'127WO2':('CBS AZNS-127WO2.xlsx','Cost Breakdown-127#2','F197'),
'166':('CBS AZNS-166.xlsx','Cost Breakdown-166','F197'),
'207':('CBS AZNS-207.xlsx','Cost Breakdown-F-20','F198'),
'40':('CBS AZNS-40.xlsx','Cost Breakdown-40','F178')}

en_names=['Rig service','Drilling fluids','Cementing','Mud logging','Wireline logging','Drilling waste management','Directional drilling','Casing running','Temporary well suspension','Slickline','Completion running','Liner hanger running','H2S service','Perforation','Wellhead services','Dry test services','Rig / completion services','Other services','Cleanup services','Drilling materials','Project management','Engineering','Tubing inspection','Wellhead / tree inspection','Downhole completion inspection']
fa_names=['خدمات دکل','خدمات سیال حفاری','خدمات سیمان‌کاری','گل‌نگاری','نمودارگیری با سیم','مدیریت پسماند حفاری','حفاری جهت‌دار','راندن جداری','تعلیق موقت چاه','خدمات اسلیک‌لاین','راندن تکمیل','نصب آویز آستری','خدمات سولفید هیدروژن','مشبک‌کاری','خدمات سرچاهی','آزمایش خشک','دکل و خدمات دوره تکمیل','سایر خدمات','خدمات پاکسازی','مواد حفاری','مدیریت پروژه','مهندسی','بازرسی لوله مغزی','بازرسی سرچاه و شیر فوران','بازرسی تکمیل درون‌چاهی']
assert len(en_names)==len(fa_names)==25
DATA=json.loads(Path('evidence/cbs_cost_observations.json').read_text(encoding='utf-8'))
books={w['file']:{s['sheet']:{c['cell']:c for c in s['cells']} for s in w['sheets']} for w in DATA['workbooks']}

def value(sheet,cell):return sheet.get(cell,{}).get('value')
def numeric(x):return isinstance(x,(int,float)) and not isinstance(x,bool) and math.isfinite(x)
def fmt(v,fa=False):
    if not numeric(v):return '—'
    s=f'{v:,.2f}'
    return s.translate(str.maketrans('0123456789,.','۰۱۲۳۴۵۶۷۸۹٬٫')) if fa else s

def leaf_groups(well):
    f,s,total=mapping[well]
    if not s:return False
    c=books[f][s]
    grand=c.get(total,{}).get('formula','').replace(' ','').upper()
    header=c.get('F6',{}).get('formula','').replace(' ','').upper()
    return bool(re.fullmatch(r'=SUM\(F5:F\d+\)',grand) and header=='=D6*E6')

def categories(well):
    f,s,_=mapping[well];cells=books[f][s];dual=well in ('10','122WO1');out=[]
    for r in range(1,250):
        n=value(cells,f'A{r}');label=value(cells,f'B{r}')
        if not numeric(n) or int(n)!=n or not 1<=n<=25 or not isinstance(label,str):continue
        a,b=('G','H') if dual else ('F',None)
        out.append({'number':int(n),'row':r,'source_label':label,'one':value(cells,f'{a}{r}'),'two':value(cells,f'{b}{r}') if b else None,'cells':f'{a}{r}' + (f' / {b}{r}' if b else '')})
    if len(out)!=25:raise ValueError((well,len(out)))
    # Some templates put zero in the numbered group header and the actual
    # line values below. Their grand-total formula sums the individual rows.
    # ST2 looks similar at the grand total but has subtotal headers; do not
    # sum its header and children again (the source total itself needs review).
    if leaf_groups(well):
        end=int(re.search(r'\d+',mapping[well][2]).group())-1
        for i,x in enumerate(out):
            first=x['row'];last=out[i+1]['row']-1 if i+1<len(out) else end
            items=[cells.get(f'F{r}') for r in range(first,last+1)]
            if any(c and ('error' in c or c.get('missing_cache')) for c in items):total=None
            else:
                vals=[c['value'] for c in items if c and numeric(c.get('value'))]
                total=sum(vals) if vals else None
            x['one']=total;x['cells']=f'F{first}:F{last}';x['derived']='sum of numeric cached values in the source row range'
    return out

def plot(well,rows):
    fig,ax=plt.subplots(figsize=(11,6),dpi=300);fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white')
    if well=='85':
        ax.barh([0,1],[48.0729166667,51.8895833333],color=['#28688d','#d2a354'],height=.45)
        ax.set_yticks([0,1],['Rig schedule total','With contingency']);ax.invert_yaxis();ax.set_xlabel('Reported duration (days)')
        ax.set_xlim(0,65)
        for i,x in enumerate([48.0729166667,51.8895833333]):ax.text(x+1,i,f'{x:.2f}',va='center')
        fig.text(.06,.90,'CBS AZNS-85.xlsx, AZNS-085 WO. Extra Work: K54 and O54.',fontsize=9,color='#9b431c')
        fig.text(.06,.055,'Schedule scope only; this is not a monetary budget or actual expenditure. Rig-less steps are not zero elapsed time.',fontsize=8.5,color='#9b431c')
        fig.suptitle('AZNS-85 | Source rig schedule, not a cost series',x=.06,y=.98,ha='left',fontweight='bold',fontsize=15,color='#223952')
        fig.subplots_adjust(left=.29,right=.95,top=.8,bottom=.20)
    else:
        usable=[r for r in rows if any(numeric(v) and v>0 for v in (r['one'],r['two']))]
        ranked=sorted(usable,key=lambda x:max(v for v in (x['one'],x['two']) if numeric(v)),reverse=True)[:8]
        dual=well in ('10','122WO1');h=.33 if dual else .62
        for i,r in enumerate(ranked):
            for idx,(v,color) in enumerate(((r['one'],'#28688d'),(r['two'],'#d0a153')) if dual else ((r['one'],'#28688d'),)):
                if not numeric(v):continue
                y=i+((idx-.5)*h if dual else 0);z=v/1e6
                ax.barh(y,z,height=h,color=color);ax.text(z+.015*max(1, max(max(vv for vv in (x['one'],x['two']) if numeric(vv)) for x in ranked)/1e6),y,f'{z:.2f}',fontsize=7.6,va='center')
        ax.set_yticks(range(len(ranked)),[en_names[r['number']-1] for r in ranked],fontsize=9);ax.invert_yaxis();ax.set_xlabel('CBS category amount (USD million)',fontsize=10)
        highest=max(max(v for v in (r['one'],r['two']) if numeric(v)) for r in ranked)/1e6;ax.set_xlim(0,max(highest*1.2,.1))
        if dual:
            from matplotlib.patches import Patch
            ax.legend(handles=[Patch(color='#28688d',label=('G: plan total' if well=='10' else 'G: source total')),Patch(color='#d0a153',label='Extra-work column')],loc='lower right',fontsize=8,frameon=False)
        fig.suptitle(f'AZNS-{well} | CBS source categories',x=.06,y=.98,ha='left',fontweight='bold',fontsize=15,color='#223952')
        f,s,_=mapping[well];fig.text(.06,.91,f'{f} | {s} | up to 8 categories; cached values / grouped sums.',fontsize=8.4,color='#9b431c')
        foot='Source has #REF! cells; cached sums may exclude error rows.' if well=='127WO1' else ('Contingency-material sheet has #REF! errors; scope impact unresolved.' if well=='122WO2' else 'Not actual expenditure. Selected category amounts may not sum to the source total; reconcile scope and formulas.')
        fig.text(.06,.055,foot,fontsize=8.5,color='#9b431c')
        fig.subplots_adjust(left=.31,right=.95,top=.82,bottom=.18)
    ax.grid(axis='x',alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0)
    out=BytesIO();fig.savefig(out,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return out.getvalue()

def set_text(p,text):
    if p.runs:
        p.runs[0].text=text
        for r in p.runs[1:]:r.text=''
    else:p.add_run(text)

def style(p,fa,header=False):
    if header:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    if fa:
        pr=p._p.get_or_add_pPr()
        if pr.find(qn('w:bidi')) is None:
            el=OxmlElement('w:bidi');before=next((x for x in pr if x.tag in [qn('w:'+n) for n in ('jc','outlineLvl','rPr','sectPr','pPrChange')]),None)
            if before is None:pr.append(el)
            else:before.addprevious(el)
    for r in p.runs:
        r.font.name='Times New Roman';r.font.size=Pt(8)
        rp=r._element.get_or_add_rPr();fonts=rp.find(qn('w:rFonts'))
        if fonts is None:fonts=OxmlElement('w:rFonts');rp.insert(0,fonts)
        fonts.set(qn('w:cs'),'B Nazanin' if fa else 'Times New Roman')
        if fa and rp.find(qn('w:rtl')) is None:rp.append(OxmlElement('w:rtl'))

def celltext(cell,text,fa,head=False):
    cell.text=text
    for p in cell.paragraphs:style(p,fa,head)

def rewrite(table,rows,fa):
    while len(table.rows)>1:table._tbl.remove(table.rows[-1]._tr)
    table.autofit=False
    n=len(table.columns);widths={3:(6.8,4.0,6.0),4:(3.2,5.2,5.2,3.2)}[n]
    for col,w in zip(table.columns,widths):col.width=Cm(w)
    for i,vals in enumerate(rows):
        row=table.rows[0] if i==0 else table.add_row()
        assert len(vals)==n
        for c,w,v in zip(row.cells,widths,vals):c.width=Cm(w);celltext(c,str(v),fa,i==0)
    if fa and table._tbl.tblPr.find(qn('w:bidiVisual')) is None:
        pr=table._tbl.tblPr;tag=OxmlElement('w:bidiVisual');w=pr.find(qn('w:tblW'))
        if w is None:pr.insert(0,tag)
        else:w.addprevious(tag)

def summary(well,fa):
    f,s,total=mapping[well]
    header=['منبع و دامنه','مقدار منبع','ارجاع'] if fa else ['Source / scope','Cached value','Cell / sheet']
    if well=='85':return [header,['مدت برنامه دکل؛ مبنای مالی نیست' if fa else 'Rig schedule; not a monetary budget', '۴۸٫۰۷ روز' if fa else '48.07 days','K54 / AZNS-085 WO. Extra Work'],['با احتیاط' if fa else 'With contingency','۵۱٫۸۹ روز' if fa else '51.89 days','O54 / AZNS-085 WO. Extra Work'],['مبنای هزینه و بودجه' if fa else 'Actual cost and budget','نیازمند تطبیق مدارک' if fa else 'Requires document reconciliation','CBS AZNS-85.xlsx']]
    cells=books[f][s];v=value(cells,total)
    name=('جمع ثبت‌شده در شیت اصلی؛ دامنه تأیید نشده' if fa else 'Cached main-sheet total; scope unverified')
    out=[header,[name,fmt(v,fa),f'{total} / {s}']]
    if well=='122WO1':out.append(['جمع ستون کار اضافی' if fa else 'Extra-work column total',fmt(value(cells,'H190'),fa),'H190 / '+s])
    if well=='10':
        s2=books[f]['Cost Breakdown-10'];out.append(['کار اضافی، شیت دیگر؛ تطبیق شود' if fa else 'Extra work in another sheet; reconcile',fmt(value(s2,'F190'),fa),'F190 / Cost Breakdown-10'])
    # Joint 116/127 ceiling is explicitly joint and must never be allocated here.
    sumsheet=books[f].get('Total Cost-EPD & Rigless')
    if sumsheet:
        label=('سقف منبع، مشترک دو چاه' if well=='116' and fa else 'Source ceiling, joint for two wells' if well=='116' else 'سقف ثبت‌شده در منبع؛ دامنه بررسی شود' if fa else 'Source ceiling; scope to reconcile')
        out.append([label,fmt(value(sumsheet,'C5'),fa),'C5 / Total Cost-EPD & Rigless'])
    return out

def audit(well,fa):
    generic=(('دامنه و اسناد هزینه','این ارقام صورت‌حساب واقعی یا سقف مصوب قابل مقایسه نیستند.','شیت منبع و فرمول‌ها','تطبیق لازم است') if fa else ('Scope / actuals','Values are not actual invoices or an approved comparable ceiling.','Source sheet and formulas','Reconcile'))
    special={
'10':(('کار اضافی','سلول F191 از ستون H جمع می‌گیرد؛ شیت دیگر کار اضافی جداگانه دارد.','F191؛ F190 شیت دیگر','دامنه‌ها جمع نشوند') if fa else ('Extra work','F191 sums H-column rows; a separate extra-work sheet also exists.','F191; other sheet F190','Do not add scopes')),
'40':(('خدمات دکل','نرخ ۲۳٬۰۰۰ و مقدار ۱۰۲ با مبلغ ثبت‌شده ۸۱۹٬۸۲۰ برابر نیست.','D5, E5, F5','مبنای مبلغ بررسی شود') if fa else ('Rig service','23,000 × 102 differs from the cached 819,820 amount.','D5, E5, F5','Reconcile basis')),
'207':(('کالای چاه','مبلغ کالا در F3 گنجانده شده و ردیف G4 صفر است؛ شناسه F-20 تأیید شود.','F3, F4, G4 / Total Cost-EPD','از دوباره‌شماری پرهیز شود') if fa else ('Material / identity','F3 includes materials while G4 is zero; confirm F-20 identity.','F3, F4, G4 / Total Cost-EPD','Avoid double count')),
'127ST2':(('جمع و احتیاط','جمع F197 هم F6 و هم اجزای F7:F12 را می‌شمارد؛ D4 صفر و G4 غیرصفر است.','F6:F12, F197؛ D4, G4','تطبیق فرمول و دامنه') if fa else ('Total / contingency','F197 sums subtotal F6 and its F7:F12 components; D4 is zero while G4 is nonzero.','F6:F12, F197; D4, G4','Reconcile formulas')),
'127WO1':(('خطای ارجاع','سلول‌های متعدد دارای خطای ارجاع‌اند؛ جمع ذخیره‌شده نهایی قابل اتکا نیست.','F30, F41, F199','اصلاح و محاسبه مجدد لازم است') if fa else ('Reference errors','Multiple cells contain reference errors; cached total is not reliable.','F30, F41, F199','Repair / recalculate')),
'122WO2':(('خطای ارجاع','ستون H شیت کالای احتیاط دارای خطای ارجاع است؛ اثر بر دامنه روشن نیست.','H5, H6, H249 / Material-Conting','اثر باید بررسی شود') if fa else ('Reference errors','Contingency-material H column contains errors; scope effect is unresolved.','H5, H6, H249 / Material-Conting','Investigate impact')),
'127WO2':(('نظرات قراردادی','حذف برخی خدمات و اقلام کالا در نظرات شیت مطرح شده و تأیید نهایی روشن نیست.','J12, Q12 / Total Cost-EPD','بررسی تأییدیه') if fa else ('Review comments','Some service/material exclusions are disputed in worksheet comments.','J12, Q12 / Total Cost-EPD','Check approvals')),
'116':(('دامنه مشترک','سقف جدول خلاصه شامل ۱۱۶ و ۱۲۷ است؛ سهم مستقل هر چاه مشخص نیست.','B3:B14, G23 / Total Cost-EPD','بدون تخصیص فرضی') if fa else ('Joint scope','Summary ceiling covers 116 and 127; standalone allocation is not established.','B3:B14, G23 / Total Cost-EPD','Do not allocate')),
'85':(('زمان بدون دکل','ردیف‌های ۴۱ تا ۴۹ زمان دکل صفر دارند؛ زمان تقویمی یا هزینه صفر نتیجه نمی‌شود.','I41:I49 / برنامه','تأیید منابع مالی') if fa else ('Rig-less time','Rows 41–49 add zero rig time, not necessarily zero elapsed time or cost.','I41:I49 / schedule','Seek cost sources'))}
    header=['موضوع','مشاهده','ارجاع منبع','اقدام'] if fa else ['Topic','Observation','Source reference','Action']
    return [header,generic]+([special[well]] if well in special else [])

def main():
 for well,(file,sheet,total) in mapping.items():
    rows=categories(well) if sheet else []
    png=plot(well,rows)
    for lang in ('EN','FA'):
        fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');doc=Document(path)
        for t in doc.tables:
            if not t.rows:continue
            head=t.rows[0].cells[0].text.strip()
            if head in ('Financial component','جزء مالی','Source / scope','منبع و دامنه'):
                rewrite(t,summary(well,fa),fa)
            elif head in ('CBS cost category','طبقه هزینه CBS','Source category','طبقه منبع'):
                hdr=['طبقه منبع','مبلغ منبع (دلار)','سلول مبلغ'] if fa else ['Source category','Source amount (USD)','Amount cell']
                new=[hdr]
                for r in rows:
                    label=(fa_names if fa else en_names)[r['number']-1]
                    amount=fmt(r['one'],fa)+(' / '+fmt(r['two'],fa) if well in ('10','122WO1') else '')
                    new.append([label,amount,r['cells']])
                if not rows:new.append(['طبقات هزینه','نیازمند بررسی مدارک','CBS AZNS-85.xlsx'] if fa else ['Cost categories','Document reconciliation needed','CBS AZNS-85.xlsx'])
                rewrite(t,new,fa)
            elif head in ('Largest CBS line items','بزرگ‌ترین ردیف‌های هزینه','Topic','موضوع'):
                rewrite(t,audit(well,fa),fa)
            elif head in ('KPI','شاخص'):
                targets=('Cost (actual vs CBS budget)','Cost variance','Average burn rate','Financial cost of NPT','هزینه (واقعی در برابر بودجه CBS)','انحراف هزینه','نرخ متوسط هزینه روزانه','هزینه مالی زمان غیرمولد')
                for r in t.rows[1:]:
                    if well=='85' and r.cells[0].text in ('Planned vs actual duration','مدت برنامه‌ای در برابر واقعی'):
                        celltext(r.cells[1], 'منبع: ۴۸٫۰۷ روز؛ با احتیاط ۵۱٫۸۹ روز؛ مقایسه با اجرا نیازمند تطبیق است' if fa else 'CBS: 48.07 days; 51.89 with contingency; actual comparison pending',fa)
                    if r.cells[0].text in targets:celltext(r.cells[1],'نیازمند تطبیق با اسناد هزینه' if fa else 'Pending reconciliation to actual cost records',fa)
            elif head in ('Reconciliation metric','شاخص تطبیق'):
                for r in t.rows[1:]:
                    if well=='85' and r.cells[0].text in ('Duration (rig days)','مدت (روز دکل)'):
                        for c in r.cells[1:]:celltext(c,'تطبیق‌نشده' if fa else 'Not reconciled',fa)
                    if r.cells[0].text in ('Cost (USD)','Cost per rig day (USD)','هزینه (دلار)','هزینه هر روز دکل (دلار)'):
                        for c in r.cells[1:]:celltext(c,'تطبیق‌نشده' if fa else 'Not reconciled',fa)
            elif head in ('Responsible party / contractor','طرف مسئول / پیمانکار'):
                for r in t.rows[1:]:celltext(r.cells[2],'تأیید نشده' if fa else 'Not verified',fa)
            elif head in ('Risk','ریسک'):
                for r in t.rows[1:]:
                    if r.cells[0].text in ('Cost overrun against CBS','اضافه‌هزینه نسبت به CBS'):
                        celltext(r.cells[0],'ریسک دامنه و بودجه' if fa else 'Budget / scope reconciliation',fa)
                        celltext(r.cells[1],'هزینه واقعی و دامنه بودجه هنوز قابل مقایسه نیستند' if fa else 'Actual cost and budget scope not yet comparable',fa)
                        celltext(r.cells[2],'تطبیق اسناد هزینه، قرارداد و دامنه' if fa else 'Reconcile invoices, contract and scope',fa)
                    elif r.cells[0].text in ('Data continuity','پیوستگی داده'):
                        celltext(r.cells[1],'پوشش و کیفیت منابع باید جداگانه بررسی شود' if fa else 'Source coverage and quality require separate review',fa)
            if len(t.rows)==1 and len(t.columns)==1:
                cell=t.cell(0,0);x=cell.text
                if x.startswith('This report consolidates every daily drilling report'):
                    celltext(cell,'This interim report compiles available records; its legacy analyses are not independently validated.',False)
                elif x.startswith('این گزارش تمامی گزارش‌های روزانه حفاری'):
                    celltext(cell,'این نسخه مقدماتی منابع در دسترس را گردآوری می‌کند؛ تحلیل‌های قدیمی هنوز به‌طور مستقل تأیید نشده‌اند.',True)
                elif x.startswith('Operation AZNS-') and 'reconstructed spend' in x:
                    celltext(cell,'Legacy duration and NPT figures in this report remain subject to source reconciliation. Actual expenditure is not established.',False)
                elif x.startswith('عملیات AZNS-') and 'هزینه بازسازی‌شده' in x:
                    celltext(cell,'مدت و زمان غیرمولد مندرج در متن قدیمی نیازمند تطبیق با منابع است. هزینه واقعی احراز نشده است.',True)
        for p in doc.paragraphs:
            x=p.text
            if well=='85' and (x.startswith('The execution consumed 51.5 rig days against a planned 47.0') or x.startswith('اجرا ۵۱٫۵ روز دکل در برابر ۴۷٫۰ روز برنامه‌ای')):
                set_text(p,'در شیت برنامه چاه ۸۵، مقدار ۴۷٫۰۳ روز جمع بخش عملیات در K50 است؛ جمع شامل ردیف‌های تکمیلی K54 برابر ۴۸٫۰۷ روز و مقدار با احتیاط O54 برابر ۵۱٫۸۹ روز است. مقایسه با مدت اجرای گزارش‌ها تا تطبیق محدوده و مبنای زمان معتبر نیست.' if fa else 'For well 85, K50 is the 47.03-day workover subtotal; K54 is the 48.07-day total including additional items, and O54 is 51.89 days with contingency. Comparison with execution remains unverified until time basis and scope are reconciled.')
            elif x.startswith('The financial weight of this unplanned time is quantified'):
                set_text(p,'No verified financial cost has been assigned to the legacy NPT attribution. Rates, scope and contractual responsibility require reconciliation.')
            elif x.startswith('وزن مالی این زمان غیربرنامه‌ریزی‌شده'):
                set_text(p,'برای انتساب قدیمی زمان غیرمولد هزینه مالی تأییدشده محاسبه نشده است. نرخ‌ها، دامنه و مسئولیت قراردادی باید تطبیق داده شوند.')
            elif x.startswith(('Use the reconstructed daily cost','Treat the reconstructed daily burn','استفاده از هزینه روزانه بازسازی‌شده')):
                set_text(p,'مبنای مالی برنامه چاه بعدی باید از قرارداد و اسناد هزینه تأییدشده استخراج و به دامنه قابل مقایسه محدود شود.' if fa else 'Set future-well cost benchmarks only from validated contracts and cost records on a comparable scope.')
            elif x.startswith(('The legacy report combines a rig day-rate estimate','Table 8-2 source:','در برآورد قبلی، نرخ روزانه دکل','منبع جدول ۸-۲:')):
                s=('دو مبلغ هر ردیف به‌ترتیب ستون مجموع و کار اضافی شیت است؛ جمع‌پذیری آنها احراز نشده است. ' if fa else 'The two values per row are source total and extra-work columns; additivity is not established. ') if well in ('10','122WO1') else ''
                group=(' طبقات از جمع مقادیر عددی ذخیره‌شده در بازه سلول هر ردیف به دست آمده‌اند؛ درستی فرمول‌ها و دامنه هنوز تأیید نشده است. ' if fa else ' Categories are derived by summing numeric cached values within each cited row range; scope and formulas remain unverified. ') if leaf_groups(well) else ''
                source=('منبع جدول ۸-۲: '+file+('، شیت '+sheet if sheet else '')+'. ' if fa else 'Table 8-2 source: '+file+(' / '+sheet if sheet else '')+'. ')
                caution=('مبالغ، مقادیر ذخیره‌شده منبع هستند، نه مخارج واقعی یا بودجه مصوب قابل مقایسه. مقادیر خالی یا خطا صفر نیستند؛ نرخ‌ها، دامنه، فرمول و تأییدیه باید تطبیق شود. سری هزینه روزانه تأیید نشده است.' if fa else 'These cached source amounts are not actual expenditure or an approved comparable budget. Missing or erroneous values are not zeros; rates, scope, formulas and approvals require reconciliation. No daily expenditure series has been verified.')
                set_text(p,source+s+group+caution)
            elif x.startswith('Table 8-1 '):set_text(p,'Table 8-1  Source financial and schedule anchors (unreconciled).')
            elif x.startswith('Table 8-2 '):set_text(p,'Table 8-2  CBS source categories and cell references (unreconciled).')
            elif x.startswith('Table 8-3 '):set_text(p,'Table 8-3  CBS reconciliation checks.')
            elif x.startswith('جدول ۸-۱ '):set_text(p,'جدول ۸-۱  مقادیر مرجع مالی یا زمانی منبع، پیش از تطبیق.')
            elif x.startswith('جدول ۸-۲ '):set_text(p,'جدول ۸-۲  طبقات منبع و نشانی سلول‌ها، پیش از تطبیق.')
            elif x.startswith('جدول ۸-۳ '):set_text(p,'جدول ۸-۳  موارد نیازمند تطبیق.')
        i=next(i for i,p in enumerate(doc.paragraphs) if p.text.startswith(('Figure W5','شکل W5')))
        bl=doc.paragraphs[i-1]._element.xpath('.//a:blip')[0];rid=bl.get(qn('r:embed'));doc.part.related_parts[rid]._blob=png
        shape=next(s for s in doc.inline_shapes if any(b.get(qn('r:embed'))==rid for b in s._inline.xpath('.//a:blip')));shape.height=int(shape.width*6/11)
        cap=('شکل W5 ـ زمان برنامه دکل طبق منبع، نه نمودار هزینه' if fa else 'Figure W5 — Source rig schedule, not a cost series') if well=='85' else (f'شکل W5 ـ طبقات هزینه مندرج در منبع، چاه {well} (تطبیق‌نشده)' if fa else f'Figure W5 — CBS source cost categories, AZNS-{well} (unreconciled)')
        set_text(doc.paragraphs[i],cap)
        for t in doc.tables:
            if t.rows:
                for c in t.rows[0].cells:
                    for p in c.paragraphs:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        doc.save(path)
    print(well,len(rows),'categories',flush=True)
if __name__=='__main__':main()
