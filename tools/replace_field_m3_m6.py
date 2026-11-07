#!/usr/bin/env python3
"""Replace unsupported portfolio cost bridge and liability plot with source QA views."""
from pathlib import Path
from io import BytesIO
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from replace_cbs_financials import mapping,books,set_text,celltext

wmap={'10':'AZNS10Pack.xlsx','12':'AZNS12Pack.xlsx','75':'AZNS75Pack.xlsx','85':'AZNS85Pack.xlsx','116':'AZNS116Pack.xlsx','122WO1':'AZNS122WO1Pack.xlsx','122WO2':'AZNS122WO2Pack.xlsx','127ST2':'AZNS127ST2Pack.xlsx','127WO1':'AZNS127WO1Pack.xlsx','127WO2':'AZNS127WOPack.xlsx','166':'AZNS166Pack.xlsx','207':'AZNS207Pack.xlsx','40':'AZNS40Pack.xlsx'}
tr={x['file']:x for x in json.loads(Path('evidence/drilling_track_observations.json').read_text())}

def m3():
    rows=[('10','Two work scopes / sheets'),('12','Scope review pending'),('75','Scope review pending'),('85','Schedule-only workbook'),('116','116 / 127 joint ceiling'),('122WO1','Plan / extra scopes'),('122WO2','Contingency material #REF!'),('127ST2','Subtotal + detail in total'),('127WO1','Source #REF! errors'),('127WO2','Procurement comments'),('166','Scope review pending'),('207','F-20 / material inclusion'),('40','Rig rate x quantity mismatch')]
    fig,ax=plt.subplots(figsize=(12,8),dpi=300);fig.patch.set_facecolor('#f7f9fc');ax.set_xlim(0,3);ax.set_ylim(len(rows)-.5,-.5);ax.axis('off')
    for i,(well,note) in enumerate(rows):
        for j,color in enumerate(('#d8eaf2' if well!='85' else '#fae9ce','#fae9ce','#f2e4e2')):
            ax.add_patch(plt.Rectangle((j,i-.43),.98,.86,facecolor=color,edgecolor='white'))
        ax.text(-.08,i,'AZNS-'+well,ha='right',va='center',fontsize=10,color='#223952')
        ax.text(.49,i,'Cost sheets' if well!='85' else 'Schedule only',ha='center',va='center',fontsize=9)
        ax.text(1.49,i,note,ha='center',va='center',fontsize=9)
        ax.text(2.49,i,'Pending',ha='center',va='center',fontsize=9,color='#a04332')
    for x,label in [(.49,'Source tabs'),(1.49,'Selected scope / formula issue'),(2.49,'Actual cost reconciliation')]:
        ax.text(x,-.9,label,ha='center',va='bottom',fontweight='bold',fontsize=9,color='#223952')
    fig.suptitle('Portfolio | Financial source review status',x=.05,y=.98,ha='left',fontsize=16,fontweight='bold',color='#223952')
    fig.text(.05,.925,'Selected source observations, not a complete audit. No actual portfolio spend or savings is established.',fontsize=9,color='#9b431c')
    fig.text(.05,.035,'A cost sheet is not proof of an approved budget. Issues require source, scope and formula reconciliation.',fontsize=8.5,color='#9b431c')
    fig.subplots_adjust(left=.18,right=.97,top=.84,bottom=.11);b=BytesIO();fig.savefig(b,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return b.getvalue()

def m6():
    wells=list(wmap);items=[]
    for well in wells:
        x=tr[wmap[well]];complete=[d for d in x['days'] if d['activity']['complete']]
        totals=[sum(d['activity'][key] for d in complete) for key in ('npt_h','unplanned_h','mixed_h')]
        items.append((well,len(complete),x['paired_sheet_candidates'],totals))
    fig,ax=plt.subplots(figsize=(12,8),dpi=300);fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white')
    for i,(well,done,total,values) in enumerate(items):
        left=0
        for val,color in zip(values,('#d0a153','#c65d42','#8a65a5')):
            ax.barh(i,val,left=left,height=.65,color=color);left+=val
        ax.text(1.015,i,f'{done}/{total}',transform=ax.get_yaxis_transform(),ha='left',va='center',fontsize=9,color='#344c61')
    maxx=max(sum(x[3]) for x in items);ax.set_xlim(0,maxx*1.05);ax.set_yticks(range(len(items)),['AZNS-'+x[0] for x in items]);ax.invert_yaxis();ax.set_xlabel('Hours in selected complete days (partial, not total NPT)',fontsize=10)
    ax.grid(axis='x',alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0)
    fig.suptitle('Portfolio | Daily source status codes',x=.05,y=.98,ha='left',fontsize=16,fontweight='bold',color='#223952')
    fig.text(.05,.93,'Hours use recognized 24-hour classified records only. Coverage = complete days / candidate sheet pairs.',fontsize=9,color='#9b431c')
    fig.legend(handles=[Patch(color='#d0a153',label='N / NPT'),Patch(color='#c65d42',label='UP / unplanned'),Patch(color='#8a65a5',label='NPT-UNPLN')],ncol=3,frameon=False,loc='upper center',bbox_to_anchor=(.55,.89),fontsize=9)
    fig.text(.05,.035,'Codes remain distinct; omitted days are not zero. No cause, contractor liability or financial cost is inferred.',fontsize=8.5,color='#9b431c')
    fig.subplots_adjust(left=.16,right=.86,top=.82,bottom=.12);b=BytesIO();fig.savefig(b,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return b.getvalue()

def main():
    figures={3:m3(),6:m6()}
    for typ in ('Master_Drilling_Engineering','Executive_Management'):
      for lang in ('EN','FA'):
        fa=lang=='FA';path=Path(f'Field_{typ}_Report_{lang}.docx');d=Document(path)
        for n,png in figures.items():
            i=next(i for i,p in enumerate(d.paragraphs) if p.text.startswith((f'Figure M{n}',f'شکل M{n}')))
            bl=d.paragraphs[i-1]._element.xpath('.//a:blip')[0];rid=bl.get(qn('r:embed'));d.part.related_parts[rid]._blob=png
            sh=next(s for s in d.inline_shapes if any(b.get(qn('r:embed'))==rid for b in s._inline.xpath('.//a:blip')));sh.height=int(sh.width*8/12)
            cap=(('شکل M3 ـ وضعیت بازبینی منابع مالی سبد' if fa else 'Figure M3 — Portfolio financial source review status') if n==3 else ('شکل M6 ـ کدهای زمان ثبت‌شده در منبع با پوشش ناقص' if fa else 'Figure M6 — Source-coded time in partially classified daily records'))
            set_text(d.paragraphs[i],cap)
            pre=d.paragraphs[i-2]
            if n==3:
                set_text(pre,'این شکل وضعیت بررسی منابع مالی را نشان می‌دهد، نه هزینه واقعی یا صرفه‌جویی سبد. وجود شیت به‌تنهایی بودجه مصوب را اثبات نمی‌کند؛ دامنه، فرمول و اسناد هزینه باید تطبیق شوند.' if fa else 'This figure shows the status of financial source review, not actual portfolio spend or savings. A cost worksheet alone does not prove an approved budget; scope, formulas and cost records require reconciliation.')
            else:
                set_text(pre,'این شکل صرفاً ساعات کدهای منبع را در گزارش‌های دارای ۲۴ ساعت طبقه‌بندی‌شده نشان می‌دهد؛ روزهای نامشخص حذف شده‌اند. کدهای NPT و غیربرنامه‌ای جدا هستند و نمودار، علت، مسئولیت قراردادی یا هزینه مالی را تعیین نمی‌کند.' if fa else 'This figure shows source status-code hours only for fully classified 24-hour records; unclassified days are omitted. NPT and unplanned codes remain distinct, and the figure does not determine cause, contractual liability or financial cost.')
        for p in d.paragraphs:
            x=p.text
            if x.startswith(('Every operation in the portfolio overran its programme','هر عملیات سبد از برنامه خود فراتر رفته')):
                set_text(p,'مدت برنامه‌ای و مدت اجرا باید برای هر عملیات از نظر دامنه و مبنای زمانی تطبیق شود؛ نمی‌توان نتیجه گرفت که همه عملیات از برنامه فراتر رفته‌اند.' if fa else 'Planned and actual duration must be reconciled by operation and time basis; it is not supported to conclude that every operation overran its programme.')
            elif x.startswith(('Two wells absorbed 37.4%','دو چاه ۳۷٫۴%')):
                set_text(p,'تمرکز هزینه در چند چاه تا تطبیق فاکتورها و دامنه عملیات اثبات نشده است.' if fa else 'The concentration of expenditure across wells is not established until invoices and operation scopes are reconciled.')
            elif x.startswith(('The data now exists to manage this. Every number','داده اکنون برای مدیریت این وضعیت وجود دارد. هر عدد')):
                set_text(p,'نمایه منابع، ارجاع سلول و صفحه را فراهم می‌کند، اما کامل‌بودن تحلیل یا هزینه واحد اندازه‌گیری‌شده را اثبات نمی‌کند.' if fa else 'The source index supplies cell and page references, but does not establish complete analysis or measured unit cost.')
            elif x.startswith(('Thirteen operations were executed between July 2023 and July 2026','سیزده عملیات میان تیر ۱۴۰۲ و تیر ۱۴۰۵')):
                set_text(p,'این نسخه ۱۳ عملیات را پوشش می‌دهد؛ جمع مالی و زمان ناوگان تا تطبیق اسناد، برنامه‌ها و پوشش گزارش‌ها تأیید نشده‌اند.' if fa else 'This interim portfolio covers 13 operations; financial totals and fleet time remain unverified pending reconciliation of records, plans and reporting coverage.')
            elif x.startswith(('Fleet utilisation is continuous from July 2023','بهره‌برداری ناوگان از تیر ۱۴۰۲ تا تیر ۱۴۰۵')):
                set_text(p,'زمان‌بندی، روزهای دکل و هزینه واحد ناوگان در متن قدیمی به تطبیق مستقل نیاز دارند؛ هزینه هر متر بدون فاکتور و تعریف دامنه قابل اتکا نیست.' if fa else 'Legacy fleet timelines, rig days and unit costs need independent reconciliation; cost per metre is not reliable without invoices and a consistent scope.')
            elif x in ('3. Where the Money Leaks','۳. نقطه نشت پول'):
                set_text(p,'۳. کدهای زمانی ثبت‌شده در منبع' if fa else '3. Source-coded time categories')
        # Mark retained legacy graphics as unverified; their underlying numbers
        # have not been rebuilt by this intervention.
        for p in d.paragraphs:
            if any(p.text.startswith((f'Figure M{n} ',f'شکل M{n} ')) for n in (1,4,5,7,8,9)) and not any(x in p.text for x in ('[UNVERIFIED]','[تأییدنشده]')):
                set_text(p,p.text+(' [تأییدنشده]' if fa else ' [UNVERIFIED]'))
        # Withdraw unsupported portfolio monetary totals in prominent summary cells.
        for t in d.tables:
            if not t.rows:continue
            head=t.rows[0].cells[0].text
            if head in ('Portfolio metric','شاخص سبد'):
                for r in t.rows[1:]:
                    if r.cells[0].text in ('Reconstructed spend','CBS budget','Portfolio variance','Financial cost of NPT','هزینه بازسازی‌شده','بودجه CBS','انحراف سبد','هزینه مالی زمان غیرمولد'):
                        celltext(r.cells[1],'تطبیق‌نشده' if fa else 'Not reconciled',fa)
                    if r.cells[0].text in ('Daily reports parsed','گزارش‌های روزانه پردازش‌شده'):
                        celltext(r.cells[1],'پوشش تحلیل روزانه متغیر است' if fa else 'Daily semantic coverage varies',fa)
            elif head in ('Measure','شاخص') and len(t.columns)==3:
                for r in t.rows[1:]:
                    if r.cells[0].text not in ('Rig days consumed','Unplanned time','روز دکل مصرف‌شده','زمان غیرمولد'):
                        for c in r.cells[1:]:celltext(c,'تطبیق‌نشده' if fa else 'Not reconciled',fa)
                    elif r.cells[0].text in ('Unplanned time','زمان غیرمولد'):
                        celltext(r.cells[2],'هزینه و انتساب تأیید نشده' if fa else 'Cost and attribution unverified',fa)
            elif head in ('Rig metric','شاخص دکل'):
                for r in t.rows[1:]:
                    if r.cells[0].text in ('Average daily burn (USD)','Cost per metre (USD)','نرخ متوسط روزانه (دلار)','هزینه هر متر (دلار)'):
                        for c in r.cells[1:]:celltext(c,'تطبیق‌نشده' if fa else 'Not reconciled',fa)
            elif head in ('Decision','تصمیم'):
                for r in t.rows[1:]:
                    celltext(r.cells[2],'تطبیق‌نشده' if fa else 'Not reconciled',fa)
                    if r.cells[0].text.startswith(('Adopt reconstructed unit cost','استفاده از هزینه واحد بازسازی')):
                        celltext(r.cells[0],'مبنای مناقصه بر اساس قرارداد و اسناد تأییدشده تعیین شود' if fa else 'Set tender benchmarks from verified contracts and cost records',fa)
            elif head in ('Service line','ردیف خدمت'):
                for r in t.rows[1:]:
                    for c in r.cells[1:]:celltext(c,'تطبیق‌نشده' if fa else 'Not reconciled',fa)
            elif head in ('Financial driver','محرکه مالی','Exposure','ریسک') and len(t.columns) in (3,4):
                for r in t.rows[1:]:
                    celltext(r.cells[1],'تطبیق‌نشده' if fa else 'Not reconciled',fa)
                    if head in ('Financial driver','محرکه مالی'):celltext(r.cells[2],'تطبیق‌نشده' if fa else 'Not reconciled',fa)
            elif len(t.rows)==1 and len(t.columns)==1:
                c=t.cell(0,0);x=c.text
                if x.startswith(('The portfolio is inside budget','سبد در مجموع زیر بودجه است','This master report treats the thirteen operations','این گزارش جامع، سیزده عملیات')):
                    celltext(c,'جمع هزینه، بودجه، انحراف و مسئولیت مالی سبد تا تطبیق اسناد و دامنه تأیید نشده است.' if fa else 'Portfolio expenditure, budget variance and financial liability are not established until records and scopes are reconciled.',fa)
            headers=[c.text for c in t.rows[0].cells]
            financecols={'Actual MUSD','Budget MUSD','Rig MUSD','Services MUSD','Total MUSD','Rig share','USD/day','Financial exposure (USD)','Share of NPT cost','ارزش در معرض','واقعی (م.دلار)','بودجه (م.دلار)','دکل (م.دلار)','خدمات (م.دلار)','جمع (م.دلار)','سهم دکل','دلار/روز','ریسک مالی (دلار)','سهم از هزینه غیرمولد'}
            if any(h in financecols for h in headers):
                for j,h in enumerate(headers):
                    if h in financecols or h in ('Var. %','Variance','انحراف ٪','انحراف'):
                        for r in t.rows[1:]:celltext(r.cells[j],'تطبیق‌نشده' if fa else 'Not reconciled',fa)
        for t in d.tables:
            if t.rows:
                for c in t.rows[0].cells:
                    for p in c.paragraphs:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        d.save(path);print(path,flush=True)
if __name__=='__main__':main()
