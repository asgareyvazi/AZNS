#!/usr/bin/env python3
"""Repair field-figure association; withdraw unsupported financial bridge.
M2/M3 are explicitly interim displays of inventory and legacy report tables.
Other retained legacy images remain unvalidated.
"""
from pathlib import Path
from io import BytesIO
import json,re,subprocess
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.oxml.ns import qn

# Fixed pre-repair commit supplies the original mapping, so reruns do not rotate again.
master=Document(BytesIO(subprocess.check_output(['git','show','060d826:Field_Master_Drilling_Engineering_Report_EN.docx'])))
original={}
for i,p in enumerate(master.paragraphs):
    m=re.match(r'^Figure M(\d+)\b',p.text)
    if not m:continue
    b=master.paragraphs[i-1]._element.xpath('.//a:blip')[0];rid=b.get(qn('r:embed'))
    original[int(m.group(1))]=master.part.related_parts[rid].blob
assert set(original)==set(range(1,11))

# The old master M2..M10 were a one-position rotation: M2 held the old M3,
# and M10 held the old M2. Capture bytes before replacing them.
remap={1:1,4:3,5:4,6:5,7:6,8:7,9:8,10:9}

inventory=json.loads(Path('evidence/source_inventory.json').read_text())
packs=[(f['file'].removeprefix('AZNS').removesuffix('Pack.xlsx'),f['sheets']) for f in inventory['files'] if f['file'].startswith('AZNS') and f['file'].endswith('Pack.xlsx')]
assert len(packs)==13

def inventory_png():
    pairs=sorted(packs,key=lambda x:-x[1]);fig,ax=plt.subplots(figsize=(11,6.8),dpi=300);fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white')
    ax.barh(range(len(pairs)),[x[1] for x in pairs],color='#28688d');ax.set_yticks(range(len(pairs)),[x[0] for x in pairs],fontsize=9);ax.invert_yaxis()
    ax.set_xlabel('Worksheets per original Pack workbook');ax.set_xlim(0,max(x[1] for x in pairs)*1.25);ax.grid(axis='x',alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0)
    for i,(_,v) in enumerate(pairs):ax.text(v+4,i,str(v),va='center',fontsize=8.5)
    fig.suptitle('AZNS | Source pack inventory',x=.06,y=.98,ha='left',fontweight='bold',fontsize=15)
    fig.text(.06,.935,'Worksheet counts are not report-day counts and do not demonstrate data completeness.',fontsize=9,color='#9b431c')
    fig.text(.06,.03,'Source: evidence/source_inventory.json; count includes all sheets, including non-DDR sheets.',fontsize=9,color='#9b431c')
    fig.subplots_adjust(left=.19,right=.96,top=.88,bottom=.12);b=BytesIO();fig.savefig(b,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return b.getvalue()

def financial_png():
    t=next(t for t in master.tables if t.rows and t.rows[0].cells[0].text=='Operation' and len(t.rows[0].cells)==8)
    rows=[]
    for r in t.rows[1:]:
        cells=r.cells
        if cells[0].text=='Portfolio':continue
        def val(s):
            try:return float(s.replace(',','').strip())
            except ValueError:return None
        rows.append((cells[0].text.replace('AZNS-',''),val(cells[5].text),val(cells[6].text)))
    assert len(rows)==13 and all(a is not None for _,a,_ in rows)
    fig,ax=plt.subplots(figsize=(11,6.8),dpi=300);fig.patch.set_facecolor('#f7f9fc');ax.set_facecolor('white');y=list(range(len(rows)))
    for i,(name,a,b) in enumerate(rows):
        ax.barh(i-.19,a,height=.34,color='#28688d')
        if b is not None:ax.barh(i+.19,b,height=.34,color='#d1a354')
        else:ax.text(.15,i+.19,'Budget not reconciled',va='center',fontsize=8,color='#9b431c')
    ax.set_yticks(y,[x[0] for x in rows],fontsize=9);ax.invert_yaxis();ax.set_xlim(0,max(max(a,b or 0) for _,a,b in rows)*1.24)
    ax.set_xlabel('Legacy report estimate (USD million)');ax.grid(axis='x',alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0)
    from matplotlib.patches import Patch
    handles=[Patch(color='#28688d',label='Reconstructed spend (unverified)'),Patch(color='#d1a354',label='Budget entry (unreconciled)')]
    fig.legend(handles=handles,loc='upper right',bbox_to_anchor=(.97,.89),ncol=2,fontsize=8,frameon=False)
    fig.suptitle('AZNS | Provisional operation cost comparison',x=.06,y=.98,ha='left',fontweight='bold',fontsize=15)
    fig.text(.06,.935,'Rounded legacy table values; no field-total saving or NPT causal bridge is implied.',fontsize=9,color='#9b431c')
    fig.text(.06,.03,'Source: field master legacy portfolio table; budgets and scope require reconciliation, including 122WO1 and 85.',fontsize=8.5,color='#9b431c')
    fig.subplots_adjust(left=.21,right=.96,top=.83,bottom=.12);b=BytesIO();fig.savefig(b,format='png',dpi=300,facecolor=fig.get_facecolor());plt.close(fig);return b.getvalue()

generated={2:inventory_png(),3:financial_png()}
for kind in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
    fa=lang=='FA';path=Path(f'Field_{kind}_Report_{lang}.docx');d=Document(path)
    found=[]
    for i,p in enumerate(d.paragraphs):
        m=re.match(r'^(?:Figure|شکل) M(\d+)\b',p.text)
        if not m:continue
        n=int(m.group(1));bl=d.paragraphs[i-1]._element.xpath('.//a:blip')
        if len(bl)!=1:raise ValueError((str(path),n))
        rid=bl[0].get(qn('r:embed'));d.part.related_parts[rid]._blob=generated[n] if n in generated else original[remap[n]]
        shape=next(s for s in d.inline_shapes if any(b.get(qn('r:embed'))==rid for b in s._inline.xpath('.//a:blip')))
        # Inline part may have changed aspect ratio after reassignment.
        from PIL import Image
        im=Image.open(BytesIO(d.part.related_parts[rid].blob));shape.height=int(shape.width*im.height/im.width)
        if n==2:
            p.runs[0].text='شکل M2 ـ تعداد شیت‌های کتاب‌کارهای مبنا (نه سنجه کامل‌بودن داده)' if fa else 'Figure M2 — Source pack worksheet inventory (not a completeness metric)'
            for r in p.runs[1:]:r.text=''
        elif n==3:
            p.runs[0].text='شکل M3 ـ مقایسه موقت هزینه و بودجه ثبت‌شده در گزارش قبلی' if fa else 'Figure M3 — Provisional cost and budget entries from the legacy report'
            for r in p.runs[1:]:r.text=''
        elif n==10:
            p.runs[0].text='شکل M10 ـ افت سیال گزارش‌شده در برابر عمق؛ انتساب سازندی تأیید نشده است' if fa else 'Figure M10 — Reported fluid losses by depth; formation attribution unverified'
            for r in p.runs[1:]:r.text=''
        found.append(n)
    assert set(found)==({1,2,3,4,5,6,7,8,9,10} if kind.startswith('Master') else {1,2,3,4,5,6,8,9,10}),path
    for p in d.paragraphs:
        text=p.text
        if text.startswith('The credibility of every number in this report depends on'):
            replacement='The inventory chart counts worksheets in each original Pack file. It does not establish completeness, semantic extraction, or the availability of any particular variable. The source index provides traceable cell and page references; engineering validation remains necessary.'
        elif text.startswith('اعتبار هر عدد در این گزارش به آنچه واقعاً در آرشیو موجود بوده'):
            replacement='نمودار موجودی، تعداد شیت‌های هر فایل اصلی Pack را نشان می‌دهد؛ از آن نمی‌توان کامل‌بودن داده یا استخراج و صحت یک متغیر خاص را نتیجه گرفت. نمایه منابع ارجاع سلول و صفحه فراهم می‌کند، اما اعتبارسنجی مهندسی همچنان لازم است.'
        elif text.startswith('The financial dashboard sets budget against reconstructed actual'):
            replacement='The chart transcribes rounded budget entries and reconstructed spend from the legacy portfolio table. Entries and scope are not reconciled, including 122WO1 and 85. The previous NPT-to-budget bridge was withdrawn because its residual was not independently evidenced and NPT cost can overlap existing cost categories.'
        elif text.startswith('داشبورد مالی، بودجه را در برابر هزینه واقعی بازسازی‌شده'):
            replacement='نمودار، ارقام گرد‌شده بودجه و هزینه بازسازی‌شده را از جدول قبلی نمایش می‌دهد. ارقام و دامنه کار، از جمله برای 122WO1 و 85، هنوز تطبیق نشده‌اند. پل قبلی بودجه و NPT حذف شد، زیرا باقیمانده آن شواهد مستقل ندارد و هزینه NPT ممکن است با طبقات هزینه موجود هم‌پوشانی داشته باشد.'
        elif text.startswith('The bridge from budget to actual runs through three terms:'):
            replacement='The prior budget-to-actual bridge has been withdrawn. The reported NPT exposure is an unverified allocation that may overlap existing cost categories; the scope residual cannot be treated as a demonstrated cause of savings. Reconcile the underlying budgets, variations and actual service records before making financial decisions.'
        elif text.startswith('پل از بودجه به واقعی از سه جمله می‌گذرد:'):
            replacement='پل قبلی بودجه به هزینه واقعی حذف شد. ریسک مالی NPT یک تخصیص تأییدنشده است و ممکن است با طبقات هزینه موجود هم‌پوشانی داشته باشد؛ باقیمانده دامنه کار را نمی‌توان علت اثبات‌شده صرفه‌جویی دانست. بودجه، تغییرات و سوابق واقعی خدمات باید پیش از تصمیم مالی تطبیق شوند.'
        elif text.startswith('Two field-wide exhibits are reproduced here because'):
            replacement='The source inventory shows worksheet counts, not data completeness. The legacy depth exhibit shows reported fluid-loss points only; it does not demonstrate formation attribution, stuck-pipe locations or a verified hazard mechanism.'
        elif text.startswith('دو نمودار میدانی در اینجا بازتولید می‌شود'):
            replacement='نمودار موجودی، تعداد شیت‌ها را نشان می‌دهد نه کامل‌بودن داده. تصویر عمقی قبلی فقط نقاط افت سیال گزارش‌شده را نمایش می‌دهد و انتساب به سازند، محل گیر رشته یا مکانیسم خطر را اثبات نمی‌کند.'
        else:continue
        p.runs[0].text=replacement
        for r in p.runs[1:]:r.text=''
    d.save(path);print(path,'relinked',len(found),'figures')
