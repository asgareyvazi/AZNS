#!/usr/bin/env python3
"""Interim safety notice and typography repair; NOT a full engineering rebuild.
Run from repository root with python-docx and openpyxl installed.
"""
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import load_workbook

root=Path('.')
a=load_workbook(root/'CBS AZNS-122WO1.xlsx', read_only=True, data_only=True)['Cost Breakdown-122']
b=load_workbook(root/'CBS AZNS-85.xlsx', read_only=True, data_only=True)['AZNS-085 WO. Extra Work']
base,extra=a['G190'].value,a['H190'].value
plan,cont=b['K54'].value,b['O54'].value
assert abs(base-4587374.5)<.01 and abs(extra-1304440.016)<.01
assert abs(plan-48.0729167)<.001 and abs(cont-51.8895833)<.001

def font(r,fa):
    r.font.name='Times New Roman'
    pr=r._element.get_or_add_rPr(); fs=pr.find(qn('w:rFonts'))
    if fs is None:
        fs=OxmlElement('w:rFonts'); pr.insert(0,fs)
    for key in ('ascii','hAnsi','eastAsia'): fs.set(qn('w:'+key),'Times New Roman')
    fs.set(qn('w:cs'),'B Nazanin' if fa else 'Times New Roman')

def paras(container):
    for p in container.paragraphs: yield p
    for t in container.tables:
        yield from table_paras(t)

def table_paras(t):
    for row in t.rows:
        for c in row.cells:
            yield from paras(c)

def tables(container):
    for t in container.tables:
        yield t
        for row in t.rows:
            for c in row.cells: yield from tables(c)

for path in sorted(root.glob('*Report_??.docx')):
    fa=path.stem.endswith('_FA'); doc=Document(path)
    # An explicit notice is preferable to presenting the legacy graphs as verified.
    first=doc.paragraphs[0]
    notice=first.insert_paragraph_before()
    notice.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=notice.add_run(('هشدار بازبینی مهندسی — نسخه فعلی برای تصمیم‌گیری عملیاتی یا مالی معتبر نیست. نمودار هزینه تجمعی برآورد خطی است، نه ثبت هزینه روزانه؛ سایر نمودارها، برآوردهای بودجه و انتساب علل و سازندها نیازمند بازسازی و تطبیق با اسناد اصلی‌اند. اصلاحات مشخص و محدود این نسخه را در یادداشت بازبینی مخزن ببینید.' if fa else 'ENGINEERING REVIEW WARNING — This legacy report is not validated for operational or financial decisions. The cumulative-cost chart is a linear allocation, not recorded daily expenditure; other charts, budgets and cause/formation attribution require reconstruction and reconciliation to source records. See the repository review note for the limited corrections made.'))
    r.bold=True; font(r,fa)
    for section in doc.sections:
        for part in (section.header,section.footer):
            for p in paras(part): p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    containers=[doc]+[part for sec in doc.sections for part in (sec.header,sec.footer)]
    for container in containers:
        for t in tables(container):
            if fa:
                pr=t._tbl.tblPr
                if pr.find(qn('w:bidiVisual')) is None:
                    el=OxmlElement('w:bidiVisual')
                    # bidiVisual precedes table width/borders/layout in ECMA-376 tblPr.
                    before=next((x for x in pr if x.tag in [qn('w:'+n) for n in ('tblW','jc','tblCellSpacing','tblInd','tblBorders','shd','tblLayout','tblCellMar','tblLook')]),None)
                    if before is not None: before.addprevious(el)
                    else: pr.append(el)
            if t.rows:
                for c in t.rows[0].cells:
                    for p in c.paragraphs: p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        for p in paras(container):
            if p.style and p.style.name in ('Heading 1','Heading 2','Heading 3'):
                p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs: font(run,fa)
    for style in doc.styles:
        if style.type in (1,2):
            style.font.name='Times New Roman'
            pr=style.element.get_or_add_rPr(); fs=pr.find(qn('w:rFonts'))
            if fs is None: fs=OxmlElement('w:rFonts'); pr.insert(0,fs)
            for key in ('ascii','hAnsi','eastAsia'): fs.set(qn('w:'+key),'Times New Roman')
            fs.set(qn('w:cs'),'B Nazanin' if fa else 'Times New Roman')
    if path.name.startswith('Field_Master'):
        for p in doc.paragraphs:
            for run in p.runs:
                text=run.text
                text=text.replace('It sits within the normal band of the portfolio on both time and cost.','A sound comparison requires a defined baseline and reconciliation; schedule and cost performance should be assessed separately.')
                text=text.replace('این عملیات در هر دو محور زمان و هزینه در محدوده طبیعی سبد قرار دارد.','مقایسه عملکرد زمان و هزینه مستلزم مبنای تعریف‌شده و تطبیق اسناد است؛ این دو محور باید جداگانه ارزیابی شوند.')
                text=text.replace('against a CBS budget of no published budget','against a CBS budget not reconciled in this report')
                text=text.replace('بودجه CBS معادل بدون بودجه منتشرشده','بودجه CBS که در این گزارش تطبیق نشده است')
                if text != run.text: run.text=text
    if path.name.startswith(('122WO1_','85_','Field_Master','Field_Executive')):
        p=notice.insert_paragraph_before()
        p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        r=p.add_run(('اصلاح منبع: در CBS چاه 122WO1، شیت Cost Breakdown-122، سلول‌های G190 و H190 به‌ترتیب ۴٬۵۸۷٬۳۷۴٫۵۰ دلار برنامه و ۱٬۳۰۴٬۴۴۰٫۰۲ دلار کار اضافی را نشان می‌دهند؛ جمع حسابی ۵٬۸۹۱٬۸۱۴٫۵۲ دلار است، نه لزوماً بودجه مصوب قابل مقایسه با هزینه واقعی. در CBS چاه 85، شیت AZNS-085 WO. Extra Work، سلول‌های K54 و O54 به‌ترتیب ۴۸٫۰۷ و ۵۱٫۸۹ روز را نشان می‌دهند؛ فعالیت‌های Rig-less در ردیف‌های ۴۱ تا ۴۹ زمان صفر در این برنامه دکل دارند، نه لزوماً مدت یا هزینه صفر.' if fa else 'SOURCE CORRECTION: CBS AZNS-122WO1, Cost Breakdown-122, G190 and H190 show USD 4,587,374.50 plan and USD 1,304,440.02 extra work. Their arithmetic sum is USD 5,891,814.52, not necessarily an approved budget comparable with actual expenditure. CBS AZNS-85, AZNS-085 WO. Extra Work, K54 and O54 show 48.07 and 51.89 days. Rig-less activities in rows 41–49 have zero time in this rig schedule, not necessarily zero elapsed duration or cost.'))
        r.bold=True; font(r,fa)
    doc.save(path)
print('Updated',len(list(root.glob('*Report_??.docx'))),'documents')
