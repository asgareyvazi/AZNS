#!/usr/bin/env python3
"""Append source-labelled, unvalidated formation-table snapshots to well reports."""
import json, re
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm,Pt

files={'10':'AZNS10Pack.xlsx','12':'AZNS12Pack.xlsx','40':'AZNS40Pack.xlsx','75':'AZNS75Pack.xlsx','85':'AZNS85Pack.xlsx','116':'AZNS116Pack.xlsx','122WO1':'AZNS122WO1Pack.xlsx','122WO2':'AZNS122WO2Pack.xlsx','127WO1':'AZNS127WO1Pack.xlsx','127WO2':'AZNS127WOPack.xlsx','127ST2':'AZNS127ST2Pack.xlsx','166':'AZNS166Pack.xlsx','207':'AZNS207Pack.xlsx'}
headers={'Forc. MD (m.)':'MD پیش‌بینی (m)','Forc. TVD (m.)':'TVD پیش‌بینی (m)','Act. MD (m.)':'MD واقعی (m)','Act. TVD (m.)':'TVD واقعی (m)','m MD':'MD (m)','m TVD':'TVD (m)','Lithology':'سنگ‌شناسی'}
names={'aghajari':'آغاجاری','gachsaran':'گچساران','asmari':'آسماری','asmary':'آسماری','pabdeh':'پابده','pabdeh / jahrum':'پابده / جهرم','jahrum':'جهرم','gurpi':'گورپی','u-gurpi':'گورپی بالایی','u.gurpi':'گورپی بالایی','l-gurpi':'گورپی پایینی','l.gurpi':'گورپی پایینی','l. gurpi':'گورپی پایینی','tarbur':'تاربور','ilam':'ایلام','laffan':'لافان','lafan':'لافان','sarvak':'سروک','kazhdumi':'کژدمی','kazhdomi':'کژدمی','dariyan':'داریان','gadvan':'گدوان','fahliyan':'فهلیان','fahlian':'فهلیان','fahlian top':'رأس فهلیان','garau':'گرو','surface':'سطح'}

def display(s):
    try:
        x=float(s)
        return f'{x:.4f}'.rstrip('0').rstrip('.')
    except (TypeError,ValueError):return s

def prepare_p(p,fa,center=False):
    if center:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    if fa:
        ppr=p._p.get_or_add_pPr()
        if ppr.find(qn('w:bidi')) is None:
            b=OxmlElement('w:bidi');before=next((x for x in ppr if x.tag in [qn('w:'+n) for n in ('jc','outlineLvl','rPr','sectPr','pPrChange')]),None)
            if before is not None:before.addprevious(b)
            else:ppr.append(b)
    for r in p.runs:
        r.font.name='Times New Roman';r.font.size=Pt(9)
        rp=r._element.get_or_add_rPr();fonts=rp.find(qn('w:rFonts'))
        if fonts is None:fonts=OxmlElement('w:rFonts');rp.insert(0,fonts)
        fonts.set(qn('w:cs'),'B Nazanin' if fa else 'Times New Roman')
        if fa and rp.find(qn('w:rtl')) is None:rp.append(OxmlElement('w:rtl'))

def set_bidi(table):
    pr=table._tbl.tblPr
    if pr.find(qn('w:bidiVisual')) is None:
        b=OxmlElement('w:bidiVisual');before=pr.find(qn('w:tblW'))
        if before is not None:before.addprevious(b)
        else:pr.append(b)

def main():
    snapshots={s['file']:s for s in json.loads(Path('evidence/formation_table_snapshots.json').read_text())}
    assert set(files.values())==set(snapshots)
    for well,f in files.items():
        snap=snapshots[f]
        for lang in ('EN','FA'):
            fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
            title='پیوست شواهد خام سازندی (نیازمند تطبیق)' if fa else 'Appendix: source formation-table snapshot (unverified)'
            if any(p.text==title for p in d.paragraphs):raise ValueError('Appendix already present in '+str(path))
            d.add_page_break();h=d.add_heading(title,level=1);prepare_p(h,fa,True)
            for r in h.runs:r.font.size=Pt(14)
            p=d.add_paragraph((f'منبع: {f}، شیت {snap["sheet"]}. انتخاب بر اساس آخرین شیت به‌ترتیب دفترکار با جدول سازندی دارای عدد عمق است؛ در نبود عدد، آخرین جدول دارای نام سازند انتخاب شده است. این انتخاب الزاماً آخرین گزارش تاریخ‌دار نیست.' if fa else f'Source: {f}, sheet {snap["sheet"]}. Selected as the last sheet in workbook order with a recognized formation table containing a numeric depth; if none, the last table with formation names. This is not necessarily the latest dated report.'))
            prepare_p(p,fa)
            p=d.add_paragraph(('نام ستون‌ها همان برچسب منبع است، نه تأیید مستقل پیش‌بینی یا اندازه‌گیری. خانه‌های فاقد عنوان نیز جداگانه نشان داده شده‌اند؛ مقادیر خالی نامعلوم‌اند، نه صفر. اعداد برای نمایش تا چهار رقم اعشار گرد شده‌اند. ترتیب/هم‌ترازی، مرجع عمق، ناسازگاری احتمالی و سنگ‌شناسی باید با برنامه و گزارش زمین‌شناسی تطبیق شوند؛ از این جدول به‌تنهایی انتساب حادثه یا دستور عملیات برداشت نشود.' if fa else 'Column roles follow the source labels, not independent confirmation of prognosis or measurement. Unlabelled populated cells are shown separately; missing cells are unknown, not zero. Values are rounded to four decimals for display. Verify row alignment, depth datum, possible inconsistencies and lithology against the programme and geological report; this table alone does not establish event attribution or operating instructions.'))
            prepare_p(p,fa)
            for tab in snap['tables']:
                p=d.add_paragraph((f'عنوان و خانه جدول در منبع: {tab["title"]} / {tab["anchor"]}' if fa else f'Source table and anchor: {tab["title"]} / {tab["anchor"]}'));prepare_p(p,fa)
                t=d.add_table(rows=1,cols=2);t.style='Table Grid';t.autofit=False
                for col,w in zip(t.columns,(4,13)):col.width=Cm(w)
                if fa:set_bidi(t)
                t.rows[0].cells[0].text='سازند / سلول' if fa else 'Formation / cell'
                t.rows[0].cells[1].text='مقادیر ثبت‌شده / سلول' if fa else 'Source values / cell'
                for row in tab['rows']:
                    cells=t.add_row().cells;name=row['formation'];namecell=row['cells'][0]['cell']
                    cells[0].text=(names.get(name.lower(),name)+f' ({name}) / {namecell}' if fa else f'{name} / {namecell}')
                    parts=[]
                    for item in row['cells'][1:]:
                        label=item['header'];
                        if fa:
                            label=headers.get(label,('ستون بدون عنوان '+item['cell'].rstrip('0123456789')) if label.startswith('Unlabelled column') else label)
                        parts.append(f'{label}: {display(item["value"])} [{item["cell"]}]')
                    cells[1].text='\n'.join(parts) if parts else ('ثبت نشده' if fa else 'Not recorded')
                for row_index,row in enumerate(t.rows):
                    for cell,width in zip(row.cells,(4,13)):
                        cell.width=Cm(width)
                        for p in cell.paragraphs:prepare_p(p,fa,row_index==0)
            d.save(path)
        print(well, snap['sheet'])
if __name__=='__main__':main()
