#!/usr/bin/env python3
"""Add a sourced, bilingual 85 rig-less schedule correction to both reports."""
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import load_workbook

sheet=load_workbook('CBS AZNS-85.xlsx',data_only=True,read_only=True)['AZNS-085 WO. Extra Work']
fa_steps={41:'استقرار تجهیزات اسلیک‌لاین، نصب غلاف محافظ و خروج و جمع‌آوری تجهیزات.',42:'راندن CT و اجرای اسیدشویی طبق برنامه و خروج رشته.',43:'استقرار اسلیک‌لاین، بازیابی غلاف محافظ و خروج و جمع‌آوری تجهیزات.',44:'پاکسازی جریانی پس از اسیدکاری: حداقل ۹۶ ساعت جریان؛ زمان آماده‌سازی و استقرار/جمع‌آوری لحاظ نشده است.',45:'راندن اسلیک‌لاین و نصب غلاف محافظ.',46:'راندن CT تا انتهای چاه، جابه‌جایی سیال با گازوئیل و خروج.',47:'راندن اسلیک‌لاین و نصب Plug و Prong در نشیمنگاه NO-GO و خروج.',48:'راندن اسلیک‌لاین و بازیابی غلاف محافظ.',49:'ایمن‌سازی چاه با نصب BPV در آویز لوله و بستن شیرهای X-mas Tree به‌جز BMV.'}
for i in range(41,50): assert str(sheet[f'I{i}'].value).lower().strip()=='rig less'
for lang in ('EN','FA'):
    fa=lang=='FA';path=f'85_Drilling_Engineering_Report_{lang}.docx';doc=Document(path)
    # insert immediately after the two existing interim review paragraphs
    anchor=doc.paragraphs[1]._p
    h=doc.add_heading('اصلاح تکمیلی: برنامه عملیات بدون دکل' if fa else 'Source correction: rig-less work sequence',level=2)
    h.alignment=WD_ALIGN_PARAGRAPH.CENTER;anchor.addnext(h._p);anchor=h._p
    p=doc.add_paragraph('منبع: CBS AZNS-85.xlsx، شیت AZNS-085 WO. Extra Work. این برنامه است، نه تأیید اجرا. سلول K54 مدت ۴۸٫۰۷ روز و O54 مدت ۵۱٫۸۹ روز با احتیاط را ثبت کرده است. ردیف ۴۰ پاکسازی پیش از اسید را ۷۲ ساعت برآورد می‌کند. ردیف‌های ۴۱ تا ۴۹ در ستون I به‌صورت Rig Less آمده‌اند و به مدت برنامه دکل زمانی اضافه نمی‌کنند؛ این موضوع به معنی صفر بودن مدت سپری‌شده یا هزینه خدمات نیست.' if fa else 'Source: CBS AZNS-85.xlsx, sheet AZNS-085 WO. Extra Work. This is a plan, not evidence of execution. K54 states 48.07 days and O54 states 51.89 days with contingency. Row 40 allows 72 hours for pre-acid cleanup. Rows 41–49 are marked Rig Less in column I and add no time to the rig schedule; this does not mean zero elapsed time or zero service cost.')
    anchor.addnext(p._p);anchor=p._p; intro=p
    t=doc.add_table(rows=1,cols=2);t.style='Table Grid'; t.autofit=False
    if fa:
        pr=t._tbl.tblPr; child=OxmlElement('w:bidiVisual'); before=pr.find(qn('w:tblW'))
        if before is not None: before.addprevious(child)
        else: pr.append(child)
    for col,width in zip(t.columns,(3,14)): col.width=Cm(width)
    t.rows[0].cells[0].text='سلول منبع' if fa else 'Source cell'
    t.rows[0].cells[1].text='فعالیت برنامه‌ریزی‌شده' if fa else 'Planned activity'
    for cell in t.rows[0].cells:
        for p0 in cell.paragraphs:p0.alignment=WD_ALIGN_PARAGRAPH.CENTER
    for i in range(41,50):
        cells=t.add_row().cells;cells[0].text=f'B{i} / I{i}';cells[1].text=fa_steps[i] if fa else str(sheet[f'B{i}'].value).replace('\n',' ')
    for row in t.rows:
        for cell,width in zip(row.cells,(3,14)): cell.width=Cm(width)
    anchor.addnext(t._tbl);anchor=t._tbl
    p=doc.add_paragraph('در یادداشت ردیف ۴۴، اجرای بدون دکل اسیدکاری و پاکسازی بعدی مشروط به تأیید کارفرما و پیمانکار است؛ در صورت شکست یکپارچگی چاه، تأمین دکل و خدمات لازم برای رفع مشکل به مسئولیت پیمانکار ذکر شده است. این شرط قراردادی باید با قرارداد مصوب تطبیق داده شود و جایگزین برنامه کنترل چاه نیست.' if fa else 'The note in row 44 makes rig-less acidizing and subsequent cleanup subject to contractor and client approval; in the event of well-integrity failure it assigns the contractor responsibility to provide a rig and required services to fix the problem. Verify this wording against the approved contract; it is not a substitute for a well-control plan.')
    anchor.addnext(p._p)
    note=p
    for paragraph in [h,intro,note]+[p for cell in t._cells for p in cell.paragraphs]:
        if fa:
            ppr=paragraph._p.get_or_add_pPr()
            if ppr.find(qn('w:bidi')) is None: ppr.append(OxmlElement('w:bidi'))
        for r in paragraph.runs:
            r.font.name='Times New Roman';pr=r._element.get_or_add_rPr();fonts=pr.find(qn('w:rFonts'))
            if fonts is None: fonts=OxmlElement('w:rFonts');pr.insert(0,fonts)
            fonts.set(qn('w:cs'),'B Nazanin' if fa else 'Times New Roman')
            if fa:
                rtl=OxmlElement('w:rtl');pr.append(rtl)
    doc.save(path)
print('Updated 85 EN/FA')
