#!/usr/bin/env python3
"""Append a small manually reviewed, source-cited program comparison.
Not a comprehensive semantic reading of any well program.
"""
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm,Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import load_workbook
from pypdf import PdfReader

pdf=PdfReader('AZNS-207 Program.pdf')
for page,phrase in [(14,'1021-1368'),(14,'1368-1583'),(14,'1994-2178'),(15,'2343-2418'),(15,'2635-3270'),(16,'3915-4180')]:
    assert phrase in pdf.pages[page-1].extract_text(),(page,phrase)
sheet=load_workbook('AZNS207Pack.xlsx',read_only=True,data_only=True)['155-2 DDR Data']
rows=[('Gachsaran','گچساران',1021,72,14),('Asmari','آسماری',1368,73,14),('Pabdeh','پابده',1994,76,14),('Upper Gurpi','گورپی بالایی',2178,77,15),('Tarbur','تاربور',2343,78,15),('Lower Gurpi','گورپی پایینی',2418,79,15),('Ilam','ایلام',2560,80,15),('Laffan','لافان',2629,81,15),('Sarvak','سروک',2635,82,15),('Kazhdumi','کژدمی',3270,83,15),('Dariyan','داریان',3464,84,15),('Gadvan','گدوان',3622,85,16),('Fahliyan','فهلیان',3915,86,16)]
assert [sheet[f'P{r}'].value for _,_,_,r,_ in rows]==[1020,1368,1991,2178,2346,2423,2563,2633,2639,3266,3467,3626,3968]

def style(p,fa,center=False):
    if center:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    if fa:
        pr=p._p.get_or_add_pPr()
        if pr.find(qn('w:bidi')) is None:
            b=OxmlElement('w:bidi');before=next((x for x in pr if x.tag in [qn('w:'+n) for n in ('jc','outlineLvl','rPr','sectPr','pPrChange')]),None)
            if before is not None:before.addprevious(b)
            else:pr.append(b)
    for run in p.runs:
        run.font.name='Times New Roman';run.font.size=Pt(9)
        rp=run._element.get_or_add_rPr();ft=rp.find(qn('w:rFonts'))
        if ft is None:ft=OxmlElement('w:rFonts');rp.insert(0,ft)
        ft.set(qn('w:cs'),'B Nazanin' if fa else 'Times New Roman')
        if fa and rp.find(qn('w:rtl')) is None:rp.append(OxmlElement('w:rtl'))

def fa_digits(s):return str(s).translate(str.maketrans('0123456789','۰۱۲۳۴۵۶۷۸۹'))

def heading(d,text,fa):
    h=d.add_heading(text,level=2);style(h,fa,True)
    for r in h.runs:r.font.size=Pt(13)

def add(d,text,fa):
    p=d.add_paragraph(text);style(p,fa)

for lang in ('EN','FA'):
    fa=lang=='FA';d=Document(f'207_Drilling_Engineering_Report_{lang}.docx')
    title='مقایسه اولیه برنامه و جدول سازندی 207' if fa else 'Preliminary 207 programme versus DDR formation comparison'
    if any(p.text==title for p in d.paragraphs):raise ValueError('already appended')
    heading(d,title,fa)
    add(d,('ستون برنامه از بخش‌های ۱٫۵٫۱ صفحات فیزیکی ۱۴ تا ۱۶ فایل AZNS-207 Program.pdf و ستون DDR از AZNS207Pack.xlsx، شیت 155-2 DDR Data، ستون P (با عنوان m MD) گرفته شده است. اختلاف صرفاً تفاضل عددی DDR منهای برنامه است: تا تأیید مرجع عمق، نسخه برنامه و ماهیت ثبت، انحراف واقعی رأس سازند محسوب نمی‌شود.' if fa else 'Programme depths are transcribed from section 1.5.1 on physical pages 14–16 of AZNS-207 Program.pdf; DDR values are from column P, labelled m MD, in AZNS207Pack.xlsx sheet 155-2 DDR Data. Delta is simply DDR minus programme: it is not a verified actual formation-top deviation until depth datum, programme revision and the nature of the DDR entry are confirmed.'),fa)
    t=d.add_table(rows=1,cols=4);t.style='Table Grid';t.autofit=False
    if fa:
        pr=t._tbl.tblPr;b=OxmlElement('w:bidiVisual');pr.find(qn('w:tblW')).addprevious(b)
    for c,text in zip(t.rows[0].cells,(['سازند','برنامه MD (m) / صفحه','DDR MD (m) / سلول','تفاضل (m)'] if fa else ['Formation','Programme MD (m) / page','DDR MD (m) / cell','Difference (m)'])):c.text=text
    for name,fa_name,plan,row,page in rows:
        actual=sheet[f'P{row}'].value;cell=t.add_row().cells
        vals=[fa_name if fa else name, f'{plan} / {page}', f'{actual} / P{row}', f'{actual-plan:+d}']
        if fa: vals=[vals[0],fa_digits(vals[1]),str(fa_digits(actual))+' / P'+str(row),fa_digits(vals[3])]
        for c,v in zip(cell,vals):c.text=v
    widths=(4.3,4.3,4.5,3.9)
    for col,w in zip(t.columns,widths):col.width=Cm(w)
    for index,row in enumerate(t.rows):
        for c,w in zip(row.cells,widths):
            c.width=Cm(w)
            for p in c.paragraphs:style(p,fa,index==0)
    add(d,('برای مثال، برنامه رأس فهلیان را ۳۹۱۵ متر و جدول DDR مقدار ۳۹۶۸ متر ثبت کرده است؛ اختلاف ۵۳ متر نیازمند بررسی زمین‌شناس است. برنامه (صفحه ۱۲) شوری آب فهلیان را ۱۴۳٬۰۰۰ ppm معادل NaCl گزارش می‌کند و به سیال خورنده اشاره دارد؛ این مقدار اندازه‌گیری جاری چاه یا تأیید H2S نیست. برنامه در صفحات ۱۴ و ۱۵ لیتولوژی و سازندهای مجاور را شرح می‌دهد؛ بدون تطبیق رخداد و عمق، هرزروی به هیچ سازند خاصی نسبت داده نمی‌شود.' if fa else 'For example, the programme gives the Fahliyan top as 3,915 m and the DDR table lists 3,968 m, a 53 m numerical difference requiring geological review. The programme (page 12) reports Fahliyan water salinity of 143,000 ppm as NaCl and notes corrosive reservoir fluids; this is not a current well measurement or confirmation of H2S. Pages 14–15 describe stratigraphy and lithology; losses are not assigned to a formation without matching the actual event and depth.'),fa)
    d.save(f'207_Drilling_Engineering_Report_{lang}.docx')
    for well in ('10','75'):
        path=f'{well}_Drilling_Engineering_Report_{lang}.docx';doc=Document(path)
        title2='یادداشت مستند برنامه چاه' if fa else 'Well-program evidence note'
        if any(p.text==title2 for p in doc.paragraphs):raise ValueError('already appended')
        heading(doc,title2,fa)
        if well=='10':
            text=('برنامه AZNS-010 Workover Program Rev.06.pdf در صفحه فیزیکی ۶ چاه را متروکه توصیف و رفع تعلیق، مشبک‌کاری آستری ۷ اینچ در فهلیان و تکمیل را برنامه‌ریزی می‌کند. صفحه ۱۷ تاریخچه حفاری تا عمق ۵۳۴۰ متر و رهاسازی با پلاگ سیمانی را ذکر می‌کند؛ همچنین از تنگی، ماهیگیری و هرزروی در عملیات تاریخی سخن می‌گوید. این موارد سابقه‌اند نه اثبات وقوع همان حوادث در عملیات تعمیری اخیر. هشدار داده‌ای: در جدول سازندی 1-2 Data نام‌های K111:K112 به‌ترتیب Aghajari و Gachsaran هستند، اما در 83-2 Data مقدار Surface به K111 افزوده و نام‌ها جابه‌جا شده‌اند؛ ارقام O112:Q112 باید از نظر هم‌ترازی پیش از استفاده تطبیق شوند.' if fa else 'AZNS-010 Workover Program Rev.06.pdf, physical page 6, describes an abandoned well and plans de-suspension, perforation of the 7-inch liner in Fahliyan and completion. Page 17 records historical drilling to 5,340 m and abandonment with cement plugs, and describes tight hole, fishing and losses. These are historical events, not proof of the same incidents during the later workover. Data warning: in 1-2 Data K111:K112 are Aghajari and Gachsaran; in 83-2 Data Surface was inserted at K111 and the names shifted. Reconcile O112:Q112 row alignment before interpreting the depths.')
        else:
            text=('برنامه AZNS-075 Workover-1 Program Rev.05.pdf در صفحه فیزیکی ۷ برداشتن تکمیل موجود و راندن رشته تکمیلی جدید را به‌عنوان اهداف اصلی ذکر می‌کند. عنوان بخش ۱۰٫۲ در صفحه ۸۶، رشته تکمیلی اولین ران را ناموفق و متعلق به Weqixin معرفی می‌کند؛ متن قابل استخراج در این صفحه علت مکانیکی خرابی را توضیح نمی‌دهد. برای تعیین مکانیسم خرابی باید شماتیک، گزارش خرابی و سوابق عملیات بررسی شوند.' if fa else 'AZNS-075 Workover-1 Program Rev.05.pdf, physical page 7, identifies removing the completion and running a new completion string as the main workover objectives. The section 10.2 heading on page 86 identifies the first-run completion string as failed and names Weqixin; the text extract on that page does not establish the mechanical failure mechanism. Review the schematic, failure investigation and operational records before attributing root cause.')
        add(doc,text,fa);doc.save(path)
print('Added program evidence to 207, 10 and 75 reports (EN/FA)')
