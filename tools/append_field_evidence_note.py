#!/usr/bin/env python3
"""Add a concise source-specific evidence note to each field-level report."""
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

for kind in ('Master_Drilling_Engineering','Executive_Management'):
    for lang in ('EN','FA'):
        fa=lang=='FA';path=f'Field_{kind}_Report_{lang}.docx';d=Document(path)
        title='یادداشت تکمیلی شواهد و نقاط تصمیم' if fa else 'Supplement: source evidence and decision gates'
        if any(p.text==title for p in d.paragraphs):raise ValueError(f'already present: {path}')
        d.add_page_break();h=d.add_heading(title,level=1);h.alignment=WD_ALIGN_PARAGRAPH.CENTER
        en=[
          'Decision gate — The existing portfolio financial totals and attribution are provisional. Reconcile scope, approved contract variations, rig-less services, invoices and daily reports before approving a budget or allocating contractor responsibility. This note is a limited source check, not a replacement for full reconciliation.',
          '122WO1 — CBS AZNS-122WO1.xlsx, Cost Breakdown-122, G190 and H190 show USD 4,587,374.50 plan and USD 1,304,440.016 extra work. They cannot be combined as a like-for-like approved budget without reconciling scope and approvals.',
          '85 — CBS AZNS-85.xlsx, AZNS-085 WO. Extra Work, K54 and O54 provide 48.07 and 51.89 days respectively; rows 41–49 classify subsequent steps as rig-less. Row 44 proposes at least 96 hours of post-acid flow and makes rig-less acidizing/cleanup subject to contractor and client approval. Exclusion from rig time is not zero elapsed time or zero cost.',
          '207 — AZNS-207 Program.pdf, physical page 16, gives a prognosed Fahliyan top of 3,915 m MD; AZNS207Pack.xlsx, 155-2 DDR Data, P86 lists 3,968 m MD. The 53 m arithmetic difference requires confirmation of datums and whether the DDR value is an actual top before it is used as a geological deviation. Page 12 reports Fahliyan water salinity of 143,000 ppm as NaCl; it is programme information, not proof of current fluid properties.',
          '10 and 75 — AZNS-010 Workover Program Rev.06.pdf pages 6 and 17 describe prior abandonment and cement plugs. AZNS-075 Workover-1 Program Rev.05.pdf page 86 labels a first-run Weqixin completion string as failed but its text does not establish the failure mechanism. The 10 pack shows a potential formation-table row shift between 1-2 Data and 83-2 Data; reconcile before use.',
          'The well-report W4 replacement plots cumulative shares only among causes shown in each rounded legacy cause table. For 207 those rows total 988 h versus the legacy report’s 1,003 h NPT; the percentage basis therefore differs from total NPT. Neither these plots nor the new formation-table snapshots validate the original cause assignments.'
        ]
        fa_text=[
          'درگاه تصمیم — جمع‌های مالی و انتساب مسئولیت در گزارش موجود موقت‌اند. پیش از تصویب بودجه یا تخصیص مسئولیت پیمانکار، دامنه کار، تغییرات قراردادی مصوب، خدمات بدون دکل، صورت‌حساب‌ها و گزارش‌های روزانه تطبیق شوند. این یادداشت کنترل محدود منبع است و جایگزین تطبیق کامل نیست.',
          '122WO1 — در CBS AZNS-122WO1.xlsx، شیت Cost Breakdown-122، سلول‌های G190 و H190 به‌ترتیب ۴٬۵۸۷٬۳۷۴٫۵۰ دلار برنامه و ۱٬۳۰۴٬۴۴۰٫۰۱۶ دلار کار اضافی را نشان می‌دهند. جمع آنها بدون تطبیق دامنه و مصوبات الزاماً بودجه مصوب قابل مقایسه نیست.',
          '85 — در CBS AZNS-85.xlsx، شیت AZNS-085 WO. Extra Work، سلول‌های K54 و O54 به‌ترتیب ۴۸٫۰۷ و ۵۱٫۸۹ روز را ثبت می‌کنند؛ ردیف‌های ۴۱ تا ۴۹ مراحل بعدی را بدون دکل مشخص می‌کنند. ردیف ۴۴ حداقل ۹۶ ساعت جریان پس از اسید را پیشنهاد و اجرای بدون دکل را مشروط به تأیید کارفرما و پیمانکار می‌کند. حذف از زمان دکل به معنی صفر بودن مدت یا هزینه نیست.',
          '207 — برنامه AZNS-207 Program.pdf، صفحه فیزیکی ۱۶، رأس پیش‌بینی‌شده فهلیان را ۳۹۱۵ متر MD و AZNS207Pack.xlsx، شیت 155-2 DDR Data، سلول P86 مقدار ۳۹۶۸ متر MD را نشان می‌دهند. اختلاف حسابی ۵۳ متر تنها پس از تأیید مبنای عمق و ماهیت واقعی یا پیش‌بینی بودن مقدار DDR قابل تفسیر است. شوری ۱۴۳٬۰۰۰ ppm معادل NaCl در صفحه ۱۲ اطلاعات برنامه است و اندازه‌گیری فعلی را اثبات نمی‌کند.',
          '10 و 75 — صفحات ۶ و ۱۷ برنامه AZNS-010 Workover Program Rev.06.pdf از متروکه بودن قبلی و پلاگ سیمانی گزارش می‌دهند. صفحه ۸۶ برنامه AZNS-075 Workover-1 Program Rev.05.pdf رشته تکمیلی اولین ران Weqixin را ناموفق نامیده، اما متن استخراجی مکانیسم خرابی را اثبات نمی‌کند. در جدول سازندی چاه 10 بین شیت‌های 1-2 Data و 83-2 Data احتمال جابه‌جایی ردیف وجود دارد و باید تطبیق شود.',
          'شکل جایگزین W4 در گزارش هر چاه درصد تجمعی را صرفاً از ردیف‌های گرد‌شده جدول علل قبلی محاسبه می‌کند. در چاه 207 جمع ردیف‌های نمایش‌یافته ۹۸۸ ساعت در برابر ۱٬۰۰۳ ساعت زمان غیرمولد گزارش قبلی است؛ بنابراین مخرج درصد با کل زمان غیرمولد متفاوت است. این نمودارها و پیوست‌های سازندی، انتساب علل را اعتبارسنجی نمی‌کنند.'
        ]
        ps=[h]+[d.add_paragraph(text) for text in (fa_text if fa else en)]
        for p in ps:
            if fa:
                pp=p._p.get_or_add_pPr();b=OxmlElement('w:bidi');before=next((x for x in pp if x.tag in [qn('w:'+n) for n in ('jc','outlineLvl','rPr','sectPr','pPrChange')]),None)
                if before is not None:before.addprevious(b)
                else:pp.append(b)
            for r in p.runs:
                r.font.name='Times New Roman';r.font.size=Pt(14 if p is h else 10)
                rp=r._element.get_or_add_rPr();ft=rp.find(qn('w:rFonts'))
                if ft is None:ft=OxmlElement('w:rFonts');rp.insert(0,ft)
                ft.set(qn('w:cs'),'B Nazanin' if fa else 'Times New Roman')
                if fa:rp.append(OxmlElement('w:rtl'))
        d.save(path);print(path)
