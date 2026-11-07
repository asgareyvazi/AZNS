#!/usr/bin/env python3
"""Source-based 122 WO1 campaign review; selected facts, not operational instructions."""
from pathlib import Path
from datetime import datetime
import json
from openpyxl import load_workbook
from pypdf import PdfReader
from docx import Document
from replace_well_bit_evidence import style,set_table

SRC='AZNS122WO1Pack.xlsx';PROGRAM='AZNS-0122 Workover-1 Program Rev.07.pdf';CBS='CBS AZNS-122WO1.xlsx'
# Each prose entry is explicitly tied to all consulted cells; multi-cell entries
# account for remarks continuing onto the next row.
ENTRIES=[
('1-1 Remark',['K23','K24','K25','K26','K27'],
 'Report 1 records rig spud at 15:00 on 14 January 2024, annulus bleed-down and subsequent pressure rebuild; return-gas readings of 12 ppm H2S and 65% LEL are written in the same report. These are historical observations, not present conditions or an acceptable operating envelope.',
 'گزارش ۱ شروع عملیات دکل را در ساعت ۱۵ روز ۱۴ ژانویه ۲۰۲۴، تخلیه فشار فضای حلقوی و افزایش مجدد بعدی را ثبت می‌کند؛ در همان گزارش مقادیر ۱۲ قسمت در میلیون سولفید هیدروژن و ۶۵ درصد حد پایین انفجار برای گاز برگشتی نوشته شده‌اند. این‌ها مشاهدات تاریخی‌اند، نه شرایط فعلی یا محدوده مجاز عملیاتی.'),
('2-1 Remark',['K22','K23','K28','K29','K30'],
 'Report 2 describes rectification of leaks and a slickline lubricator/BOP pressure test labelled OK, followed by a failed depth counter and calibration during shifting-tool work. The source entry states the SSD was opened after several attempts; signed tests and complete tool records are needed to verify scope and outcome.',
 'گزارش ۲ رفع نشتی‌ها و آزمون فشار روان‌ساز و فوران‌گیر اسلیک‌لاین با نتیجه «مطلوب» و سپس خرابی و کالیبراسیون شمارنده عمق در کار با ابزار تغییر وضعیت را شرح می‌دهد. منبع باز شدن شیر کشویی پس از چند تلاش را ذکر می‌کند؛ دامنه و نتیجه به نمودار آزمون امضاشده و پرونده کامل ابزار نیاز دارند.'),
('6-1 Remark',['K22','K23','K24','K25'],
 'Report 6 records circulation toward a 106 pcf return and a source-stated 40 bbl loss during circulation; a later static check is labelled static. The loss figure must be reconciled to the fluid ledger and reporting interval before inclusion in a total.',
 'گزارش ۶ گردش تا رسیدن وزن برگشتی به ۱۰۶ پوند بر فوت مکعب و هرزروی اعلام‌شده ۴۰ بشکه در حین گردش را ثبت می‌کند؛ کنترل ایستایی بعدی در منبع ایستا نامیده شده است. پیش از هر جمع، رقم هرزروی باید با دفتر سیال و بازه ثبت تطبیق شود.'),
('8-1 Remark',['K22','K23','K24','K25','K26','K30','K31'],
 'Report 8 describes a tubing-running-tool leak, a separated crossover and later release of the packer, and records gas-cut returns during circulation. The entry does not by itself resolve the leak path or establish the condition of the whole completion string.',
 'گزارش ۸ نشتی ابزار راندن لوله، جداشدن رابط و سپس آزاد شدن پکر را شرح می‌دهد و بازگشت گل گاززده در گردش را ثبت می‌کند. این مدخل به‌تنهایی مسیر نشتی یا وضعیت کل رشته تکمیلی را مشخص نمی‌کند.'),
('11-1 Remark',['K22','K25','K26','K27'],
 'Report 11 observes a 2 by 4 cm washout on a retrieved 4½-inch tubing coupling and separately notes adherent asphaltene and difficulty breaking connections. The tally position, damage mechanism and causal relationship to the annulus pressure require physical inspection and records; they are not established by this excerpt.',
 'گزارش ۱۱ شستگی ۲ در ۴ سانتی‌متری روی یک کوپلینگ لوله ۴½ اینچ بازیابی‌شده را مشاهده کرده و جداگانه به آسفالتین چسبیده و دشواری بازکردن اتصالات اشاره می‌کند. محل مطابق فهرست، مکانیسم آسیب و رابطه علّی با فشار فضای حلقوی به بازرسی فیزیکی و اسناد نیاز دارند و از این گزیده احراز نمی‌شوند.'),
('15-1 Remark',['K24','K25','K26','K27','K29','K30','K31'],
 'Report 15 describes two corrosion-logging runs: the first logged from 3,924 to 3,752 m, and the second started at 3,780 m and logged upward. These are recorded logging intervals, not a corrosion diagnosis; the complete logs and interpretation are required.',
 'گزارش ۱۵ دو ران نمودارگیری خوردگی را شرح می‌دهد: ران اول از ۳۹۲۴ تا ۳۷۵۲ متر نمودار گرفته و ران دوم از ۳۷۸۰ متر به سمت بالا شروع شده است. این‌ها بازه‌های ثبت‌شده نمودارگیری‌اند، نه تشخیص خوردگی؛ لاگ کامل و تفسیر آن لازم است.'),
('19-1 Remark',['K24','K25','K26','K27','K28','K29','K30','K31','K32'],
 'Report 19 records installation work on a 7-inch seal-nipple packer assembly near the existing liner lap and a source-labelled satisfactory pull test after repeated setting attempts. This historical entry alone is not evidence of present-day isolation or an approved barrier.',
 'گزارش ۱۹ کار نصب مجموعه پکر نیپل آب‌بندی ۷ اینچ نزدیک هم‌پوشانی آستری موجود و آزمون کشش با نتیجه مطلوبِ ثبت‌شده پس از چند تلاش نصب را بیان می‌کند. این ثبت تاریخی به‌تنهایی اثبات جداسازی کنونی یا مانع مصوب نیست.'),
('25-1 Remark',['K22','K23','K25'],
 'Report 25 describes milling the old completion packer from 3,928.35 to 3,928.8 m and stopping because of low progress and repeated string stalls. The underlying parameters are historical, not a recommended milling recipe.',
 'گزارش ۲۵ آسیاب‌کاری پکر تکمیلی قدیمی از ۳۹۲۸٫۳۵ تا ۳۹۲۸٫۸ متر و توقف به‌علت پیشروی کم و توقف مکرر رشته را شرح می‌دهد. پارامترهای ثبت‌شده تاریخی‌اند، نه دستورالعمل توصیه‌شده آسیاب‌کاری.'),
('29-1 Remark',['K22','K23','K24','K28','K29'],
 'Report 29 records an unsuccessful taper-tap fishing attempt and states that a packer segment was found on the tap. The same entry reports running to 3,935.6 m but then says top of fish 3,565.6 m before stating the fish dropped to 3,945 m. The incompatible depth sequence is preserved as a source discrepancy; do not silently replace the suspect value.',
 'گزارش ۲۹ یک تلاش ناموفق ماهیگیری با تِیپر تَپ را ثبت کرده و می‌گوید قطعه‌ای از پکر روی ابزار دیده شد. همان مدخل راندن تا ۳۹۳۵٫۶ متر و سپس رأس قطعه در ۳۵۶۵٫۶ متر را پیش از ذکر افتادن قطعه تا ۳۹۴۵ متر بیان می‌کند. این توالی عمق ناسازگار به‌عنوان اختلاف منبع حفظ می‌شود؛ مقدار مشکوک نباید بی‌صدا اصلاح شود.'),
('34-1 Remark',['K22','K23','K24','K25'],
 'Report 34 states the fish dropped, reports washing and reaming to 4,148.5 m and then describes milling to 4,150 m with the source claim that the fish tail was completely milled. That claim concerns the tail; it does not independently verify recovery or removal of the complete fish.',
 'گزارش ۳۴ افتادن قطعه، شست‌وشو و گشادکاری تا ۴۱۴۸٫۵ متر و سپس آسیاب‌کاری تا ۴۱۵۰ متر را گزارش می‌کند و در منبع ادعا می‌کند انتهای قطعه کاملاً آسیاب شده است. این ادعا درباره انتهای قطعه است و مستقلاً بازیابی یا حذف کل قطعه را ثابت نمی‌کند.'),
('40-1 Remark',['K24','K25','K26','K27','K30'],
 'Report 40 describes a 7-inch liner tie-back assembly, including a source-stated 156 liner joints and a tie-back sleeve. The count is a source tally, not verified independent joints added to other reported totals; the approved string tally is required.',
 'گزارش ۴۰ مجموعه بازبست آستری ۷ اینچ، شامل ۱۵۶ شاخه آستری مطابق شمارش منبع و غلاف بازبست را شرح می‌دهد. این عدد شمارش منبع است، نه تعداد مستقلی که بتوان به سایر جمع‌ها اضافه کرد؛ فهرست مصوب رشته لازم است.'),
('41-1 Remark',['K22','K27','K28','K29','K36'],
 'Report 41 places the tie-back assembly at a 3,775.4 m liner-lap reference and records pumping 713 sacks of class G cement with a reported slurry volume of 172 bbl. A surface sample is later called hard; a surface sample and pumped quantity do not establish downhole cement placement or isolation.',
 'گزارش ۴۱ مجموعه بازبست را در مرجع هم‌پوشانی آستری ۳۷۷۵٫۴ متر قرار داده و پمپاژ ۷۱۳ کیسه سیمان رده G با حجم دوغاب گزارش‌شده ۱۷۲ بشکه را ثبت می‌کند. نمونه سطحی بعداً سخت توصیف شده است؛ نمونه سطحی و مقدار پمپ‌شده جای‌گیری سیمان درون چاه یا جداسازی را ثابت نمی‌کنند.'),
('42-1 Remark',['K22','K23','K24'],
 'Report 42 identifies a new 7-inch liner lap at 1,998 m after drilling out cement. This is a different reported reference from the deeper 3,775.4 m lap used during the tie-back run; do not treat them as interchangeable or assume that one displaced the other.',
 'گزارش ۴۲ پس از حفاری سیمان، هم‌پوشانی جدید آستری ۷ اینچ را در ۱۹۹۸ متر معرفی می‌کند. این مرجع با هم‌پوشانی عمیق‌تر ۳۷۷۵٫۴ متر در هنگام راندن بازبست متفاوت است؛ آن دو را یکسان یا جایگزین یکدیگر فرض نکنید.'),
('47-1 Remark',['K22','K23','K24','K25'],
 'Report 47 records a 104 pcf completion fluid and states 620 bbl was displaced. Programme pages 29–30 describe a 106 pcf fluid basis, with properties subject to the hole and a detailed programme to be issued. The 2 pcf numerical difference is not, by itself, proof of noncompliance, equal test conditions, final density or approval.',
 'گزارش ۴۷ سیال تکمیلی با وزن ۱۰۴ پوند بر فوت مکعب و جابه‌جایی ۶۲۰ بشکه را ثبت می‌کند. صفحات ۲۹ و ۳۰ برنامه مبنای ۱۰۶ پوند بر فوت مکعب را با امکان تغییر خواص بر حسب شرایط چاه و صدور برنامه تفصیلی ذکر می‌کنند. تفاوت عددی دو پوند بر فوت مکعب به‌تنهایی اثبات عدم انطباق، شرایط یکسان اندازه‌گیری، چگالی نهایی یا تأیید نیست.'),
('53-1 Remark',['K22','K26','K27','K29','K30','K31'],
 'Report 53 records repeated attempts before CIV 2 disk rupture, a 7,000 psi completion-string test labelled OK, and a plug left in the well then retrieved. The test scope and sequence must be checked against the signed test chart; it is not a current integrity certification.',
 'گزارش ۵۳ تلاش‌های مکرر پیش از شکستن دیسک شیر تزریق شماره ۲، آزمون رشته تکمیلی در ۷۰۰۰ پوند بر اینچ مربع با نتیجه «مطلوب» و جاماندن و سپس بازیابی یک پلاگ را ثبت می‌کند. دامنه و توالی آزمون باید با نمودار امضاشده تطبیق شود؛ این ثبت گواهی یکپارچگی فعلی نیست.'),
('54-1 Remark',['K25','K26','K27','K28','K29','K39','K40'],
 'Report 54 lists set-packer completion depths and several source-labelled satisfactory tests. These are source-reported historical depths rather than a current barrier schematic; the test chart, tally, datum and final completion drawing must be reviewed together.',
 'گزارش ۵۴ عمق اجزای تکمیل پس از نصب پکر و چند آزمون با نتیجه مطلوبِ ثبت‌شده را فهرست می‌کند. این‌ها عمق‌های تاریخی گزارش‌شده‌اند نه شماتیک مانع فعلی؛ نمودار آزمون، فهرست تجهیزات، مرجع عمق و نقشه نهایی تکمیل باید با هم بررسی شوند.'),
('56-1 Remark',['K22','K23','K24','K25','K26','K27','K30'],
 'Report 56 describes a coiled-tubing run to 3,650 m and displacement with 130 bbl NaCl brine, followed by retrieval and rig release at 18:00 on 9 March 2024. This establishes that the report describes a coiled-tubing activity; it does not establish execution of the separate programme acid wash or post-acid flow.',
 'گزارش ۵۶ راندن لوله‌مغزی سیار تا ۳۶۵۰ متر و جابه‌جایی با ۱۳۰ بشکه آب‌نمک کلرید سدیم، سپس بیرون‌کشیدن و ترخیص دکل در ساعت ۱۸ روز ۹ مارس ۲۰۲۴ را شرح می‌دهد. این ثبت شرح یک فعالیت لوله‌مغزی سیار است؛ اجرای اسیدشویی یا جریان‌دهی پس از اسید در برنامه جداگانه را ثابت نمی‌کند.')]

COMPLETION=[('End of completion string','انتهای رشته تکمیلی','3737.5','K39'),('No-go landing nipple','نیپل نشیمن بدون گذر','3736.88','K39'),('7-inch packer','پکر ۷ اینچ','3721.11','K39'),('Sliding side door','شیر کشویی جانبی','3708.35','K39'),('Traveling joint','لوله لغزشی','3689.60','K40'),('CIV 1','شیر تزریق شیمیایی ۱','3677.52','K40'),('CIV 2','شیر تزریق شیمیایی ۲','3665.64','K40'),('3½ × 4½-inch crossover','رابط ۳½ به ۴½ اینچ','1967.65','K40'),('TR3SV','شیر ایمنی زیرسطحی','62.89','K40')]

def insert_para(d,anchor,text,fa):
 p=d.add_paragraph(text);style(p,fa);anchor._p.addprevious(p._p)
def head(d,anchor,title):
 h=d.add_heading(title,level=2);h.alignment=1;anchor._p.addprevious(h._p)

def main():
 wb=load_workbook(SRC,read_only=True,data_only=True);records=[]
 for sheet,cells,en,fa in ENTRIES:
  values={c:wb[sheet][c].value for c in cells};assert all(isinstance(x,str) and x.strip() for x in values.values()),(sheet,values)
  records.append({'sheet':sheet,'cells':values,'english':en,'persian':fa})
 extras={'1-1 Remark': ['K34','K40'],'2-1 Remark':['K40']}
 extra={s:{c:wb[s][c].value for c in cs} for s,cs in extras.items()};wb.close()
 assert '14-Jan-2024' in extra['1-1 Remark']['K34'] and '09-Mar-2024, 18:00' in records[-1]['cells']['K30']
 assert '2023-Jan-19' in extra['1-1 Remark']['K40'] and '2023-Jan-09' in extra['2-1 Remark']['K40']
 start=datetime(2024,1,14,15);end=datetime(2024,3,9,18);hours=(end-start).total_seconds()/3600;assert hours==1323
 pdf=PdfReader(PROGRAM);pages={p:' '.join(pdf.pages[p-1].extract_text().split()) for p in (7,26,27,28,29,30)}
 for page,needle in [(7,'45 (Main Plan) + 7 (Contingency Plan)'),(26,'wireline corrosion logging'),(27,'post -acid clean -up'),(28,'5-7/8" Mill'),(29,'Down Hole losses'),(30,'Calcium Bromide')]:assert needle.lower() in pages[page].lower(),(page,needle)
 financial={}
 for cached in (False,True):
  book=load_workbook(CBS,read_only=True,data_only=cached);s=book['Cost Breakdown-122'];financial['cached' if cached else 'formula_or_value']={k:s[k].value for k in ('B5','C5','D5','E5','F5','G5','H5','G190','H190')};book.close()
 assert financial['cached']['G5']==1035000 and financial['cached']['H5']==547170
 for lang in ('EN','FA'):
  fa=lang=='FA';path=Path(f'122WO1_Drilling_Engineering_Report_{lang}.docx');d=Document(path);title='بازسازی منتخب عملیات تعمیر اول چاه ۱۲۲' if fa else 'Selected AZNS-122 workover 1 campaign reconstruction'
  if any(p.text==title for p in d.paragraphs):continue
  anchor=next(p for p in d.paragraphs if p.text in ('Performance and Parameter Analysis','تحلیل عملکرد و پارامترهای عملیاتی'));head(d,anchor,title)
  insert_para(d,anchor,('این توالی گزیده‌ای از گزارش‌های روزانه است، نه بازسازی کامل تمام ساعت‌ها یا تأیید مستقل رویدادها. ارجاعات محدوده سلول، ادامه متن در چند ردیف را نیز نشان می‌دهند. تاریخ‌ها، ساعات و نام خدمت‌دهندگان باید با گزارش امضاشده و قرارداد مصوب تطبیق شوند.' if fa else 'This sequence selects daily-report evidence; it is not a complete hour-by-hour reconstruction or independent verification. Cell ranges include continuation lines. Dates, hours and service names require signed-report and contract reconciliation.'),fa)
  for r in records:
   cells=list(r['cells']);ref=f"{SRC}, {r['sheet']}!{cells[0]}"+(f":{cells[-1]}" if len(cells)>1 else '')
   insert_para(d,anchor,f"{r['persian' if fa else 'english']} [{ref}]",fa)
  head(d,anchor,'مقایسه محدود مدت، دامنه و منابع' if fa else 'Limited duration, scope and source comparison')
  insert_para(d,anchor,('ثبت شروع در ۱۴ ژانویه ۲۰۲۴ ساعت ۱۵ و ترخیص در ۹ مارس ۲۰۲۴ ساعت ۱۸ فاصله تقویمی ۱۳۲۳ ساعت، معادل ۵۵ روز و ۳ ساعت، دارند. برنامه صفحه ۷ عبارت ۴۵ روز برنامه اصلی + ۷ روز برنامه احتیاطی را درج می‌کند؛ جمع حسابی ۵۲ روز است. این مقایسه به‌تنهایی اضافه‌زمان قراردادی یا اختلاف معتبر برنامه و اجرا نیست: مرز زمان، دامنه دکل و خدمات، تأیید تغییرات، هم‌پوشانی و زمان بدون دکل باید جداگانه تطبیق شوند. [AZNS122WO1Pack.xlsx، 1-1 Remark!K34 و 56-1 Remark!K30؛ AZNS-0122 Workover-1 Program Rev.07.pdf، صفحه ۷]' if fa else 'The reported spud at 15:00 on 14 January 2024 and rig release at 18:00 on 9 March 2024 are 1,323 calendar hours apart, or 55 days and 3 hours. Programme page 7 prints 45 main-plan days + 7 contingency-plan days, an arithmetic 52 days. This does not itself establish contractual overrun or a valid planned/actual variance: time boundaries, rig and service scope, approved changes, overlaps and rig-less time must be reconciled. [AZNS122WO1Pack.xlsx, 1-1 Remark!K34 and 56-1 Remark!K30; AZNS-0122 Workover-1 Program Rev.07.pdf, p.7]'),fa)
  insert_para(d,anchor,('تاریخ رسید واحد سیمان در 1-1 Remark!K40 برابر ۱۹ ژانویه ۲۰۲۳ و در 2-1 Remark!K40 برابر ۹ ژانویه ۲۰۲۳ نوشته شده است، در حالی که همان گزارش‌ها شروع دکل را ژانویه ۲۰۲۴ ثبت می‌کنند. اختلاف روز و سال باید با دفاتر ورود و گزارش‌های امضاشده کنترل شود؛ تصحیح فرضی انجام نشده است.' if fa else 'The cement-unit receipt note in 1-1 Remark!K40 says 19 January 2023, while 2-1 Remark!K40 says 9 January 2023; these reports also record rig spud in January 2024. The day and year require reconciliation to signed reports and arrival logs; no assumed correction is made.'),fa)
  insert_para(d,anchor,('برنامه صفحه ۲۷ پس از بخش دکل، فعالیت‌های بدون دکل از جمله اسیدشویی و جریان‌دهی پیش و پس از اسید را شرح می‌دهد؛ انجام اسیدکاری و پاکسازی پس از آن به‌صورت بدون دکل در متن به درخواست پیمانکار و تأیید کارفرما مشروط است. گزارش ۵۶ جابه‌جایی آب‌نمک را ثبت کرده، نه اجرای قطعی اسیدشویی یا جریان‌دهی برنامه‌ای. مدت این مراحل را از دوره دکل حذف یا هزینه آن را صفر فرض نکنید. [AZNS-0122 Workover-1 Program Rev.07.pdf، صفحه ۲۷؛ AZNS122WO1Pack.xlsx، 56-1 Remark!K22:K27]' if fa else 'Programme p.27 describes a separate rig-less sequence, including acid wash and pre- and post-acid flow; performing acidizing and post-acid cleanup rig-less is subject to contractor request and client approval in the source. Report 56 records brine displacement, not confirmed execution of the planned acid wash or flow. Do not infer zero duration or cost merely because a step is outside the rig period. [AZNS-0122 Workover-1 Program Rev.07.pdf, p.27; AZNS122WO1Pack.xlsx, 56-1 Remark!K22:K27]'),fa)
  head(d,anchor,'تکمیل، رشته حفاری و کنترل مالی' if fa else 'Completion, BHA and financial checks')
  insert_para(d,anchor,('عمق‌های زیر عین ثبت گزارش ۵۴ پس از نصب پکر هستند؛ این جدول نقشه نهایی مانع یا مدرک مستقل سلامت تکمیل نیست. مرجع عمق، فهرست تجهیزات، نقشه تأییدشده و اسناد آزمون باید تطبیق شوند. [AZNS122WO1Pack.xlsx، 54-1 Remark!K39:K40]' if fa else 'The following depths reproduce report 54 after setting the packer; this is not a final barrier diagram or independent evidence of completion integrity. Reconcile the depth datum, tally, approved drawing and test documentation. [AZNS122WO1Pack.xlsx, 54-1 Remark!K39:K40]'),fa)
  table=d.add_table(rows=1,cols=3);table.style='Table Grid';anchor._p.addprevious(table._tbl)
  rows=[(['جزء گزارش‌شده','عمق منبع، متر','سلول'] if fa else ['Reported component','Source depth, m','Cell'])]+[[pers if fa else en,depth,cell] for en,pers,depth,cell in COMPLETION];set_table(table,rows,[8.0,4.,4.4],fa)
  p=d.add_paragraph('جدول ـ عمق‌های منتخب تکمیل پس از نصب پکر در گزارش ۵۴' if fa else 'Table — Selected completion depths after packer setting in report 54');style(p,fa,True);anchor._p.addprevious(p._p)
  insert_para(d,anchor,('برنامه صفحه ۲۸ یک مجموعه آسیاب ۵⅞ اینچ با طوق، جار، لوله سنگین و لوله حفاری را نشان می‌دهد. مجموعه‌های عملیاتی مشاهده‌شده و اختلاف مستقل اندازه خراشنده در جدول رشته باید با شماره ران، فاز، طول، قطر داخلی و خارجی، اتصالات و فهرست واقعی مقایسه شوند؛ شباهت اسم ابزار انطباق طراحی و اجرا را ثابت نمی‌کند. بدون هندسه تأییدشده، خواص هم‌زمان سیال و افت سطحی، افت فشار هیدرولیکی و حاشیه ایمن محاسبه نشده‌اند. [AZNS-0122 Workover-1 Program Rev.07.pdf، صفحه ۲۸؛ AZNS122WO1Pack.xlsx، 12-2 Data و 26-1 Remark!K25:K27]' if fa else 'Programme p.28 depicts a 5⅞-inch milling assembly with collars, jar, heavy-weight and drill pipe. Observed assemblies and the separate scraper-size discrepancy in the BHA table need run, phase, length, inside/outside diameter, connection and actual-tally reconciliation; a similar tool name does not prove the planned assembly was run. Hydraulic pressure drop and safe margins are not calculated without verified geometry, synchronized fluid properties and surface losses. [AZNS-0122 Workover-1 Program Rev.07.pdf, p.28; AZNS122WO1Pack.xlsx, 12-2 Data and 26-1 Remark!K25:K27]'),fa)
  insert_para(d,anchor,('در ساختار هزینه، ردیف خدمات دکل نرخ ۲۳۰۰۰ دلار در روز، ۴۵ روز برنامه و ۲۳٫۷۹ روز در ستون کار اضافی را نشان می‌دهد؛ مبالغ ذخیره‌شده آن به‌ترتیب ۱٬۰۳۵٬۰۰۰ و ۵۴۷٬۱۷۰ دلارند. این ستون‌ها برآورد منبع‌اند، نه صورت‌حساب یا الزاماً روزهای واقعی. ردیف‌های هزینه نباید دوباره به سرجمع‌های G190 و H190 اضافه شوند. [CBS AZNS-122WO1.xlsx، Cost Breakdown-122!B5:H5 و G190:H190]' if fa else 'The CBS rig-service row lists USD 23,000/day, 45 planned days and 23.79 days in the extra-work column; cached amounts are USD 1,035,000 and USD 547,170 respectively. These columns are source estimates, not invoices or necessarily actual days. Do not add the line a second time to totals G190 and H190. [CBS AZNS-122WO1.xlsx, Cost Breakdown-122!B5:H5 and G190:H190]'),fa)
  d.save(path)
 out={'well':'122WO1','source':SRC,'programme':PROGRAM,'events':records,'additional_date_cells':extra,'spud':'2024-01-14T15:00','rig_release':'2024-03-09T18:00','calendar_elapsed_hours':hours,'programme_physical_pages':list(pages),'completion_depths':[{'english':en,'persian':fa,'source_depth_text':depth,'sheet':'54-1 Remark','cell':cell} for en,fa,depth,cell in COMPLETION],'cbs':financial,'limitations':'Selected historical source observations, not complete chronology, approved plan variance, present barrier status, invoices or execution of rig-less acid wash.'}
 Path('evidence/122wo1_campaign_review.json').write_text(json.dumps(out,ensure_ascii=False,indent=2)+'\n');print('events',len(records),'event cells',sum(len(x['cells']) for x in records),'completion components',len(COMPLETION),'elapsed hours',hours)
if __name__=='__main__':main()
