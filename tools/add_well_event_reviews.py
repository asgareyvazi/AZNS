#!/usr/bin/env python3
"""Add manually reviewed, source-cell daily observations, not incident verdicts."""
import json
from pathlib import Path
from docx import Document
from replace_well_bit_evidence import DATA,style

# (worksheet prefix, cell, exact source anchor, English paraphrase, Persian paraphrase)
SEL={
 '207':[
 (32,'S32','Fish length: 19.84m','The note lists a 17½-inch PDC bit, bit sub and two 9½-inch drill collars left in hole, with a reported fish length of 19.84 m.','یادداشت، مته پلی‌کریستال ۱۷½ اینچ، رابط مته و دو شاخه طوق حفاری ۹½ اینچ را به‌عنوان قطعات جامانده با طول گزارش‌شده ۱۹٫۸۴ متر ذکر می‌کند.'),
 (106,'S24','1-3 BPH Losses','The report describes drilling the 12¼-inch hole from 3,741 to 3,743 m with losses stated as 1–3 barrels per hour.','گزارش حفاری حفره ۱۲¼ اینچ از ۳۷۴۱ تا ۳۷۴۳ متر را با هرزروی اعلام‌شده ۱ تا ۳ بشکه در ساعت شرح می‌دهد.'),
 (139,'S25','2500 psi/10min; OK','The surface dry-test line pressure test is recorded as 2,500 psi for ten minutes, with the source result “OK”; the same entry records setting a test packer at 3,730 m.','آزمون فشار خط سطحی آزمون خشک با فشار ۲۵۰۰ پوند بر اینچ مربع به مدت ده دقیقه و نتیجه منبع «مطلوب» ثبت شده است؛ همان مدخل نصب پکر آزمون در ۳۷۳۰ متر را ذکر می‌کند.')],
 '127ST2':[
 (8,'S37','Window Milling','The entry describes window milling with an 8½-inch cased-hole sidetrack assembly; this cell alone does not establish the final window dimensions.','مدخل آسیاب‌کاری پنجره با مجموعه ۸½ اینچی انحراف در چاه جداری‌گذاری‌شده را شرح می‌دهد؛ این سلول به‌تنهایی ابعاد نهایی پنجره را ثابت نمی‌کند.'),
 (31,'S41','TDS Gooseneck Leakage','The report attributes an interruption to a top-drive gooseneck leak and describes pulling the directional assembly from 3,788.9 to 3,743.5 m; contractual responsibility remains unverified.','گزارش توقفی را به نشتی گلویی تاپ‌درایو نسبت داده و بیرون کشیدن مجموعه جهت‌دار از ۳۷۸۸٫۹ تا ۳۷۴۳٫۵ متر را شرح می‌دهد؛ مسئولیت قراردادی تأیید نشده است.'),
 (44,'S27','2500 Psi/10min','A surface line test is reported at 2,500 psi for ten minutes with the source result “OK”; it is not proof of current barrier integrity.','آزمون خط سطحی با فشار ۲۵۰۰ پوند بر اینچ مربع به مدت ده دقیقه و نتیجه «مطلوب» گزارش شده است؛ این ثبت یکپارچگی فعلی موانع را ثابت نمی‌کند.')],
 '40':[
 (33,'K22','8-10 BPH Losses','The report describes low-progress drilling from 2,944 to 2,945 m with losses of 8–10 barrels per hour and two pill treatments.','گزارش پیشروی کم از ۲۹۴۴ تا ۲۹۴۵ متر، هرزروی ۸ تا ۱۰ بشکه در ساعت و دو نوبت پمپاژ پیل را شرح می‌دهد.'),
 (53,'K22','3-5 BPH losses','During the reported gamma-ray log-up from 300 m to surface, the same cell notes circulation and losses of 3–5 barrels per hour.','در نمودارگیری گاما از ۳۰۰ متر تا سطح، همان سلول گردش سیال و هرزروی ۳ تا ۵ بشکه در ساعت را ذکر می‌کند.'),
 (105,'K22','Catch fish confirmed','The entry reports an overshot trip to 2,455 m, tagging the top of fish and a confirmed catch; a complete recovery outcome requires subsequent records.','مدخل راندن اورشات تا ۲۴۵۵ متر، تماس با رأس قطعه جامانده و تأیید گرفتن آن را گزارش می‌کند؛ نتیجه بازیابی کامل به سوابق بعدی نیاز دارد.')],
 '127WO1':[
 (26,'K22','No Succ','The report states that attempts to free the fish were unsuccessful and circulation was established; the cell does not establish the cause of sticking.','گزارش می‌گوید تلاش‌ها برای آزادکردن قطعه جامانده ناموفق بوده و گردش برقرار شده است؛ این سلول علت گیر را ثابت نمی‌کند.'),
 (42,'K23','Top of Fish is  3270.77m','The cell gives the top of fish as 3,270.77 m alongside recovered tubing details; the depth reference and full assembly need reconciliation.','سلول در کنار جزئیات لوله بازیابی‌شده، رأس قطعه جامانده را ۳۲۷۰٫۷۷ متر ذکر می‌کند؛ مرجع عمق و مجموعه کامل نیازمند تطبیق‌اند.'),
 (60,'K23','Pressure Test CMT Line to 3500 psi; Ok','The cement line pressure test is recorded at 3,500 psi with result “OK”; the entry also refers to a whipstock packer at 3,339.08 m.','آزمون فشار خط سیمان با فشار ۳۵۰۰ پوند بر اینچ مربع و نتیجه «مطلوب» ثبت شده است؛ مدخل همچنین به پکر ویپ‌استاک در عمق ۳۳۳۹٫۰۸ متر اشاره می‌کند.')],
 '10':[
 (7,'K33','Leakage Stand pipe ground line','A standpipe ground-line leak is reported, followed by stopping the operation and changing the mud-pump ground manifold; confirm the full event sequence separately.','نشتی خط زمینی استندپایپ و سپس توقف عملیات و تعویض منیفولد زمینی پمپ گل گزارش شده است؛ توالی کامل رویداد باید جداگانه تطبیق شود.'),
 (11,'K24','Pressure test Line to 3000 psi, Ok','The cement-pump truck line pressure test is recorded at 3,000 psi with result “OK”; this is a historical source observation.','آزمون فشار خط کامیون پمپ سیمان با فشار ۳۰۰۰ پوند بر اینچ مربع و نتیجه «مطلوب» ثبت شده است؛ این یک مشاهده تاریخی منبع است.'),
 (64,'K25','f/3559 to 4077m','The entry reports running a 5⅞-inch milling assembly from 3,559 to 4,077 m; it does not by itself establish milling success.','مدخل راندن مجموعه آسیاب ۵⅞ اینچ از ۳۵۵۹ تا ۴۰۷۷ متر را گزارش می‌کند؛ به‌تنهایی موفقیت آسیاب‌کاری را ثابت نمی‌کند.')],
 '12':[
 (19,'S38','asphaltene and sand','While pulling the fish, the report notes two 4½-inch tubing joints plugged with asphaltene and sand, and observed box damage while breaking connections.','گزارش هنگام بیرون کشیدن قطعه جامانده به گرفتگی دو شاخه لوله ۴½ اینچ با آسفالتین و ماسه و مشاهده آسیب مادگی هنگام بازکردن اتصال اشاره می‌کند.'),
 (28,'S24','5" L.Lap','The report records running a 6-inch junk mill to 3,778 m, identified as the 5-inch liner lap; the loss-rate text in this cell is ambiguous and is not normalized.','گزارش راندن آسیاب ضایعات ۶ اینچ تا ۳۷۷۸ متر، معرفی‌شده به‌عنوان هم‌پوشانی آستری ۵ اینچ، را ثبت می‌کند؛ متن نرخ هرزروی این سلول مبهم است و نرمال‌سازی نشده است.'),
 (50,'S34','5000 psi for 15min','An offline pressure test of the stated completion components is recorded at 5,000 psi for 15 minutes with an “OK” result for the first listed test.','آزمون فشار خارج از خط برای اجزای تکمیلی ذکرشده با فشار ۵۰۰۰ پوند بر اینچ مربع به مدت ۱۵ دقیقه ثبت شده و نتیجه نخستین آزمون «مطلوب» است.')],
 '75':[
 (17,'S28','3000 psi for 15 minutes','The report records a surface-line pressure test at 3,000 psi for 15 minutes, with the test confirmed “OK”.','گزارش آزمون فشار خطوط سطحی با فشار ۳۰۰۰ پوند بر اینچ مربع به مدت ۱۵ دقیقه و نتیجه تأییدشده «مطلوب» را ثبت می‌کند.'),
 (29,'S26','to 2266m','The report states that window milling and dressing with an 8½-inch mill started and reached 2,266 m.','گزارش آغاز آسیاب و پرداخت پنجره با آسیاب ۸½ اینچ تا عمق ۲۲۶۶ متر را ذکر می‌کند.'),
 (38,'S44','8-10 bph losses','The report describes sliding directional drilling from 2,609 to 2,610 m with 8–10 barrels per hour of reported losses.','گزارش حفاری جهت‌دار به حالت لغزشی از ۲۶۰۹ تا ۲۶۱۰ متر را با هرزروی گزارش‌شده ۸ تا ۱۰ بشکه در ساعت شرح می‌دهد.')],
 '85':[
 (23,'K24','no succ','The report describes two slickline magnet runs to catch the fish as unsuccessful; the source classification or vendor name does not itself establish contractual liability.','گزارش دو نوبت راندن آهنربا با اسلیک‌لاین برای گرفتن قطعه جامانده را ناموفق توصیف می‌کند؛ کد منبع یا نام خدمت‌دهنده به‌تنهایی مسئولیت قراردادی را ثابت نمی‌کند.'),
 (27,'K27','f/ 3975 to 4066m','The report describes reaming down with a 5⅞-inch milling assembly from 3,975 to 4,066 m; the parameters in the source are historical observations, not an operating window.','گزارش گشادکاری رو به پایین با مجموعه آسیاب ۵⅞ اینچ از ۳۹۷۵ تا ۴۰۶۶ متر را شرح می‌دهد؛ پارامترهای منبع مشاهده تاریخی‌اند، نه پنجره عملیاتی.'),
 (30,'K30','7" Liner Cement','The entry records running a wireline tool string for 7-inch liner cement evaluation; the full log and interpretation are needed to determine cement quality.','مدخل راندن مجموعه ابزار نمودارگیری با سیم برای ارزیابی سیمان آستری ۷ اینچ را ثبت می‌کند؛ تعیین کیفیت سیمان به لاگ کامل و تفسیر آن نیاز دارد.')],
 '116':[
 (33,'K24','3846.1m to 3876.1m','The cell describes pushing the fish from 3,846.1 to 3,876.1 m and resuming milling; it does not independently prove the fish was recovered.','سلول راندن قطعه جامانده از ۳۸۴۶٫۱ تا ۳۸۷۶٫۱ متر و ازسرگیری آسیاب‌کاری را شرح می‌دهد؛ به‌تنهایی بازیابی قطعه را ثابت نمی‌کند.'),
 (51,'K24','Catch The Fish','The report states that the fish was caught, worked on and released while pulling out; reconcile the surrounding entries before establishing the final outcome.','گزارش گرفتن قطعه جامانده، کار روی آن و رهاکردن آن هنگام بیرون کشیدن را ذکر می‌کند؛ نتیجه نهایی به تطبیق مدخل‌های پیرامونی نیاز دارد.'),
 (91,'K23','1200psi for 15min, Ok','The report records a pressure test of the 7-inch shooting nipple and wireline blowout preventer at 1,200 psi for 15 minutes, result “OK”.','گزارش آزمون فشار نیپل مشبک‌کاری ۷ اینچ و فوران‌گیر سیمی با فشار ۱۲۰۰ پوند بر اینچ مربع به مدت ۱۵ دقیقه و نتیجه «مطلوب» را ثبت می‌کند.')],
 '122WO1':[
 (20,'K25','corrosion of  9-5/8','The entry notes waiting for a client decision concerning corrosion of the 9⅝-inch casing while pulling setting tools; the extent and mechanism of corrosion are not established by this cell.','مدخل انتظار برای تصمیم کارفرما درباره خوردگی جداری ۹⅝ اینچ هم‌زمان با بیرون کشیدن ابزار نصب را ذکر می‌کند؛ میزان و مکانیسم خوردگی از این سلول احراز نمی‌شود.'),
 (34,'K24','Fish dropped','The report says the fish dropped and describes washing and reaming from 4,083 to 4,148.5 m with 1–2 barrels per hour of losses.','گزارش افتادن قطعه جامانده و شست‌وشو و گشادکاری از ۴۰۸۳ تا ۴۱۴۸٫۵ متر با هرزروی ۱ تا ۲ بشکه در ساعت را شرح می‌دهد.'),
 (36,'K24','Obs. LG mark','After pulling a 7-5/16-inch polish mill to surface, the report records an observed “LG” mark and 0–1 barrel per hour losses; the mark is not diagnosed here.','پس از بیرون کشیدن آسیاب پرداخت ۷-۵/۱۶ اینچ به سطح، گزارش نشانه «LG» و هرزروی صفر تا یک بشکه در ساعت را ثبت می‌کند؛ این نشانه در اینجا تفسیر خرابی نشده است.')],
 '122WO2':[
 (18,'S24','F/4077m To 4080.30m','The report describes starting milling and clean-out from 4,077 to 4,080.30 m; the recorded parameters are not recommended settings.','گزارش آغاز آسیاب‌کاری و تمیزکاری از ۴۰۷۷ تا ۴۰۸۰٫۳۰ متر را شرح می‌دهد؛ پارامترهای ثبت‌شده مقادیر توصیه‌شده نیستند.'),
 (37,'S27','Positive Pressure Test CMT Plug','A positive pressure test of a cement plug is recorded at 2,000 psi with the source result “OK”.','آزمون فشار مثبت پلاگ سیمانی با فشار ۲۰۰۰ پوند بر اینچ مربع و نتیجه منبع «مطلوب» ثبت شده است.'),
 (74,'S24','Totally : 29 jts 4-1/2','The report records running 10 additional 4½-inch liner joints, for a stated total of 29 joints, and making up the liner-hanger assembly.','گزارش راندن ۱۰ شاخه دیگر آستری ۴½ اینچ با مجموع اعلام‌شده ۲۹ شاخه و بستن مجموعه آویز آستری را ثبت می‌کند.')],
 '127WO2':[
 (15,'S34','Tool String is stuck','The entry describes a slickline tool string becoming stuck during retrieval and subsequently being freed; it does not establish the sticking mechanism.','مدخل گیرکردن رشته ابزار اسلیک‌لاین هنگام بازیابی و سپس آزادشدن آن را شرح می‌دهد؛ مکانیسم گیر را ثابت نمی‌کند.'),
 (44,'S42','Severe corrosion and washout','The report notes severe corrosion and washout observed on the retrieved fish; metallurgical mechanism and responsibility require separate evidence.','گزارش خوردگی شدید و شستگی مشاهده‌شده روی قطعه جامانده بازیابی‌شده را ذکر می‌کند؛ مکانیسم متالورژیکی و مسئولیت به شواهد جداگانه نیاز دارند.'),
 (58,'S26','24.Feb.2026','The report describes pulling a wireline logging string for the 7-inch liner from 3,147 m to surface; the full logs are needed to interpret liner condition.','گزارش بیرون کشیدن رشته نمودارگیری با سیم برای آستری ۷ اینچ از ۳۱۴۷ متر تا سطح را شرح می‌دهد؛ تفسیر وضعیت آستری به لاگ‌های کامل نیاز دارد.')],
 '166':[
 (4,'S30','No Succ','The report describes repeated unsuccessful blowout-preventer test attempts and work to fix leaks; no current barrier condition is inferred.','گزارش تلاش‌های مکرر ناموفق برای آزمون فوران‌گیر و کار برای رفع نشتی را شرح می‌دهد؛ وضعیت فعلی موانع از آن استنباط نمی‌شود.'),
 (5,'S27','4500Psi, OK','The report records rectification of top- and bottom-ram leaks and a 4,500 psi test with source result “OK”.','گزارش رفع نشتی رام‌های بالا و پایین و آزمون با فشار ۴۵۰۰ پوند بر اینچ مربع با نتیجه منبع «مطلوب» را ثبت می‌کند.'),
 (23,'S39','1,000 Psi / 10 Min','The tubing hanger and tie-down screws are reported pressure-tested at 1,000 psi for ten minutes, with result “OK”.','آویز لوله و پیچ‌های مهاری طبق گزارش با فشار ۱۰۰۰ پوند بر اینچ مربع به مدت ده دقیقه آزمون شده و نتیجه «مطلوب» بوده است.')]
}

FOLLOW={
 '207':('Reconcile the fish dimensions, BHA tally, depth and the complete fishing/sidetrack records before calculating lost-hole performance or assigning a cause. Compare loss-rate narratives with the separately recorded loss-volume field; the bases may differ.','پیش از محاسبه عملکرد بخش از‌دست‌رفته یا انتساب علت، ابعاد قطعه، فهرست رشته، عمق و سوابق کامل ماهیگیری و انحراف تطبیق شوند. روایت نرخ هرزروی با ستون جداگانه حجم هرزروی مقایسه شود؛ مبنای ثبت ممکن است متفاوت باشد.'),
 '127ST2':('The window record, leak entry and test relate to different source worksheets and equipment. Verify event times, phase boundaries, test chart and repair evidence before drawing a causal or contractor conclusion.','ثبت پنجره، مدخل نشتی و آزمون به شیت‌های منبع و تجهیزات متفاوت مربوط‌اند. پیش از نتیجه‌گیری علّی یا قراردادی، زمان رخداد، مرز فاز، نمودار آزمون و شواهد تعمیر بررسی شوند.'),
 '40':('Loss rate is not a formation assignment or cumulative volume. Reconcile report-end depths with the actual event depth, fluid ledger and fishing record before building a phase balance.','نرخ هرزروی انتساب سازندی یا حجم تجمعی نیست. پیش از تهیه تراز هر فاز، عمق پایان گزارش با عمق واقعی رویداد، دفتر سیال و سابقه ماهیگیری تطبیق شود.'),
 '127WO1':('These fragments cover fish work, a depth entry and a line test; they do not prove a continuous sequence, final recovery, well status or sidetrack execution without the full records.','این گزیده‌ها کار روی قطعه، یک مقدار عمق و آزمون خط را پوشش می‌دهند؛ بدون سوابق کامل، توالی پیوسته، بازیابی نهایی، وضعیت چاه یا اجرای انحراف را ثابت نمی‌کنند.'),
 '10':('The leak, line test and milling trip concern different systems. A pressure-test result should be checked against the signed chart, test envelope and later barrier records.','نشتی، آزمون خط و ران آسیاب به سامانه‌های متفاوت مربوط‌اند. نتیجه آزمون فشار باید با نمودار امضاشده، محدوده آزمون و سوابق بعدی موانع تطبیق شود.'),
 '12':('The source has an ambiguous “8-2” loss expression; do not reorder or sum it as a measured rate. Reconcile tubing condition and test scope with photographs, tallies and signed test records.','منبع عبارت مبهم هرزروی «8-2» دارد؛ آن را به‌صورت نرخ اندازه‌گیری‌شده بازمرتب یا جمع نکنید. وضعیت لوله و دامنه آزمون با تصاویر، فهرست لوله و سوابق امضاشده تطبیق شوند.'),
 '75':('Depths from the original and sidetrack intervals require phase and datum reconciliation. A logged loss rate does not identify a formation or a validated total-loss volume.','عمق‌های شاخه اصلی و انحرافی به تطبیق فاز و مبنای عمق نیاز دارند. نرخ هرزروی ثبت‌شده، سازند یا حجم کل تأییدشده را تعیین نمی‌کند.'),
 '85':('The magnet, reaming and logging observations must be reconciled with the actual workover sequence; they do not demonstrate execution of the separately planned rig-less acid treatment.','مشاهدات آهنربا، گشادکاری و نمودارگیری باید با توالی واقعی تعمیر تطبیق شوند؛ اجرای اسیدکاری بدون دکلِ برنامه‌ریزی‌شده جداگانه را ثابت نمی‌کنند.'),
 '116':('Fish movements do not establish final recovery. Review the fishing tally, full sequence, test record and approved barrier diagram before drawing an operational conclusion.','حرکت قطعه جامانده به‌معنای بازیابی نهایی نیست. پیش از نتیجه‌گیری عملیاتی، فهرست ماهیگیری، توالی کامل، سابقه آزمون و نمودار مصوب موانع بازبینی شوند.'),
 '122WO1':('The corrosion observation needs inspection records and well-specific history; the separate scraper-size conflict in the BHA table also requires vendor/tally confirmation.','مشاهده خوردگی به اسناد بازرسی و تاریخچه اختصاصی چاه نیاز دارد؛ اختلاف مستقل اندازه خراشنده در جدول رشته نیز به تأیید سازنده و فهرست تجهیزات نیاز دارد.'),
 '122WO2':('A single plug test and liner tally do not prove long-term isolation or final configuration; reconcile cement, pressure, liner and completion records.','یک آزمون پلاگ و شمارش آستری، جداسازی بلندمدت یا وضعیت نهایی را ثابت نمی‌کند؛ سوابق سیمان، فشار، آستری و تکمیل تطبیق شوند.'),
 '127WO2':('The corrosion and washout note describes a retrieved object, not automatically the casing or the cause of the incident. Review photographs, identification and inspection before attributing a mechanism.','یادداشت خوردگی و شستگی به یک قطعه بازیابی‌شده مربوط است، نه خودکار به جداری یا علت حادثه. پیش از انتساب مکانیسم، تصاویر، شناسایی و بازرسی بررسی شوند.'),
 '166':('A later successful test does not erase the earlier unsuccessful attempts or prove current integrity. Reconcile test diagrams, repairs and accepted barriers across the complete record.','آزمون موفق بعدی تلاش‌های ناموفق قبلی را حذف نمی‌کند و یکپارچگی فعلی را ثابت نمی‌کند. نمودارهای آزمون، تعمیرات و موانع پذیرفته‌شده در کل سابقه تطبیق شوند.')}

def main():
 raw={x['file']:x for x in json.loads(Path('evidence/activity_observations.json').read_text())};review=[]
 for well,w in DATA.items():
  source=raw[w['source']];events=[]
  for day,cell,anchor,en,fa in sorted(SEL[well],key=lambda x:(x[0],x[1])):
   entry=next(x for x in source['days'] if x['day']==day);z=next(x for x in entry['detail_cells'] if x['cell']==cell)
   assert anchor.lower() in z['raw'].lower(),(well,day,cell,anchor,z['raw'])
   seg=next((x for x in entry['activity']['segments'] if x.get('detail') and x['detail']['cell']==cell),None)
   events.append({'day':day,'sheet':entry['remark_sheet'],'cell':cell,'raw':z['raw'],'anchor':anchor,'english':en,'persian':fa,'associated_status_segment':seg})
  review.append({'well':well,'file':w['source'],'events':events,'engineering_follow_up':{'english':FOLLOW[well][0],'persian':FOLLOW[well][1]}})
  for lang in ('EN','FA'):
   fa=lang=='FA';path=Path(f'{well}_Drilling_Engineering_Report_{lang}.docx');d=Document(path)
   title='شواهد تکمیلی منتخب از گزارش‌های روزانه' if fa else 'Additional selected daily-report evidence'
   if any(p.text==title for p in d.paragraphs):continue
   # W2 caption follows an image paragraph; insert before the image.
   cap=next(p for p in d.paragraphs if p.text.startswith(('Figure W2','شکل W2')));image=cap._p.getprevious();assert image.xpath('.//a:blip')
   h=d.add_heading(title,level=2);h.alignment=1;image.addprevious(h._p)
   intro=('این سه مدخل به‌ترتیب شماره شیت، نه شماره تأییدشده گزارش، نمایش داده شده‌اند؛ شماره شیت به‌تنهایی زمان وقوع، توالی کامل، تکمیل عملیات یا وضعیت کنونی چاه را اثبات نمی‌کند. نتیجه‌های آزمون همان برچسب گزارش‌اند و باید با نمودار آزمون و محدوده تأییدشده تطبیق شوند.' if fa else 'These three entries are ordered by worksheet prefix, not verified report number; sheet order alone does not establish event time, complete chronology, operational completion or current well condition. Test outcomes are source labels and require reconciliation with signed charts and the tested scope.')
   p=d.add_paragraph(intro);style(p,fa);image.addprevious(p._p)
   for e in events:
    para=(f"شیت منبع {e['day']} ـ {e['persian']} [{w['source']}، {e['sheet']}!{e['cell']}]" if fa else f"Source worksheet {e['day']} — {e['english']} [{w['source']}, {e['sheet']}!{e['cell']}]")
    p=d.add_paragraph(para);style(p,fa);image.addprevious(p._p)
   p=d.add_paragraph(FOLLOW[well][1 if fa else 0]);style(p,fa);image.addprevious(p._p);d.save(path)
 Path('evidence/well_event_reviews_selected.json').write_text(json.dumps(review,indent=2,ensure_ascii=False)+'\n')
 print('Wells:',len(review),'events:',sum(len(x['events']) for x in review))
if __name__=='__main__':main()
