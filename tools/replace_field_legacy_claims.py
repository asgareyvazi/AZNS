#!/usr/bin/env python3
"""Replace unsupported field claims/figures with source-scoped, bilingual evidence."""
from pathlib import Path
import json,re
from docx import Document
from docx.oxml.ns import qn
from replace_well_bit_evidence import set_table,text

D=Path('evidence');load=lambda s:json.loads((D/(s+'.json')).read_text())
meta=load('field_figure_observations')['figures'];ph=load('program_phases_services_reviewed');program=load('program_summary_reviewed');fluid=load('fluid_observation_summary');acts=load('activity_observations')
P={x['well']:x for x in program};F={x['well']:x for x in fluid};A={x['file']:x for x in acts};C=meta['M4'];
CAP={2:('Original pack worksheet inventory, not completeness','موجودی شیت‌های بسته‌های اصلی، نه کامل‌بودن داده'),1:('Programme-stated main/estimated durations, not actual time','مدت اصلی یا برآوردی درج‌شده در برنامه، نه زمان واقعی'),4:('Recognized source-page coverage, not rig performance','پوشش شناسایی صفحات منبع، نه عملکرد دکل'),5:('Recognized BHA, viscosity and loss fields','پوشش شناسایی رشته ته‌چاهی، گرانروی و هرزروی'),7:('Report-end depth samples in four operations','نمونه‌های عمق پایان گزارش در چهار عملیات'),8:('Selected programme phases for wells 122 and 127','مراحل منتخب برنامه‌های چاه‌های ۱۲۲ و ۱۲۷'),9:('Examples of unresolved source reconciliation','نمونه اختلاف‌ها و ابهام‌های نیازمند تطبیق'),10:('Loss entries against report-end depth, not event or formation attribution','مقادیر هرزروی در برابر عمق پایان گزارش؛ نه عمق رویداد یا انتساب سازندی')}
INTRO={
 2:('The inventory counts worksheets in each original daily pack, including all sheet types; it does not establish days, completeness or semantic coverage. The 127WO2 label refers to AZNS127WOPack.xlsx.', 'موجودی تعداد همه شیت‌های هر بسته روزانه اصلی را نشان می‌دهد؛ تعداد روز، کامل‌بودن داده یا پوشش معنایی را اثبات نمی‌کند. برچسب 127WO2 به فایل AZNS127WOPack.xlsx اشاره دارد.'),
 1:('The chart compares the main-plan or estimated well duration stated in each of 14 programmes. Moving and contingency figures are not added; these are neither verified actual durations nor a fleet timeline. Scope and phase boundaries differ.', 'نمودار مدت اصلی یا برآوردی چاه را از ۱۴ برنامه نشان می‌دهد. ارقام جابه‌جایی و احتیاطی جمع نشده‌اند؛ این مقادیر زمان واقعی تأییدشده یا زمان‌بندی ناوگان نیستند و دامنه و مرز فازها یکسان نیست.'),
 4:('The chart distinguishes independently recognized status pages, complete classified status pages and recognized data pages. It is a coverage view, not a comparison of rig availability, efficiency, responsibility or cost.', 'نمودار صفحات وضعیت، صفحات وضعیت کامل و صفحات داده شناسایی‌شده را جدا نشان می‌دهد. این نمودار پوشش داده است، نه مقایسه دسترس‌پذیری، کارایی، مسئولیت یا هزینه دکل.'),
 5:('The figure shows selected field recognition within the identified data pages, not service prices or expenditure. An unrecognized field is not evidence of zero loss, no equipment or a missing physical event.', 'شکل پوشش شناسایی چند فیلد را در صفحات داده نشان می‌دهد، نه قیمت خدمت یا هزینه واقعی. شناسایی‌نشدن یک فیلد به معنی صفر بودن هرزروی، نبود تجهیز یا رخ‌ندادن رویداد نیست.'),
 7:('Four operations with changing source depth are shown as individual report-end samples. Report numbers are not drilling hours; points are not continuous trajectories or evidence of a learning curve. A depth reset or plateau requires operational context.', 'چهار عملیات دارای تغییر عمق در منبع به‌صورت نمونه‌های مستقل پایان گزارش نمایش داده شده‌اند. شماره گزارش ساعت حفاری نیست؛ نقاط مسیر پیوسته یا شاهد منحنی یادگیری نیستند. تغییر ناگهانی یا توقف عمق باید در متن عملیات بررسی شود.'),
 8:('Selected phases from six distinct programmes for wells 122 and 127 are shown with physical-page anchors. These are plans, not confirmed execution, a complete lifecycle or a cost comparison; the first and second sidetrack programmes remain distinct.', 'مراحل منتخب شش برنامه مستقل چاه‌های ۱۲۲ و ۱۲۷ همراه صفحه فیزیکی منبع نمایش داده شده‌اند. این‌ها برنامه‌اند، نه اجرای تأییدشده، چرخه کامل مداخلات یا مقایسه هزینه؛ دو برنامه انحراف اول و دوم جدا هستند.'),
 9:('The figure lists selected source differences and scope cautions for review, not a ranked risk matrix. Resolve well identity, datum, reporting basis and approvals before design or attribution.', 'شکل چند اختلاف منبع و محدودیت دامنه را برای تطبیق نشان می‌دهد، نه ماتریس رتبه‌بندی ریسک. پیش از طراحی یا انتساب باید هویت چاه، مبنای عمق، مبنای ثبت و مصوبات روشن شوند.'),
 10:('The loss plot pairs numeric entries and report-end depth on the same daily record where both are recognized. Values are not summed and the report-end depth does not locate the loss event. This is not a formation hazard or stuck-pipe map.', 'نمودار فقط مقادیر عددی هرزروی و عمق پایان گزارش همان روز را در صورت شناسایی هر دو نمایش می‌دهد. مقادیر جمع نشده‌اند و عمق پایان گزارش محل رویداد هرزروی نیست. این شکل نقشه مخاطره سازندی یا گیر رشته نیست.')}

MASTER={
1:('ENGINEERING REVIEW WARNING — This report is provisional. Source-coded coverage and selected source observations do not establish total schedule, actual expenditure, risk severity, event causation or contractor liability. Full reconciliation and page-rendered review are pending.','هشدار بازبینی مهندسی — این گزارش مقدماتی است. پوشش کدهای منبع و مشاهدات منتخب، زمان کل، هزینه واقعی، شدت ریسک، علت رویداد یا مسئولیت پیمانکار را اثبات نمی‌کند. تطبیق کامل و بازبینی صفحات رندرشده باقی است.'),
4:('Source-scoped analysis of thirteen daily packs and fourteen programmes','تحلیل مبتنی بر شواهد سیزده بسته روزانه و چهارده برنامه'),
5:('Based on original daily workbooks, CBS files and well programmes; extraction and engineering reconciliation are incomplete.','بر اساس کتاب‌کارهای روزانه، فایل‌های ساختار شکست هزینه و برنامه‌های چاه؛ استخراج و تطبیق مهندسی کامل نشده است.'),
11:('The archive contains 13 daily packs, 13 CBS workbooks and 14 programmes. It supports source-specific comparisons, but neither approved comparable budgets nor actual portfolio expenditure has been established. Cached CBS values, scope changes, rig-less work and invoices must be reconciled before any cost variance is reported.','آرشیو شامل ۱۳ بسته روزانه، ۱۳ کتاب‌کار هزینه و ۱۴ برنامه است. این منابع مقایسه‌های مشخص و قابل ارجاع را ممکن می‌کنند، اما بودجه‌های مصوب هم‌دامنه و هزینه واقعی کل سبد هنوز احراز نشده‌اند. پیش از گزارش انحراف هزینه باید مقادیر ذخیره‌شده، تغییر دامنه، خدمات بدون دکل و صورت‌حساب‌ها تطبیق شوند.'),
12:('Status codes are available on 1,301 recognized status pages, with incomplete coverage and distinct NPT, unplanned and mixed classifications. Their source labels do not establish productive time, technical causes or monetary liability. Selected program, BHA and fluid observations require independent engineering review.','کدهای وضعیت در ۱۳۰۱ صفحه وضعیت شناسایی‌شده با پوشش ناقص و تفکیک غیرمولد، غیربرنامه‌ای و کد ترکیبی موجودند. برچسب منبع زمان مولد، علت فنی یا مسئولیت مالی را ثابت نمی‌کند. مشاهدات منتخب برنامه، رشته ته‌چاهی و سیال به بازبینی مستقل مهندسی نیاز دارند.'),
13:('Table 0-1 — Source inventory and outstanding verification.','جدول ۰-۱ ـ موجودی منابع و موارد نیازمند تأیید.'),
14:('1. Programme scope and duration','۱. دامنه و مدت برنامه‌ها'),
18:('Programme duration labels and page references are retained in the detailed table below. Do not add contingency, rig moves or rig-less durations unless the source time basis and scope have been reconciled.','برچسب مدت و ارجاع صفحه برنامه در جدول تفصیلی زیر حفظ شده است. زمان احتیاطی، جابه‌جایی دکل یا خدمات بدون دکل بدون تطبیق مبنای زمانی و دامنه جمع نشوند.'),
23:('Coverage depends on the workbook family and parser. Counts of worksheets and recognized fields do not prove semantic completeness, successful execution or absence of unrecognized information.','پوشش به ساختار کتاب‌کار و روش استخراج وابسته است. تعداد شیت‌ها و فیلدهای شناسایی‌شده کامل‌بودن معنا، اجرای موفق یا نبود اطلاعات شناسایی‌نشده را ثابت نمی‌کند.'),
28:('3.1 Programme duration and actual schedule','۳-۱. مدت برنامه و زمان‌بندی واقعی'),
29:('Table 3-1 — Selected programme duration references.','جدول ۳-۱ ـ ارجاعات منتخب مدت برنامه.'),
31:('Table 3-2 — Selected CBS source values and review gates.','جدول ۳-۲ ـ مقادیر منتخب منابع هزینه و نقاط بررسی.'),
32:('4. Daily-record coverage','۴. پوشش گزارش‌های روزانه'),
36:('Table 4-1 — Source coverage and interpretation boundaries.','جدول ۴-۱ ـ پوشش منابع و محدودیت تفسیر.'),
37:('4.1 Source-page coverage by operation','۴-۱. پوشش صفحات منبع به تفکیک عملیات'),
38:('Table 4-2 — Recognized status and data pages.','جدول ۴-۲ ـ صفحات وضعیت و داده شناسایی‌شده.'),
39:('5. Data fields and programme service assignments','۵. فیلدهای داده و انتساب خدمات در برنامه'),
43:('5.1 Selected programme service assignments','۵-۱. انتساب خدمات منتخب در برنامه'),
44:('Table 5-1 — Selected service providers as printed in programme pages.','جدول ۵-۱ ـ خدمت‌دهندگان منتخب طبق صفحات برنامه.'),
45:('These are programme listings, not proof of a signed contract, actual presence, performance, payment or responsibility. TBA is retained as unresolved. The 207 programme lists PANIZ/SEALAND for electrical services and PGD for mud logging.','این موارد درج برنامه‌اند، نه اثبات قرارداد امضاشده، حضور واقعی، عملکرد، پرداخت یا مسئولیت. مورد تعیین‌نشده باز می‌ماند. برنامه ۲۰۷ برای خدمات الکتریکی پانیذ/سی‌لند و برای گل‌نگاری پی‌جی‌دی را درج کرده است.'),
46:('5.2 Source fluid-field coverage','۵-۲. پوشش فیلدهای سیال در منبع'),
47:('Table 5-2 — Numeric observations in recognized data pages.','جدول ۵-۲ ـ تعداد مشاهدات عددی در صفحات داده شناسایی‌شده.'),
48:('The counts refer to source fields, not total operation days or validated measurements. Zero recognized numeric loss entries means no numeric entry was recovered by the parser, not zero downhole loss. Individual loss entries are not summed.','تعدادها به فیلدهای منبع مربوط‌اند، نه کل روزهای عملیات یا اندازه‌گیری‌های تأییدشده. صفر مشاهده عددی هرزروی به معنی بازیابی‌نشدن مقدار عددی توسط ابزار است، نه هرزروی صفر؛ مقادیر هرزروی جمع نشده‌اند.'),
49:('6. Source status codes and reconciliation','۶. کدهای وضعیت منبع و تطبیق'),
53:('Table 6-1 — Selected source-review issues.','جدول ۶-۱ ـ موضوعات منتخب بازبینی منابع.'),
54:('6.2 Source status hours within complete records','۶-۲. ساعت کدهای منبع در گزارش‌های کامل'),
55:('Table 6-2 — Source status hours with partial coverage.','جدول ۶-۲ ـ ساعت کدهای منبع با پوشش جزئی.'),
56:('These hours are source-code sums over recognized complete pages only. Codes are not merged, omitted or ambiguous days are not zero, and no cause, contractor liability or total operational duration is inferred.','این ساعت‌ها جمع کدهای منبع فقط در صفحات کامل شناسایی‌شده‌اند. کدها ادغام نشده‌اند، روزهای حذف‌شده یا مبهم صفر نیستند و علت، مسئولیت پیمانکار یا مدت کل عملیات از آنها استنباط نمی‌شود.'),
57:('7. Selected report-end depth observations','۷. مشاهدات منتخب عمق پایان گزارش'),
61:('8. Programme phases for repeated well identifiers','۸. مراحل برنامه برای شناسه‌های تکراری چاه'),
65:('9. Source reconciliation examples','۹. نمونه‌های تطبیق منابع'),
69:('10. Source loss entries and report-end depth','۱۰. مقادیر هرزروی منبع و عمق پایان گزارش'),
74:('Each paragraph identifies a programme objective, source duration and data-sheet recognition. These are not a complete execution history, cost record, formation attribution or proven operational outcome.','هر بند هدف برنامه، مدت منبع و تعداد شیت‌های داده شناسایی‌شده را بیان می‌کند. این‌ها تاریخچه کامل اجرا، ثبت هزینه، انتساب سازندی یا نتیجه اثبات‌شده عملیات نیستند.'),
89:('Reconcile the approved scope, dates, units, daily activities, invoices and CBS formula errors before calculating contract or portfolio variances.','پیش از محاسبه انحراف قرارداد یا سبد، دامنه مصوب، تاریخ‌ها، واحدها، فعالیت‌های روزانه، صورت‌حساب‌ها و خطاهای فرمول هزینه تطبیق شوند.'),
90:('Reconstruct activity chronology and incident evidence before classifying technical causes, preventive actions or contractual responsibility.','پیش از طبقه‌بندی علل فنی، اقدامات پیشگیرانه یا مسئولیت قراردادی، توالی فعالیت‌ها و شواهد رخداد بازسازی شود.'),
91:('Keep rig, rig-less and service scopes separate; reconcile planned, contingency and actual records without assuming they are additive.','دامنه دکل، بدون دکل و خدمات جدا نگه داشته شود؛ سوابق برنامه، احتیاطی و واقعی بدون فرض جمع‌پذیری تطبیق شوند.'),
92:('Compare alternative campaign plans only after verifying intervention order, equipment, approvals, mobilisation and well-specific barriers.','برنامه‌های جایگزین کمپین فقط پس از تأیید ترتیب مداخلات، تجهیزات، مصوبات، بسیج و موانع اختصاصی هر چاه مقایسه شوند.'),
93:('Review loss and formation records against depth datum, individual event time, formation tops and approved well-control and loss-response procedures.','سوابق هرزروی و سازند با مبنای عمق، زمان هر رویداد، رأس‌های سازند و رویه‌های مصوب کنترل چاه و پاسخ به هرزروی تطبیق شوند.'),
94:('Do not use legacy unit costs, rig comparisons or inferred learning curves as tender benchmarks without verified comparable records.','هزینه واحد قدیمی، مقایسه دکل‌ها یا منحنی یادگیری استنباط‌شده بدون اسناد هم‌دامنه و تأییدشده مبنای مناقصه قرار نگیرد.'),
98:('Each operation has an interim bilingual report, not a complete validated daily history, cost reconstruction or causal attribution.','برای هر عملیات گزارش مقدماتی دوزبانه وجود دارد، نه تاریخچه روزانه کامل تأییدشده، بازسازی هزینه یا انتساب علّی.'),
101:('The evidence index preserves original file hashes, spreadsheet cells and physical PDF pages. Status-page and data-page extractions are separate; chart inputs and limitations are recorded in the evidence directory. Formula caches are not recalculated. Further reconciliation and page-level visual QA are necessary.','نمایه شواهد هش فایل‌های اصلی، سلول‌های صفحه‌گسترده و صفحات فیزیکی PDF را حفظ می‌کند. استخراج صفحات وضعیت و داده جداست؛ ورودی نمودارها و محدودیت‌ها در پوشه شواهد ثبت شده‌اند. مقادیر ذخیره‌شده فرمول‌ها بازمحاسبه نشده‌اند. تطبیق تکمیلی و کنترل بصری صفحه‌ای لازم است.'),
109:('Well-report W4 plots shares of recognized complete-day source status codes, not verified technical causes; formation-table snapshots likewise require depth and row reconciliation.','شکل W4 گزارش هر چاه سهم کدهای وضعیت منبع را در روزهای کامل شناسایی‌شده نشان می‌دهد، نه علل فنی تأییدشده؛ تصویر خام جدول سازند نیز به تطبیق عمق و ردیف نیاز دارد.')}
EXEC={
1:MASTER[1],4:('South Azadegan Field Development — Source evidence review','توسعه میدان آزادگان جنوبی ـ بازبینی شواهد منابع'),
11:('Portfolio actual spending and budget variance remain unverified; CBS cached values are not invoices or approved comparable budgets.','هزینه واقعی و انحراف بودجه کل سبد تأیید نشده‌اند؛ مقادیر ذخیره‌شده فایل هزینه، صورت‌حساب یا بودجه مصوب هم‌دامنه نیستند.'),
12:('Status codes cover recognized complete daily records only; NPT, unplanned and mixed codes remain distinct and cannot be priced without verified scope and rates.','کدهای وضعیت فقط گزارش‌های کامل شناسایی‌شده را پوشش می‌دهند؛ کدهای غیرمولد، غیربرنامه‌ای و ترکیبی جدا هستند و بدون دامنه و نرخ تأییدشده ارزش‌گذاری نمی‌شوند.'),
13:('Technical causes, contractor liability and proposed commercial remedies require incident reconstruction and contract review.','علل فنی، مسئولیت پیمانکار و راهکارهای تجاری پیشنهادی به بازسازی رویداد و بررسی قرارداد نیاز دارند.'),
14:('Rig comparison is not established by the current source coverage; recognize missing pages and reconcile assignments and time bases first.','مقایسه دکل‌ها با پوشش فعلی منابع احراز نشده است؛ ابتدا صفحات مفقود، تخصیص‌ها و مبنای زمانی باید تطبیق شوند.'),
15:('Programme references for wells 122 and 127 describe distinct proposed phases; they do not establish execution or a single campaign saving.','ارجاعات برنامه‌ای چاه‌های ۱۲۲ و ۱۲۷ مراحل پیشنهادی مستقل را شرح می‌دهند؛ اجرا یا صرفه‌جویی یک کمپین مشترک اثبات نشده است.'),
16:('Selected fluid loss and depth entries require unit, reporting-basis and event-depth review; the current exhibits do not establish formation causes.','مقادیر منتخب هرزروی و عمق به بررسی واحد، مبنای ثبت و عمق رویداد نیاز دارند؛ نمودارهای فعلی علت سازندی را اثبات نمی‌کنند.'),
29:('4. Source field coverage and economics gate','۴. پوشش فیلدهای منبع و درگاه بررسی اقتصادی'),
33:('5. Daily-record coverage','۵. پوشش گزارش‌های روزانه'),
37:('6. Programme phases for wells 122 and 127','۶. مراحل برنامه‌ای چاه‌های ۱۲۲ و ۱۲۷'),
41:('7. Source reconciliation examples','۷. نمونه‌های تطبیق منابع'),
45:('8. Decisions and evidence gates','۸. تصمیم‌ها و درگاه‌های شواهد'),
46:('Table 7-1 — Proposed evidence-review gates, not approved actions.','جدول ۷-۱ ـ درگاه‌های پیشنهادی بازبینی شواهد، نه اقدامات مصوب.'),
47:('9. Financial and technical exposure','۹. ریسک مالی و فنی'),
48:('The financial cost of nonproductive or unplanned time is not established. Quantification requires complete time records, contract terms, invoices, approval and a consistent cost scope; no annual projection is supported by the available review.','هزینه مالی زمان غیرمولد یا غیربرنامه‌ای احراز نشده است. محاسبه آن به سابقه کامل زمان، مفاد قرارداد، صورت‌حساب، مصوبات و دامنه هزینه سازگار نیاز دارد؛ بر اساس بازبینی فعلی پیش‌بینی سالانه قابل اتکا نیست.'),
49:('Table 8-1 — Selected review issues and evidence gates.','جدول ۸-۱ ـ موضوعات منتخب و درگاه‌های بررسی.'),
50:('10. Evidence needed for management decisions','۱۰. شواهد مورد نیاز تصمیم مدیریتی'),
51:('Review approved scope changes, complete daily reports, source classification and contract terms, plus validated invoices and equipment records. Only then evaluate performance, causation, cost, liability and any proposed contractual remedy.','تغییر دامنه مصوب، گزارش‌های روزانه کامل، طبقه‌بندی منبع، مفاد قرارداد، صورت‌حساب و سوابق معتبر تجهیزات بررسی شوند. ارزیابی عملکرد، علت، هزینه، مسئولیت و راهکار قراردادی فقط پس از آن انجام شود.'),
53:('11.1 Source inventory and loss/depth evidence','۱۱-۱. موجودی منابع و شواهد هرزروی و عمق'),
54:('The inventory figure counts workbook sheets, not data completeness. The loss figure pairs numeric source observations with report-end depths, not event depths or formations; missing entries are not zero.','شکل موجودی تعداد شیت‌های کتاب‌کار است، نه کامل‌بودن داده. شکل هرزروی، مقدار عددی منبع را با عمق پایان همان گزارش جفت می‌کند، نه عمق رویداد یا سازند؛ مقدار مفقود صفر نیست.'),
59:('Appendix A. Financial source gates','پیوست الف. درگاه‌های منابع مالی'),
60:('Table A-1 — Selected CBS source values and unresolved scope.','جدول الف-۱ ـ مقادیر منتخب منابع هزینه و دامنه حل‌نشده.'),
61:('Cached values and source labels alone do not establish actual spending, approved budget, allocation or savings. Recalculate formulas and reconcile invoices and scope before comparison.','مقادیر ذخیره‌شده و برچسب‌های منبع به‌تنهایی هزینه واقعی، بودجه مصوب، تخصیص یا صرفه‌جویی را ثابت نمی‌کنند. پیش از مقایسه باید فرمول‌ها بازمحاسبه و صورت‌حساب‌ها و دامنه تطبیق شوند.'),
64:('Table A-2 — Source-page recognition by operation.','جدول الف-۲ ـ پوشش شناسایی صفحات منبع به تفکیک عملیات.'),
66:MASTER[101],74:MASTER[109]}

# Bilingual, source-grounded financial examples, not additive cost or actual.
FIN=[
 ('122WO1','CBS AZNS-122WO1.xlsx / Cost Breakdown-122!G190','4,587,374.50 USD','Plan label; approval and scope pending','برچسب برنامه؛ مصوبه و دامنه نامشخص'),
 ('122WO1','CBS AZNS-122WO1.xlsx / Cost Breakdown-122!H190','1,304,440.016 USD','Extra work; not automatically additive','کار اضافی؛ جمع‌پذیری خودکار اثبات نشده'),
 ('85','CBS AZNS-85.xlsx / AZNS-085 WO. Extra Work!K54','48.0729167 days','Rig schedule; rig-less work separate','زمان‌بندی دکل؛ فعالیت بدون دکل جداست'),
 ('85','CBS AZNS-85.xlsx / AZNS-085 WO. Extra Work!O54','51.8895833 days','With contingency; scope pending','با احتیاط؛ دامنه نیازمند تطبیق'),
 ('116 / 127WO1','CBS AZNS-127WO1 and 116.xlsx / Total Cost-EPD & Rigless!C5','13,909,986.21 USD','Joint scope; do not allocate without evidence','دامنه مشترک؛ تخصیص بدون مدرک مجاز نیست')]
ISSUES=[
 ('40','Programme pp.13 / 149','20-inch casing depth 152.5 / 150 m','Reconcile depth and approved design','عمق جداری ۲۰ اینچ ۱۵۲٫۵ / ۱۵۰ متر','تطبیق عمق و طرح مصوب'),
 ('12 / 127WO1','Programme p.28 / p.58','Identical 3,889–3,911 m collapse wording','Verify well-specific source and event','عبارت یکسان فروریختگی ۳۸۸۹–۳۹۱۱ متر','احراز انتساب به چاه و رویداد'),
 ('127ST1','Programme pp.11 / 70','Sidetrack 3,300 m; window 3,337 m','Reconcile definitions and approved design','انحراف ۳۳۰۰ متر؛ پنجره ۳۳۳۷ متر','تطبیق تعریف و طرح مصوب'),
 ('122WO1','12-2 Data!I43 / K43','7-inch scraper description / 5-3/8-inch OD','Verify component specification','شرح خراشنده ۷ اینچ / قطر خارجی ۵⅜','تأیید مشخصات ابزار'),
 ('207','Programme p.16; 155-2 DDR Data!P86','Prognosed 3,915 m / DDR 3,968 m','Verify datum and actual/predicted status','پیش‌بینی ۳۹۱۵ متر / ثبت روزانه ۳۹۶۸ متر','تأیید مبنای عمق و ماهیت مقدار'),
 ('85','CBS: AZNS-085 WO. Extra Work rows 41–49','Rig-less steps excluded from rig schedule','Reconcile elapsed time and costs','مراحل بدون دکل خارج از زمان دکل','تطبیق زمان سپری‌شده و هزینه')]

def replace_table(d,old,rows,widths,fa):
 table=d.add_table(rows=1,cols=len(widths));table.style='Table Grid';old._tbl.addprevious(table._tbl);old._tbl.getparent().remove(old._tbl);set_table(table,rows,widths,fa)

def status_rows(fa):
 rows=[['عملیات','صفحات وضعیت','صفحات کامل','صفحات داده'] if fa else ['Operation','Status pages','Complete pages','Data pages']]
 return rows+[[x['well'],str(x['status_pages']),str(x['complete_status_pages']),str(x['recognized_data_pages'])] for x in C]

def financial_rows(fa):
 rows=[['عملیات','شیت / سلول منبع','مقدار ثبت‌شده','محدودیت'] if fa else ['Operation','Source sheet / cell','Cached value','Limitation']]
 for w,ref,val,en,fr in FIN:rows.append([w,ref,val.replace('USD','دلار').replace('days','روز') if fa else val,fr if fa else en])
 return rows

def issue_rows(fa):
 rows=[['عملیات','منبع','اختلاف یا ابهام','درگاه بررسی'] if fa else ['Operation','Source','Issue','Review gate']]
 for w,src,issue,gate,if_,gf in ISSUES:rows.append([w,src,if_ if fa else issue,gf if fa else gate])
 return rows

def status_hours(fa):
 rows=[['عملیات','غیرمولد','غیربرنامه‌ای','ترکیبی'] if fa else ['Operation','N / NPT h','UP h','Mixed h']]
 for c in C:
  complete=[d['activity'] for d in A[c['file']]['days'] if d['activity']['complete']]
  rows.append([c['well']]+[f"{sum(x[k] for x in complete):,.1f}" for k in ('npt_h','unplanned_h','mixed_h')])
 return rows

def service_rows(fa):
 translations={'VIRA':'ویرا','Pars Hafaran Arvand':'پارس حفاران اروند','SEALAND':'سی‌لند','PANIZ':'پانیذ','PANIZ/SEALAND':'پانیذ / سی‌لند','MAPSA KISH':'مپسا کیش','GEO DATA':'ژئودیتا','Pars Geo Data':'پارس ژئودیتا','PGD':'پی‌جی‌دی','APD':'ای‌پی‌دی','TBA':'تعیین‌نشده'}
 rows=[['عملیات','سیال','پسماند','نمودارگیری / مشبک','گل‌نگاری','صفحه'] if fa else ['Operation','Fluids','Waste','Logging / perf.','Mud logging','Page']]
 for x in ph:
  y=x['service_assignment_as_printed'];vals=[y[k] for k in ('drilling_completion_fluids','waste_management','electrical_logging_perforation','mud_logging')]
  if fa:vals=[translations[z] for z in vals]
  rows.append([x['well']]+vals+[str(x['services_physical_page'])])
 return rows

def fluid_rows(fa):
 rows=[['عملیات','صفحات داده','وزن گل','گرانروی','نقطه تسلیم','هرزروی'] if fa else ['Operation','Data pages','Mud weight','PV','YP','Loss']]
 for well in [x['well'] for x in C]:
  z=F[well];rows.append([well,str(z['recognized_data_sheets'])]+[str(z['fields'][k]['numeric_accepted_count']) for k in ('mud_weight','pv','yp','downhole_loss')])
 return rows

def main():
 for typ in ('Master_Drilling_Engineering','Executive_Management'):
  master=typ.startswith('Master')
  for lang in ('EN','FA'):
   fa=lang=='FA';path=Path(f'Field_{typ}_Report_{lang}.docx');d=Document(path);pars=d.paragraphs;tables=list(d.tables)
   # Fixed-index checks prevent silent edits to a different document layout.
   assert pars[10].style.name=='Heading 1' and len(tables)==(13 if master else 7)
   replacement=MASTER if master else EXEC
   for i,pair in replacement.items():text(pars[i],pair[1 if fa else 0])
   for n in (1,2,4,5,7,8,9,10):
    candidates=[(i,p) for i,p in enumerate(pars) if re.match(rf'^(?:Figure|شکل) M{n}(?:\s|\b)',p.text)]
    if n==7 and not master:assert not candidates;continue
    assert len(candidates)==1,(path,n,len(candidates));i,cap=candidates[0];bl=pars[i-1]._element.xpath('.//a:blip');assert len(bl)==1
    rid=bl[0].get(qn('r:embed'));d.part.related_parts[rid]._blob=(D/'field_figures'/f'M{n}.png').read_bytes()
    shape=next(s for s in d.inline_shapes if any(q.get(qn('r:embed'))==rid for q in s._inline.xpath('.//a:blip')));shape.height=int(shape.width*8/12)
    text(cap,('شکل' if fa else 'Figure')+f' M{n} — '+CAP[n][1 if fa else 0]);cap.alignment=1
    # Executive M10 has two preceding figures; preserve the shared M2 intro,
    # which was replaced explicitly above.
    if not (not master and n in (2,10)):text(pars[i-2],INTRO[n][1 if fa else 0])
   if master:
    replace_table(d,tables[0],([['شاخص','وضعیت']]+[['بسته روزانه','۱۳'],['کتاب‌کار هزینه','۱۳'],['برنامه چاه','۱۴'],['صفحات وضعیت شناسایی‌شده','۱۳۰۱'],['هزینه واقعی و زمان کل','نیازمند تطبیق']] if fa else [['Metric','Status'],['Daily packs','13'],['CBS workbooks','13'],['Well programmes','14'],['Recognized status pages','1,301'],['Actual expenditure and total time','Reconciliation pending']]),[7.3,9.2],fa)
    replace_table(d,tables[2],([['موضوع','وضعیت','اقدام']]+[['هزینه و بودجه','احراز نشده','تطبیق مصوبه، فرمول، دامنه و صورت‌حساب'],['زمان عملیات','احراز نشده','تطبیق گزارش، تاریخ و مبنای زمان'],['مسئولیت و علت','احراز نشده','بازسازی رخداد و بررسی قرارداد']] if fa else [['Topic','Status','Required review'],['Cost and budget','Not established','Reconcile approvals, formulas, scope and invoices'],['Operation duration','Not established','Reconcile reports, dates and time basis'],['Cause and liability','Not established','Reconstruct events and review contracts']]),[4,4,8.5],fa)
    subset=['207','40','75','127WO1','85'];z=[['برنامه','مدت درج‌شده','منبع / صفحه','محدودیت'] if fa else ['Programme','Stated duration','Source / page','Limitation']]
    for w in subset:
     x=P[w];z.append([w,x['duration_days_fa' if fa else 'duration_days_en'],f"{x['file']} / {x['physical_page']}",'نه زمان واقعی' if fa else 'Not actual time'])
    replace_table(d,tables[3],z,[2.0,4.0,7.2,3.3],fa)
    replace_table(d,tables[4],financial_rows(fa),[2.0,5.0,3.4,6.1],fa)
    coverage=([['سنجه','تعداد','محدودیت'],['بسته روزانه','۱۳','خانواده‌های مختلف شیت'],['صفحات وضعیت شناسایی‌شده','۱۳۰۱','کامل‌بودن زمانی هر صفحه جدا بررسی می‌شود'],['برنامه چاه','۱۴','برنامه‌ها اثبات اجرا نیستند']] if fa else [['Metric','Count','Limitation'],['Daily packs','13','Different sheet families'],['Recognized status pages','1,301','Completeness is checked per page'],['Well programmes','14','Plans are not proof of execution']])
    replace_table(d,tables[5],coverage,[5,2.5,9],fa)
    replace_table(d,tables[6],status_rows(fa),[4,4,4,4.5],fa)
    replace_table(d,tables[7],service_rows(fa),[2.0,2.4,3.4,3.5,3.1,1.4],fa)
    replace_table(d,tables[8],fluid_rows(fa),[3.0,2.7,2.7,2.7,2.7,2.7],fa)
    replace_table(d,tables[9],issue_rows(fa),[2.2,4.6,5,4.7],fa)
    replace_table(d,tables[10],status_hours(fa),[4.1,4.1,4.1,4.2],fa)
    replace_table(d,tables[11],status_rows(fa),[4,4,4,4.5],fa)
    text(pars[97],'جدول الف-۲ ـ پوشش صفحات منبع به تفکیک عملیات.' if fa else 'Table A-2 — Source-page coverage by operation.')
   else:
    kpi=([['شاخص','مقدار','خوانش'],['بسته روزانه','۱۳','پوشش متغیر'],['برنامه چاه','۱۴','شامل دو انحراف مستقل ۱۲۷'],['صفحات وضعیت','۱۳۰۱','نه مدت عملیات'],['هزینه واقعی','احراز نشده','نیازمند تطبیق'],['علت و مسئولیت','احراز نشده','نیازمند بررسی رخداد و قرارداد']] if fa else [['Measure','Value','Interpretation'],['Daily packs','13','Coverage varies'],['Well programmes','14','Includes two distinct 127 sidetracks'],['Recognized status pages','1,301','Not operation duration'],['Actual expenditure','Not established','Reconciliation pending'],['Cause and liability','Not established','Event and contract review required']])
    replace_table(d,tables[0],kpi,[5,3.5,8],fa)
    decisions=([['موضوع','شواهد لازم','درگاه','زمان'],['زمان و عملکرد','گزارش روزانه و تخصیص دکل','تطبیق پوشش و توالی','پیش از ارزیابی'],['بودجه و هزینه','مصوبه، فرمول، دامنه و صورت‌حساب','تطبیق مالی','پیش از مقایسه'],['علت و مسئولیت','روایت رویداد و قرارداد','بررسی مستقل','پیش از انتساب'],['طرح چاه','برنامه، وضعیت واقعی و موانع','بازبینی مهندسی','پیش از تصمیم']] if fa else [['Topic','Required evidence','Gate','Timing'],['Time and performance','Daily reports and rig assignment','Coverage and sequence review','Before assessment'],['Budget and cost','Approvals, formulas, scope and invoices','Financial reconciliation','Before comparison'],['Cause and liability','Event narrative and contract','Independent review','Before attribution'],['Well plan','Programme, as-built and barriers','Engineering review','Before decision']])
    replace_table(d,tables[2],decisions,[3.5,5,4.6,3.4],fa)
    replace_table(d,tables[3],issue_rows(fa),[2.2,4.6,5,4.7],fa)
    replace_table(d,tables[4],financial_rows(fa),[2.0,5.0,3.4,6.1],fa)
    replace_table(d,tables[5],status_rows(fa),[4,4,4,4.5],fa)
   d.save(path);print(path)
if __name__=='__main__':main()
