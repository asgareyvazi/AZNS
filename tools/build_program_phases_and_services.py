#!/usr/bin/env python3
"""Review selected plan phases and source programme service assignments.

Physical pages and exact text anchors are validated against the source PDFs.
Plans are not evidence of execution, contractual award, liability, or current
well condition. The full services table remains in each programme.
"""
import json,re
from pathlib import Path
from pypdf import PdfReader

# Each row: page, exact anchor, English summary, Persian summary.
PHASES={
'127ST1':[(70,'open window @ 3337m','Plan a cased-hole window at 3,337 m and drill the 8½-inch section to 3,939 m, then run and cement a 7-inch liner.','برنامه گشودن پنجره در عمق ۳۳۳۷ متر، حفاری بخش ۸½ اینچ تا ۳۹۳۹ متر و نصب و سیمان‌کاری آستری ۷ اینچ را بیان می‌کند.'),(71,'perform 7" liner lap drey test','After the liner, the programme describes a liner-lap dry test and cement-quality logging.','پس از آستری، برنامه آزمون خشک هم‌پوشانی آستری و نمودارگیری کیفیت سیمان را شرح می‌دهد.')],
'127ST2':[(34,'KOP will be @3055m','Plan a whipstock setting near 3,060 m, a 3,055 m kickoff and window opening in 9⅝-inch casing.','برنامه نصب ویپ‌استاک نزدیک ۳۰۶۰ متر، نقطه انحراف ۳۰۵۵ متر و گشودن پنجره در جداری ۹⅝ اینچ را ذکر می‌کند.'),(29,'Casing Running Fah. 3940','The time breakdown includes drilling to 3,940 m, liner running, cementing and a dry test; it is a schedule, not execution evidence.','جدول زمان، حفاری تا ۳۹۴۰ متر، راندن آستری، سیمان‌کاری و آزمون خشک را در برنامه دارد؛ جدول زمان سند اجرای عملیات نیست.')],
'10':[(29,'D/O cement plug#1','The plan calls for drilling out cement plugs and cleaning the liner before completing the well.','برنامه حفاری پلاگ‌های سیمانی و تمیزکاری آستری را پیش از تکمیل چاه پیش‌بینی می‌کند.'),(29,'pressure test same','The opening sequence includes wellhead checks and BOP pressure testing.','توالی آغازین شامل بررسی سرچاه و آزمون فشار فوران‌گیر است.')],
'12':[(29,'try to catch & retrieve fish','The programme describes pulling the completion and attempting to catch and retrieve the fish.','برنامه برداشت رشته تکمیل و تلاش برای گرفتن و بازیابی قطعه جامانده را شرح می‌دهد.'),(29,'mill completion PKR','The next planned operations include clean-out and milling the completion packer.','عملیات بعدی برنامه شامل تمیزکاری و آسیاب‌کاری پکر تکمیلی است.')],
'75':[(33,'the completion string should be pulled out','The stated main sequence is to pull completion, sidetrack the well and recomplete it.','توالی اصلی اعلام‌شده برداشت رشته تکمیل، انحراف چاه و تکمیل مجدد است.'),(27,'Nipple up BOP and test same','The time breakdown includes installation and testing of the BOP before the listed completion-string removal.','جدول زمان نصب و آزمون فوران‌گیر را پیش از برداشت رشته تکمیل درج می‌کند.')],
'85':[(26,'perform 7" tied-back liner lap dry test','The programme includes a 7-inch tied-back liner-lap dry test following cement-plug drilling and clean-out.','برنامه پس از حفاری پلاگ سیمانی و تمیزکاری، آزمون خشک هم‌پوشانی آستری ۷ اینچ بازگردان را پیش‌بینی می‌کند.'),(26,'perform 7" liner lap dry test','The proposed sequence also includes logging, tentative Fahliyan perforation and completion; a test outcome is not established by the plan.','توالی پیشنهادی نمودارگیری، مشبک‌کاری موقت فهلیان و تکمیل را نیز در بر دارد؛ نتیجه آزمون از برنامه اثبات نمی‌شود.')],
'116':[(46,'resume mill fish (CTU pipe)','The plan describes milling a coiled-tubing fish and permanent packer, followed by a fishing attempt.','برنامه آسیاب‌کاری قطعه جامانده لوله مغزی سیار و پکر دائمی و سپس تلاش ماهیگیری را شرح می‌دهد.'),(46,'high risk of String stuck','The programme identifies low-clearance metal-to-metal sticking risk during clean-out inside the 4½-inch liner.','برنامه برای تمیزکاری درون آستری ۴½ اینچ، خطر گیر رشته به‌علت فاصله کم فلز با فلز را ذکر می‌کند.')],
'122WO1':[(26,'mill Upper packer','The programme describes removal of the completion and milling and fishing the upper packer.','برنامه برداشت رشته تکمیل و آسیاب‌کاری و ماهیگیری پکر بالایی را شرح می‌دهد.'),(26,'Liner Lap Positive Test','The sequence also includes corrosion logging, a liner-lap positive test and recompletion.','توالی برنامه نمودارگیری خوردگی، آزمون مثبت هم‌پوشانی آستری و تکمیل مجدد را نیز شامل می‌شود.')],
'122WO2':[(29,'Milling Completion PKR','The plan describes removal of the completion and milling and fishing the completion packer.','برنامه برداشت رشته تکمیل و آسیاب‌کاری و ماهیگیری پکر تکمیلی را شرح می‌دهد.'),(29,'Impression Block','The same sequence proposes an impression-block run and a fishing assembly.','همان توالی راندن بلوک اثرپذیر و مجموعه ماهیگیری را پیشنهاد می‌کند.')],
'127WO1':[(58,'collapse interval from 3889-3911m','The programme identifies an interval at 3,889–3,911 m to check and smooth, and plans completion removal and recompletion.','برنامه بازه ۳۸۸۹ تا ۳۹۱۱ متر را برای بررسی و هموارسازی ذکر و برداشت و تکمیل مجدد را پیش‌بینی می‌کند.'),(59,'Follow contingency plan #4','The planned clean-out includes a contingency if the stated interval cannot be passed; prior client approval is mentioned for certain remedial work.','در تمیزکاری برنامه‌ریزی‌شده، برای عبور نکردن از بازه ذکرشده مسیر احتیاطی و برای برخی اقدامات اصلاحی تأیید کارفرما مطرح شده است.')],
'127WO2':[(45,'de – complete he well','The main objective is completion removal and recompletion; the opening steps include wellhead checks and slickline operations.','هدف اصلی برداشت و نصب مجدد تکمیل است؛ مراحل آغازین بررسی سرچاه و عملیات سیم نازک را شامل می‌شود.'),(45,'perform BOP pressure test','The programme describes installing and testing BOPs before retrieving the completion.','برنامه نصب و آزمون فوران‌گیر را پیش از بازیابی رشته تکمیل شرح می‌دهد.')],
'166':[(30,'Rig Less Operation','The programme distinguishes an initial rig-less sequence and conditional transition to operations with a rig.','برنامه توالی اولیه بدون دکل و انتقال مشروط به عملیات با دکل را از هم تفکیک می‌کند.'),(31,'Rig less Sequence of Operation','After the planned rig release, the programme includes further rig-less slickline, coiled-tubing and stimulation activities.','پس از ترخیص برنامه‌ریزی‌شده دکل، برنامه عملیات بدون دکل شامل سیم نازک، لوله مغزی سیار و تحریک را پیش‌بینی می‌کند.')],
'207':[(50,'Drill ahead formation down to 150 m','The 26-inch section is planned to 150 m with a 20-inch casing point.','بخش ۲۶ اینچ تا ۱۵۰ متر با نقطه جداری ۲۰ اینچ در برنامه ذکر شده است.'),(55,'Resume DRLG 12-1/4” Hole Section to 3915m','The 12¼-inch section is planned to 3,915 m for 9⅝-inch casing, subject to formation-top review.','بخش ۱۲¼ اینچ برای جداری ۹⅝ اینچ تا ۳۹۱۵ متر در برنامه آمده و مشروط به بررسی رأس سازند است.')],
'40':[(75,'D/Out CMT, Shoe and Rat hole to 1493m','The re-entry sequence includes drilling out cement and shoe to the stated 1,493 m.','توالی بازورود، حفاری سیمان و کفشک تا عمق ذکرشده ۱۴۹۳ متر را شامل می‌شود.'),(78,'Full Suite /CDR/GR log Tools','The plan describes wireline logging across 1,493–3,986 m before casing operations.','برنامه نمودارگیری کابلی در بازه ۱۴۹۳ تا ۳۹۸۶ متر را پیش از عملیات جداری شرح می‌دهد.')],
}
SERVICES={'127ST1':65,'127ST2':28,'10':26,'12':24,'122WO1':23,'122WO2':25,'40':21,'75':25,'116':42,'127WO2':42,'127WO1':55,'166':26,'207':180,'85':23}
# Only a narrow set is transcribed; no blanket inference from service labels.
PATTERN=re.compile(r'2 Drilling Fluids (?:and|&) Completion Fluids (.*?) 3 Waste Management (.*?) 4 Electrical Logging (?:and|&) Perforation (.*?) 5 Mud Logging (.*?) 6 ',re.I)

def main():
 summary=json.loads(Path('evidence/program_summary_reviewed.json').read_text());out=[]
 for x in summary:
  well=x['well'];pdf=PdfReader(x['file']);ph=[]
  for page,anchor,en,fa in PHASES[well]:
   text=' '.join((pdf.pages[page-1].extract_text() or '').split());assert anchor in text,(well,page,anchor)
   ph.append({'physical_page':page,'source_anchor':anchor,'english':en,'persian':fa})
  p=SERVICES[well];text=' '.join((pdf.pages[p-1].extract_text() or '').split());m=PATTERN.search(text);assert m,(well,p,text)
  service=dict(zip(('drilling_completion_fluids','waste_management','electrical_logging_perforation','mud_logging'),m.groups()))
  out.append({'well':well,'file':x['file'],'selected_plan_phases':ph,'services_physical_page':p,'service_assignment_as_printed':service,'services_raw_excerpt':m.group(0),'caveat':'Selected programme statements, not evidence of execution or current assignment. No liability inferred.'})
 Path('evidence/program_phases_services_reviewed.json').write_text(json.dumps(out,ensure_ascii=False,indent=2)+'\n');print('Validated',len(out),'programme PDFs,',len(out)*2,'phase anchors and',len(out)*4,'service assignments.')
if __name__=='__main__':main()
