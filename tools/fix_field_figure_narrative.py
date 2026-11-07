#!/usr/bin/env python3
"""Correct unsupported narrative around the redrawn field figures."""
from docx import Document
replacements={
'The bubble matrix positions every operation by schedule ratio against cost variance': 'The chart compares the legacy schedule ratios and budget variances for operations with a reported budget entry. The 122 WO1 and 85 entries are excluded because their budget scopes are not reconciled. A point right of one and above zero indicates both a longer reported duration and a positive reported variance; this descriptive comparison does not prove learning or support a contracting decision without scope and baseline validation.',
'ماتریس حبابی هر عملیات را برحسب نسبت زمان اجرا به برنامه': 'نمودار نسبت زمان و انحراف هزینه گزارش قبلی را برای عملیات دارای بودجه قابل نمایش مقایسه می‌کند. 122 WO1 و 85 به‌دلیل تطبیق‌نشدن دامنه بودجه حذف شده‌اند. نقطه سمت راست نسبت یک و بالای صفر نشان‌دهنده مدت بیشتر و انحراف هزینه مثبت در گزارش قبلی است؛ این مقایسه بدون اعتبارسنجی دامنه و خط مبنا یادگیری یا تصمیم قراردادی را اثبات نمی‌کند.',
'The bubble matrix is the one-page picture of the portfolio': 'The chart compares the legacy schedule and budget figures for 11 operations. 122 WO1 and 85 are excluded because their budgets have not been reconciled. Differences between points do not by themselves establish causation, learning or a basis to extend a contract.',
'ماتریس حبابی تصویر یک‌صفحه‌ای سبد است': 'نمودار ارقام زمان و بودجه گزارش قبلی را برای ۱۱ عملیات مقایسه می‌کند. 122 WO1 و 85 به‌دلیل تطبیق‌نشدن بودجه حذف شده‌اند. تفاوت نقاط به‌تنهایی رابطه علّی، یادگیری یا مبنای تمدید قرارداد را اثبات نمی‌کند.',
'Mud loss and stuck-pipe events plotted against measured depth show': 'The legacy exhibit shows reported fluid-loss points against a depth value; it does not display stuck-pipe events or identify formations. Before interpreting clusters or reservoir proximity, reconcile the depths, loss records, stratigraphic tops and pressure/fracture-gradient data with the source reports.',
'رسم رویدادهای افت سیال و گیر رشته برحسب عمق اندازه‌گیری‌شده': 'نمودار قبلی نقاط افت سیال گزارش‌شده را در برابر مقدار عمق نمایش می‌دهد؛ رویدادهای گیر رشته یا نام سازندها در آن نشان داده نشده‌اند. پیش از تفسیر تمرکز رویدادها یا نزدیکی مخزن، عمق، سوابق هرزروی، رأس سازند و داده‌های فشار و گرادیان شکست با منابع تطبیق شوند.'
}
for kind in ('Master_Drilling_Engineering','Executive_Management'):
 for lang in ('EN','FA'):
  path=f'Field_{kind}_Report_{lang}.docx';d=Document(path)
  for p in d.paragraphs:
   for prefix,replacement in replacements.items():
    if p.text.startswith(prefix):
     p.runs[0].text=replacement
     for r in p.runs[1:]:r.text=''
     break
  d.save(path)
