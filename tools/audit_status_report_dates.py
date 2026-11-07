#!/usr/bin/env python3
"""Inventory recognizable source report dates; never equate sheet order with time."""
from pathlib import Path
from datetime import date,datetime
import json,re,calendar
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
MONTHS={k.lower():v for v in range(1,13) for k in (calendar.month_name[v],calendar.month_abbr[v])}

def raw(v):
 return v.isoformat() if isinstance(v,(datetime,date)) else v

def legacy(v):
 if isinstance(v,(datetime,date)):return v.date().isoformat() if isinstance(v,datetime) else v.isoformat()
 if not isinstance(v,str):return None
 # Only year-first Gregorian representations; leave ambiguous locale dates unresolved.
 m=re.fullmatch(r'\s*(\d{4})[./-](\d{1,2})[./-](\d{1,2})\s*',v)
 if not m or not 1900<=int(m.group(1))<=2100:return None
 try:return date(*map(int,m.groups())).isoformat()
 except ValueError:return None

def ddr(year,month,day):
 try:
  y=int(year);d=int(day)
  if not 1900<=y<=2100 or float(year)!=y or float(day)!=d:return None
  if isinstance(month,str):
   z=month.strip().lower();m=MONTHS.get(z) if not z.isdigit() else int(z)
  else:m=int(month)
  return date(y,m,d).isoformat()
 except (ValueError,TypeError,OverflowError):return None

def main():
 audits=json.loads((ROOT/'evidence/sheet_report_number_audit.json').read_text())['workbooks'];activity={x['file']:x for x in json.loads((ROOT/'evidence/activity_observations.json').read_text())};books=[]
 for book in audits:
  wb=load_workbook(ROOT/book['file'],read_only=True,data_only=True);formulas=load_workbook(ROOT/book['file'],read_only=True,data_only=False);family={x['remark_sheet']:x['family'] for x in activity[book['file']]['days']};rows=[]
  for r in book['rows']:
   s=wb[r['sheet']];kind=family[r['sheet']];addr=['F11','H11','J11'] if kind=='ddr' else ['N5','S9'];values={c:raw(s[c].value) for c in addr}
   fs=formulas[r['sheet']];formula_cells={c:fs[c].value for c in addr if fs[c].data_type=='f'}
   volatile=any(re.search(r'(?i)(?:TODAY|NOW|RAND|RANDBETWEEN|OFFSET|INDIRECT)\s*\(',v) for v in formula_cells.values())
   dt=None if volatile else (ddr(*(s[c].value for c in addr)) if kind=='ddr' else legacy(s['N5'].value))
   rows.append({'sheet':r['sheet'],'sheet_prefix':r['sheet_number'],'report_number':r['header_report_number'],'family':kind,'source_cells':values,'formula_cells':formula_cells,'volatile_date_formula':volatile,'report_date_iso_from_header':dt,'interpretation':'DDR split year-month-day' if kind=='ddr' else 'legacy N5 datetime or year-first Gregorian text; S9 recorded without conversion'})
  wb.close();formulas.close();rows.sort(key=lambda r:r['sheet_prefix']);comparisons=[]
  # Only immediately consecutive recognized rows with dates, not inferred missing dates.
  for x,y in zip(rows,rows[1:]):
   if not x['report_date_iso_from_header'] or not y['report_date_iso_from_header']:continue
   delta=(date.fromisoformat(y['report_date_iso_from_header'])-date.fromisoformat(x['report_date_iso_from_header'])).days
   if delta<=0:comparisons.append({'earlier_sheet':x['sheet'],'earlier_date':x['report_date_iso_from_header'],'later_sheet':y['sheet'],'later_date':y['report_date_iso_from_header'],'delta_days':delta,'kind':'backwards' if delta<0 else 'same_date'})
  books.append({'file':book['file'],'recognized_status_pages':len(rows),'parseable_date_headers':sum(r['report_date_iso_from_header'] is not None for r in rows),'unparsed_date_headers':sum(r['report_date_iso_from_header'] is None for r in rows),'volatile_date_formulas':sum(r['volatile_date_formula'] for r in rows),'adjacent_backwards':sum(x['kind']=='backwards' for x in comparisons),'adjacent_same_date':sum(x['kind']=='same_date' for x in comparisons),'rows':rows,'non_increasing_adjacent_headers':comparisons})
 out={'method':'Source report-date header inventory; legacy N5 accepts Excel datetime or unambiguous year-first Gregorian text only, retaining S9 raw; DDR uses F11/H11/J11; volatile date formulas are retained with cached values but excluded from date comparisons. Non-increasing comparisons are between adjacent recognized status worksheets in numeric-prefix order with both dates parseable. Date readings and prefix ordering are not independently verified, and matching/equal dates are not necessarily errors. No elapsed campaign durations inferred.','workbooks':books}
 (ROOT/'evidence/status_report_date_audit.json').write_text(json.dumps(out,indent=2,ensure_ascii=False)+'\n')
 for b in books:print(b['file'],b['recognized_status_pages'],b['parseable_date_headers'],'backwards',b['adjacent_backwards'],'same',b['adjacent_same_date'],'unparsed',b['unparsed_date_headers'],'volatile',b['volatile_date_formulas'])
if __name__=='__main__':main()
