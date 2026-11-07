#!/usr/bin/env python3
"""Extract daily status pages independently of the presence of a data page.

A workbook may put the daily Remark page in the '-2' slot, or have only a
Lookahead in the data slot. Neither case invalidates the actual remark page.
Preserve source rows and narrative as raw text; do not infer root cause.
"""
import argparse,json,re,sqlite3
from pathlib import Path
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles.numbers import BUILTIN_FORMATS,is_date_format
from extract_drilling_tracks import FILES,activity_day,get_row

def time_style_ids(file):
 wb=load_workbook(file,read_only=True,data_only=True);ids=set()
 for n,s in enumerate(wb._cell_styles):
  fmt=BUILTIN_FORMATS.get(s.numFmtId) if s.numFmtId<164 else (wb._number_formats[s.numFmtId-164] if s.numFmtId-164<len(wb._number_formats) else None)
  if fmt and is_date_format(fmt) and 'h' in fmt.lower() and 'y' not in fmt.lower() and 'd' not in fmt.lower():ids.add(str(n))
 wb.close();return ids

def extract(db,file):
 sheets=db.execute('select s.id,s.name from sheets s join files f on f.id=s.file_id where f.name=? order by s.id',(file,)).fetchall();candidates=defaultdict(list)
 for sid,name in sheets:
  match=re.match(r'^(\d+)-(1|2)\b',name)
  if not match:continue
  # Remark may be in slot 2; exclude data and Lookahead by checking the
  # activity header, not sheet name alone.
  r=db.execute('select address,raw_value,style_id,type from cells where sheet_id=? and raw_value is not null',(sid,)).fetchall()
  cells={a:v for a,v,style,typ in r}
  family='legacy' if get_row(cells,'J',r'^Code 2 P-N-UP') is not None else ('ddr' if get_row(cells,'M',r'^Status$') is not None else None)
  if family:candidates[int(match[1])].append((int(match[2]),sid,name,family,r))
 styles=time_style_ids(file);days=[];ambiguous=[]
 for day,options in sorted(candidates.items()):
  options.sort(key=lambda x:x[0]);_,sid,name,family,rows=options[0]
  if len(options)>1:ambiguous.append({'day':day,'sheets':[x[2] for x in options],'selected':name})
  cells={a:v for a,v,style,typ in rows};stylemap={a:style for a,v,style,typ in rows if typ in (None,'n')}
  a=activity_day(cells,name,family,stylemap,styles);detailscol='K' if family=='legacy' else 'S';codecol='I'
  for segment in a['segments']:
   row=segment['row'];address=f'{detailscol}{row}';code=f'{codecol}{row}'
   segment['detail']={'sheet':name,'cell':address,'raw':cells[address]} if address in cells else None
   segment['code']={'sheet':name,'cell':code,'raw':cells[code]} if code in cells else None
  # Save all narrative cells from the active block (not attached to an
  # adjacent time segment unless they share its row).
  header=get_row(cells,'J',r'^Code 2 P-N-UP') if family=='legacy' else get_row(cells,'M',r'^Status$')
  stop=next((r for r in range(header+1,min(header+55,90)) if cells.get(f'A{r}','').strip().lower().startswith('total')),min(header+55,90))
  texts=[{'sheet':name,'cell':f'{detailscol}{r}','raw':cells[f'{detailscol}{r}']} for r in range(header+1,stop) if f'{detailscol}{r}' in cells]
  days.append({'day':day,'remark_sheet':name,'family':family,'activity':a,'detail_cells':texts})
 return {'file':file,'remark_days':len(days),'duplicate_status_page_days':ambiguous,'days':days,
  'limitations':'Remark pages identified independently of data pages. Duplicated day numbers select the first slot and are flagged. Status codes remain source codes, not independently verified causes or liability. Complete means recognized durations and statuses total approximately 24 hours; it does not validate their classification. Detail cells are not automatically assigned across rows. Report day number is not verified elapsed duration.'}

def main():
 p=argparse.ArgumentParser();p.add_argument('database',type=Path);p.add_argument('output',type=Path);a=p.parse_args();db=sqlite3.connect(f'file:{a.database.resolve()}?mode=ro',uri=True)
 result=[extract(db,file) for file in FILES];a.output.write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n')
 for w in result:print(w['file'],'remark pages',w['remark_days'],'complete',sum(d['activity']['complete'] for d in w['days']),'duplicate',len(w['duplicate_status_page_days']))
if __name__=='__main__':main()
