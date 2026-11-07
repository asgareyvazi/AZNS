#!/usr/bin/env python3
"""Raw, cell-addressed BHA blocks in recognized daily data pages.

Layout recognition checks the component/OD/length header. Two concurrent
DDR BHA columns are independent and retained separately. Rows are not grouped
as unique runs or assumed to be chronological, installed or effective at MD.
"""
import argparse,json,re,sqlite3
from pathlib import Path
from collections import Counter

def cell(c,a,sheet):
 return {'sheet':sheet,'cell':a,'raw':c[a]} if a in c else None

def norm(s):return re.sub(r'\s+',' ',s.strip()).lower()
def extract(db,tracks):
 result=[]
 for well in tracks:
  file=well['file'];sheets=[];unrecognized=[]
  for day in well['days']:
   name=day['data_sheet'];sid=db.execute('select s.id from sheets s join files f on f.id=s.file_id where f.name=? and s.name=?',(file,name)).fetchone()[0]
   c=dict(db.execute('select address,raw_value from cells where sheet_id=? and raw_value is not null',(sid,)));blocks=[]
   for r in range(35,76):
    # Legacy blocks: row number, textual BHA label, count, OD, ID, length.
    if norm(c.get(f'H{r}',''))=='row' and norm(c.get(f'K{r}','')).startswith('od') and norm(c.get(f'M{r}','')).startswith('len'):
     entries=[]
     for rr in range(r+1,min(r+16,90)):
      item=cell(c,f'I{rr}',name)
      if not item:break
      entries.append({'row':rr,'item':item,'quantity':cell(c,f'J{rr}',name),'od':cell(c,f'K{rr}',name),'id':cell(c,f'L{rr}',name),'length':cell(c,f'M{rr}',name),'cumulative_length':cell(c,f'N{rr}',name)})
     if entries:blocks.append({'family':'legacy','header_row':r,'block_description':cell(c,f'I{r}',name),'identifier':cell(c,f'H{r-2}',name),'components':entries})
    # DDR blocks in normal B/F and shifted A/E column families.
    for itemcol,odcol,lencol,cumcol,identcol in [('B','C','D','E','B'),('F','G','H','I','F'),('A','B','C','D','A'),('E','F','G','H','E')]:
     if norm(c.get(f'{itemcol}{r}',''))!='item' or not norm(c.get(f'{odcol}{r}','')).startswith('od') or not norm(c.get(f'{lencol}{r}','')).startswith('length'):continue
     entries=[]
     for rr in range(r+1,min(r+16,90)):
      item=cell(c,f'{itemcol}{rr}',name)
      if not item:break
      entries.append({'row':rr,'item':item,'quantity':None,'od':cell(c,f'{odcol}{rr}',name),'id':None,'length':cell(c,f'{lencol}{rr}',name),'cumulative_length':cell(c,f'{cumcol}{rr}',name)})
     if entries:blocks.append({'family':'ddr','header_row':r,'block_description':None,'identifier':cell(c,f'{identcol}{r-1}',name),'components':entries})
   if blocks:sheets.append({'day':day['day'],'sheet':name,'blocks':blocks})
   else:unrecognized.append(name)
  result.append({'file':file,'data_sheets_considered':len(well['days']),'sheets_with_nonempty_bha_blocks':len(sheets),'sheets_without_recognized_nonempty_bha_blocks':unrecognized,'sheets':sheets,'limitations':'Raw display cells from data-linked sheets only. Blank rows terminate a block; later or differently formatted blocks may be missed. Multiple columns, repeating identifiers and sheet order do not prove unique BHA runs, measured downhole state, chronology, integrity or current configuration. No pressure-drop or hydraulic-limit calculation is inferred.'})
 return result

def main():
 p=argparse.ArgumentParser();p.add_argument('database',type=Path);p.add_argument('output',type=Path);a=p.parse_args();db=sqlite3.connect(f'file:{a.database.resolve()}?mode=ro',uri=True)
 tracks=json.loads(Path('evidence/drilling_track_observations.json').read_text());result=extract(db,tracks);a.output.write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n')
 for w in result:print(w['file'],w['sheets_with_nonempty_bha_blocks'],'/',w['data_sheets_considered'],'blocks',sum(len(s['blocks']) for s in w['sheets']),'components',sum(len(b['components']) for s in w['sheets'] for b in s['blocks']))
if __name__=='__main__':main()
