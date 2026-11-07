#!/usr/bin/env python3
"""Extract source-labelled daily mud-material fields without summing inventory.

Recognizes material header locations in legacy and DDR formats, including
shifted layouts. Original field meanings and units remain raw. Excludes
cement/additive tables outside the mud-material blocks.
"""
import argparse,json,re,sqlite3
from pathlib import Path

def norm(v):return re.sub(r'\s+',' ',v.strip()).lower().rstrip('.')
def scan(c,sheet):
 def data(address):
  if address not in c:return None
  raw,formula=c[address];return {'sheet':sheet,'cell':address,'raw':raw,'formula_xml':formula}
 for header_row,start,stop in [(7,8,49),(39,40,90)]:
  headers={}
  for i in range(1,35):
   col='';n=i
   while n:n,rem=divmod(n-1,26);col=chr(65+rem)+col
   z=data(f'{col}{header_row}')
   if not z or not z['raw']:continue
   h=norm(z['raw'])
   key={'material name':'material','product type':'material','unit':'unit','daily used':'used','used':'used','today rec':'received','receive':'received','received':'received','total rec':'total_received','stock':'stock','on hand':'stock'}.get(h)
   if key:
    if key in headers:return None
    headers[key]={'column':col,'cell':z}
  if not all(k in headers for k in ('material','unit','used','stock')):continue
  # Validate anchors co-located in the same block, not header-like cells
  # in unrelated tables elsewhere in the same row.
  def idx(c):
   n=0
   for x in c:n=n*26+ord(x)-64
   return n
  cs=[idx(headers[k]['column']) for k in ('material','used','unit','stock')]
  if max(cs)-min(cs)>20:continue
  entries=[]
  for r in range(start,stop+1):
   name=data(f"{headers['material']['column']}{r}")
   if not name or not (name['raw'] or '').strip():continue
   fields={k:data(f"{v['column']}{r}") for k,v in headers.items() if k not in ('material','unit')}
   if not any(v and ((v['raw'] is not None and v['raw']!='') or v['formula_xml']) for v in fields.values()):continue
   entries.append({'row':r,'material':name,'unit':data(f"{headers['unit']['column']}{r}"),'fields':fields})
  return {'family':'ddr' if header_row==7 else 'legacy','header_row':header_row,'headers':headers,'entries':entries}
 return None

def main():
 ap=argparse.ArgumentParser();ap.add_argument('database',type=Path);ap.add_argument('output',type=Path);arg=ap.parse_args()
 db=sqlite3.connect(f'file:{arg.database.resolve()}?mode=ro',uri=True);tracks=json.loads(Path('evidence/drilling_track_observations.json').read_text());out=[]
 for w in tracks:
  sheets=[];missing=[]
  for d in w['days']:
   name=d['data_sheet'];sid=db.execute('select s.id from sheets s join files f on f.id=s.file_id where f.name=? and s.name=?',(w['file'],name)).fetchone()[0]
   c={address:(raw,formula) for address,raw,formula in db.execute('select address,raw_value,formula_xml from cells where sheet_id=?',(sid,))}
   block=scan(c,name)
   if block:sheets.append({'day':d['day'],'sheet':name,**block})
   else:missing.append(name)
  out.append({'file':w['file'],'recognized_data_sheets':len(w['days']),'recognized_material_blocks':len(sheets),'unrecognized_sheets':missing,'sheets':sheets,'limitations':'Raw cached values, units and formula XML only. Header recognition and row limits may miss blocks. Stock, total received, received and used are distinct; blank and errors are not zero. Units, package sizes, source quality, duplicate rows and daily versus cumulative basis require reconciliation. No conversion or sum is calculated; a row does not prove downhole usage or a causal outcome.'})
  print(w['file'],len(sheets),'/',len(w['days']),'entries',sum(len(s['entries']) for s in sheets))
 arg.output.write_text(json.dumps(out,ensure_ascii=False,indent=2)+'\n')
if __name__=='__main__':main()
