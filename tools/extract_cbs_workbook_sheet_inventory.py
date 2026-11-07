#!/usr/bin/env python3
"""Inventory all CBS sheets and their indexed formula/error coverage.

Sheet names alone are not evidence of a particular work scope, value,
execution, approvals or cost. Numeric cells are not recalculated.
"""
import argparse,json,sqlite3,re
from pathlib import Path

def main():
 p=argparse.ArgumentParser();p.add_argument('database',type=Path);p.add_argument('output',type=Path);a=p.parse_args();db=sqlite3.connect(f'file:{a.database.resolve()}?mode=ro',uri=True)
 files=db.execute("select id,name from files where name like 'CBS AZNS-%' and kind='xlsx' order by name").fetchall();assert len(files)==13;out=[]
 for file_id,name in files:
  sheets=[]
  for sid,title,state,cell_count in db.execute('select id,name,state,cell_count from sheets where file_id=? order by id',(file_id,)):
   row=db.execute('select sum(case when formula_xml is not null then 1 else 0 end),sum(case when type=\'e\' then 1 else 0 end),sum(case when formula_xml is not null and raw_value is null then 1 else 0 end) from cells where sheet_id=?',(sid,)).fetchone()
   title_cell=db.execute("select raw_value from cells where sheet_id=? and address='A1'",(sid,)).fetchone()
   sheets.append({'sheet':title,'state':state,'indexed_nonempty_cells':cell_count,'formula_cells':row[0] or 0,'cached_error_cells':row[1] or 0,'formula_cells_missing_cache':row[2] or 0,'A1':title_cell[0] if title_cell else None})
  out.append({'file':name,'sheet_count':len(sheets),'sheets':sheets})
 a.output.write_text(json.dumps({'description':'Inventory of all indexed CBS sheets, not a semantic financial audit or evidence of work execution. A1 titles may be blank or stale. Counts depend on indexed cached OOXML.','workbooks':out},ensure_ascii=False,indent=2)+'\n')
 print('workbooks',len(out),'sheets',sum(x['sheet_count'] for x in out),'cached error cells',sum(s['cached_error_cells'] for x in out for s in x['sheets']))
if __name__=='__main__':main()
