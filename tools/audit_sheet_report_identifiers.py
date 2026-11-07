#!/usr/bin/env python3
"""Compare worksheet numeric prefix against the source report-number header.

The SQLite index must have been built directly from the original packs.
Neither workbook ordering nor the report-number header proves chronology.
"""
import argparse,json,sqlite3
from pathlib import Path

def main():
 p=argparse.ArgumentParser();p.add_argument('database',type=Path);p.add_argument('out',type=Path);a=p.parse_args();db=sqlite3.connect(f'file:{a.database.resolve()}?mode=ro',uri=True)
 observations=json.loads(Path('evidence/activity_observations.json').read_text());books=[]
 for w in observations:
  rows=[]
  for d in w['days']:
   address='AQ5' if d['family']=='ddr' else 'S5';label='AL5' if d['family']=='ddr' else 'P5'
   data=db.execute('select address,raw_value,type,formula_xml from cells join sheets on sheets.id=cells.sheet_id join files on files.id=sheets.file_id where files.name=? and sheets.name=? and address in (?,?)',(w['file'],d['remark_sheet'],address,label)).fetchall();data={r[0]:r for r in data};header=data.get(address);lbl=data.get(label)
   number=None
   if header:
    try:
     n=float(header[1]);number=int(n) if n.is_integer() else None
    except (TypeError,ValueError):pass
   label_text=lbl[1] if lbl else None
   if not label_text or 'report no' not in label_text.lower():number=None
   rows.append({'sheet':d['remark_sheet'],'sheet_number':d['day'],'header_cell':address,'header_label_cell':label,'header_label':label_text,'header_raw':header[1] if header else None,'header_type':header[2] if header else None,'header_formula_xml':header[3] if header else None,'header_report_number':number,'matches_prefix':None if number is None else number==d['day']})
  books.append({'file':w['file'],'status_pages':len(rows),'unreadable_header':sum(r['matches_prefix'] is None for r in rows),'mismatched':sum(r['matches_prefix'] is False for r in rows),'rows':rows})
 a.out.write_text(json.dumps({'method':'Compare numeric worksheet prefix with indexed report-number header in recognized status pages. Header labels and cached values only; neither field proves chronological ordering, date, completeness or execution.','workbooks':books},ensure_ascii=False,indent=2)+'\n')
 for w in books:print(w['file'],'pages',w['status_pages'],'mismatches',w['mismatched'],'unreadable',w['unreadable_header'])
if __name__=='__main__':main()
