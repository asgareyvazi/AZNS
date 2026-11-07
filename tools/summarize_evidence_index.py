#!/usr/bin/env python3
"""Summarize a raw evidence index; do not treat keyword hits as verified facts."""
import argparse, json, sqlite3
from pathlib import Path

def summary(db):
    files=[]
    for fid,name,digest,size,kind,status in db.execute('SELECT id,name,sha256,size,kind,status FROM files ORDER BY name'):
        item={'file':name,'sha256':digest,'bytes':size,'kind':kind,'status':status}
        if kind=='xlsx':
            item['sheets'],item['nonempty_cells']=db.execute('SELECT count(*),coalesce(sum(cell_count),0) FROM sheets WHERE file_id=?',(fid,)).fetchone()
        else:
            item['pages'],item['blank_text_pages']=db.execute("SELECT count(*),coalesce(sum(CASE WHEN length(trim(text))=0 THEN 1 ELSE 0 END),0) FROM pages WHERE file_id=?",(fid,)).fetchone()
        files.append(item)
    sheets=[{'file':a,'sheet':b,'nonempty_cells':c} for a,b,c in db.execute('''SELECT f.name,s.name,s.cell_count FROM sheets s JOIN files f ON f.id=s.file_id
        WHERE f.name LIKE 'CBS%' AND (lower(s.name) LIKE '%rigless%' OR lower(s.name) LIKE '%rig-less%' OR lower(s.name) LIKE '%acid%' OR lower(s.name) LIKE '%coil%' OR lower(s.name) LIKE '%testing%' OR lower(s.name) LIKE '%slick%' OR lower(s.name) LIKE '%sampling%' OR lower(s.name) LIKE '%wireline prod%' OR lower(s.name) LIKE '%memory gauge%' OR lower(s.name) LIKE '%oil transportation%' OR lower(s.name) LIKE '%pvt samp%' OR lower(s.name) LIKE '%r-less stim%') ORDER BY f.name,s.name''')]
    return {'description':'Raw inventory only; extracted text and cached Excel values are not verified engineering facts. PDF physical page numbers may differ from printed page numbers. Numeric spreadsheet values may be Excel date serials; units and formulas require review.',
            'files':files,'stimulation_sheet_candidates':sheets,
            'errors':[{'file':a,'location':b,'message':c} for a,b,c in db.execute('SELECT f.name,e.location,e.message FROM errors e JOIN files f ON f.id=e.file_id ORDER BY f.name,e.location')]}

def main():
    p=argparse.ArgumentParser();p.add_argument('database',type=Path);p.add_argument('output',type=Path);a=p.parse_args()
    db=sqlite3.connect(f'file:{a.database.resolve()}?mode=ro',uri=True)
    out=summary(db);a.output.parent.mkdir(parents=True,exist_ok=True);a.output.write_text(json.dumps(out,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(f"{len(out['files'])} files, {len(out['stimulation_sheet_candidates'])} candidate service tabs, {len(out['errors'])} extraction errors")
if __name__=='__main__':main()
