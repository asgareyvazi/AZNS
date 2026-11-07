#!/usr/bin/env python3
"""Extract a traceable snapshot of source formation tables, not a geological model.
Select the last worksheet in workbook order with a populated formation table;
if no depths are populated, select the last table containing formation names.
Column roles are taken from visible source headers, never assumed actual/planned.
"""
import argparse,json,re,sqlite3
from pathlib import Path

cell_re=re.compile(r'^([A-Z]+)(\d+)$')
formation=re.compile(r'aghaj|mishan|gach|asmar|pabdeh|jahrum|gurpi|tarbur|ilam|laf{1,2}an|sarvak|s[ae]rv[ae]k|kazh|k[ae]j|dari|gadvan|gadv|fahl|garau|surface',re.I)

def colno(s):
    n=0
    for c in s:n=n*26+ord(c)-64
    return n

def parse_sheet(c,sid):
    data={}
    for addr,val in c.execute('SELECT address,raw_value FROM cells WHERE sheet_id=? AND raw_value IS NOT NULL',(sid,)):
        m=cell_re.match(addr or '')
        if m:data[(int(m[2]),m[1])]=val
    anchors=[(r,col,v) for (r,col),v in data.items() if v.strip().lower() in ('formation top data','formation geology data')]
    result=[]
    for r,acol,title in sorted(anchors):
        headers={col:val for (row,col),val in data.items() if row==r+1 and colno(col)>=colno('K') and colno(col)<=colno('S')}
        namecol=next((col for col,val in headers.items() if val.strip().lower() in ('name','formation')),None)
        if not namecol: continue
        # These headers belong to the table, not the parallel POB/weather block.
        use={col:val for col,val in headers.items() if colno(col)>=colno(namecol)}
        rows=[];missing=0
        for n in range(r+2,r+22):
            if any(rr==n for rr,_,_ in anchors):break
            name=data.get((n,namecol),'')
            if not name or not formation.search(name):
                missing+=1
                if missing>=2:break
                continue
            missing=0
            vals=[{'header':use.get(col, 'Unlabelled column '+col),'cell':f'{col}{n}','value':data[(n,col)]} for col in [chr(v) for v in range(ord(namecol),ord('S')+1)] if (n,col) in data]
            rows.append({'formation':name,'cells':vals})
        if rows:result.append({'anchor':f'{acol}{r}','title':title,'header_row':r+1,'rows':rows})
    return result

def extract(db):
    output=[]
    for fid,name in db.execute("SELECT id,name FROM files WHERE name LIKE 'AZNS%Pack.xlsx' ORDER BY name").fetchall():
        candidates=db.execute("""SELECT DISTINCT s.id,s.name FROM sheets s JOIN cells c ON c.sheet_id=s.id
          WHERE s.file_id=? AND lower(c.raw_value) IN ('formation top data','formation geology data') ORDER BY s.id DESC""",(fid,)).fetchall()
        chosen=None;fallback=None
        for sid,sheet in candidates:
            tables=parse_sheet(db,sid)
            if not tables:continue
            if fallback is None:fallback=(sheet,tables)
            def is_number(x):
                try:float(x);return True
                except ValueError:return False
            if any(is_number(cell['value']) for t in tables for row in t['rows'] for cell in row['cells'] if 'md' in cell['header'].lower() or 'tvd' in cell['header'].lower()):
                chosen=(sheet,tables);break
        chosen=chosen or fallback
        output.append({'file':name,'sheet':chosen[0] if chosen else None,'tables':chosen[1] if chosen else [],
            'selection':'Last in workbook order with a recognized formation table containing a numeric value; otherwise last with recognized names. This may not be the latest dated report.'})
    return output

def main():
    p=argparse.ArgumentParser();p.add_argument('database',type=Path);p.add_argument('output',type=Path);args=p.parse_args()
    c=sqlite3.connect(f'file:{args.database.resolve()}?mode=ro',uri=True)
    data=extract(c);args.output.parent.mkdir(parents=True,exist_ok=True)
    args.output.write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    for w in data:print(w['file'],w['sheet'],sum(len(t['rows']) for t in w['tables']))
if __name__=='__main__':main()
