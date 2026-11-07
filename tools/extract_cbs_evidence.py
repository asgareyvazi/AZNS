#!/usr/bin/env python3
"""Transcribe CBS primary cost worksheets with cached values and formula provenance.

This is source evidence, not an approved budget, an actual-cost ledger or a
formula recalculation. Workbooks are opened read-only and never modified.
"""
import argparse,glob,json,hashlib,re
from pathlib import Path
from openpyxl import load_workbook


def interesting(name):
    n=re.sub(r'\s+',' ',name.lower()).strip()
    return n.startswith(('cost breakdown','cost b.d','total cost','rig-less stimulation'))

def serial(value):
    if value is None or isinstance(value,(str,int,float,bool)):return value
    return value.isoformat()

def extract(file):
    path=Path(file)
    with path.open('rb') as f:sha=hashlib.file_digest(f,'sha256').hexdigest()
    cached=load_workbook(path,read_only=True,data_only=True)
    formulas=load_workbook(path,read_only=True,data_only=False)
    sheets=[]
    for title in cached.sheetnames:
        if not interesting(title):continue
        a=cached[title];b=formulas[title];cells=[]
        for row_c,row_f in zip(a.iter_rows(),b.iter_rows()):
            for c,f in zip(row_c,row_f):
                if c.value is None and f.value is None:continue
                x={'cell':f.coordinate,'value':serial(c.value)}
                if f.data_type=='f':x['formula']=f.value
                if c.data_type=='e' or f.data_type=='e':x['error']=serial(c.value if c.data_type=='e' else f.value)
                if f.data_type=='f' and c.value is None:x['missing_cache']=True
                cells.append(x)
        sheets.append({'sheet':title,'cells':cells})
    cached.close();formulas.close()
    return {'file':path.name,'sha256':sha,'sheets':sheets}

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--output',required=True);args=ap.parse_args()
    data=[extract(p) for p in sorted(glob.glob('CBS AZNS-*.xlsx'))]
    assert len(data)==13
    result={'description':'Source CBS primary cost sheets, raw cached values and formulas. Amounts are not validated actual expenditures or approved comparable budgets; categories and scopes may overlap.','workbooks':data}
    Path(args.output).write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    for x in data:
        cells=[c for s in x['sheets'] for c in s['cells']];print(x['file'],len(x['sheets']),'sheets',len(cells),'cells',sum('error' in c for c in cells),'errors',sum('missing_cache' in c for c in cells),'missing caches')
if __name__=='__main__':main()
