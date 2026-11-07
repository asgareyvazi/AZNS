#!/usr/bin/env python3
"""Extract two-column bit observations from AZNS-207 daily data sheets.
Counters are reported as recorded; a maximum is NOT treated as run metres or hours.
"""
import argparse, json, re, sqlite3
from collections import defaultdict
from pathlib import Path

def norm(v):return re.sub(r'\s+',' ',v.strip().lower()).replace('no.','no')
def numeric(v):
    try:return float(v)
    except (ValueError,TypeError):return None

def extract(db):
    sheets=db.execute("""select s.id,s.name from sheets s join files f on f.id=s.file_id
       where f.name='AZNS207Pack.xlsx' and s.name GLOB '*-2 *' order by s.id""").fetchall()
    observations=[]
    for sid,sh in sheets:
        cells={a:v for a,v in db.execute('SELECT address,raw_value FROM cells WHERE sheet_id=? AND raw_value IS NOT NULL',(sid,))}
        labels={}
        for addr,value in cells.items():
            m=re.fullmatch(r'B(\d+)',addr)
            if m:labels[int(m.group(1))]=norm(value)
        def row(pattern):return next((r for r,t in sorted(labels.items()) if re.search(pattern,t)),None)
        fields={'bit_no':row(r'^bit no(?:\b|\.)'), 'rerun':row(r'^bit rerun'),
                'size':row(r'^bit size'), 'type':row(r'^bit type'), 'manufacturer':row(r'^bit manufacture'),
                'iadc':row(r'^iadc code'), 'serial':row(r'^bit serial'), 'tfa':row(r'^tfa'),
                'daily_m':row(r'^bit drilled'), 'cumulative_m':row(r'^bit cum\.?(?:ulative)? drilled'),
                'daily_h':row(r'^bit hours on bottom'), 'cumulative_h':row(r'^bit cum\.?(?:ulative)? hrs on bottom'),
                'dull':row(r'^bit dull condition')}
        if fields['bit_no'] is None:continue
        for col in ('E','G'):
            bit=cells.get(f"{col}{fields['bit_no']}");number=numeric(bit)
            # Header strings, blanks, and malformed bit identifiers do not identify runs.
            if number is None or number<=0 or not number.is_integer():continue
            obs={'sheet':sh,'column':col,'fields':{}}
            for key,r in fields.items():
                if r is not None and (address:=f'{col}{r}') in cells:
                    obs['fields'][key]={'cell':address,'value':cells[address]}
            observations.append(obs)
    groups=defaultdict(list)
    for obs in observations:
        f=obs['fields'];bit=int(float(f['bit_no']['value']));rr=numeric(f.get('rerun',{}).get('value'))
        key=(bit,int(rr) if rr is not None and rr>=0 and rr.is_integer() else None)
        groups[key].append(obs)
    summary=[]
    for (bit,rerun),obs in sorted(groups.items(),key=lambda kv:(kv[0][0],-1 if kv[0][1] is None else kv[0][1])):
        item={'bit_no':bit,'rerun':rerun,'observations':len(obs)}
        for field in ('size','type','manufacturer','iadc','serial'):
            item[field]=list(dict.fromkeys(o['fields'][field]['value'] for o in obs if field in o['fields']))
        for field in ('cumulative_m','cumulative_h'):
            values=[(numeric(o['fields'][field]['value']),o) for o in obs if field in o['fields']]
            values=[(n,o) for n,o in values if n is not None and n>=0]
            if values:
                n,o=max(values,key=lambda x:x[0]);item['maximum_'+field]={'value':n,'sheet':o['sheet'],'cell':o['fields'][field]['cell']}
            else:item['maximum_'+field]=None
        dull=[o for o in obs if o['fields'].get('dull',{}).get('value','').strip() not in ('','-')]
        item['last_nonblank_dull']={'value':dull[-1]['fields']['dull']['value'],'sheet':dull[-1]['sheet'],'cell':dull[-1]['fields']['dull']['cell']} if dull else None
        item['conflicts']=[key for key in ('size','type','manufacturer','iadc','serial') if len(item[key])>1]
        summary.append(item)
    return {'source':'AZNS207Pack.xlsx','sheet_count':len(sheets),'observations':observations,'groups':summary,
            'limitations':'Grouping requires numeric bit number and uses explicitly recorded rerun, or unknown; possible tool-number reuse and metadata conflicts require review. Maxima of cumulative fields are observed cumulative counters, not incremental run totals; counters can include previous reruns. Dull value is last nonblank in sheet order, not necessarily final. Zero values are not evidence of no progress.'}

def main():
    p=argparse.ArgumentParser();p.add_argument('database',type=Path);p.add_argument('output',type=Path);a=p.parse_args()
    db=sqlite3.connect(f'file:{a.database.resolve()}?mode=ro',uri=True);result=extract(db)
    a.output.write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    print('Sheets',result['sheet_count'],'observations',len(result['observations']),'groups',len(result['groups']))
    for x in result['groups']:
        print(x['bit_no'],x['rerun'],x['observations'],(x['maximum_cumulative_m'] or {}).get('value'),(x['maximum_cumulative_h'] or {}).get('value'),'conflicts',x['conflicts'])
if __name__=='__main__':main()
