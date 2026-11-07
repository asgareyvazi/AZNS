#!/usr/bin/env python3
"""Transcribe bit/mill column observations across legacy and DDR data layouts.

Never infer an absent bit number or absent rerun, merge by tool name, or add
cumulative counters. Nozzle rows are not bit numbers.
"""
import argparse,json,math,re,sqlite3
from collections import defaultdict
from pathlib import Path

PACKS={'10':'AZNS10Pack.xlsx','12':'AZNS12Pack.xlsx','75':'AZNS75Pack.xlsx','85':'AZNS85Pack.xlsx','116':'AZNS116Pack.xlsx','122WO1':'AZNS122WO1Pack.xlsx','122WO2':'AZNS122WO2Pack.xlsx','127ST2':'AZNS127ST2Pack.xlsx','127WO1':'AZNS127WO1Pack.xlsx','127WO2':'AZNS127WOPack.xlsx','166':'AZNS166Pack.xlsx','207':'AZNS207Pack.xlsx','40':'AZNS40Pack.xlsx'}

def numeric(raw):
    try:x=float(raw)
    except (ValueError,TypeError):return None
    return x if math.isfinite(x) else None

def bit_id(raw):
    if raw is None:return None,None
    # A standalone RR#3 does not establish whether 3 is the bit or rerun.
    m=re.fullmatch(r'\s*#?0*(\d+)(?:\s*(?:\|\s*)?RR\s*#?\s*(\d+))?\s*',str(raw),re.I)
    if not m or int(m.group(1))<=0:return None,None
    return int(m.group(1)),int(m.group(2)) if m.group(2) is not None else None

def rerun_id(raw):
    if raw is None:return None
    x=numeric(raw)
    if x is not None and x>=0 and x.is_integer():return int(x)
    m=re.fullmatch(r'\s*RR\s*#?\s*(\d+)\s*',str(raw),re.I)
    return int(m.group(1)) if m else None

def norm(v):return re.sub(r'\s+',' ',v.strip().lower()).replace('no.','no')

def layout(cells):
    # Anchors distinguish bit data from report, look-ahead and nozzle tables.
    for label,cols,family in [('H',('K','M'),'legacy'),('B',('E','G'),'ddr'),('A',('D','F'),'ddr-shifted')]:
        labels={int(m.group(1)):norm(v) for a,v in cells.items() if (m:=re.fullmatch(label+r'(\d+)',a))}
        def row(pattern):return next((r for r,t in sorted(labels.items()) if re.search(pattern,t)),None)
        fields={
            'bit_no':row(r'^bit no(?:\b|\.)'), 'rerun':row(r'^bit rerun'),
            'size':row(r'^bit size'), 'type':row(r'^bit type'), 'manufacturer':row(r'^bit manufacture'),
            'iadc':row(r'^iadc code'), 'serial':row(r'^bit serial'), 'tfa':row(r'^tfa'),
            'daily_m':row(r'^bit drilled'), 'cumulative_m':row(r'^bit cum\.?(?:ulative)? drilled'),
            'daily_h':row(r'^(?:bit hours on bottom|on btm bit hours)'),
            'cumulative_h':row(r'^bit cum\.?(?:ulative)? (?:hrs|hours)'),
            'dull':row(r'^bit dull condition')}
        if fields['bit_no'] is not None and fields['size'] is not None and fields['type'] is not None:
            return family,cols,fields
    return None

def extract(db,pack,well,pair_counts):
    sheet_rows=db.execute('select s.id,s.name from sheets s join files f on f.id=s.file_id where f.name=? order by s.id',(pack,)).fetchall()
    # Only true data pages; e.g. `-2 Remark` and `-2 Lookahead` are excluded.
    sheets=[(sid,n) for sid,n in sheet_rows if re.fullmatch(r'\d+-2 (?:Data|DDR Data|DDR P 2)',n,re.I)]
    obs=[];unassigned=[];families=defaultdict(int);examined=[]
    for sid,name in sheets:
        cells=dict(db.execute('select address,raw_value from cells where sheet_id=? and raw_value is not null',(sid,)))
        result=layout(cells)
        if result is None:continue
        family,cols,fields=result;families[family]+=1;examined.append(name)
        for col in cols:
            f={key:{'cell':f'{col}{r}','value':cells[f'{col}{r}']} for key,r in fields.items() if r is not None and f'{col}{r}' in cells}
            meaningful={k:v for k,v in f.items() if str(v['value']).strip().lower() not in ('','-','--','n/a','na','none','nil')}
            if not meaningful:continue
            bit,inline_rr=bit_id(f.get('bit_no',{}).get('value'))
            explicit_rr=rerun_id(f.get('rerun',{}).get('value'))
            mismatch=inline_rr is not None and explicit_rr is not None and inline_rr!=explicit_rr
            rr=None if mismatch else (inline_rr if inline_rr is not None else explicit_rr)
            entry={'sheet':name,'column':col,'family':family,'fields':f,'bit_no':bit,'rerun':rr,'rerun_conflict':mismatch}
            if bit is None:unassigned.append(entry)
            else:obs.append(entry)
    groups=defaultdict(list)
    for o in obs:groups[(o['bit_no'],o['rerun'])].append(o)
    summary=[]
    for (bit,rr),records in sorted(groups.items(),key=lambda x:(x[0][0],-1 if x[0][1] is None else x[0][1])):
        g={'bit_no':bit,'rerun':rr,'observations':len(records)}
        for field in ('size','type','manufacturer','iadc','serial'):
            g[field]=list(dict.fromkeys(o['fields'][field]['value'] for o in records if field in o['fields'] and o['fields'][field]['value'].strip() not in ('','-','--')))
        for field in ('cumulative_m','cumulative_h'):
            vals=[(n,o) for o in records if field in o['fields'] and (n:=numeric(o['fields'][field]['value'])) is not None and n>=0]
            if vals:
                n,o=max(vals,key=lambda x:x[0]);g['maximum_'+field]={'value':n,'sheet':o['sheet'],'cell':o['fields'][field]['cell']}
            else:g['maximum_'+field]=None
        dull=[o for o in records if o['fields'].get('dull',{}).get('value','').strip().lower() not in ('','-','--','n/a','na')]
        g['last_nonblank_dull']={'value':dull[-1]['fields']['dull']['value'],'sheet':dull[-1]['sheet'],'cell':dull[-1]['fields']['dull']['cell']} if dull else None
        g['conflicts']=[key for key in ('size','type','manufacturer','iadc','serial') if len(g[key])>1]
        if any(o['rerun_conflict'] for o in records):g['conflicts'].append('rerun')
        summary.append(g)
    return {'well':well,'source':pack,'candidate_pairs':pair_counts.get(pack),'named_data_sheets':len(sheets),'recognized_data_sheets':len(examined),'layouts':dict(families),'observations':obs,'unassigned_observations':unassigned,'groups':summary,
            'limitations':'Sheet-order grouping is not a verified chronological run history. A number may be reused, explicit reruns can share a cumulative counter, and metadata can conflict. Maximum reported cumulative counter is not incremental metres or hours. Unassigned entries can include mills and bits without numeric IDs. Missing entries are not zero, and the last nonblank dull value in workbook order is not necessarily final.'}

def main():
    p=argparse.ArgumentParser();p.add_argument('database',type=Path);p.add_argument('output',type=Path);a=p.parse_args()
    tracks=json.loads(Path('evidence/drilling_track_observations.json').read_text());counts={w['file']:w['paired_sheet_candidates'] for w in tracks}
    db=sqlite3.connect(f'file:{a.database.resolve()}?mode=ro',uri=True)
    data=[extract(db,pack,well,counts) for well,pack in PACKS.items()];a.output.write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    for w in data:print(w['well'],w['layouts'],'data',w['recognized_data_sheets'],'/',w['named_data_sheets'],'pairs',w['candidate_pairs'],'assigned',len(w['observations']),'unassigned',len(w['unassigned_observations']),'groups',len(w['groups']))
if __name__=='__main__':main()
