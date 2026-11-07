#!/usr/bin/env python3
"""Extract traceable daily operational observations for DDR-family wells.
No interpolation, unit conversion (except chart pcf -> ppg), or semantic event
attribution is performed here. Non-numeric and missing entries remain raw.
"""
import argparse,json,re,sqlite3
from pathlib import Path
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format

FILES=('AZNS207Pack.xlsx','AZNS127ST2Pack.xlsx','AZNS40Pack.xlsx','AZNS127WO1Pack.xlsx','AZNS10Pack.xlsx','AZNS12Pack.xlsx','AZNS75Pack.xlsx','AZNS85Pack.xlsx','AZNS116Pack.xlsx','AZNS122WO1Pack.xlsx','AZNS122WO2Pack.xlsx','AZNS127WOPack.xlsx','AZNS166Pack.xlsx')

def scalar(cells,addr,sheet):
    if addr not in cells:return None
    raw=cells[addr]
    try:value=float(raw)
    except (ValueError,TypeError):value=None
    return {'sheet':sheet,'cell':addr,'raw':raw,'value':value}

def get_row(cells,col,pattern):
    for address,text in cells.items():
        match=re.fullmatch(col+r'(\d+)',address)
        if match and re.search(pattern,re.sub(r'\s+',' ',text.strip()),re.I):return int(match[1])
    return None

def activity_day(remark,name,family,styles,time_styles):
    # Source statuses P/Plan/PLN are reported as planned, not automatically
    # productive; unknown/ambiguous values invalidate the 24-hour composition.
    statuscol,hourscol,startcol,endcol=('J','H','F','G') if family=='legacy' else ('M','E','A','C')
    header=get_row(remark,statuscol,r'^Code 2 P-N-UP') if family=='legacy' else get_row(remark,statuscol,r'^Status$')
    if header is None:return {'segments':[],'complete':False,'planned_h':None,'unplanned_h':None,'reason':'status header missing'}
    out=[]
    stop=next((r for r in range(header+1,min(header+55,90)) if remark.get(f'A{r}','').strip().lower().startswith('total')),min(header+55,90))
    for row in range(header+1,stop):
        st=remark.get(f'{statuscol}{row}','').strip();hour=remark.get(f'{hourscol}{row}');start=remark.get(f'{startcol}{row}');end=remark.get(f'{endcol}{row}')
        if hour is None and start is None and end is None:continue
        duration=scalar(remark,f'{hourscol}{row}',name) or {'sheet':name,'cell':f'{hourscol}{row}','raw':None,'value':None}
        if duration['value'] is not None and styles.get(f'{hourscol}{row}') in time_styles:
            duration['value']*=24
            duration['derived']='Excel time-formatted day fraction converted to hours'
        label=st.upper().replace(' ','')
        if label in ('P','PLAN','PLN','PLANNED'):kind='planned'
        elif label in ('N','NPT'):kind='npt'
        elif label in ('UP','UNPLAN','UNPLANNED','UNPLN'):kind='unplanned'
        elif label in ('NPT-UNPLN','NPT/UNPLN'):kind='mixed'
        else:kind='unknown'
        if start is None or end is None:kind='unknown'
        out.append({'row':row,'duration':duration,'status_cell':f'{statuscol}{row}','status':st,'classification':kind})
    valid=all(x['duration']['value'] is not None and 0<=x['duration']['value']<=24 and x['classification']!='unknown' for x in out)
    total=sum(x['duration']['value'] for x in out if x['duration']['value'] is not None)
    complete=bool(out) and valid and 23.9<=total<=24.1
    return {'segments':out,'complete':complete,'planned_h':sum(x['duration']['value'] for x in out if x['classification']=='planned') if complete else None,
            'npt_h':sum(x['duration']['value'] for x in out if x['classification']=='npt') if complete else None,
            'unplanned_h':sum(x['duration']['value'] for x in out if x['classification']=='unplanned') if complete else None,
            'mixed_h':sum(x['duration']['value'] for x in out if x['classification']=='mixed') if complete else None,
            'reason':('complete recognized 24-hour record' if complete else 'incomplete, unknown status, non-numeric duration or total not 24 hours')}

def extract(db,files=FILES):
    out=[]
    for file in files:
        sheets=db.execute('''select s.id,s.name from sheets s join files f on f.id=s.file_id where f.name=? order by s.id''',(file,)).fetchall()
        days=defaultdict(dict)
        for sid,name in sheets:
            m=re.match(r'^(\d+)-(1|2)\b',name)
            if m:days[int(m[1])][int(m[2])]=(sid,name)
        workbook=load_workbook(file,read_only=True,data_only=True)
        time_styles=set()
        for idx,style in enumerate(workbook._cell_styles):
            code=BUILTIN_FORMATS.get(style.numFmtId) if style.numFmtId<164 else (workbook._number_formats[style.numFmtId-164] if style.numFmtId-164<len(workbook._number_formats) else None)
            if code and is_date_format(code) and 'h' in code.lower() and 'y' not in code.lower() and 'd' not in code.lower():time_styles.add(str(idx))
        workbook.close()
        observations=[]
        for day,pair in sorted(days.items()):
            if 1 not in pair or 2 not in pair:continue
            rid,rname=pair[1];did,dname=pair[2]
            rcells=db.execute('select address,raw_value,style_id,type from cells where sheet_id=? and raw_value is not null',(rid,)).fetchall()
            remark={a:v for a,v,style,typ in rcells};styles={a:style for a,v,style,typ in rcells if typ in (None,'n')}
            data=dict(db.execute('select address,raw_value from cells where sheet_id=? and raw_value is not null',(did,)))
            # The two families carry different label/value columns. Never classify
            # a Remark-only sheet as a Data sheet just because it has a -2 prefix.
            if get_row(data,'A',r'^Mid\s+Nig(?:h)?t?h?\s+Depth'):
                family='legacy'
                def datum(key, col, pattern, valuecol):
                    r=get_row(data,col,pattern)
                    return scalar(data,f'{valuecol}{r}',dname) if r else None
                fields={'md_end':datum('md_end','A',r'^Mid\s+Nig(?:h)?t?h?\s+Depth','D'),
                        'daily_rop':datum('daily_rop','A',r'^Average ROP','D'),
                        'active_formation':datum('active_formation','A',r'^Formation Name','D'),
                        'mud_weight':datum('mud_weight','O',r'^Mud Weight','R'),
                        'pv':datum('pv','O',r'^PV\b','R'),'yp':datum('yp','O',r'^YP\b','R'),
                        'downhole_loss':datum('downhole_loss','O',r'^Formation Loss Vol','R')}
                for key,pattern in [('wob',r'^W\.\s*O\.\s*B'),('rpm',r'^RPM\s*/'),('flow',r'^Pump Output'),('spp',r'^Pump Pressure')]:
                    r=get_row(data,'H',pattern);x=scalar(data,f'K{r}',dname) if r else None
                    lo=hi=None
                    if x:
                        # A plain scalar is not promoted to a range; a strictly
                        # hyphenated non-negative pair can be split with provenance.
                        if x['value'] is not None:lo=x
                        else:
                            m=re.fullmatch(r'\s*(\d+(?:\.\d+)?)\s*[-–]\s*(\d+(?:\.\d+)?)\s*',x['raw'])
                            if m and float(m[1])<=float(m[2]):
                                lo={**x,'value':float(m[1]),'derived':'lower endpoint of explicit range'}
                                hi={**x,'value':float(m[2]),'derived':'upper endpoint of explicit range'}
                    fields[f'{key}_a_min']=lo;fields[f'{key}_a_max']=hi
                    fields[f'{key}_b_min']=None;fields[f'{key}_b_max']=None
            elif get_row(data,'J',r'^Mud Weight\b') or get_row(data,'I',r'^Mud Weight\b'):
                family='ddr'
                mdrow=get_row(remark,'L',r'^MD\s*\(m\)\s*@\s*24:00')
                roprow=get_row(remark,'L',r'^AVG\.\s*ROP')
                formrow=get_row(remark,'AF',r'^Formation\s*:')
                fields={'md_end':scalar(remark,f'R{mdrow}',rname) if mdrow else None,
                        'daily_rop':scalar(remark,f'R{roprow}',rname) if roprow else None,
                        'active_formation':scalar(remark,f'AI{formrow}',rname) if formrow else None}
                shifted=bool(get_row(data,'I',r'^Mud Weight\b'))
                fluid_label,fluid_value,param_col=('I','K','A') if shifted else ('J','L','B')
                for key,col,pattern,valuecol in [
                    ('mud_weight',fluid_label,r'^Mud Weight\b',fluid_value),('pv',fluid_label,r'^PV\s*\(',fluid_value),('yp',fluid_label,r'^YP\s*\(',fluid_value),
                    ('downhole_loss',fluid_label,r'^Mud Lost Down Hole',fluid_value)]:
                    r=get_row(data,col,pattern);fields[key]=scalar(data,f'{valuecol}{r}',dname) if r else None
                for key,pattern in [('wob',r'^W\.\s*O\.\s*B'),('rpm',r'^Surf\.\s*(?:RPM|Rot)'),('spp',r'^Pump Pressure')]:
                    r=get_row(data,param_col,pattern)
                    cols=('D','E','F','H') if shifted else ('E','F','G','I')
                    for stream,c in zip(('a_min','a_max','b_min','b_max'),cols):fields[f'{key}_{stream}']=scalar(data,f'{c}{r}',dname) if r else None
                r=get_row(data,param_col,r'^Pump Output\s*\(gpm\)')
                cols=('D','E','F','H') if shifted else ('E','F','G','I')
                for stream,col in zip(('a_min','a_max','b_min','b_max'),cols):
                    fields[f'flow_{stream}']=scalar(data,f'{col}{r+1}',dname) if r else None
            else:continue
            observations.append({'day':day,'remark_sheet':rname,'data_sheet':dname,'fields':fields,'activity':activity_day(remark,rname,family,styles,time_styles)})
        out.append({'file':file,'paired_sheets':len(observations),'paired_sheet_candidates':sum(1 for p in days.values() if 1 in p and 2 in p),'days':observations,
            'limitations':'End-of-day MD is not necessarily event or tool depth. Values describe report entries, may be missing/non-numeric, and are not validated or synchronous. DDR range channels are two separate bit columns. Legacy Data ranges are parsed only when explicit numeric pairs; single scalars are not asserted as minima. Some paired sheets are not Data sheets and are excluded.'})
    return out

def main():
    p=argparse.ArgumentParser();p.add_argument('database',type=Path);p.add_argument('output',type=Path);a=p.parse_args()
    db=sqlite3.connect(f'file:{a.database.resolve()}?mode=ro',uri=True);out=extract(db)
    a.output.write_text(json.dumps(out,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    for w in out:
        print(w['file'],w['paired_sheets'],'complete activities',sum(d['activity']['complete'] for d in w['days']))
        for k in ('md_end','daily_rop','active_formation','mud_weight','wob_a_min','wob_a_max','rpm_a_min','flow_a_min','spp_a_min','pv','yp','downhole_loss'):
            print(' ',k,sum(1 for d in w['days'] if (x:=d['fields'].get(k)) and x['value'] is not None))
if __name__=='__main__':main()
