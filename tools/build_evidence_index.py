#!/usr/bin/env python3
"""Build a local, raw, traceable search index; this is not semantic interpretation.

Dependencies: lxml and pypdf. Example:
  python tools/build_evidence_index.py --source . --output /tmp/azns-evidence.sqlite
Cells retain raw OOXML values and formulas: numbers, including date serials, are
NOT interpreted as physical units, dates, cost, or validated observations.
The database is generated data and should not be committed to the repository.
"""
import argparse, hashlib, json, posixpath, sqlite3, sys, time, zipfile
from pathlib import Path
from lxml import etree
from pypdf import PdfReader

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'
T=lambda n: '{%s}%s'%(NS,n)

def sha256(p):
    h=hashlib.sha256()
    with p.open('rb') as f:
        for chunk in iter(lambda:f.read(1048576),b''): h.update(chunk)
    return h.hexdigest()

def resolve(target):
    if target.startswith('/'): return target.lstrip('/')
    return posixpath.normpath(posixpath.join('xl',target))

def strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist(): return []
    out=[]
    with z.open('xl/sharedStrings.xml') as f:
        for _,el in etree.iterparse(f,events=('end',),tag=T('si')):
            out.append(''.join(el.itertext(tag=T('t'))))
            el.clear()
            while el.getprevious() is not None: del el.getparent()[0]
    return out

def excel(db, file_id, path):
    with zipfile.ZipFile(path) as z:
        wb=etree.fromstring(z.read('xl/workbook.xml'))
        rels=etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        mapping={r.get('Id'):resolve(r.get('Target')) for r in rels if r.get('TargetMode')!='External'}
        ss=strings(z)
        for sheet in wb.findall('.//'+T('sheet')):
            name=sheet.get('name'); dest=mapping.get(sheet.get('{%s}id'%REL)); count=0
            cur=db.execute('INSERT INTO sheets(file_id,name,xml_path,state) VALUES(?,?,?,?)', (file_id,name,dest,sheet.get('state','visible')))
            sid=cur.lastrowid
            if not dest or dest not in z.namelist():
                db.execute('INSERT INTO errors(file_id,location,message) VALUES(?,?,?)',(file_id,name,'Sheet relationship missing or unsupported'))
                continue
            try:
                batch=[]
                with z.open(dest) as stream:
                    for _,el in etree.iterparse(stream,events=('end',),tag=T('c')):
                        typ=el.get('t','n'); addr=el.get('r'); v=el.find(T('v')); f=el.find(T('f'))
                        val=v.text if v is not None else None
                        if typ=='s' and val is not None:
                            try: val=ss[int(val)]
                            except (IndexError,ValueError): raise ValueError('Invalid shared-string index at '+str(addr))
                        elif typ=='inlineStr': val=''.join(el.itertext(tag=T('t')))
                        form=etree.tostring(f,encoding='unicode') if f is not None else None
                        if val is not None or form is not None:
                            batch.append((sid,addr,typ,val,form,el.get('s'))); count+=1
                        el.clear()
                        while el.getprevious() is not None: del el.getparent()[0]
                        if len(batch)>=5000:
                            db.executemany('INSERT INTO cells VALUES(?,?,?,?,?,?)',batch);batch=[]
                if batch: db.executemany('INSERT INTO cells VALUES(?,?,?,?,?,?)',batch)
                db.execute('UPDATE sheets SET cell_count=? WHERE id=?',(count,sid))
            except Exception as exc:
                db.execute('INSERT INTO errors(file_id,location,message) VALUES(?,?,?)',(file_id,name,str(exc)))
            db.commit()

def pdf(db,file_id,path):
    reader=PdfReader(str(path))
    if reader.is_encrypted and not reader.decrypt(''):
        raise ValueError('Encrypted PDF could not be opened with empty password')
    for n,page in enumerate(reader.pages,1):
        try:
            text=page.extract_text() or ''
            db.execute('INSERT INTO pages(file_id,page,text) VALUES(?,?,?)',(file_id,n,text))
        except Exception as exc:
            db.execute('INSERT INTO errors(file_id,location,message) VALUES(?,?,?)',(file_id,'page '+str(n),str(exc)))
    db.commit()

def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--source',type=Path,default=Path('.'));ap.add_argument('--output',type=Path,required=True)
    args=ap.parse_args();out=args.output.resolve();source=args.source.resolve()
    if out.exists(): ap.error('output already exists; choose a new path rather than overwriting')
    out.parent.mkdir(parents=True,exist_ok=True)
    files=sorted([*source.glob('AZNS*Pack.xlsx'),*source.glob('CBS*.xlsx'),*source.glob('*.pdf')])
    db=sqlite3.connect(out)
    db.executescript('''PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL;
    CREATE TABLE files(id INTEGER PRIMARY KEY,name TEXT UNIQUE,sha256 TEXT,size INTEGER,kind TEXT,status TEXT);
    CREATE TABLE sheets(id INTEGER PRIMARY KEY,file_id INTEGER,name TEXT,xml_path TEXT,state TEXT,cell_count INTEGER DEFAULT 0);
    CREATE TABLE cells(sheet_id INTEGER,address TEXT,type TEXT,raw_value TEXT,formula_xml TEXT,style_id TEXT);
    CREATE TABLE pages(file_id INTEGER,page INTEGER,text TEXT);
    CREATE TABLE errors(file_id INTEGER,location TEXT,message TEXT);
    CREATE TABLE meta(key TEXT PRIMARY KEY,value TEXT);
    ''')
    db.execute('INSERT INTO meta VALUES(?,?)',('interpretation','Raw OOXML cached values/formulas and extracted PDF text; not validated engineering data.'))
    for i,path in enumerate(files,1):
        typ='pdf' if path.suffix.lower()=='.pdf' else 'xlsx'
        file_id=db.execute('INSERT INTO files(name,sha256,size,kind,status) VALUES(?,?,?,?,?)',(path.name,sha256(path),path.stat().st_size,typ,'started')).lastrowid
        db.commit(); start=time.monotonic()
        try:
            (pdf if typ=='pdf' else excel)(db,file_id,path)
            errors=db.execute('SELECT count(*) FROM errors WHERE file_id=?',(file_id,)).fetchone()[0]
            status='partial' if errors else 'ok'
        except Exception as exc:
            db.execute('INSERT INTO errors(file_id,location,message) VALUES(?,?,?)',(file_id,'file',repr(exc)));status='failed'
        db.execute('UPDATE files SET status=? WHERE id=?',(status,file_id));db.commit()
        print(f'{i}/{len(files)} {path.name}: {status} ({time.monotonic()-start:.1f}s)',flush=True)
    db.executescript('CREATE INDEX cells_sheet_addr ON cells(sheet_id,address); CREATE INDEX pages_file_page ON pages(file_id,page); CREATE INDEX sheets_file_name ON sheets(file_id,name);')
    db.commit()
    summary={'files':db.execute('select count(*) from files').fetchone()[0], 'statuses':dict(db.execute('select status,count(*) from files group by status')),'sheets':db.execute('select count(*) from sheets').fetchone()[0], 'cells':db.execute('select count(*) from cells').fetchone()[0], 'pages':db.execute('select count(*) from pages').fetchone()[0], 'blank_pdf_pages':db.execute("select count(*) from pages where length(trim(text))=0").fetchone()[0], 'errors':db.execute('select count(*) from errors').fetchone()[0]}
    print(json.dumps(summary,indent=2),flush=True)
    db.close()
if __name__=='__main__': main()
