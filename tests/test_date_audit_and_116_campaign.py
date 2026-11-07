import json,unittest,sys
from pathlib import Path
from datetime import datetime,date
from openpyxl import load_workbook
from docx import Document
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tools'))
from audit_status_report_dates import legacy,ddr

def plain(v):return v.isoformat() if isinstance(v,(date,datetime)) else v
class DateAuditAnd116Campaign(unittest.TestCase):
 def test_source_date_headers_and_formulas(self):
  books=json.loads((ROOT/'evidence/status_report_date_audit.json').read_text())['workbooks']
  self.assertEqual(sum(x['recognized_status_pages'] for x in books),1301);self.assertEqual(sum(x['parseable_date_headers'] for x in books),1042);self.assertEqual(sum(x['volatile_date_formulas'] for x in books),13)
  for b in books:
   wb=load_workbook(ROOT/b['file'],read_only=True,data_only=True);f=load_workbook(ROOT/b['file'],read_only=True,data_only=False)
   for r in b['rows']:
    ws=wb[r['sheet']];wf=f[r['sheet']]
    addresses=list(r['source_cells']);got={a:plain(ws[a].value) for a in addresses};self.assertEqual(got,r['source_cells'],(b['file'],r['sheet']))
    formulas={a:wf[a].value for a in addresses if wf[a].data_type=='f'};self.assertEqual(formulas,r['formula_cells'],(b['file'],r['sheet']))
    expected=None if r['volatile_date_formula'] else (ddr(*(ws[a].value for a in addresses)) if r['family']=='ddr' else legacy(ws['N5'].value))
    self.assertEqual(expected,r['report_date_iso_from_header'])
   wb.close();f.close()
  eightyfive=next(b for b in books if b['file']=='AZNS85Pack.xlsx');r=eightyfive['rows'][0];self.assertEqual(r['formula_cells']['N5'],'=TODAY()-1');self.assertIsNone(r['report_date_iso_from_header']);self.assertEqual(r['source_cells']['S9'],'05/11/1402')
 def test_unambiguous_date_parsing(self):
  self.assertEqual(legacy('2024.09.22'),'2024-09-22');self.assertIsNone(legacy('01/03/2024'));self.assertIsNone(legacy('13/7/2023'));self.assertIsNone(legacy('1403.07.01'));self.assertEqual(ddr(2024,'Oct',1),'2024-10-01');self.assertIsNone(ddr(2024,'Oct',32))
 def test_116_selected_events_are_sourced_and_in_both_documents(self):
  data=json.loads((ROOT/'evidence/116_campaign_review.json').read_text());events=data['events'];self.assertEqual(len(events),16);self.assertEqual(sum(len(e['cells']) for e in events),109);self.assertEqual([e['header_report_number'] for e in events],sorted(e['header_report_number'] for e in events));self.assertEqual(events[-1]['header_report_number'],118)
  wb=load_workbook(ROOT/data['source'],read_only=True,data_only=True)
  for e in events:
   self.assertEqual(wb[e['sheet']]['S5'].value,e['header_report_number']);self.assertIn(e['anchor'].lower(),' '.join(e['cells'].values()).lower())
   for cell,value in e['cells'].items():self.assertEqual(wb[e['sheet']][cell].value,value,(e['sheet'],cell))
  wb.close()
  for lang,key in [('EN','english'),('FA','persian')]:
   d=Document(ROOT/f'116_Drilling_Engineering_Report_{lang}.docx');p='\n'.join(x.text for x in d.paragraphs)
   for e in events:self.assertIn(e[key],p);self.assertIn(e['sheet']+'!'+next(iter(e['cells'])),p)
 def test_date_notes_in_all_reports(self):
  for p in ROOT.glob('*Report_*.docx'):
   full='\n'.join(x.text for x in Document(p).paragraphs);self.assertIn('status_report_date_audit.json',full,p.name)
  for lang in ('EN','FA'):
   p='\n'.join(x.text for x in Document(ROOT/f'85_Drilling_Engineering_Report_{lang}.docx').paragraphs);self.assertIn('=TODAY()-1',p)
if __name__=='__main__':unittest.main()
