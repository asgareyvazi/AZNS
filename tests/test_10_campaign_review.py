import json,unittest
from pathlib import Path
from openpyxl import load_workbook
from pypdf import PdfReader
from docx import Document
ROOT=Path(__file__).resolve().parents[1]
class Campaign10(unittest.TestCase):
 def test_cells_headers_and_reports(self):
  d=json.loads((ROOT/'evidence/10_campaign_review.json').read_text());events=d['events'];self.assertEqual(len(events),27);self.assertEqual(sum(len(e['cells']) for e in events),219)
  self.assertEqual([(e['header_report_number'],e['worksheet_prefix']) for e in events],sorted((e['header_report_number'],e['worksheet_prefix']) for e in events))
  w=load_workbook(ROOT/d['source'],read_only=True,data_only=True)
  for e in events:
   self.assertEqual(w[e['sheet']]['S5'].value,e['header_report_number']);self.assertIn(e['anchor'].lower(),' '.join(e['cells'].values()).lower())
   for c,v in e['cells'].items():self.assertEqual(w[e['sheet']][c].value,v,(e['sheet'],c))
  self.assertEqual(w['27-1 Remark']['S5'].value,27);self.assertEqual(w['28-1 Remark']['S5'].value,27);w.close()
  for lang,key in [('EN','english'),('FA','persian')]:
   text='\n'.join(p.text for p in Document(ROOT/f'10_Drilling_Engineering_Report_{lang}.docx').paragraphs)
   for e in events:self.assertIn(e[key],text);self.assertIn(e['sheet']+'!'+next(iter(e['cells'])),text)
 def test_programme_financial_sources_and_field(self):
  d=json.loads((ROOT/'evidence/10_campaign_review.json').read_text());pg=' '.join(PdfReader(ROOT/d['programme']['file']).pages[6].extract_text().split());self.assertEqual(pg,d['programme']['text_excerpt']);self.assertIn('39 (Main Plan) + 9 (Contingency Plan)',pg)
  c=load_workbook(ROOT/d['cbs']['file'],read_only=True,data_only=True);f=load_workbook(ROOT/d['cbs']['file'],read_only=True,data_only=False)
  for name,cells in d['cbs']['cells'].items():
   for addr,v in cells.items():self.assertEqual(c[name][addr].value,v['cached']);self.assertEqual(f[name][addr].value,v['formula_or_literal'])
  c.close();f.close()
  for kind in ('Master_Drilling_Engineering','Executive_Management'):
   for lang in ('EN','FA'):
    text='\n'.join(p.text for p in Document(ROOT/f'Field_{kind}_Report_{lang}.docx').paragraphs);self.assertIn('evidence/10_campaign_review.json',text)
if __name__=='__main__':unittest.main()
