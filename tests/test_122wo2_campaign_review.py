import json,unittest
from pathlib import Path
from openpyxl import load_workbook
from pypdf import PdfReader
from docx import Document
ROOT=Path(__file__).resolve().parents[1]
class Campaign122WO2(unittest.TestCase):
 def test_cells_and_reports(self):
  d=json.loads((ROOT/'evidence/122wo2_campaign_review.json').read_text());events=d['events'];self.assertEqual(len(events),31);self.assertEqual(sum(len(e['cells']) for e in events),213)
  self.assertIsNone(events[-1]['header_report_number']);self.assertEqual(events[-1]['worksheet_prefix'],121)
  w=load_workbook(ROOT/d['source'],read_only=True,data_only=True)
  for e in events:
   self.assertEqual(w[e['sheet']]['AQ5'].value,e['header_report_number']);self.assertIn(e['anchor'].lower(),' '.join(e['cells'].values()).lower())
   for c,v in e['cells'].items():self.assertEqual(w[e['sheet']][c].value,v,(e['sheet'],c))
  self.assertIn('2025.02.22',w['119-1 DDR Remark']['S32'].value);self.assertIn('2025.02.23',w['120-1 DDR Remark']['S29'].value);self.assertIn('2025.02.23',w['121-1 DDR Remark']['S30'].value);w.close()
  for lang,key in [('EN','english'),('FA','persian')]:
   text='\n'.join(p.text for p in Document(ROOT/f'122WO2_Drilling_Engineering_Report_{lang}.docx').paragraphs)
   for e in events:self.assertIn(e[key],text);self.assertIn(e['sheet']+'!'+next(iter(e['cells'])),text)
 def test_programme_cbs_and_field(self):
  d=json.loads((ROOT/'evidence/122wo2_campaign_review.json').read_text());pg=' '.join(PdfReader(ROOT/d['programme']['file']).pages[6].extract_text().split());self.assertEqual(pg,d['programme']['text_excerpt']);self.assertIn('45 (Main Plan) + 7 (Contingency Plan)',pg)
  c=load_workbook(ROOT/d['cbs']['file'],read_only=True,data_only=True);f=load_workbook(ROOT/d['cbs']['file'],read_only=True,data_only=False);sheet=d['cbs']['sheet']
  for addr,v in d['cbs']['cells'].items():self.assertEqual(c[sheet][addr].value,v['cached']);self.assertEqual(f[sheet][addr].value,v['formula_or_literal'])
  c.close();f.close()
  for kind in ('Master_Drilling_Engineering','Executive_Management'):
   for lang in ('EN','FA'):
    txt='\n'.join(p.text for p in Document(ROOT/f'Field_{kind}_Report_{lang}.docx').paragraphs);self.assertIn('evidence/122wo2_campaign_review.json',txt)
if __name__=='__main__':unittest.main()
