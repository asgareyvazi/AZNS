import json, unittest
from pathlib import Path
from openpyxl import load_workbook
from pypdf import PdfReader
from docx import Document
ROOT=Path(__file__).resolve().parents[1]
class Campaign75(unittest.TestCase):
 def test_cells_and_bilingual_reports(self):
  d=json.loads((ROOT/'evidence/75_campaign_review.json').read_text());events=d['events'];self.assertEqual(len(events),26);self.assertEqual(sum(len(x['cells']) for x in events),202)
  nums=[x['header_report_number'] for x in events];self.assertEqual(nums,sorted(nums));w=load_workbook(ROOT/d['source'],read_only=True,data_only=True)
  for e in events:
   self.assertEqual(w[e['sheet']]['AQ5'].value,e['header_report_number']);self.assertIn(e['anchor'].lower(),' '.join(e['cells'].values()).lower())
   for c,v in e['cells'].items():self.assertEqual(w[e['sheet']][c].value,v,(e['sheet'],c))
  w.close()
  for lang,key in [('EN','english'),('FA','persian')]:
   text='\n'.join(p.text for p in Document(ROOT/f'75_Drilling_Engineering_Report_{lang}.docx').paragraphs)
   for e in events:self.assertIn(e[key],text);self.assertIn(e['sheet']+'!'+next(iter(e['cells'])),text)
 def test_programme_and_field(self):
  d=json.loads((ROOT/'evidence/75_campaign_review.json').read_text());pg=' '.join(PdfReader(ROOT/d['programme']['file']).pages[7].extract_text().split());self.assertEqual(pg,d['programme']['text_excerpt']);self.assertIn('141.4 days (Main Plan) + 149.2 Days (with Contingency Plan)',pg)
  for kind in ('Master_Drilling_Engineering','Executive_Management'):
   for lang in ('EN','FA'):
    txt='\n'.join(p.text for p in Document(ROOT/f'Field_{kind}_Report_{lang}.docx').paragraphs);self.assertIn('evidence/75_campaign_review.json',txt)
if __name__=='__main__':unittest.main()
