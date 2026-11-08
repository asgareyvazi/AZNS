import json,unittest
from pathlib import Path
from openpyxl import load_workbook
from pypdf import PdfReader
from docx import Document
ROOT=Path(__file__).resolve().parents[1]
class Campaign85(unittest.TestCase):
 def test_source_cells_report_headers_and_bilingual_report(self):
  data=json.loads((ROOT/'evidence/85_campaign_review.json').read_text());events=data['events'];self.assertEqual(len(events),20);self.assertEqual(sum(len(e['cells']) for e in events),154)
  self.assertEqual([e['header_report_number'] for e in events],sorted(e['header_report_number'] for e in events))
  w=load_workbook(ROOT/data['source'],read_only=True,data_only=True)
  for e in events:
   self.assertEqual(w[e['sheet']]['S5'].value,e['header_report_number']);self.assertIn(e['anchor'].lower(),' '.join(e['cells'].values()).lower())
   for c,v in e['cells'].items():self.assertEqual(w[e['sheet']][c].value,v,(e['sheet'],c))
  w.close()
  for lang,key in [('EN','english'),('FA','persian')]:
   text='\n'.join(p.text for p in Document(ROOT/f'85_Drilling_Engineering_Report_{lang}.docx').paragraphs)
   for e in events:self.assertIn(e[key],text);self.assertIn(e['sheet']+'!'+next(iter(e['cells'])),text)
   self.assertIn('=TODAY()-1',text)
 def test_programme_and_rigless_rows(self):
  data=json.loads((ROOT/'evidence/85_campaign_review.json').read_text());pg=' '.join(PdfReader(ROOT/data['programme']['file']).pages[7].extract_text().split());self.assertEqual(pg,data['programme']['text_excerpt']);self.assertIn('42 (Main Plan) + 5.5 (Contingency Plan)',pg)
  c=data['cbs'];w=load_workbook(ROOT/c['file'],read_only=True,data_only=True);s=w[c['sheet']]
  for addr,value in c['cached_cells'].items():self.assertEqual(s[addr].value,value)
  self.assertTrue(all(s[f'I{i}'].value.strip().lower()=='rig less' for i in range(41,50)));w.close()
 def test_field_summaries_and_volatile_formulas(self):
  for kind in ('Master_Drilling_Engineering','Executive_Management'):
   for lang in ('EN','FA'):
    txt='\n'.join(p.text for p in Document(ROOT/f'Field_{kind}_Report_{lang}.docx').paragraphs);self.assertIn('evidence/85_campaign_review.json',txt)
  w=load_workbook(ROOT/'AZNS85Pack.xlsx',read_only=True,data_only=False)
  self.assertEqual(sum(w[f'{i}-1 Remark']['N5'].value=='=TODAY()-1' for i in range(1,54)),13);w.close()
if __name__=='__main__':unittest.main()
