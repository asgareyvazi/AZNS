import json,unittest
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from pypdf import PdfReader
from docx import Document
ROOT=Path(__file__).resolve().parents[1]
class CampaignReview(unittest.TestCase):
 @classmethod
 def setUpClass(cls):cls.data=json.loads((ROOT/'evidence/122wo1_campaign_review.json').read_text())
 def test_all_remarks_match_source_and_bilingual_reports(self):
  x=self.data;self.assertEqual(len(x['events']),17);self.assertEqual(sum(len(e['cells']) for e in x['events']),90);w=load_workbook(ROOT/x['source'],read_only=True,data_only=True)
  docs={lang:Document(ROOT/f'122WO1_Drilling_Engineering_Report_{lang}.docx') for lang in ('EN','FA')}
  for e in x['events']:
   for c,v in e['cells'].items():self.assertEqual(w[e['sheet']][c].value,v,(e['sheet'],c))
   for lang,key in [('EN','english'),('FA','persian')]:
    pars=docs[lang].paragraphs
    self.assertTrue(any(e[key] in p.text and e['sheet']+'!' in p.text for p in pars),(e['sheet'],lang))
  for sheet,cells in x['additional_date_cells'].items():
   for c,v in cells.items():self.assertEqual(w[sheet][c].value,v)
  for v in x['completion_depths']:self.assertIn(v['source_depth_text'],w[v['sheet']][v['cell']].value)
  w.close()
  self.assertEqual((datetime.fromisoformat(x['rig_release'])-datetime.fromisoformat(x['spud'])).total_seconds()/3600,x['calendar_elapsed_hours']);self.assertEqual(x['calendar_elapsed_hours'],1323)
  for lang in ('EN','FA'):
   d=docs[lang]
   self.assertTrue(any(all(any(row.cells[1].text==v['source_depth_text'] and row.cells[2].text==v['cell'] for row in t.rows) for v in x['completion_depths']) for t in d.tables if len(t.columns)>=3))
 def test_program_and_cbs_anchor_values(self):
  x=self.data;p=PdfReader(ROOT/x['programme']);self.assertIn('45 (Main Plan) + 7 (Contingency Plan)',' '.join(p.pages[6].extract_text().split()))
  for mode in (True,False):
   w=load_workbook(ROOT/'CBS AZNS-122WO1.xlsx',read_only=True,data_only=mode);s=w['Cost Breakdown-122']
   for cell,value in x['cbs']['cached' if mode else 'formula_or_value'].items():self.assertEqual(s[cell].value,value)
   w.close()
  self.assertEqual(x['cbs']['cached']['G5'],1035000);self.assertEqual(x['cbs']['cached']['H5'],547170)
if __name__=='__main__':unittest.main()

class CampaignFigure(unittest.TestCase):
 def test_figure_embedded_in_both_reports(self):
  from io import BytesIO
  from PIL import Image
  from docx.oxml.ns import qn
  data=(ROOT/'evidence/well_figures/122WO1_completion_depths.png').read_bytes();image=Image.open(BytesIO(data));self.assertAlmostEqual(image.info['dpi'][0],300,delta=.1)
  for lang in ('EN','FA'):
   d=Document(ROOT/f'122WO1_Drilling_Engineering_Report_{lang}.docx');p=d.paragraphs;ix=next(i for i,v in enumerate(p) if v.text.startswith(('Figure — Selected completion-component depths','شکل ـ عمق‌های منتخب اجزای تکمیل')));rid=p[ix-1]._p.xpath('.//a:blip')[0].get(qn('r:embed'));self.assertEqual(d.part.related_parts[rid].blob,data)
