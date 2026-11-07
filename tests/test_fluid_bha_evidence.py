import json,unittest
from pathlib import Path
from docx import Document

class FluidBhaEvidenceTest(unittest.TestCase):
 @classmethod
 def setUpClass(cls):
  cls.fluid={x['well']:x for x in json.loads(Path('evidence/fluid_observation_summary.json').read_text())}
  cls.bha={x['well']:x for x in json.loads(Path('evidence/bha_selected_reviewed.json').read_text())}
 def test_coverage_and_missing_loss(self):
  f=self.fluid['12']['fields'];self.assertEqual((f['pv']['numeric_accepted_count'],f['downhole_loss']['numeric_accepted_count'],f['downhole_loss']['maximum']['value']),(47,54,130))
  for well in ('10','116','127WO1','127WO2','166'):
   x=self.fluid[well]['fields']['downhole_loss'];self.assertEqual(x['numeric_accepted_count'],0);self.assertIsNone(x['minimum']);self.assertIsNone(x['maximum'])
  z=self.fluid['85']['fields']['downhole_loss'];self.assertEqual(z['numeric_accepted_count'],7);self.assertEqual((z['minimum']['value'],z['maximum']['value']),(0,0))
 def test_bha_selection_carries_source_cells(self):
  self.assertEqual(len(self.bha),13)
  for w in self.bha.values():
   self.assertEqual(len(w['components']),4)
   for row in w['components']:
    self.assertEqual(row['item']['sheet'],w['sheet']);self.assertTrue(row['item']['cell'])
    if row['od']:self.assertEqual(row['od']['sheet'],w['sheet'])
  c=self.bha['122WO1']['components'][1];self.assertIn('7"',c['item']['raw']);self.assertIn('5-3/8',c['od']['raw'])
 def test_reports_replace_unsupported_summaries(self):
  for w in self.fluid:
   for lang in ('EN','FA'):
    d=Document(f'{w}_Drilling_Engineering_Report_{lang}.docx'); heads=[t.rows[0].cells[0].text for t in d.tables if t.rows]
    self.assertIn('Source observation' if lang=='EN' else 'مشاهده منبع',heads)
    self.assertIn('Selected component and cell' if lang=='EN' else 'جزء منتخب و سلول',heads)
    self.assertNotIn('Mud weight envelope',heads)
    bit=next(t for t in d.tables if t.rows[0].cells[0].text in ('Bit','شماره') and len(t.rows[0].cells)==6)
    following=bit._tbl.getnext();self.assertIn('4-2' if lang=='EN' else '۴-۲',''.join(following.xpath('.//w:t/text()')))
if __name__=='__main__':unittest.main()
