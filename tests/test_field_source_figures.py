import unittest,json,hashlib,re
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from PIL import Image

class FieldSourceFiguresTest(unittest.TestCase):
 @classmethod
 def setUpClass(cls):
  cls.meta=json.loads(Path('evidence/field_figure_observations.json').read_text())['figures']
 def test_programme_and_independent_coverage(self):
  plans={x['well']:x for x in self.meta['M1']};self.assertEqual(len(plans),14)
  self.assertEqual(plans['127ST1']['days'],80);self.assertEqual(plans['207']['days'],125.4)
  inventory={x['well']:x for x in self.meta['M2']};self.assertEqual(inventory['127WO2']['file'],'AZNS127WOPack.xlsx');self.assertEqual(inventory['127WO2']['sheets'],136)
  coverage={x['well']:x for x in self.meta['M4']};self.assertEqual(sum(x['status_pages'] for x in coverage.values()),1301)
  self.assertEqual((coverage['127WO1']['status_pages'],coverage['127WO1']['complete_status_pages'],coverage['127WO1']['recognized_data_pages']),(137,131,88))
  self.assertEqual((coverage['12']['status_pages'],coverage['12']['complete_status_pages'],coverage['12']['recognized_data_pages']),(72,70,65))
 def test_missing_loss_is_not_zero(self):
  points={x['well']:x['points'] for x in self.meta['M10']}
  self.assertEqual(len(points['12']),54);self.assertEqual(len(points['127WO1']),0)
  self.assertEqual(len(points['85']),7);self.assertTrue(all(z['loss']['value']==0 for z in points['85']))
  for pts in points.values():
   for z in pts:self.assertTrue(z['md']['cell'] and z['loss']['cell'])
 def test_images_embedded_at_300_dpi(self):
  for n in (1,2,4,5,7,8,9,10):
   file=Path(f'evidence/field_figures/M{n}.png')
   with Image.open(file) as im:
    self.assertTrue(all(abs(x-300)<1 for x in im.info['dpi']));self.assertGreater(im.width,2500)
   for kind in ('Master_Drilling_Engineering','Executive_Management'):
    if n==7 and kind.startswith('Executive'):continue
    for lang in ('EN','FA'):
     d=Document(f'Field_{kind}_Report_{lang}.docx');pars=d.paragraphs
     i=next(i for i,p in enumerate(pars) if re.match(rf'^(?:Figure|شکل) M{n}\b',p.text))
     rid=pars[i-1]._p.xpath('.//a:blip')[0].get(qn('r:embed'));self.assertEqual(hashlib.sha256(d.part.related_parts[rid].blob).digest(),hashlib.sha256(file.read_bytes()).digest())
 def test_removed_specific_unsupported_field_claims(self):
  for kind in ('Master_Drilling_Engineering','Executive_Management'):
   for lang in ('EN','FA'):
    d=Document(f'Field_{kind}_Report_{lang}.docx');s=' '.join(p.text for p in d.paragraphs)+' '.join(c.text for t in d.tables for r in t.rows for c in r.cells)
    for false in ('14,781,123','72,675,486','۲۵٫۸%','۱۴٬۷۸۱٬۱۲۳','۵۶٬۲۹۴','28.9%','69.1%'):
     self.assertNotIn(false,s)
if __name__=='__main__':unittest.main()
