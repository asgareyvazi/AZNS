import json,unittest
from hashlib import sha256
from pathlib import Path
from pypdf import PdfReader
from docx import Document
ROOT=Path(__file__).resolve().parents[1]
class FormationSchematicReviewTest(unittest.TestCase):
 @classmethod
 def setUpClass(cls):cls.data=json.loads((ROOT/'evidence/program_schematic_tops_reviewed.json').read_text())
 def test_selected_source_pages_and_bilingual_tables(self):
  records=self.data['records'];self.assertEqual(len(records),7);self.assertEqual(sum(len(r['tops']) for r in records),81)
  for r in records:
   page=PdfReader(ROOT/r['file']).pages[r['physical_page']-1]
   if r['method']=='embedded-image':
    image=max(page.images,key=lambda i:len(i.data));self.assertEqual(sha256(image.data).hexdigest(),r['embedded_image_sha256'])
   else:
    t=page.extract_text()
    for top in r['tops']:self.assertIn(top['source_depth_as_printed'],t,(r['well'],top))
   for lang in ('EN','FA'):
    doc=Document(ROOT/f"{r['well']}_Drilling_Engineering_Report_{lang}.docx"); tables=[t for t in doc.tables if t.cell(0,0).text in ('Formation / member','سازند یا عضو')]
    self.assertTrue(any(all(any(row.cells[0].text==top['formation' if lang=='EN' else 'formation_fa'] and row.cells[1].text==top['source_depth_as_printed'] for row in t.rows) for top in r['tops']) for t in tables),r['well'])
 def test_same_st1_st2_image_and_column_distinction(self):
  p=PdfReader(ROOT/'AZNS -127 Side Track#01 Program Rev.03.pdf').pages[61];h=sha256(max(p.images,key=lambda i:len(i.data)).data).hexdigest();self.assertEqual(h,self.data['st1_same_image_sha256']);self.assertEqual(h,next(r for r in self.data['records'] if r['well']=='127ST2')['embedded_image_sha256'])
  r=next(r for r in self.data['records'] if r['well']=='127ST2');self.assertEqual(r['tops'][0]['source_depth_as_printed'],'1084')
 def test_slashes_not_normalized(self):
  r=next(r for r in self.data['records'] if r['well']=='85');self.assertIn('1279/9',[t['source_depth_as_printed'] for t in r['tops']])
if __name__=='__main__':unittest.main()

class FigureReviewTest(unittest.TestCase):
 def test_three_w3_pngs_match_both_embedded_reports(self):
  from io import BytesIO
  from PIL import Image
  from docx.oxml.ns import qn
  for well in ('40','127WO1','127ST2'):
   data=(ROOT/f'evidence/well_figures/{well}_W3.png').read_bytes()
   image=Image.open(BytesIO(data));self.assertGreaterEqual(image.width,4000);self.assertAlmostEqual(image.info['dpi'][0],300,delta=0.1)
   for lang in ('EN','FA'):
    doc=Document(ROOT/f'{well}_Drilling_Engineering_Report_{lang}.docx');p=doc.paragraphs;ix=next(i for i,x in enumerate(p) if x.text.startswith(('Figure W3','شکل W3')));rid=p[ix-1]._p.xpath('.//a:blip')[0].get(qn('r:embed'));self.assertEqual(doc.part.related_parts[rid].blob,data)
