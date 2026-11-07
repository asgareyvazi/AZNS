import unittest,json,re,hashlib
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from PIL import Image
from pypdf import PdfReader

class MaterialAndGeologyTests(unittest.TestCase):
 @classmethod
 def setUpClass(cls):
  cls.raw={w['file']:w for w in json.loads(Path('evidence/mud_material_observations.json').read_text())}
  cls.chosen={w['well']:w for w in json.loads(Path('evidence/mud_material_selected_reviewed.json').read_text())}
  cls.geo=json.loads(Path('evidence/program_207_stratigraphy_reviewed.json').read_text())
 def test_shifted_headers_and_scope(self):
  s={s['day']:s for s in self.raw['AZNS207Pack.xlsx']['sheets']};self.assertEqual(s[2]['headers']['used']['cell']['cell'],'P7');self.assertEqual(s[72]['headers']['used']['cell']['cell'],'Q7')
  s={s['day']:s for s in self.raw['AZNS116Pack.xlsx']['sheets']};self.assertEqual(s[13]['headers']['used']['cell']['cell'],'L39');self.assertEqual(s[13]['headers']['unit']['cell']['cell'],'E39')
  for w in self.raw.values():self.assertEqual(w['recognized_material_blocks']+len(w['unrecognized_sheets']),w['recognized_data_sheets'])
 def test_selected_materials_keep_original_cells(self):
  self.assertEqual(len(self.chosen),13)
  for w in self.chosen.values():
   self.assertEqual(len(w['samples']),2)
   for s in w['samples']:
    self.assertGreaterEqual(len(s['rows']),2)
    for row in s['rows']:
     self.assertEqual(row['material']['sheet'],s['sheet']);self.assertEqual(row['fields']['used']['sheet'],s['sheet']);self.assertEqual(row['unit']['sheet'],s['sheet']);self.assertGreater(float(row['fields']['used']['raw']),0)
 def test_207_programme_rows_and_delta(self):
  rows=self.geo['rows'];self.assertEqual(len(rows),16)
  pdf=' '.join(PdfReader(self.geo['programme_file']).pages[16].extract_text().split())
  for row in rows:self.assertIn(row['programme_source_text'],pdf)
  top=next(r for r in rows if r['name']=='Fahliyan');self.assertEqual(top['numeric_difference_ddr_minus_programme_m'],53);self.assertEqual(top['ddr_cell'],'P86')
 def test_document_summaries_and_image(self):
  file=Path('evidence/well_figures/207_W3.png')
  with Image.open(file) as im:self.assertTrue(all(abs(v-300)<1 for v in im.info['dpi']))
  for well in self.chosen:
   for lang in ('EN','FA'):
    d=Document(f'{well}_Drilling_Engineering_Report_{lang}.docx');heads=[t.rows[0].cells[0].text for t in d.tables]
    self.assertIn('Selected sheet' if lang=='EN' else 'شیت منتخب',heads)
    s=' '.join(p.text for p in d.paragraphs)+' '.join(c.text for t in d.tables for r in t.rows for c in r.cells)
    for forbidden in ('Issued for information','Personnel on board averaged','Every quantitative statement in this report resolves','صادرشده برای اطلاع','هر گزاره کمّی این گزارش','۱۵٬۹۱۸ بشکه'):
     self.assertNotIn(forbidden,s)
    if well=='207':
     i=next(i for i,p in enumerate(d.paragraphs) if p.text.startswith(('Figure W3','شکل W3')));rid=d.paragraphs[i-1]._p.xpath('.//a:blip')[0].get(qn('r:embed'))
     self.assertEqual(hashlib.sha256(file.read_bytes()).digest(),hashlib.sha256(d.part.related_parts[rid].blob).digest())
 def test_85_rigless_table_follows_cover(self):
  for lang in ('EN','FA'):
   d=Document(f'85_Drilling_Engineering_Report_{lang}.docx');nodes=list(d.element.body)
   cover=next(i for i,n in enumerate(nodes) if n.tag.endswith('}p') and ''.join(n.xpath('.//w:t/text()'))=='AZNS-85')
   table=next(t for t in d.tables if t.rows[0].cells[0].text in ('Source cell','سلول منبع'));self.assertGreater(nodes.index(table._tbl),cover)
if __name__=='__main__':unittest.main()
