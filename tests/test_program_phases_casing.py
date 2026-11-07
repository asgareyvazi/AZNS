import json,unittest
from pathlib import Path
from pypdf import PdfReader
from docx import Document
ROOT=Path(__file__).resolve().parents[1]
class ProgrammePhaseAndCasingTests(unittest.TestCase):
 @classmethod
 def setUpClass(cls):
  cls.phases=json.loads((ROOT/'evidence/program_phases_services_reviewed.json').read_text())
  cls.casings=json.loads((ROOT/'evidence/program_casing_intervals_reviewed.json').read_text())
 def test_all_source_phase_and_service_anchors(self):
  self.assertEqual(len(self.phases),14)
  for w in self.phases:
   pdf=PdfReader(ROOT/w['file'])
   for phase in w['selected_plan_phases']:
    text=' '.join(pdf.pages[phase['physical_page']-1].extract_text().split());self.assertIn(phase['source_anchor'],text)
   text=' '.join(pdf.pages[w['services_physical_page']-1].extract_text().split());self.assertIn(w['services_raw_excerpt'],text)
   self.assertEqual(len(w['service_assignment_as_printed']),4)
 def test_all_casing_source_anchors_and_document_values(self):
  self.assertEqual(len(self.casings),14);self.assertEqual(sum(len(x['intervals']) for x in self.casings),71)
  for w in self.casings:
   source=' '.join(PdfReader(ROOT/w['file']).pages[w['physical_page']-1].extract_text().split())
   for row in w['intervals']:
    self.assertIn(row['source_anchor'],source);self.assertLess(row['from_m'],row['to_m'])
   if w['well']=='127ST1':continue
   for lang in ('EN','FA'):
    doc=Document(ROOT/f"{w['well']}_Drilling_Engineering_Report_{lang}.docx")
    table=next(t for t in doc.tables if t.rows and t.rows[0].cells[0].text in ('Programme element','جزء در برنامه'))
    self.assertEqual(len(table.rows)-1,len(w['intervals']))
    for actual,expected in zip(table.rows[1:],w['intervals']):
     self.assertEqual(float(actual.cells[1].text),expected['from_m']);self.assertEqual(float(actual.cells[2].text),expected['to_m'])
    self.assertTrue(any(w['file'] in p.text for p in doc.paragraphs if p.text.startswith(('Table 6-1','جدول ۶-۱'))))
 def test_conflicts_and_stages(self):
  c={w['well']:w for w in self.casings};self.assertEqual(c['127WO1']['intervals'][0]['to_m'],146);self.assertEqual(c['12']['intervals'][3]['from_m'],3285);self.assertIn('Planned',c['207']['scope']);self.assertIn('before sidetrack 2',c['127ST2']['scope'])
  d=Document(ROOT/'12_Drilling_Engineering_Report_EN.docx');self.assertTrue(any('wording about checking a 3,889–3,911 m' in p.text for p in d.paragraphs))
