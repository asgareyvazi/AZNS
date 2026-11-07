import json,re,unittest
from pathlib import Path
from pypdf import PdfReader

class ProgramReviewTests(unittest.TestCase):
 def test_all_fourteen_source_pages_and_labels(self):
  items=json.loads(Path('evidence/program_summary_reviewed.json').read_text())
  self.assertEqual(len(items),14);self.assertEqual(len({x['file'] for x in items}),14)
  for p in items:
   raw=' '.join(PdfReader(p['file']).pages[p['physical_page']-1].extract_text().split())
   self.assertIn(p['source_objective_excerpt'],raw,p['well']);self.assertIn(p['source_duration_excerpt'],raw,p['well'])
   self.assertIn(p['programme_stated_md_m'],raw,p['well'])
  d={x['well']:x for x in items}
  self.assertIn('125.4',d['207']['duration_days_en']);self.assertIn('60.4',d['127WO1']['duration_days_en'])
  self.assertIn('149.2 with contingency',d['75']['duration_days_en'])
  self.assertNotEqual(d['127ST1']['file'],d['127ST2']['file'])
 def test_selected_formation_rows_are_directly_present(self):
  entries=json.loads(Path('evidence/program_formation_tops_reviewed.json').read_text())
  for w in entries:
   page=' '.join(PdfReader(w['file']).pages[w['physical_page']-1].extract_text().split())
   for x in w['tops']:self.assertIn(x['source_text'],page,(w['well'],x['name']))
  d={x['well']:x for x in entries}
  self.assertEqual(d['122WO2']['source_basis'],'pre_wo1')
  self.assertEqual(d['40']['source_basis'],'expected')
  self.assertEqual(next(x['md_m'] for x in d['116']['tops'] if x['name']=='Fahliyan'),'3900')
if __name__=='__main__':unittest.main()
