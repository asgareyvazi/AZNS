import json,sys,unittest
from pathlib import Path
from docx import Document
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT/'tools'))
from extract_drilling_tracks import activity_day

class ActivityEvidenceTests(unittest.TestCase):
 @classmethod
 def setUpClass(cls):
  cls.activity={x['file']:x for x in json.loads((ROOT/'evidence/activity_observations.json').read_text())}
 def test_recovered_remark_only_days(self):
  a=self.activity['AZNS127WO1Pack.xlsx'];day=next(d for d in a['days'] if d['day']==100)
  self.assertEqual(day['remark_sheet'],'100-2 Remark');self.assertTrue(day['activity']['complete'])
  self.assertTrue(any(x['cell']=='K22' and '3480' in x['raw'] for x in day['detail_cells']))
  a=self.activity['AZNS12Pack.xlsx'];day=next(d for d in a['days'] if d['day']==16)
  self.assertEqual(day['remark_sheet'],'16-1 DDR Remark');self.assertTrue(day['activity']['complete'])
  self.assertEqual(sum(d['activity']['complete'] for d in a['days']),70)
  self.assertEqual(sum(d['activity']['complete'] for d in self.activity['AZNS127WO1Pack.xlsx']['days']),131)
 def test_separate_source_codes_and_unknown(self):
  cells={'J21':'Code 2 P-N-UP','J22':'N','H22':'12','F22':'0','G22':'12','J23':'UP','H23':'12','F23':'12','G23':'24'}
  day=activity_day(cells,'x','legacy',{},set());self.assertTrue(day['complete']);self.assertEqual(day['npt_h'],12);self.assertEqual(day['unplanned_h'],12)
  cells['J23']='?';self.assertFalse(activity_day(cells,'x','legacy',{},set())['complete'])
 def test_all_selected_examples_match_sources(self):
  docs=json.loads((ROOT/'evidence/daily_observation_examples.json').read_text());self.assertEqual(len(docs),13)
  for w in docs:
   self.assertEqual(len(w['selected_examples']),2)
   for ex in w['selected_examples']:
    day=next(d for d in self.activity[w['file']]['days'] if d['day']==ex['day'])
    self.assertIn({'sheet':ex['sheet'],'cell':ex['cell'],'raw':ex['raw']},day['detail_cells'])
   for lang in ('EN','FA'):
    d=Document(ROOT/f"{w['well']}_Drilling_Engineering_Report_{lang}.docx")
    for ex in w['selected_examples']:self.assertTrue(any(f"{ex['sheet']}!{ex['cell']}" in p.text for p in d.paragraphs))
 def test_complete_days_no_duplicates(self):
  for a in self.activity.values():
   days=a['days'];self.assertEqual(len({d['day'] for d in days}),len(days));self.assertFalse(a['duplicate_status_page_days'])
   for d in days:
    x=d['activity']
    if x['complete']:
     total=sum(x[k] for k in ('planned_h','npt_h','unplanned_h','mixed_h'));self.assertGreaterEqual(total,23.9);self.assertLessEqual(total,24.1);self.assertFalse(any(s['classification']=='unknown' for s in x['segments']))
