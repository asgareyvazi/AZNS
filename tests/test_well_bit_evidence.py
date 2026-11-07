import json,sqlite3,unittest
from pathlib import Path
from tools.extract_well_bit_observations import bit_id,rerun_id,layout,extract

class BitEvidenceTest(unittest.TestCase):
 def test_id_requires_an_unambiguous_numeric_bit(self):
  for raw,want in [('#03',(3,None)),('4 RR2',(4,2)),('2 | RR2',(2,2)),('#01RR1',(1,1)),('1 RR#2',(1,2)),('RR#3',(None,None)),('J.Mill#02',(None,None)),('2 R#1',(None,None)),('Nozzle No',(None,None)),('',(None,None))]:
   self.assertEqual(bit_id(raw),want,raw)
 def test_rerun_only_when_explicit(self):
  for raw,want in [(None,None),('RR',None),('RR2',2),('RR#2',2),('0',0),('1',1),('1.5',None),('-1',None)]:self.assertEqual(rerun_id(raw),want,raw)
 def test_layout_requires_bit_anchors_not_nozzle(self):
  self.assertIsNone(layout({'H11':'Nozzle No','H12':'Bit Size','H13':'Bit Type'}))
  legacy=layout({'H11':'Bit No','H12':'Bit Size','H13':'Bit Type','H26':'Bit Cum. Drilled','H28':'Bit Cum. Hours'})
  self.assertEqual(legacy[1],('K','M'));self.assertEqual(legacy[2]['cumulative_m'],26)
  shifted=layout({'A9':'Bit No','A11':'Bit Size','A12':'Bit Type','A28':'Bit Cum. Drilled'})
  self.assertEqual(shifted[1],('D','F'));self.assertEqual(shifted[2]['cumulative_m'],28)
 def test_nozzle_and_ambiguous_identifier_do_not_become_bits(self):
  db=sqlite3.connect(':memory:');db.executescript('create table files(id integer,name text);create table sheets(id integer,file_id integer,name text);create table cells(sheet_id integer,address text,raw_value text);')
  db.execute('insert into files values(1,"sample.xlsx")');db.execute('insert into sheets values(1,1,"1-2 Data")');db.execute('insert into sheets values(2,1,"2-2 Remark")')
  rows={'H11':'Bit No','H12':'Bit Size','H13':'Bit Type','H26':'Bit Cum. Drilled','K11':'RR#3','K12':'6 in','K13':'mill','K26':'100','M11':'#3','M12':'8 in','M13':'PDC','M26':'20'}
  db.executemany('insert into cells values(1,?,?)',rows.items());db.executemany('insert into cells values(2,?,?)',rows.items())
  w=extract(db,'sample.xlsx','test',{'sample.xlsx':2});self.assertEqual(w['recognized_data_sheets'],1);self.assertEqual(len(w['unassigned_observations']),1);self.assertEqual(len(w['groups']),1)
  self.assertEqual(w['groups'][0]['bit_no'],3);self.assertIsNone(w['groups'][0]['rerun']);self.assertEqual(w['groups'][0]['maximum_cumulative_m']['value'],20)
 def test_known_source_samples(self):
  d={x['well']:x for x in json.loads(Path('evidence/well_bit_observations.json').read_text())}
  self.assertEqual(d['12']['recognized_data_sheets'],65);self.assertEqual(d['12']['candidate_pairs'],72)
  self.assertEqual(d['127WO1']['recognized_data_sheets'],88);self.assertEqual(d['127WO1']['candidate_pairs'],137)
  self.assertEqual(len(d['166']['groups']),0);self.assertGreater(len(d['166']['unassigned_observations']),0)
  g=next(g for g in d['207']['groups'] if g['bit_no']==3 and g['rerun']==0)
  self.assertEqual(g['maximum_cumulative_m'],{'value':580.8,'sheet':'15-2 DDR Data','cell':'E28'})
  self.assertEqual(g['maximum_cumulative_h'],{'value':62.0,'sheet':'15-2 DDR Data','cell':'E30'})
if __name__=='__main__':unittest.main()
