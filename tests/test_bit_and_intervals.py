import sqlite3,sys,unittest
from pathlib import Path
from docx import Document
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'tools'))
from extract_207_bit_observations import extract
from replace_w6_intervals import rows

class BitTests(unittest.TestCase):
 def test_bit_no_and_rerun_are_separate_and_header_is_not_a_bit(self):
  c=sqlite3.connect(':memory:')
  c.executescript('CREATE TABLE files(id INTEGER PRIMARY KEY,name TEXT); CREATE TABLE sheets(id INTEGER PRIMARY KEY,file_id INTEGER,name TEXT); CREATE TABLE cells(sheet_id INTEGER,address TEXT,raw_value TEXT);')
  c.execute("insert into files values(1,'AZNS207Pack.xlsx')")
  c.executemany('insert into sheets values(?,?,?)',[(1,1,'1-2 DDR Data'),(2,1,'2-2 DDR Data')])
  labels=[('B9','Bit No.'),('B10','Bit Rerun NO (First Run=0)'),('B28','Bit Cum. Drilled (m)'),('B30','Bit Cum. Hrs on Bottom')]
  for sid in (1,2):c.executemany('insert into cells values(?,?,?)',[(sid,a,v) for a,v in labels])
  c.executemany('insert into cells values(?,?,?)',[(1,'E9','3'),(1,'E10','0'),(1,'E28','250'),(1,'G9','Nozzle No'),(2,'E9','3'),(2,'E10','0'),(2,'E28','275'),(2,'G9','3'),(2,'G10','1'),(2,'G28','300')])
  result=extract(c);self.assertEqual(len(result['groups']),2)
  self.assertEqual([g['maximum_cumulative_m']['value'] for g in result['groups']],[275,300])
  self.assertEqual(len(result['observations']),3)

class IntervalTests(unittest.TestCase):
 def test_ambiguous_values_and_non_casing_are_not_invented(self):
  d=Document();t=d.add_table(rows=1,cols=6)
  for label,top,shoe in [('9-5/8"','Surface','3854'),('7"','3703','4922'),('9-5/8"','Surface','3909/5'),('1. Rigging & Moving','14','P')]:
   cells=t.add_row().cells;cells[0].text=label;cells[4].text=top;cells[5].text=shoe
  data,excluded=rows(t)
  self.assertEqual(data,[('9-5/8"',0.0,3854.0),('7"',3703.0,4922.0)])
  self.assertEqual(len(excluded),2)
if __name__=='__main__':unittest.main()
