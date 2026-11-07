import sqlite3, sys, unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'tools'))
from replace_w4_pareto import pareto
from extract_formation_tables import parse_sheet,extract

class ParetoTests(unittest.TestCase):
    def test_top_down_cumulative_is_monotonic(self):
        out=pareto([('A',294),('B',280),('C',241)])
        self.assertAlmostEqual(out[0][2],294/815*100)
        self.assertLess(out[0][2],out[1][2]);self.assertAlmostEqual(out[-1][2],100)

class FormationTests(unittest.TestCase):
    def test_dynamic_anchor_and_unlabelled_cells(self):
        c=sqlite3.connect(':memory:')
        c.executescript('CREATE TABLE files(id INTEGER PRIMARY KEY,name TEXT); CREATE TABLE sheets(id INTEGER PRIMARY KEY,file_id INTEGER,name TEXT); CREATE TABLE cells(sheet_id INTEGER,address TEXT,raw_value TEXT);')
        c.execute("INSERT INTO files VALUES(1,'AZNS1Pack.xlsx')")
        c.executemany('INSERT INTO sheets VALUES(?,?,?)',[(1,1,'1-2 Data'),(2,1,'2-2 Data')])
        values=[('M69','Formation Top Data'),('M70',' Name'),('O70','m MD'),('R70','Lithology'),('M71','Gachsaran'),('O71','1020'),('P71','1021'),('R71','Anhy')]
        c.executemany('INSERT INTO cells VALUES(?,?,?)',[(1,a,v) for a,v in values])
        # Later table lacks depth; select the latest sheet with a depth instead.
        c.executemany('INSERT INTO cells VALUES(?,?,?)',[(2,'N80','Formation Top Data'),(2,'N81','Name'),(2,'P81','m MD'),(2,'N82','Gadvan')])
        result=extract(c);self.assertEqual(result[0]['sheet'],'1-2 Data')
        row=result[0]['tables'][0]['rows'][0]
        self.assertEqual(row['formation'],'Gachsaran')
        unlabelled=next(x for x in row['cells'] if x['cell']=='P71')
        self.assertEqual(unlabelled['header'],'Unlabelled column P')
        self.assertEqual(unlabelled['value'],'1021')
if __name__=='__main__':unittest.main()
