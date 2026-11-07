import sys,unittest,tempfile
from pathlib import Path
from openpyxl import Workbook,load_workbook
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'tools'))
from extract_cbs_evidence import extract
import replace_cbs_financials as finance

class CbsEvidenceTests(unittest.TestCase):
    def test_missing_formula_cache_and_error_are_not_zero(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'CBS AZNS-test.xlsx';w=Workbook();w.active.title='Cost Breakdown-test'
            w.active['A1']='scope';w.active['B1']='=1+2';w.active['C1']='#REF!';w.save(path)
            obj=extract(path);cells={x['cell']:x for x in obj['sheets'][0]['cells']}
            self.assertIsNone(cells['B1']['value']);self.assertTrue(cells['B1']['missing_cache'])
            self.assertEqual(cells['B1']['formula'],'=1+2');self.assertEqual(cells['C1']['error'],'#REF!')
    def test_grouped_and_subtotal_families_are_distinct(self):
        self.assertTrue(finance.leaf_groups('12'));self.assertFalse(finance.leaf_groups('127ST2'))
        rows=finance.categories('12');fluids=next(x for x in rows if x['number']==2)
        sheet=load_workbook('CBS AZNS-12.xlsx',read_only=True,data_only=True)['Cost Breakdown-12']
        expected=sum(sheet[f'F{i}'].value for i in range(6,13) if isinstance(sheet[f'F{i}'].value,(int,float)))
        self.assertAlmostEqual(fluids['one'],expected);self.assertEqual(fluids['cells'],'F6:F12');self.assertGreater(fluids['one'],0)
        st2=finance.categories('127ST2');self.assertEqual(next(x for x in st2 if x['number']==2)['cells'],'F6')
    def test_formula_risks_and_scope(self):
        sheets=finance.books['CBS AZNS-127 ST 2.xlsx'];s=sheets['Cost Breakdown-127']
        self.assertEqual(s['F6']['formula'],'=SUM(F7:F12)');self.assertEqual(s['F197']['formula'],'=SUM(F5:F196)')
        self.assertEqual(finance.books['CBS AZNS-127 WO1.xlsx']['Cost Breakdown-127']['F30']['error'],'#REF!')
        self.assertEqual(finance.books['CBS AZNS-122WO2.xlsx']['Cost Breakdown-Material-Conting']['H249']['error'],'#REF!')
        d=finance.books['CBS AZNS-40.xlsx']['Cost Breakdown-40']
        self.assertNotEqual(d['D5']['value']*d['E5']['value'],d['F5']['value'])
        self.assertEqual(finance.mapping['85'][1],None)

if __name__=='__main__':unittest.main()
