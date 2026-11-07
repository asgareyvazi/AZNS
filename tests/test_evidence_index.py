import sqlite3, subprocess, sys, tempfile, unittest
from pathlib import Path
from openpyxl import Workbook
from pypdf import PdfWriter

SCRIPT=Path(__file__).resolve().parents[1]/'tools'/'build_evidence_index.py'

class EvidenceIndexTest(unittest.TestCase):
    def test_cells_formulas_page_and_refuse_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); w=Workbook();s=w.active;s.title='DDR Data';s['A1']='سروک';s['B2']=42.5;s['C3']='=B2*2';s['D4']=True
            w.save(root/'AZNSTestPack.xlsx')
            pdf=PdfWriter();pdf.add_blank_page(width=100,height=100)
            with (root/'plan.pdf').open('wb') as f: pdf.write(f)
            output=root/'evidence.sqlite'
            cmd=[sys.executable,str(SCRIPT),'--source',str(root),'--output',str(output)]
            r=subprocess.run(cmd,capture_output=True,text=True);self.assertEqual(r.returncode,0,r.stderr)
            c=sqlite3.connect(output)
            self.assertEqual(c.execute("select raw_value from cells where address='A1'").fetchone()[0],'سروک')
            self.assertEqual(float(c.execute("select raw_value from cells where address='B2'").fetchone()[0]),42.5)
            self.assertIn('B2*2',c.execute("select formula_xml from cells where address='C3'").fetchone()[0])
            self.assertEqual(c.execute('select count(*) from pages').fetchone()[0],1)
            self.assertEqual(c.execute('select count(*) from errors').fetchone()[0],0)
            self.assertEqual(c.execute("select count(*) from files where status='ok'").fetchone()[0],2)
            c.close()
            self.assertNotEqual(subprocess.run(cmd,capture_output=True).returncode,0)
if __name__=='__main__':unittest.main()
