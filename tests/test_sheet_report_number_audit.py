import hashlib,json,re,unittest
from pathlib import Path
from openpyxl import load_workbook
from docx import Document
from docx.oxml.ns import qn
ROOT=Path(__file__).resolve().parents[1]

class SheetNumberAudit(unittest.TestCase):
 def test_header_values_match_original_workbooks(self):
  books=json.loads((ROOT/'evidence/sheet_report_number_audit.json').read_text())['workbooks']
  self.assertEqual(sum(b['status_pages'] for b in books),1301)
  self.assertEqual(sum(b['mismatched'] for b in books),447)
  self.assertEqual(sum(b['unreadable_header'] for b in books),1)
  for b in books:
   book=load_workbook(ROOT/b['file'],read_only=True,data_only=True)
   for r in b['rows']:
    sh=book[r['sheet']];cells=next(sh.iter_rows(min_row=5,max_row=5,max_col=43));n=cells[42 if r['header_cell']=='AQ5' else 18].value;label=cells[37 if r['header_cell']=='AQ5' else 15].value
    self.assertEqual(label,r['header_label'],(b['file'],r['sheet']))
    self.assertEqual(None if n is None else int(n),r['header_report_number'],(b['file'],r['sheet']))
    self.assertEqual(None if n is None else int(n)==r['sheet_number'],r['matches_prefix'])
   book.close()
 def test_critical_116_dates_and_122_headers(self):
  w=load_workbook(ROOT/'AZNS116Pack.xlsx',read_only=True,data_only=True)
  for s,num,date in [('101',107,'2024.09.21'),('102',117,'2024.10.01'),('103',118,'2024.10.02')]:
   self.assertEqual(w[f'{s}-1 Remark']['S5'].value,num);self.assertEqual(w[f'{s}-1 Remark']['N5'].value,date)
  self.assertEqual([w[f'{i}-1 Remark']['S5'].value for i in range(104,113)],list(range(108,117)))
  self.assertIn('08-06-2024',w['1-1 Remark']['K22'].value);self.assertIn('2024.06.07',w['103-1 Remark']['N6'].value);w.close()
  w=load_workbook(ROOT/'AZNS122WO1Pack.xlsx',read_only=True,data_only=True);self.assertEqual(w['24-1 Remark']['S5'].value,23);self.assertEqual(str(w['56-1 Remark']['N5'].value),'2024-09-03 00:00:00');self.assertIn('09-Mar-2024',w['56-1 Remark']['K30'].value);w.close()
  w=load_workbook(ROOT/'AZNS122WO2Pack.xlsx',read_only=True,data_only=True);self.assertIn('2025.02.22',w['119-1 DDR Remark']['S32'].value);self.assertIn('2025.02.23',w['120-1 DDR Remark']['S29'].value);w.close()
 def test_figures_and_corrections_in_both_languages(self):
  wells={'10':'AZNS10Pack.xlsx','12':'AZNS12Pack.xlsx','75':'AZNS75Pack.xlsx','85':'AZNS85Pack.xlsx','116':'AZNS116Pack.xlsx','122WO1':'AZNS122WO1Pack.xlsx','122WO2':'AZNS122WO2Pack.xlsx','127WO1':'AZNS127WO1Pack.xlsx','127WO2':'AZNS127WOPack.xlsx','127ST2':'AZNS127ST2Pack.xlsx','166':'AZNS166Pack.xlsx','207':'AZNS207Pack.xlsx','40':'AZNS40Pack.xlsx'}
  for well in wells:
   for lang in ('EN','FA'):
    d=Document(ROOT/f'{well}_Drilling_Engineering_Report_{lang}.docx');paras=d.paragraphs;full='\n'.join(p.text for p in paras)
    self.assertIn('sheet_report_number_audit.json',full)
    self.assertNotRegex(full,r'(?m)^(?:Daily report|Report) \d+ —')
    self.assertNotRegex(full,r'(?m)^گزارش \d+ ـ')
    for n in (2,3):
     idx=next(i for i,p in enumerate(paras) if p.text.startswith((f'Figure W{n}',f'شکل W{n}')));rid=paras[idx-1]._p.xpath('.//a:blip')[0].get(qn('r:embed'));expected=(ROOT/f'evidence/well_figures/{well}_W{n}.png').read_bytes();self.assertEqual(hashlib.sha256(expected).digest(),hashlib.sha256(d.part.related_parts[rid].blob).digest(),(well,lang,n))
  for typ in ('Master_Drilling_Engineering','Executive_Management'):
   for lang in ('EN','FA'):
    d=Document(ROOT/f'Field_{typ}_Report_{lang}.docx');self.assertIn('sheet_report_number_audit.json','\n'.join(p.text for p in d.paragraphs))
    if typ.startswith('Master'):
     pars=d.paragraphs;i=next(i for i,p in enumerate(pars) if p.text.startswith(('Figure M7','شکل M7')));rid=pars[i-1]._p.xpath('.//a:blip')[0].get(qn('r:embed'));self.assertEqual(hashlib.sha256((ROOT/'evidence/field_figures/M7.png').read_bytes()).digest(),hashlib.sha256(d.part.related_parts[rid].blob).digest())
if __name__=='__main__':unittest.main()
