import json,unittest
from pathlib import Path
from collections import Counter
from openpyxl import load_workbook
from docx import Document

ROOT=Path(__file__).resolve().parents[1]

class EventReviewsAndCbsServices(unittest.TestCase):
 def test_all_selected_events_against_source_and_documents(self):
  data=json.loads((ROOT/'evidence/well_event_reviews_selected.json').read_text());self.assertEqual(len(data),13)
  for w in data:
   self.assertEqual(len(w['events']),3);book=load_workbook(ROOT/w['file'],read_only=True,data_only=True)
   for e in w['events']:
    self.assertEqual(str(book[e['sheet']][e['cell']].value),e['raw'],(w['well'],e['cell']))
    self.assertIn(e['anchor'].lower(),e['raw'].lower())
    for lang,key in [('EN','english'),('FA','persian')]:
     doc=Document(ROOT/f"{w['well']}_Drilling_Engineering_Report_{lang}.docx");paragraphs=[p.text for p in doc.paragraphs]
     self.assertTrue(any(e[key] in p and f"{e['sheet']}!{e['cell']}" in p for p in paragraphs))
     self.assertLess(next(i for i,p in enumerate(paragraphs) if e[key] in p),next(i for i,p in enumerate(paragraphs) if p.startswith(('Figure W2','شکل W2'))))
   book.close()
 def test_cbs_selected_cells_and_formula_cache(self):
  data=json.loads((ROOT/'evidence/cbs_service_spot_checks.json').read_text());self.assertEqual(len(data),10)
  for x in data:
   for cached in (True,False):
    book=load_workbook(ROOT/x['file'],read_only=True,data_only=cached);ws=book[x['sheet']]
    for addr,values in x['cells'].items():self.assertEqual(ws[addr].value,values['cached' if cached else 'formula_or_value'],(x['file'],x['sheet'],addr))
    book.close()
   for lang in ('EN','FA'):
    doc=Document(ROOT/f"{x['well']}_Drilling_Engineering_Report_{lang}.docx");text='\n'.join(p.text for p in doc.paragraphs);self.assertIn(x['sheet'].strip(),text)
  forty=next(x for x in data if x['well']=='40');self.assertAlmostEqual(sum(forty['cells'][f'C{r}']['cached'] for r in range(3,16)),forty['cells']['C16']['cached'])
 def test_inventory_matches_source_sheets_and_errors(self):
  data=json.loads((ROOT/'evidence/cbs_workbook_sheet_inventory.json').read_text())['workbooks'];self.assertEqual(len(data),13);self.assertEqual(sum(x['sheet_count'] for x in data),196)
  for x in data:
   wb=load_workbook(ROOT/x['file'],read_only=True,data_only=True);self.assertEqual(wb.sheetnames,[s['sheet'] for s in x['sheets']])
   for s in x['sheets']:
    self.assertEqual(wb[s['sheet']]['A1'].value,s['A1'])
    if s['cached_error_cells']:
     count=sum(c.data_type=='e' for row in wb[s['sheet']].iter_rows() for c in row);self.assertEqual(count,s['cached_error_cells'],(x['file'],s['sheet']))
   wb.close()
if __name__=='__main__':unittest.main()
