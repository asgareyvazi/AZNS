import json, unittest
from pathlib import Path
from openpyxl import load_workbook
from docx import Document
ROOT=Path(__file__).resolve().parents[1]
class Campaign12(unittest.TestCase):
 def test_selected_cells_and_order(self):
  data=json.loads((ROOT/'evidence/12_campaign_review.json').read_text());events=data['events'];assert len(events)==24;assert sum(len(e['cells']) for e in events)==120
  assert [e['header_report_number'] for e in events]==sorted(e['header_report_number'] for e in events)
  wb=load_workbook(ROOT/data['source'],read_only=True,data_only=True)
  for e in events:
   assert e['anchor'].lower() in ' '.join(e['cells'].values()).lower()
   for c,v in e['cells'].items():assert wb[e['sheet']][c].value==v
  wb.close()
 def test_check_valve_and_scope(self):
  events=json.loads((ROOT/'evidence/12_campaign_review.json').read_text())['events'];e=next(x for x in events if x['worksheet_prefix']==70)
  assert 'Successfully' in e['cells']['S26'] and 'retrieved successfully' in e['english'];assert 'بازیابی موفق' in e['persian']
  for lang in ('EN','FA'):
   d=Document(ROOT/f'12_Drilling_Engineering_Report_{lang}.docx');text='\n'.join(p.text for p in d.paragraphs);assert '70-1 DDR Remark!' in text;assert '38-1 DDR Remark!' in text
  for typ in ('Master_Drilling_Engineering','Executive_Management'):
   for lang in ('EN','FA'):
    d=Document(ROOT/f'Field_{typ}_Report_{lang}.docx');assert any('70-1 DDR Remark!S26:S28' in p.text for p in d.paragraphs)
