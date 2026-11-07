import sys
import unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'tools'))
from extract_drilling_tracks import activity_day, scalar, get_row


class TrackTests(unittest.TestCase):
    def test_activity_preserves_status_and_stops_at_first_total(self):
        remark={'M22':'Status','A24':'0','C24':'1','E24':'12','M24':'NPT',
                'A25':'1','C25':'2','E25':'6','M25':'UNPLN',
                'A26':'2','C26':'3','E26':'6','M26':'NPT-UNPLN',
                'A30':'Total','M32':'Status','A33':'0','C33':'1','E33':'9','M33':'PLN'}
        a=activity_day(remark,'1-1 DDR Remark','ddr',{},set())
        assert a['complete'] and (a['npt_h'],a['unplanned_h'],a['mixed_h'],a['planned_h'])==(12,6,6,0)
        assert len(a['segments'])==3


    def test_time_style_and_ambiguous_duration(self):
        remark={'J21':'Code 2 P-N-UP','F22':'0','G22':'1','H22':'0.25','J22':'P',
                'F23':'1','G23':'2','H23':'18','J23':'UP','A68':'Total Hrs.'}
        a=activity_day(remark,'1-1 Remark','legacy',{'H22':'12'},{'12'})
        assert a['complete'] and a['planned_h']==6 and a['unplanned_h']==18
        remark['H22']='2/5'
        a=activity_day(remark,'1-1 Remark','legacy',{'H22':'12'},{'12'})
        assert not a['complete'] and a['planned_h'] is None


    def test_range_parser_refuses_slashes(self):
        assert scalar({'D1':'2-4'},'D1','x')['value'] is None
        assert scalar({'D1':'2/5'},'D1','x')['value'] is None
        assert get_row({'I9':' Mud Weight','J9':'Unit'},'I',r'^Mud Weight\b')==9
