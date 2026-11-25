"""차량 배정·방문 일정 불러오기 속도(2026-11-25 부장님 '느려') — 결과는 같고 덜 계산하는지.
- 직원 명단: 화면 요청 하나 안에서는 한 번만 만든다(요청 밖은 그때그때)
- 리드 행: 화면 불러오기(GET)만 10초 메모, 저장(POST)은 늘 새로, 행은 요청마다 복사
- 방문 일정 날짜 고르기: 날짜 하나만 적힌 칸은 바로 비교(여러 날짜·다른 모양은 예전처럼)
이름은 모두 가짜."""
from datetime import date

import pandas as pd
from flask import Flask

from dashboard.services import vehicle_assign as va
from dashboard.services import visit_plan as vp


def test_staff_built_once_per_request(monkeypatch):
    calls = []
    monkeypatch.setattr(va, '_staff_now', lambda: calls.append(1) or [{'code': 'TH', 'name': '서준', 'slack_id': 'UTH', 'dept': ''}])
    app = Flask(__name__)
    with app.test_request_context('/api/vehicles/assign'):
        a = va.staff()
        a[0]['name'] = '바뀜'                       # 받은 쪽이 고쳐도 메모는 그대로
        assert va.staff()[0]['name'] == '서준' and len(calls) == 1
    with app.test_request_context('/api/vehicles/assign'):
        va.staff()
        assert len(calls) == 2                    # 다음 요청은 새로
    va.staff()
    va.staff()
    assert len(calls) == 4                        # 요청 밖(스케줄러)은 그때그때


def test_lead_rows_memo_get_only(monkeypatch):
    from dashboard.blueprints import gw_vehicles as bp
    import dashboard.services.lead_service as ls
    reads = []
    monkeypatch.setattr(bp, '_leads_memo', {'at': 0.0, 'val': None})
    monkeypatch.setattr(ls, 'load_leads_data', lambda *a, **k: reads.append(1) or pd.DataFrame([{'리드 No': 'L-1', '고객명': '가상'}]))
    app = Flask(__name__)
    with app.test_request_context('/api/vehicles/assign', method='GET'):
        r = bp._leads()
        r[0]['고객명'] = '바뀜'
        assert bp._leads()[0]['고객명'] == '가상' and len(reads) == 1      # 10초 안 GET 은 메모, 행은 복사
    with app.test_request_context('/api/vehicles/plan', method='POST'):
        bp._plan_leads()
        assert len(reads) == 2                                              # 저장은 늘 새로
    monkeypatch.setattr(bp, '_LEADS_TTL', 0.0)
    with app.test_request_context('/api/vehicles/assign', method='GET'):
        bp._leads()
        assert len(reads) == 3                                              # 오래되면 새로
    monkeypatch.setattr(ls, 'load_leads_data', lambda *a, **k: None)
    with app.test_request_context('/api/vehicles/plan', method='POST'):
        try:
            bp._plan_leads()
            assert False, '못 읽으면 멈춰야 함'
        except bp._LeadsUnavailable:
            pass


def test_on_day_fast_path_same_result():
    day = '2026-11-25'
    leads = [{'리드 No': n, '방문 예정일': v} for n, v in [
        ('A', '2026-11-25'), ('B', '2026-11-24'), ('C', "'2026.11.25"), ('D', '2026-11-24~2026-11-26'),
        ('E', '2026-11-25 14:00'), ('F', ''), ('G', None), ('H', '2026-13-40'), ('I', '2026/11/25')]]
    got = [l['리드 No'] for l in vp._on_day(day, leads)]
    assert got[:2] == ['A', 'C'] and 'D' in got and 'I' in got and 'B' not in got and 'H' not in got
    assert date.fromisoformat(day)
