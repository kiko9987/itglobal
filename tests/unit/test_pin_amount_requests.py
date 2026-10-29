"""경영지원 오늘 할 일 — 금액 반영 대기 줄 (2026-10-30)."""
from dashboard.services import pin_remind as pr


def test_amount_req_summary():
    s = pr._amount_req_summary('amount', {
        'code': 'G4152-JW', 'before': {'총액 1': '11242000', '부가세': 'True'},
        'updates': {'부가세': False}, 'requester_initial': 'JW', 'requested_at': '10.26 13:39'})
    assert s == '`G4152-JW` 부가세 있음→없음 · 요청 JW 10.26 13:39'
    s = pr._amount_req_summary('amount', {
        'code': 'G4166-YM', 'before': {'총액 1': 6300000.0}, 'updates': {'총액 1': '9000000'},
        'requester_initial': 'YM', 'requested_at': '10.23 15:40'})
    assert s == '`G4166-YM` 총액 1 6,300,000원→9,000,000원 · 요청 YM 10.23 15:40'
    s = pr._amount_req_summary('cancel', {'code': 'R4181-JSH', 'amt_disp': '2,420,000원',
                                          'requester_initial': 'JSH', 'requested_at': '10.20 09:00'})
    assert s == '`R4181-JSH` 공사 취소 매출 0원 (2,420,000원 → 0원) · 요청 JSH 10.20 09:00'


def test_collect_skips_already_applied(monkeypatch):
    """PM 에 이미 반영된 요청(✅ 안 눌러 pending 이 남은 것)은 뺀다 (10-30 G4168·R4140)."""
    import json
    store = {
        'project_amount_req:C1:1.0': {'code': 'R4140-YM', 'before': {'총액 1': '2000000'},
                                      'updates': {'총액 1': '222000'}},      # 2,220,000 으로 반영됨
        'project_amount_req:C1:2.0': {'code': 'R4068-TH', 'before': {'부가세': 'True'},
                                      'updates': {'부가세': False}},          # 아직 그대로 → 남김
        'project_cancel_money_req:C1:3.0': {'code': 'G4168-TH', 'amt_disp': '2,000,000원'},  # 0원 반영됨
        'project_cancel_money_req:C1:4.0': {'code': 'R9999-AA', 'amt_disp': '1,000,000원'},  # 아직
    }

    class _R:
        def scan_iter(self, match, count=0):
            p = match.rstrip('*')
            return [k for k in store if k.startswith(p)]

        def get(self, k):
            return json.dumps(store[k])

    class _RC:
        redis = _R()

    monkeypatch.setattr('dashboard.utils.redis_client.get_redis_client', lambda: _RC())
    monkeypatch.setattr(pr, '_client', lambda: None)
    recs = [
        {'프로젝트 코드': 'R4140-YM', '총액 1': 2220000.0, '부가세': True, '공사 확정': '2026-09-30'},
        {'프로젝트 코드': 'R4068-TH', '총액 1': 3244500.0, '부가세': True, '공사 확정': '2026-09-02'},
        {'프로젝트 코드': 'G4168-TH', '총액 1': 0.0, '부가세': True, '공사 확정': ''},
        {'프로젝트 코드': 'R9999-AA', '총액 1': 1000000.0, '부가세': True, '공사 확정': ''},
    ]
    monkeypatch.setattr('dashboard.services.project_service.get_project_records', lambda *a, **k: recs)
    got = [e['summary'].split('`')[1] for e in pr.collect_amount_requests()]
    assert got == ['R4068-TH', 'R9999-AA']


def test_remind_text_has_amount_section_and_new_header():
    t = pr.build_pin_remind_text({'amount_reqs': [{'summary': '`G1` 금액 수정', 'permalink': 'https://x'}]})
    assert '경영지원 오늘 할 일 1건' in t
    assert '금액 반영 대기 — PM 반영 후 ✅ (1건)' in t
    assert '• `G1` 금액 수정  |  <https://x|바로가기>' in t
