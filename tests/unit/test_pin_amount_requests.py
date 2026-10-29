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


def test_remind_text_has_amount_section_and_new_header():
    t = pr.build_pin_remind_text({'amount_reqs': [{'summary': '`G1` 금액 수정', 'permalink': 'https://x'}]})
    assert '경영지원 오늘 할 일 1건' in t
    assert '금액 반영 대기 — PM 반영 후 ✅ (1건)' in t
    assert '• `G1` 금액 수정  |  <https://x|바로가기>' in t
