# -*- coding: utf-8 -*-
"""입금 반환(환불) 자동 짝짓기 (2026-10-08 R4163-TH 실데이터).

R4163-TH 계약금 200,000원(10/03 하나·김보라(오라라(O) → 현금거래 전환 → 10/08 하나 출금 200,000원.
조건 ①같은 계좌 ②같은 금액 ③같은 이름 ④90일 이내 ⑤받은 돈−총액=출금액(또는 취소·전액).
"""
import sys
from datetime import datetime

sys.path.insert(0, '.')

import pytest

from dashboard.services import refund_match as rm

DEP = '2026/10/03\n하나,10/03, 11:20\n255******31304\n입금200,000원\n김보라(오라라(O'
OUT_TEXT = '[Web발신]\n하나,10/08, 14:30\n255******31304\n출금200,000원\n김보라(오라라(O'
OUT = {'id': 'o1', 'text': OUT_TEXT, 'amount': 200000, 'partner': '김보라(오라라(O',
       'bank': '하나', 'acct_code': 'R', 'ts': int(datetime(2026, 10, 8, 14, 30).timestamp())}


def _row(u=200000, v=0, w=2500000, total2=2500000, notes=(DEP, '', ''), inv=('미발행', '-', 'N입금')):
    vals = [''] * 30
    vals[0], vals[19], vals[20], vals[21], vals[22] = 'R4163-TH', total2, u, v, w
    vals[25], vals[26], vals[27] = inv
    return {'code': 'R4163-TH', 'row': 4164, 'vals': vals, 'notes': list(notes)}


def test_r4163_after_cash_recorded_is_single_candidate():
    c = rm.find_candidates(OUT, [_row()])
    assert len(c) == 1
    c = c[0]
    assert (c['code'], c['stage'], c['col'], c['amount'], c['stage_value']) == ('R4163-TH', '계약금', 'U', 200000, 200000)
    assert c['reason'] == '받은 돈 2,700,000원 − 총액 2,500,000원 = 200,000원'
    assert c['stage_invoice'] == '미발행' and c['dep_date'] == '2026-10-03'


def test_before_cash_recorded_no_candidate():
    # 오늘 실제 상태: 받은 돈 200,000 < 총액 2,500,000 → ⑤ 불충족
    assert rm.find_candidates(OUT, [_row(w=0)]) == []


@pytest.mark.parametrize('change', [
    {'acct_code': 'G', 'bank': '기업'},          # ① 다른 계좌
    {'amount': 210000},                           # ② 다른 금액
    {'partner': '홍길동'},                         # ③ 다른 이름
    {'ts': int(datetime(2027, 1, 20).timestamp())},  # ④ 90일 초과
])
def test_each_condition_required(change):
    o = dict(OUT, **change)
    # ② 금액이 다르면 ⑤도 다르게 맞춰 '금액 조건'만 깨지게
    row = _row(w=2510000) if 'amount' in change else _row()
    assert rm.find_candidates(o, [row]) == []


def test_cancelled_project_full_refund():
    c = rm.find_candidates(OUT, [_row(w=0, total2=0)])
    assert len(c) == 1 and '공사 취소' in c[0]['reason']


def test_already_refunded_not_proposed_again():
    refunded = DEP + '\n\n' + OUT_TEXT.replace('[Web발신]\n', '')
    assert rm.find_candidates(OUT, [_row(u=0, notes=(refunded, '', ''))]) == []


def test_amount_check_reasons():
    assert rm.amount_check(2700000, 2500000, 200000).startswith('받은 돈 2,700,000원')
    assert rm.amount_check(200000, 0, 200000).startswith('공사 취소')
    assert rm.amount_check(200000, 2500000, 200000) == ''


def test_candidate_card_has_buttons_and_value_change():
    c = rm.find_candidates(OUT, [_row()])[0]
    blocks = rm.build_candidate_blocks('o1', OUT, c)
    acts = [e['action_id'] for e in blocks[1]['elements']]
    assert acts == ['payment_refund_record', 'payment_refund_dismiss']
    assert '계약금 200,000원 → 0원' in blocks[2]['elements'][0]['text']
    assert '받은 돈 2,700,000원 − 총액 2,500,000원 = 200,000원' in blocks[0]['text']['text']


def test_row_baseline_mapping_full_refund_rules():
    from dashboard.services.payment_sync import row_baseline_mapping
    refunded = DEP + '\n\n' + OUT_TEXT.replace('[Web발신]\n', '')
    vals = _row(u=0)['vals']
    m = row_baseline_mapping(vals, [refunded, '', ''])
    assert m['u'] == 0 and m['w'] == 2500000
    assert m['u_phash'] == ''            # 값 0 단계는 phash 미인정 (폴러 게이트와 동일)
    assert m['u_ref'] != ''              # 반환 블록 phash 는 기록 → 폴러가 반환 재감지 안 함


class _FakeMgr:
    F2L = {'계약금': 'U', '중도금': 'V', '잔금': 'W', '계산서': 'Y', '총액 2': 'T', '수금 확인': 'AD',
           '계약금 계산서': 'Z', '중도금 계산서': 'AA', '잔금 계산서': 'AB'}

    def __init__(self):
        self.cells = {'U4164': 200000, 'V4164': 0, 'W4164': 2500000, 'T4164': 2500000,
                      'Y4164': '계약금 - 미발행', 'Z4164': '미발행', 'AA4164': '-', 'AB4164': 'N입금',
                      'AD4164': 'FALSE'}
        self.notes = {'U4164': DEP}

    def find_row_by_project_code(self, *a):
        return 4164

    def get_cell_value(self, sid, sn, cell):
        return self.cells.get(cell, '')

    def update_cell_value(self, sid, sn, cell, v):
        self.cells[cell] = v
        return True

    def get_cell_note(self, sid, sn, cell):
        return self.notes.get(cell, '')

    def update_cell_note(self, sid, sn, cell, note):
        self.notes[cell] = note
        return True

    def get_field_to_letter(self):
        return dict(self.F2L)


def test_commit_refund_mode_writes_note_net_value_and_invoice(monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    import dashboard.services.lead_service as ls
    import dashboard.utils.user_database as udb
    mgr = _FakeMgr()
    monkeypatch.setenv('GOOGLE_SHEET_ID', 'sid')
    monkeypatch.setenv('GOOGLE_SHEET_NAME', '공사 현황')
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: mgr)
    monkeypatch.setattr(udb, 'get_audit_repository', lambda: type('A', (), {'log_action': lambda *a, **k: None})())
    ok, old, new, err = sb._commit_intake_to_sheet('R4163-TH', '계약금', -200000, OUT_TEXT, 'U_SB', refund=True)
    assert ok and (old, new) == (200000, 0), err
    assert mgr.cells['U4164'] == 0
    assert mgr.notes['U4164'].startswith(DEP + '\n\n') and '출금200,000원' in mgr.notes['U4164']
    assert '[Web발신]' not in mgr.notes['U4164']
    assert mgr.cells['Z4164'] == '-'                     # 전액 반환 → 계약금 계산서 '-'
    assert '계약금' not in mgr.cells['Y4164']             # Y 요약에서 계약금 미발행 사라짐


def test_commit_refund_keeps_issued_invoice(monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    import dashboard.services.lead_service as ls
    import dashboard.utils.user_database as udb
    mgr = _FakeMgr()
    mgr.cells['Z4164'] = '발행'
    monkeypatch.setenv('GOOGLE_SHEET_ID', 'sid')
    monkeypatch.setenv('GOOGLE_SHEET_NAME', '공사 현황')
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: mgr)
    monkeypatch.setattr(udb, 'get_audit_repository', lambda: type('A', (), {'log_action': lambda *a, **k: None})())
    sb._commit_intake_to_sheet('R4163-TH', '계약금', -200000, OUT_TEXT, 'U_SB', refund=True)
    assert mgr.cells['Z4164'] == '발행'                  # 발행 계산서는 사람이 판단


# ── 받은 돈 돌려줌 (2026-10-15 R4101-JSH) ──
# 계약금 1,200,000원(9/11 하나·김영은) → 무부가세 현금 공사로 바뀌어 10/15 하나에서 그대로 출금 반환.
# 총액 12,000,000 = 받은 돈(계약금 1.2M + 상계 0.4M + 잔금 10.4M) → ⑤(과입금) 불충족.
R4101_DEP = '2026/09/11\n하나,09/11, 13:00\n255******31304\n입금1,200,000원\n김영은'
R4101_OUT = {'id': 'o4101', 'text': '[Web발신]\n하나,10/15, 15:00\n255******31304\n출금1,200,000원\n김영은',
             'amount': 1200000, 'partner': '김영은', 'bank': '하나', 'acct_code': 'R',
             'ts': int(datetime(2026, 10, 15, 15, 0).timestamp())}


def _r4101(u=1200000):
    vals = [''] * 30
    vals[0], vals[19], vals[20], vals[21], vals[22] = 'R4101-JSH', 12000000, u, 400000, 10400000
    vals[25], vals[26], vals[27] = 'N입금', 'N입금', 'N입금'
    notes = (R4101_DEP, '대체전표\n2026-10-15\n400,000원\n중고매입 상계처리',
             '2026/10/12\n농협, 10/12 10:45\n352-****-1682-33\n입금 10,400,000원\n김영은')
    return {'code': 'R4101-JSH', 'row': 4102, 'vals': vals, 'notes': list(notes)}


def test_reroute_only_with_flag():
    assert rm.find_candidates(R4101_OUT, [_r4101()]) == []            # 기존 규칙: ⑤ 불충족
    c = rm.find_candidates(R4101_OUT, [_r4101()], include_reroute=True)
    assert len(c) == 1 and c[0]['kind'] == 'reroute' and c[0]['stage'] == '계약금'
    assert (c[0]['unpaid_before'], c[0]['unpaid_after']) == (0, 1200000)


def test_reroute_card_wording():
    c = rm.find_candidates(R4101_OUT, [_r4101()], include_reroute=True)[0]
    blocks = rm.build_candidate_blocks('o4101', R4101_OUT, c)
    body = blocks[0]['text']['text']
    assert '받은 돈을 돌려준 출금' in body
    assert '받은 돈 : 12,000,000원 / 총액 12,000,000원 (더 받은 돈 없음)' in body
    assert '기록하면 : 계약금 1,200,000원 → 0원  ·  미수금 0원 → 1,200,000원 (다시 받을 돈)' in body
    assert '미수금이 1,200,000원이 됩니다' in blocks[1]['elements'][0]['confirm']['text']['text']
    assert '[반환 아님]' in blocks[2]['elements'][0]['text']


def test_overpay_card_wording_unchanged():
    c = rm.find_candidates(OUT, [_row()], include_reroute=True)
    assert len(c) == 1 and c[0]['kind'] == 'overpay'
    assert '금액 검증 : 받은 돈 2,700,000원 − 총액 2,500,000원 = 200,000원' in \
        rm.build_candidate_blocks('o1', OUT, c[0])[0]['text']['text']


def test_reroute_requires_name_amount_account():
    for change in ({'partner': '홍길동'}, {'amount': 1100000}, {'acct_code': 'G', 'bank': '기업'}):
        assert rm.find_candidates(dict(R4101_OUT, **change), [_r4101()], include_reroute=True) == []


def test_reroute_skipped_when_already_refunded_or_overpaid_mismatch():
    done = _r4101(u=0)
    done['notes'][0] = R4101_DEP + '\n\n' + R4101_OUT['text'].replace('[Web발신]\n', '')
    assert rm.find_candidates(R4101_OUT, [done], include_reroute=True) == []     # 이미 반환 기록
    over = _r4101()
    over['vals'][19] = 10000000                                                    # 과입금인데 금액이 다름
    assert rm.find_candidates(R4101_OUT, [over], include_reroute=True) == []


def test_scan_prefers_overpay_over_reroute(monkeypatch):
    posted = []
    import dashboard.blueprints.sms_inbound as si

    class _R:
        def set(self, *a, **k): return True
        def delete(self, *a): pass
        def exists(self, *a): return False
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': _R()})())
    monkeypatch.setattr(si, 'recent_outflows', lambda since: [R4101_OUT])
    monkeypatch.setattr(rm, 'post_candidate', lambda oid, o, c, slack=None: posted.append(c) or True)
    monkeypatch.setattr(rm, 'load_rows', lambda: [_r4101()])
    assert rm.scan_and_propose()['proposed'] == 1 and posted[0]['kind'] == 'reroute'
