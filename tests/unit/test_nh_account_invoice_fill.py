# -*- coding: utf-8 -*-
"""농협 N통장 입금 → 계산서 칸 'N입금' 자동 채움 + 직접 고른 단계 발행 시 N입금→발행 (2026-10-15).

계기: R4101-JSH 잔금 10,400,000원(농협 352-****-1682-33)이 '현금' 글자가 없어 '미발행'으로 채워짐.
시트 실적상 농협 통장 입금 계산서 칸은 97%가 'N입금'(계산서 없이 받는 돈). 판정은 계좌번호로 —
'농협' 글자로 보면 하나 통장 입금의 입금자 '수원원예농협'(R2346-TH)에 걸린다.
예외(G2918-YG: 농협 입금 계약금에 실제 계산서 발행)는 계산서 요청에서 그 단계를 고르면 '발행'으로.
"""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb

NH = '[Web발신]\n농협, 10/12 10:45\n352-****-1682-33\n입금 10,400,000원\n김영은'
HANA_NH_PAYER = '[Web발신]\n일시 05/26, 13:42\n입금 16,170,000원\n계좌번호 255******31304\n적요 수원원예농협'
HANA = '[Web발신]\n하나,09/30, 17:02\n255******31304\n입금 3,014,000원\n대연이엔지주식회'


class _FakeMgr:
    F2L = {'계약금': 'U', '중도금': 'V', '잔금': 'W', '계산서': 'Y', '총액 1': 'R', '총액 2': 'T',
           '수금 확인': 'AD', '계약금 계산서': 'Z', '중도금 계산서': 'AA', '잔금 계산서': 'AB'}

    def __init__(self, **cells):
        self.cells = {'R10': 12000000, 'T10': 13200000, 'U10': 1200000, 'V10': 0, 'W10': 0,
                      'Y10': '계약금 - 미발행', 'Z10': '미발행', 'AA10': '-', 'AB10': '',
                      'AD10': 'FALSE'}
        self.cells.update(cells)
        self.notes = {}

    def find_row_by_project_code(self, *a, **k):
        return 10

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


@pytest.fixture
def mgr(monkeypatch):
    m = _FakeMgr()
    monkeypatch.setenv('GOOGLE_SHEET_ID', 'sid')
    monkeypatch.setenv('GOOGLE_SHEET_NAME', '공사 현황')
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: m)
    import dashboard.utils.user_database as udb     # ⚠️ 운영 users.db 감사 로그 차단
    monkeypatch.setattr(udb, 'get_audit_repository',
                        lambda: type('A', (), {'log_action': lambda *a, **k: None})())
    return m


def test_nh_account_deposit_fills_n_ipgeum(mgr):
    ok, old, new, err = sb._commit_intake_to_sheet('R4101-JSH', '잔금', 10400000, NH, 'U_SB')
    assert ok, err
    assert mgr.cells['W10'] == 10400000
    assert mgr.cells['AB10'] == 'N입금'


def test_hana_deposit_from_nonghyup_payer_stays_mibalhaeng(mgr):
    # 입금자 이름의 '농협'은 우리 N통장이 아님 → 계좌(하나) 기준 '미발행'
    ok, *_ = sb._commit_intake_to_sheet('R2346-TH', '잔금', 16170000, HANA_NH_PAYER, 'U_SB')
    assert ok and mgr.cells['AB10'] == '미발행'


def test_hana_deposit_unchanged(mgr):
    ok, *_ = sb._commit_intake_to_sheet('G4125-YM', '잔금', 3014000, HANA, 'U_SB')
    assert ok and mgr.cells['AB10'] == '미발행'


def test_split_memo_uses_original_sms_account(mgr):
    # 분할 노트엔 계좌번호가 없음 → 원 입금 문자(acct_text)로 판정
    memo = sb._build_split_memo({'date_md': '10/12', 'partner': '김영은', 'bank': '농협'},
                                5200000, 10400000)
    assert '352' not in memo
    ok, *_ = sb._commit_intake_to_sheet('R4101-JSH', '잔금', 5200000, memo, 'U_SB', acct_text=NH)
    assert ok and mgr.cells['AB10'] == 'N입금'


def test_existing_invoice_cell_not_overwritten(mgr):
    mgr.cells['AB10'] = '발행'
    ok, *_ = sb._commit_intake_to_sheet('R4101-JSH', '잔금', 10400000, NH, 'U_SB')
    assert ok and mgr.cells['AB10'] == '발행'


def test_invoice_on_selected_n_stage_marks_issued(mgr):
    # G2918-YG 형: 농협 입금 계약금(N입금)에 실제로 계산서를 끊음 → 고른 단계는 '발행'
    mgr.cells.update({'Z10': 'N입금', 'AA10': '미발행', 'V10': 5500000})
    sb._mark_invoice_issued_in_sheet('G2918-YG', '계약금', invoice_amt='1090909', vat_val='sep')
    assert mgr.cells['Z10'] == '발행'
    assert mgr.cells['AA10'] == '미발행'          # 부분 발행 — 다른 단계는 그대로


def test_invoice_on_other_stage_keeps_n_ipgeum(mgr):
    # 직접 고르지 않은 단계의 N입금은 전체 발행(통합)이어도 그대로
    mgr.cells.update({'Z10': 'N입금', 'W10': 12000000, 'AB10': '미발행'})
    sb._mark_invoice_issued_in_sheet('R4101-JSH', '잔금', invoice_amt='12000000', vat_val='sep')
    assert mgr.cells['AB10'] == '발행'
    assert mgr.cells['Z10'] == 'N입금'


class _RowMgr(_FakeMgr):
    """행 한 번 읽기(get_row_map) 지원 매니저 — 칸별 조회 횟수 집계 (2026-10-15 속도 개선)."""
    def __init__(self, **cells):
        super().__init__(**cells)
        self.cell_reads = 0

    def get_cell_value(self, sid, sn, cell):
        self.cell_reads += 1
        return super().get_cell_value(sid, sn, cell)

    def get_row_map(self, sid, sn, row):
        return {k[:-len(str(row))]: v for k, v in self.cells.items() if k.endswith(str(row))}


@pytest.mark.parametrize('memo, stage, amount', [(NH, '잔금', 10400000), (HANA, '잔금', 3014000)])
def test_row_snapshot_same_result_fewer_reads(monkeypatch, memo, stage, amount):
    import dashboard.services.lead_service as ls
    import dashboard.utils.user_database as udb
    monkeypatch.setenv('GOOGLE_SHEET_ID', 'sid'); monkeypatch.setenv('GOOGLE_SHEET_NAME', '공사 현황')
    monkeypatch.setattr(udb, 'get_audit_repository',
                        lambda: type('A', (), {'log_action': lambda *a, **k: None})())
    a, b = _FakeMgr(), _RowMgr()
    for m in (a, b):
        monkeypatch.setattr(ls, 'get_sheets_manager', lambda m=m: m)
        assert sb._commit_intake_to_sheet('X1', stage, amount, memo, 'U_SB')[0]
    assert a.cells == b.cells and a.notes == b.notes          # 같은 결과
    assert b.cell_reads <= 1                                  # 덮어쓰기 직전 재확인만 실시간 조회


def test_refund_record_restores_card_on_failure(monkeypatch):
    """[반환으로 기록] → 즉시 '기록 중' 카드, 재검증 실패면 원래 후보 카드로 되돌림."""
    from dashboard.services import refund_match as rm
    import dashboard.utils.redis_client as rcm

    class _R:
        def set(self, *a, **k): return True
        def delete(self, *a): pass
        def exists(self, *a): return False
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': _R()})())
    cand = {'code': 'R4101-JSH', 'stage': '계약금', 'amount': 1200000, 'stage_value': 1200000, 'row': 4102,
            'kind': 'reroute', 'paid': 12000000, 'total2': 12000000, 'unpaid_before': 0, 'unpaid_after': 1200000,
            'dep_date': '2026-09-11', 'dep_bank': '하나', 'dep_partner': '김영은', 'reason': 'x'}
    out = {'id': 'o1', 'text': '출금1,200,000원\n김영은', 'ts': 1792000000, 'bank': '하나', 'partner': '김영은'}
    monkeypatch.setattr(rm, 'load_candidate', lambda oid: {'outflow': out, 'cand': cand})
    monkeypatch.setattr(rm, 'load_row', lambda code, row: [{'code': code}])
    monkeypatch.setattr(rm, 'find_candidates', lambda o, rows, include_reroute=False: [])   # 조건 바뀜
    monkeypatch.setattr(sb, '_intake_ephemeral', lambda *a, **k: None)
    updates = []
    client = type('Cl', (), {'chat_update': lambda self, **k: updates.append(k)})()
    sb._do_refund_record(client, sb._SETTLEMENT_CHECKER_ID, 'C1', '1.1', 'o1')
    assert '반환 기록 중' in updates[0]['text']
    assert updates[-1]['blocks'][1]['elements'][0]['action_id'] == 'payment_refund_record'   # 버튼 복원
