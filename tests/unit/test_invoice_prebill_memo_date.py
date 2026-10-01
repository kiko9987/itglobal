# -*- coding: utf-8 -*-
"""계산서 선발행 자동 메모에 발행일 기록 (2026-09-30 SB 요청).

선발행(입금 0) 단계를 '발행'으로 자동기록할 때 계산서_메모(Y 노트)에 발행액을 남기는데,
날짜가 없어 SB 수기 메모('2026-08-14 41,600,000원 … 발행')와 달리 언제 발행했는지 몰랐다.
→ 'YYYY-MM-DD {단계} 선발행 X원'. 재호출 시 중복 기록 없음 + 금액 파서는 금액 1개만 인식.
"""
import re
import sys
from datetime import datetime

sys.path.insert(0, '.')

import dashboard.blueprints.slack_bot as sb

F2L = {'계약금': 'U', '중도금': 'V', '잔금': 'W', '계산서': 'Y',
       '계약금 계산서': 'Z', '중도금 계산서': 'AA', '잔금 계산서': 'AB',
       '수금 확인': 'AD', '총액 1': 'R', '총액 2': 'T'}


class _FakeManager:
    def __init__(self, row=10):
        self.row = row
        self.vals = {f'R{row}': 2620000, f'T{row}': 2882000, f'U{row}': 0, f'V{row}': 0,
                     f'W{row}': 0, f'Z{row}': '-', f'AA{row}': '-', f'AB{row}': '미발행',
                     f'AD{row}': False, f'Y{row}': '잔금 - 미발행'}
        self.notes = {}

    def find_row_by_project_code(self, *a, **k):
        return self.row

    def get_field_to_letter(self):
        return F2L

    def get_cell_value(self, sid, sn, cell):
        return self.vals.get(cell, '')

    def update_cell_value(self, sid, sn, cell, v):
        self.vals[cell] = v
        return True

    def get_cell_note(self, sid, sn, cell):
        return self.notes.get(cell, '')

    def update_cell_note(self, sid, sn, cell, note):
        self.notes[cell] = note
        return True


class _FakeRedis:
    """invoice_memo_card:{card} nx 마킹만 흉내 (카드 한 장 = 메모 한 줄)."""
    def __init__(self):
        self.store = {}

    def set(self, key, val, nx=False, ex=None):
        if nx and key in self.store:
            return None
        self.store[key] = val
        return True


def _run(monkeypatch, fm, biz='', card_key='', amt='2620000', fake_redis=None):
    monkeypatch.setenv('GOOGLE_SHEET_ID', 'SID')
    monkeypatch.setenv('GOOGLE_SHEET_NAME', '공사 현황')
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: fm)
    # ⚠️ 감사 로그는 운영 users.db 에 직접 기록됨 — 반드시 가짜로 대체 (운영 DB 오염 방지)
    import dashboard.utils.user_database as udb
    monkeypatch.setattr(udb, 'get_audit_repository',
                        lambda: type('A', (), {'log_action': lambda self, **k: None})())
    if fake_redis is not None:
        import dashboard.utils.redis_client as rcm
        monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': fake_redis})())
    sb._mark_invoice_issued_in_sheet('R4100-MS', '잔금', invoice_amt=amt, vat_val='sep',
                                     biz=biz, card_key=card_key)


def test_prebill_memo_has_issue_date(monkeypatch):
    fm = _FakeManager()
    _run(monkeypatch, fm)
    today = f'{datetime.now():%Y-%m-%d}'
    assert fm.vals['AB10'] == '발행'
    assert fm.notes['Y10'] == f'{today} 잔금 선발행 2,882,000원'
    # 금액 파서(billStatus/_project_issued_invoice 미러)는 금액 하나만 인식 — 날짜 무영향
    assert re.findall(r'[\d,]+\s*원', fm.notes['Y10']) == ['2,882,000원']


def test_normal_issue_writes_dated_line_without_amount(monkeypatch):
    fm = _FakeManager()
    fm.vals['W10'] = 2882000                        # 입금 있는 단계 → 일반 발행
    fm.notes['Y10'] = '2026-08-14 41,600,000원 부가세 별도 발행'   # SB 수기 메모 보존
    _run(monkeypatch, fm)
    today = f'{datetime.now():%Y-%m-%d}'
    assert fm.vals['AB10'] == '발행'
    assert fm.notes['Y10'] == f'2026-08-14 41,600,000원 부가세 별도 발행\n{today} 잔금 발행'
    # '원' 금액 없는 줄 → 금액 파서 결과 불변
    assert re.findall(r'[\d,]+\s*원', fm.notes['Y10']) == ['41,600,000원']


def test_already_issued_stage_records_history_only(monkeypatch):
    """이미 '발행'인 단계에 또 발행(수정발행·사업자 분할 두 번째 장) — 계산서 칸·Y 요약은 그대로,
    메모엔 이력 한 줄 (2026-10-01, 예전엔 아무것도 안 써서 재발행 이력 유실)."""
    fm = _FakeManager()
    fm.vals['AB10'] = '발행'
    _run(monkeypatch, fm, biz='SM CORPORATION')
    today = f'{datetime.now():%Y-%m-%d}'
    assert fm.vals['AB10'] == '발행'
    assert fm.vals['Y10'] == '잔금 - 미발행'          # Y 요약 재계산 안 함 (칸 변경 없음)
    assert fm.notes['Y10'] == f'{today} 잔금 선발행 2,882,000원 · SM CORPORATION'


def test_split_issue_same_day_keeps_both_lines(monkeypatch):
    """고객 요청 사업자 분할 — 같은 날 같은 금액 두 장도 사업자가 다르면 둘 다 기록 (G3991-YM)."""
    fm = _FakeManager()
    fr = _FakeRedis()
    _run(monkeypatch, fm, biz='SM CORPORATION', card_key='C:1', amt='1310000', fake_redis=fr)
    _run(monkeypatch, fm, biz='(주)설린', card_key='C:2', amt='1310000', fake_redis=fr)
    today = f'{datetime.now():%Y-%m-%d}'
    assert fm.notes['Y10'] == (f'{today} 잔금 선발행 1,441,000원 · SM CORPORATION\n'
                               f'{today} 잔금 선발행 1,441,000원 · (주)설린')


def test_same_card_retrigger_writes_once(monkeypatch):
    """같은 요청 카드 스레드에 파일을 또 올려 재처리돼도 메모는 한 줄 (카드 한 장 = 한 줄)."""
    fm = _FakeManager()
    fr = _FakeRedis()
    _run(monkeypatch, fm, biz='오쿠드', card_key='C:9', fake_redis=fr)
    fm.notes['Y10'] = fm.notes['Y10']            # 다음 날 재첨부를 흉내: 날짜가 달라도
    first = fm.notes['Y10']
    _run(monkeypatch, fm, biz='오쿠드', card_key='C:9', fake_redis=fr)
    assert fm.notes['Y10'] == first


def test_prebill_memo_not_duplicated_for_old_dateless_line(monkeypatch):
    fm = _FakeManager()
    fm.notes['Y10'] = '잔금 선발행 2,882,000원'     # 날짜 도입 전 기록
    fm.vals['AB10'] = '미발행'
    _run(monkeypatch, fm)
    assert fm.notes['Y10'] == '잔금 선발행 2,882,000원'
