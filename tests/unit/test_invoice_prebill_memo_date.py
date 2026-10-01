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


def _run(monkeypatch, fm, biz='', card_key='', amt='2620000', fake_redis=None, replace=None):
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
                                     biz=biz, card_key=card_key, replace=replace)


def test_prebill_memo_has_issue_date(monkeypatch):
    fm = _FakeManager()
    _run(monkeypatch, fm)
    today = f'{datetime.now():%Y-%m-%d}'
    assert fm.vals['AB10'] == '발행'
    assert fm.notes['Y10'] == f'{today} 잔금 선발행 2,882,000원'
    # 금액 파서(billStatus/_project_issued_invoice 미러)는 금액 하나만 인식 — 날짜 무영향
    assert re.findall(r'[\d,]+\s*원', fm.notes['Y10']) == ['2,882,000원']


def test_normal_issue_writes_dated_line_with_amount(monkeypatch):
    """일반 발행(입금 있는 단계)도 이 장 금액(VAT 포함)·사업자 기록 — 분할 발행 사업자별 금액 (2026-10-01)."""
    fm = _FakeManager()
    fm.vals['W10'] = 2882000                        # 입금 있는 단계 → 일반 발행
    fm.notes['Y10'] = '2026-08-14 41,600,000원 부가세 별도 발행'   # SB 수기 메모 보존
    _run(monkeypatch, fm, biz='(주)설린')
    today = f'{datetime.now():%Y-%m-%d}'
    assert fm.vals['AB10'] == '발행'
    assert fm.notes['Y10'] == (f'2026-08-14 41,600,000원 부가세 별도 발행\n'
                               f'{today} 잔금 발행 2,882,000원 · (주)설린')
    # 'X원 1개' 폴백은 단계 표시 줄을 빼고 셈 → SB 수기 금액 그대로
    assert sb._pre_issued_amount(fm.notes['Y10'], '잔금') == 41600000


def test_normal_issue_line_not_taken_as_preissued_amount():
    """일반 발행 줄의 금액은 선발행 금액으로 쓰지 않음 (다른 단계 선발행 판정 오염 방지)."""
    memo = '2026-09-01 계약금 발행 1,000,000원 · A'
    assert sb._pre_issued_amount(memo, '잔금') == 0
    memo2 = memo + '\n2026-09-05 잔금 선발행 2,000,000원 · A'
    assert sb._pre_issued_amount(memo2, '잔금') == 2000000


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


G3991 = '\n'.join([
    '2026-08-27 SM/설린 50:50 발행',
    '2026-08-27 잔금 선발행 3,410,000원 · SM CORPORATION',
    '2026-09-02 잔금 취소 -3,410,000원 · SM CORPORATION (수정발행)',
    '2026-09-02 잔금 선발행 1,705,000원 · (주)설린 (수정발행: 고객 요청 사업자 분할)',
    '2026-09-02 잔금 선발행 1,705,000원 · SM CORPORATION (수정발행: 고객 요청 사업자 분할)',
])


def test_active_lines_cancel_removes_replaced_invoice():
    """수정발행 취소 줄 = 같은 단계·금액·사업자 장을 지움 → 현재 = 분할 두 장 (2026-10-01)."""
    act = sb._active_invoice_lines(G3991)
    assert [(a['biz'], a['amt']) for a in act] == [('(주)설린', 1705000), ('SM CORPORATION', 1705000)]
    assert sb._pre_issued_amount(G3991, '잔금') == 3410000


def test_split_on_different_days_sums():
    """분할 두 장을 다른 날 처리해도 합산 (날짜 묶음 방식의 한계 해소)."""
    memo = ('2026-10-02 잔금 선발행 1,000,000원 · A상사\n'
            '2026-10-05 잔금 선발행 1,000,000원 · B상사')
    assert sb._pre_issued_amount(memo, '잔금') == 2000000


def test_cancel_without_match_is_ignored_and_all_cancelled_is_zero():
    memo = '2026-10-02 잔금 선발행 1,000,000원 · A상사\n2026-10-03 잔금 취소 -999,000원 · A상사 (수정발행)'
    assert sb._pre_issued_amount(memo, '잔금') == 1000000          # 금액 안 맞는 취소 → 무시
    memo2 = '2026-10-02 잔금 선발행 1,000,000원 · A상사\n2026-10-03 잔금 취소 -1,000,000원 · A상사 (수정발행)'
    assert sb._pre_issued_amount(memo2, '잔금') == 0                # 전부 취소 → 0 (폴백 X)


def test_replace_writes_cancel_line_before_new_line(monkeypatch):
    """수정발행 요청 첨부 완료 → 메모에 취소 줄 + 새 장 줄 (G3991 첫 분할 카드 흐름)."""
    fm = _FakeManager()
    fm.vals['AB10'] = '발행'
    fm.notes['Y10'] = '2026-08-27 잔금 선발행 2,882,000원 · SM CORPORATION'
    _run(monkeypatch, fm, biz='(주)설린', amt='1310000',
         replace=[{'date': '2026-08-27', 'stage': '잔금', 'amt': 2882000, 'biz': 'SM CORPORATION'}])
    today = f'{datetime.now():%Y-%m-%d}'
    assert fm.notes['Y10'] == ('2026-08-27 잔금 선발행 2,882,000원 · SM CORPORATION\n'
                               f'{today} 잔금 취소 -2,882,000원 · SM CORPORATION (수정발행)\n'
                               f'{today} 잔금 선발행 1,441,000원 · (주)설린')
    act = sb._active_invoice_lines(fm.notes['Y10'])
    assert [(a['biz'], a['amt']) for a in act] == [('(주)설린', 1441000)]


def test_project_issued_invoice_flags_over_issue(monkeypatch):
    """대체 표시 없이 다시 발행 → 장 합계 > 총액2 → over_issued 경고."""
    rec = {'프로젝트 코드': 'X1', '총액 1': 3100000, '총액 2': 3410000, '부가세': True,
           '계약금': 0, '중도금': 0, '잔금': 0, '미수금': 3410000,
           '계약금 계산서': '', '중도금 계산서': '', '잔금 계산서': '발행', '수금 확인': False,
           '계산서_메모': ('2026-08-27 잔금 선발행 3,410,000원 · SM\n'
                         '2026-09-02 잔금 선발행 1,705,000원 · 설린\n'
                         '2026-09-02 잔금 선발행 1,705,000원 · SM')}
    import dashboard.services.project_service as ps
    monkeypatch.setattr(ps, 'get_project_records', lambda: [rec])
    s = sb._project_issued_invoice('X1')
    assert s['over_issued'] == 6820000 and len(s['active']) == 3
    assert '총액' in sb._fmt_issued_warn(s) and '대체 표시' in sb._fmt_issued_warn(s)


def _issued_summary(monkeypatch, memo, **over):
    rec = {'프로젝트 코드': 'P1', '총액 1': 3100000, '총액 2': 3410000, '부가세': True,
           '계약금': 0, '중도금': 1705000, '잔금': 0, '미수금': 1705000,
           '계약금 계산서': '-', '중도금 계산서': '-', '잔금 계산서': '발행', '수금 확인': False,
           '계산서_메모': memo}
    rec.update(over)
    import dashboard.services.project_service as ps
    monkeypatch.setattr(ps, 'get_project_records', lambda: [rec])
    return sb._project_issued_invoice('P1')


def test_issued_warn_full(monkeypatch):
    """전액 발행 → '추가로 발행할 금액 없음' + 기존 계산서 있으면 수정발행 안내 (사용자 확정 문구)."""
    s = _issued_summary(monkeypatch, G3991)
    assert sb._fmt_issued_warn(s) == (
        ':white_check_mark: 총액 3,410,000원(VAT 포함) 전액 발행됨 — 추가로 발행할 금액이 없습니다.\n'
        '수정발행이 필요할 때만 아래 기존 발행 계산서를 선택하세요.')


def test_issued_warn_partial(monkeypatch):
    s = _issued_summary(monkeypatch, '2026-10-01 중도금 발행 1,705,000원 · (주)설린',
                        **{'중도금 계산서': '발행', '잔금 계산서': ''})
    assert sb._fmt_issued_warn(s) == (
        ':clipboard: 발행됨 1,705,000원 / 총액 3,410,000원 (VAT 포함)\n'
        '→ 남은 공급가액 1,550,000원이 발행 금액에 입력되었습니다. 확인 후 조정하세요.')


def test_issued_warn_uncertain(monkeypatch):
    s = _issued_summary(monkeypatch, '')        # 입금 0 발행인데 메모에 금액 없음
    assert s['uncertain']
    assert sb._fmt_issued_warn(s) == (
        ':warning: 이미 발행한 계산서가 있지만 금액을 모두 확인하지 못했습니다. 발행 금액을 직접 확인하세요.')


def test_modal_replace_checkbox_and_submit_roundtrip(monkeypatch):
    """요청 모달: 기존 계산서 있으면 '대체할 기존 계산서' 체크박스(선택) → 제출 시 스냅샷으로 복원."""
    import json as _json
    active = sb._active_invoice_lines(G3991)
    meta = _json.dumps({'code': 'G3991-YM', 'active': [[a['date'], a['stage'], int(a['amt']), a['biz']]
                                                       for a in active]}, ensure_ascii=False)
    view = sb._build_invoice_modal_view('G3991-YM', 'SM', '-', '0', '-', meta, active_invoices=active)
    blk = next(b for b in view['blocks'] if b.get('block_id') == 'replace')
    assert blk['optional'] is True and blk['element']['type'] == 'checkboxes'
    assert [o['text']['text'] for o in blk['element']['options']] == [
        '09-02 잔금 1,705,000원 · (주)설린', '09-02 잔금 1,705,000원 · SM CORPORATION']
    # 모달 없을 때(기존 계산서 없음)는 블록 자체가 없음
    assert not any(b.get('block_id') == 'replace'
                   for b in sb._build_invoice_modal_view('X', 'A', '-', '0', '-', '{}')['blocks'])
    captured = {}
    monkeypatch.setattr(sb, 'post_invoice_request', lambda **k: captured.update(k))
    monkeypatch.setattr(sb, '_slack_user_to_initial', lambda c, u: 'YM')
    values = {'biz': {'value': {'value': '(주)설린'}}, 'addr': {'value': {'value': '-'}},
              'amt': {'value': {'value': '1,550,000'}}, 'email': {'value': {'value': 'a@b.c'}},
              'memo': {'value': {'value': ''}},
              'vat': {'value': {'selected_option': {'value': 'sep'}}},
              'stages': {'value': {'selected_option': {'value': '잔금'}}},
              'replace': {'value': {'selected_options': [{'value': '1'}]}}}
    sb._process_invoice_submission(None, {'user': {'id': 'U1'}},
                                   {'private_metadata': meta, 'state': {'values': values}})
    assert captured['replace'] == [{'date': '2026-09-02', 'stage': '잔금', 'amt': 1705000,
                                    'biz': 'SM CORPORATION'}]


def test_prebill_memo_not_duplicated_for_old_dateless_line(monkeypatch):
    fm = _FakeManager()
    fm.notes['Y10'] = '잔금 선발행 2,882,000원'     # 날짜 도입 전 기록
    fm.vals['AB10'] = '미발행'
    _run(monkeypatch, fm)
    assert fm.notes['Y10'] == '잔금 선발행 2,882,000원'
