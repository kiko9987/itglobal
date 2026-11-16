# -*- coding: utf-8 -*-
"""수금 메모 안전 (2026-11-16 상태 머신 점검 묶음 A).

① 메모를 '읽지 못함'과 '메모 없음'을 가른다 — 못 읽은 채 새 블록만 쓰면 입금 이력이 통째로 사라졌다.
② 메모 쓰기가 실패하면 금액도 쓰지 않는다 — 값만 있고 누가 낸 돈인지 없는 칸이 생겼다.
③ 폴러: 메모가 덜 적힌 행이 한 회차 5자리를 계속 차지해 뒤쪽 공사 입금 카드가 멈추지 않게.
④ 폴러: 행이 몇 줄만 밀려도(20행 이하) 옆 공사 기준값과 비교하지 않게.
"""
import sys
import time

sys.path.insert(0, '.')

import pytest

from dashboard.services import payment_sync as ps
from dashboard.utils import google_sheets as gs
from dashboard.utils.google_sheets import NoteReadError, read_note_for_append

DEP = '2026/10/03\n하나,10/03, 11:20\n255******31304\n입금200,000원\n김보라(오라라(O'
OUT_TEXT = '[Web발신]\n하나,10/08, 14:30\n255******31304\n출금200,000원\n김보라(오라라(O'

pytestmark = pytest.mark.usefixtures('sheet_layout_45')


# ── ① 읽기 실패 구분 ──

class _Plain:
    def get_cell_note(self, *a):
        return None


class _Checked(_Plain):
    def __init__(self, note=None, fail=False):
        self.note, self.fail = note, fail

    def get_cell_note_checked(self, *a):
        if self.fail:
            raise NoteReadError('502')
        return self.note


def test_read_note_for_append():
    assert read_note_for_append(_Checked('메모'), 's', 'n', 'U2') == '메모'
    assert read_note_for_append(_Checked(None), 's', 'n', 'U2') == ''          # 메모 없음
    with pytest.raises(NoteReadError):
        read_note_for_append(_Checked(fail=True), 's', 'n', 'U2')               # 못 읽음
    assert read_note_for_append(_Plain(), 's', 'n', 'U2') == ''                 # 예전 매니저(가짜 등)


def _bare_manager(execute):
    import threading
    m = object.__new__(gs.GoogleSheetsManager)
    m._lock = threading.RLock()
    m._execute_with_retry = execute
    return m


def test_get_cell_note_checked_raises_instead_of_none():
    def boom(factory, name):
        raise RuntimeError('HttpError 502')
    with pytest.raises(NoteReadError):
        _bare_manager(boom).get_cell_note_checked('s', 'n', 'U2')
    found = {'sheets': [{'data': [{'rowData': [{'values': [{'note': '입금 메모'}]}]}]}]}
    assert _bare_manager(lambda f, n: found).get_cell_note_checked('s', 'n', 'U2') == '입금 메모'
    assert _bare_manager(lambda f, n: {'sheets': [{'data': [{}]}]}).get_cell_note_checked('s', 'n', 'U2') is None


# ── ② 입금·반환 기록: 메모 못 읽음 / 못 씀 → 금액도 안 씀 ──

class _Mgr:
    F2L = {'계약금': 'U', '중도금': 'V', '잔금': 'W', '계산서': 'Y', '총액 2': 'T', '수금 확인': 'AD',
           '계약금 계산서': 'Z', '중도금 계산서': 'AA', '잔금 계산서': 'AB'}

    def __init__(self, read_fail=False, write_result=True, write_really=True):
        self.cells = {'U4164': 200000, 'V4164': 0, 'W4164': 2500000, 'T4164': 2500000,
                      'Y4164': '계약금 - 미발행', 'Z4164': '미발행', 'AA4164': '-', 'AB4164': 'N입금',
                      'AD4164': 'FALSE'}
        self.notes = {'U4164': DEP}
        self.read_fail, self.write_result, self.write_really = read_fail, write_result, write_really

    def find_row_by_project_code(self, *a):
        return 4164

    def get_cell_value(self, sid, sn, cell):
        return self.cells.get(cell, '')

    def update_cell_value(self, sid, sn, cell, v):
        self.cells[cell] = v
        return True

    def get_cell_note(self, sid, sn, cell):
        return None if self.read_fail else self.notes.get(cell, '')

    def get_cell_note_checked(self, sid, sn, cell):
        if self.read_fail:
            raise NoteReadError('502')
        return self.notes.get(cell)

    def update_cell_note(self, sid, sn, cell, note):
        if self.write_really:
            self.notes[cell] = note
        return self.write_result

    def get_field_to_letter(self):
        return dict(self.F2L)


def _commit(monkeypatch, mgr):
    import dashboard.blueprints.slack_bot as sb
    import dashboard.services.lead_service as ls
    import dashboard.utils.user_database as udb
    monkeypatch.setenv('GOOGLE_SHEET_ID', 'sid')
    monkeypatch.setenv('GOOGLE_SHEET_NAME', '공사 현황')
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: mgr)
    monkeypatch.setattr(udb, 'get_audit_repository', lambda: type('A', (), {'log_action': lambda *a, **k: None})())
    return sb._commit_intake_to_sheet('R4163-TH', '계약금', -200000, OUT_TEXT, 'U_SB', refund=True)


def test_commit_stops_when_note_unreadable(monkeypatch):
    mgr = _Mgr(read_fail=True)
    ok, old, new, err = _commit(monkeypatch, mgr)
    assert not ok and (old, new) == (200000, 200000) and '메모를 읽지 못해' in err
    assert mgr.cells['U4164'] == 200000 and mgr.notes['U4164'] == DEP        # 아무것도 안 바뀜


def test_commit_stops_when_note_write_fails(monkeypatch):
    mgr = _Mgr(write_result=False, write_really=False)
    ok, old, new, err = _commit(monkeypatch, mgr)
    assert not ok and '금액도 기록하지 않았습니다' in err
    assert mgr.cells['U4164'] == 200000 and mgr.notes['U4164'] == DEP


def test_commit_continues_when_write_reported_failure_but_landed(monkeypatch):
    """쓰기 응답만 실패하고 실제로는 써진 경우 — 다시 읽어 확인하고 이어서 금액 기록."""
    mgr = _Mgr(write_result=False, write_really=True)
    ok, old, new, err = _commit(monkeypatch, mgr)
    assert ok and (old, new) == (200000, 0), err
    assert mgr.cells['U4164'] == 0 and '출금200,000원' in mgr.notes['U4164']


def test_commit_normal_path_unchanged(monkeypatch):
    mgr = _Mgr()
    ok, old, new, err = _commit(monkeypatch, mgr)
    assert ok and mgr.notes['U4164'].startswith(DEP + '\n\n')


def test_apply_transfer_does_not_write_when_note_unreadable(monkeypatch):
    from dashboard.services import fund_transfer as ft
    import dashboard.services.lead_service as ls
    mgr = _Mgr(read_fail=True)
    monkeypatch.setenv('GOOGLE_SHEET_ID', 'sid')
    monkeypatch.setenv('GOOGLE_SHEET_NAME', '공사 현황')
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: mgr)
    req = {'code': 'G4166-YM', 'stage': '계약금', 'col': 'U', 'from': 'R', 'to': 'G', 'amount': 200000,
           'partner': '김보라', 'bank': '하나', 'date': '10/03', 'req_line': '⚠️ … 자금 이동 요청'}
    pv = {'amount': 200000, 'partner': '김보라', 'bank': '기업', 'acct_code': 'G', 'date_md': '10/04'}
    with pytest.raises(NoteReadError):
        ft.apply_transfer(req, '입금 200,000원 김보라', pv, 'SB')
    assert mgr.notes['U4164'] == DEP


# ── ③·④ 폴러 ──

class _RC(dict):
    def get(self, k):
        return dict.get(self, k)

    def set(self, k, v, nx=False, ex=None):
        if nx and k in self:
            return False
        self[k] = v
        return True

    def delete(self, *ks):
        for k in ks:
            self.pop(k, None)

    def exists(self, k):
        return int(k in self)

    def pipeline(self, transaction=False):
        rc = self

        class _P:
            def __init__(self):
                self.q = []

            def exists(self, k):
                self.q.append(rc.exists(k))

            def execute(self):
                return self.q
        return _P()


def test_incomplete_rows_go_last_only_when_over_limit():
    rc = _RC()
    rows = [{'row': r} for r in range(2, 10)]
    rc[f'{ps._INCOMPLETE_KEY}2'] = '1|A'
    rc[f'{ps._INCOMPLETE_KEY}3'] = '1|B'
    assert [c['row'] for c in ps._incomplete_last(rc, rows, 5)] == [4, 5, 6, 7, 8, 9, 2, 3]
    assert ps._incomplete_last(rc, rows[:5], 5) == rows[:5]                       # 자리 남으면 그대로


def test_incomplete_row_alerts_once_after_6_hours(caplog):
    rc = _RC()
    ps._note_incomplete_row(rc, 7, 'G4100-AA')
    assert rc[f'{ps._INCOMPLETE_KEY}7'].endswith('|G4100-AA')
    rc[f'{ps._INCOMPLETE_KEY}7'] = f'{int(time.time()) - 7 * 3600}|G4100-AA'
    with caplog.at_level('ERROR'):
        ps._note_incomplete_row(rc, 7, 'G4100-AA')
        ps._note_incomplete_row(rc, 7, 'G4100-AA')
    assert sum('6시간 넘게' in r.getMessage() for r in caplog.records) == 1         # 하루 한 번
    ps._note_incomplete_row(rc, 7, 'G4200-BB')                                      # 행이 다른 공사로 → 새로 셈
    assert rc[f'{ps._INCOMPLETE_KEY}7'].endswith('|G4200-BB')
    ps._clear_incomplete_row(rc, 7)
    assert f'{ps._INCOMPLETE_KEY}7' not in rc


# 실제 폴러 한 회차(_sync_payments_locked) — test_payment_sync_mass_change 와 같은 가짜 시트·슬랙, Redis db15
_MEMO = '2023-03-14\n22,200,000원\n김수정\n플렌트입금'
_FIRST = ('G8999-TT', 22_200_000)
# 처리될 행 메모는 다른 테스트와 겹치지 않게 — 같은 입금(금액·입금자·날짜)이면 앞 카드에 통합돼 '발송'으로 안 셈
_MEMO_DONE = '2026-11-10\n1,000,000원\n홍길순테스트\n기업입금'


class _Exec:
    def __init__(self, v):
        self.v = v

    def execute(self):
        return self.v


class _FakeService:
    def __init__(self, rows, notes):
        self.rows, self.notes = rows, notes

    def spreadsheets(self):
        return self

    def values(self):
        return self

    def get(self, **kw):
        if 'includeGridData' in kw:
            return _Exec({'sheets': [{'data': [{'rowData': [
                {'values': [{'note': n} for n in r]} for r in self.notes]}]}]})
        return _Exec({'values': self.rows})


class _FakeSlack:
    def __init__(self, *a, **k):
        pass

    def __getattr__(self, name):
        return lambda *a, **k: {'ok': True, 'ts': '1.000001', 'channel': 'C1', 'messages': []}


@pytest.fixture
def run(monkeypatch):
    from dashboard.utils.redis_client import get_redis_client
    rc = get_redis_client().redis

    def _clear():
        for pat in ('payment_sync:*', 'payment_pending_group:*', 'payment_grace_ref:*', 'payment_slack:ts:G9*'):
            for k in rc.scan_iter(pat):
                rc.delete(k)
    _clear()
    monkeypatch.setattr(ps, '_poll_header_ok', lambda *a, **k: True)
    monkeypatch.setattr('slack_sdk.WebClient', _FakeSlack)
    I = ps._sheet_idx()
    W = ps._poll_width(I)
    memo_rows = set()
    monkeypatch.setattr(ps, '_fetch_row_notes',
                        lambda sid, sn, row, *a, **k: [_MEMO_DONE, '', ''] if row in memo_rows else ['', '', ''])

    def _run(pairs, with_memo=()):
        memo_rows.clear()
        memo_rows.update(with_memo)
        rows, notes = [], []
        for i, (code, u) in enumerate([_FIRST] + pairs):
            r = [''] * W
            r[I['프로젝트 코드']], r[I['총액 2']], r[I['계약금']] = code, 50_000_000, u
            rows.append(r)
            notes.append([_MEMO if i == 0 else (_MEMO_DONE if i + 2 in memo_rows else ''), '', ''])
        monkeypatch.setattr(ps, '_get_payment_service', lambda: _FakeService(rows, notes))
        return ps._sync_payments_locked({'processed': 0, 'sent': 0, 'errors': 0}, 'sid', '공사 현황', 'C1', 'x')
    _run.rc = rc
    yield _run
    _clear()


def test_small_row_shift_is_not_compared(run):
    """위쪽에 한 줄이 끼어 아래 3행이 한 칸씩 밀림 — 20행 이하라도 옆 공사 기준값과 비교하지 않음."""
    pairs = [('G9001-TT', 3_000_000), ('G9002-TT', 2_000_000), ('G9003-TT', 1_000_000)]
    run(pairs)
    shifted = [('G7777-TT', 5_000_000)] + pairs
    assert run(shifted)['processed'] == 0
    assert run.rc.hget(f'{ps.REDIS_KEY_PREFIX}4', 'p') == 'G9001-TT'               # 기준값은 새 자리로
    assert run.rc.hget(f'{ps.REDIS_KEY_PREFIX}4', 'u') == '3000000'
    assert run(shifted)['processed'] == 0


def test_incomplete_memo_rows_do_not_block_others(run):
    """메모가 덜 적힌 입금 5행이 앞에 있어도, 다음 회차엔 뒤쪽 공사의 완성된 입금이 먼저 처리된다."""
    pairs = [(f'G{9100 + i}-TT', 0) for i in range(6)]
    run(pairs)                                                         # 기준값
    paid = [(c, 1_000_000) for c, _ in pairs]
    first = run(paid, with_memo={8})                                   # 8행 = 6번째(G9105) 만 메모 완성
    assert first['processed'] == 5 and first['sent'] == 0               # 앞 5행(미완성)이 자리 차지
    second = run(paid, with_memo={8})
    assert second['sent'] >= 1                                         # 이번엔 G9105 가 앞으로
    assert run.rc.hget(f'{ps.REDIS_KEY_PREFIX}8', 'u') == '1000000'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
