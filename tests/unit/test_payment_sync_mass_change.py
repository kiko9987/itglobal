# -*- coding: utf-8 -*-
"""입금 폴러 '변경 과다' 안전장치 (2026-10-30).

기준값(payment_sync:row:{N})은 행 번호로 저장돼, 행 삽입·정렬로 밀리면 수십 행이 한꺼번에 '변경'으로 보인다.
그래서 21행 이상이면 기준값만 다시 만들고 발송을 건너뛰는데, 경영지원이 입금 20여 건을 한 번에 붙여 넣은
진짜 변경도 같이 묻혔다. 기준값에 그 행의 공사 코드(p)를 함께 두어 코드가 그대로인 행은 발송한다.
"""
import sys
sys.path.insert(0, '.')

import pytest

from dashboard.services import payment_sync as ps
from dashboard.utils.redis_client import get_redis_client

N = 25
_MEMO = '2023-03-14\n22,200,000원\n김수정\n플렌트입금'
_FIRST = ('G8999-TT', 22_200_000)       # 메모 있는 행 — 메모 0행 응답 안전장치에 안 걸리게


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
    rc = get_redis_client().redis

    def _clear():
        for k in rc.scan_iter('payment_sync:*'):
            rc.delete(k)
    _clear()
    monkeypatch.setattr(ps, '_poll_header_ok', lambda *a, **k: True)
    monkeypatch.setattr('slack_sdk.WebClient', _FakeSlack)
    monkeypatch.setattr(ps, '_fetch_row_notes', lambda *a, **k: ['', '', ''])
    I = ps._sheet_idx()
    W = ps._poll_width(I)

    def _run(pairs):
        rows = []
        for code, u in [_FIRST] + pairs:
            r = [''] * W
            r[I['프로젝트 코드']], r[I['총액 2']], r[I['계약금']] = code, 50_000_000, u
            rows.append(r)
        notes = [[_MEMO, '', '']] + [['', '', ''] for _ in pairs]
        monkeypatch.setattr(ps, '_get_payment_service', lambda: _FakeService(rows, notes))
        return ps._sync_payments_locked({'processed': 0, 'sent': 0, 'errors': 0}, 'sid', '공사 현황', 'C1', 'x')
    _run.rc = rc
    yield _run
    _clear()


def _pairs(u_of):
    return [(f'G{9000 + i}-TT', u_of(i)) for i in range(N)]


def test_baseline_stores_project_code(run):
    run(_pairs(lambda i: 0))
    assert run.rc.hget(f'{ps.REDIS_KEY_PREFIX}3', 'p') == 'G9000-TT'


def test_bulk_real_deposits_still_processed(run):
    """같은 공사 행 25개가 한꺼번에 입금 — 예전엔 0행 처리(전부 재구축), 이제 한 폴링 5행씩 발송 진행."""
    run(_pairs(lambda i: 0))
    res = run(_pairs(lambda i: 1_000_000))
    assert res['processed'] == 5


def test_row_shift_rebaselines_without_sending(run):
    """위쪽에 행이 끼어 아래 행이 전부 한 칸씩 밀림 — 코드가 다 바뀌었으니 발송 없이 기준값만 재구축."""
    u_of = lambda i: 1_000 * (100 - i)          # 아래로 갈수록 작음 → 밀리면 각 행 값이 '증가'로 보임
    run(_pairs(u_of))
    shifted = [('G7777-TT', 5_000_000)] + _pairs(u_of)
    res = run(shifted)
    assert res['processed'] == 0
    assert run.rc.hget(f'{ps.REDIS_KEY_PREFIX}4', 'p') == 'G9000-TT'   # 기준값은 새 자리로
    assert run(shifted)['processed'] == 0                             # 다음 폴링도 조용


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
