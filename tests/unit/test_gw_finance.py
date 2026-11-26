# -*- coding: utf-8 -*-
"""재무 > 입출금 화면 API — 농협(N통장)은 관리하지 않는 계좌라 화면·합계에서 빠진다(11-26 부장님). 장부엔 남음."""
import inspect
import sys

sys.path.insert(0, '.')

from flask import Flask

from dashboard.blueprints import gw_finance as g
from dashboard.services import fund_daily as fd

import pytest


@pytest.fixture(autouse=True)
def _no_real_db(monkeypatch):
    """운영 users.db 를 건드리지 않게 — 결재 짝짓기·장부 조회는 가짜로."""
    monkeypatch.setattr(fd, 'transfers_done', lambda s, e: [])
    monkeypatch.setattr(g.bl, 'link_approvals', lambda s, e, a: 0)
    monkeypatch.setattr(g.bl, 'sms_start', lambda: '')
    monkeypatch.setattr(g.bl, 'linked_approval_ids', lambda: set())


def test_nonghyup_hidden_from_screen(monkeypatch):
    rows = [{'id': 1, 'acct_code': '농협', 'direction': 'out', 'amount': 900_000, 'memo': ''},
            {'id': 2, 'acct_code': '1011', 'direction': 'in', 'amount': 5_000, 'memo': ''}]
    monkeypatch.setattr(g.bl, 'list_tx', lambda *a, **k: list(rows))
    monkeypatch.setattr(g.bl, 'summary', lambda day: [{'code': '1011'}, {'code': '농협'}, {'code': '하도급'}])
    view = inspect.unwrap(g.api_list)          # 로그인·권한 데코레이터 없이 본문만
    with Flask('t').test_request_context('/api/finance/bank'):
        j = view().get_json()
    assert [r['acct_code'] for r in j['rows']] == ['1011']
    assert j['total'] == {'in': 5_000, 'out': 0, 'unassigned': 1, 'no_approval': 0, 'no_out': 0}   # 농협 출금은 합계에도 안 들어감
    assert [a['code'] for a in j['accounts']] == ['1011', '하도급']


def test_report_balance_rows_hidden_unless_bal_filter(monkeypatch):
    rows = [{'id': 1, 'acct_code': '1012', 'direction': 'bal', 'amount': 0, 'memo': '', 'source': 'report'},
            {'id': 2, 'acct_code': '1013', 'direction': 'bal', 'amount': 0, 'memo': '', 'source': 'manual'}]
    monkeypatch.setattr(g.bl, 'list_tx', lambda *a, **k: list(rows))
    monkeypatch.setattr(g.bl, 'summary', lambda day: [])
    view = inspect.unwrap(g.api_list)
    with Flask('t').test_request_context('/api/finance/bank'):
        assert [r['id'] for r in view().get_json()['rows']] == [2]          # 자금일보 잔액 줄은 거래 내역에서 숨김
    with Flask('t').test_request_context('/api/finance/bank?direction=bal'):
        assert [r['id'] for r in view().get_json()['rows']] == [1, 2]



def test_approval_pairs_and_transfers_without_sms(monkeypatch):
    """출금 줄에 결재 짝 요약, 짝 없는 문자 출금은 '결재 없음', 이체 완료인데 출금 문자 없는 G·R 결재는 따로(N 은 뺌)."""
    rows = [{'id': 1, 'acct_code': '1011', 'direction': 'out', 'amount': 300_000, 'memo': 'x', 'source': 'sms', 'approval_id': 84},
            {'id': 2, 'acct_code': '9012', 'direction': 'out', 'amount': 50_000, 'memo': '', 'source': 'sms', 'approval_id': None},
            {'id': 3, 'acct_code': '1011', 'direction': 'in', 'amount': 1_000, 'memo': '', 'source': 'sms', 'approval_id': None}]
    appr = [{'id': 84, 'form': 'fund', 'amount': 300_000, 'grn': 'G', 'partner': '최만호', 'content': '앵글', 'done_at': '2026-11-26 10:00:00'},
            {'id': 90, 'form': 'fund', 'amount': 700_000, 'grn': 'R', 'partner': '세광', 'content': '자재', 'done_at': '2026-11-26 11:00:00'},
            {'id': 91, 'form': 'fund', 'amount': 900_000, 'grn': 'N', 'partner': '일당', 'content': '', 'done_at': '2026-11-26 11:00:00'},
            {'id': 92, 'form': 'fund', 'amount': 100_000, 'grn': 'G', 'partner': '옛날', 'content': '', 'done_at': '2026-11-25 23:00:00'}]
    monkeypatch.setattr(g.bl, 'list_tx', lambda *a, **k: [dict(r) for r in rows])
    monkeypatch.setattr(g.bl, 'summary', lambda day: [])
    monkeypatch.setattr(fd, 'transfers_done', lambda s, e: appr)
    monkeypatch.setattr(g.bl, 'sms_start', lambda: '2026-11-26 02:30')
    monkeypatch.setattr(g.bl, 'linked_approval_ids', lambda: {84})
    view = inspect.unwrap(g.api_list)
    with Flask('t').test_request_context('/api/finance/bank?start=2026-11-20&end=2026-11-26'):
        j = view().get_json()
    by = {r['id']: r for r in j['rows']}
    assert by[1]['approval']['id'] == 84 and by[1]['approval']['form'] == '자금집행' and not by[1]['no_approval']
    assert by[2]['no_approval'] and not by[3]['no_approval']                 # 입금은 대조 대상 아님
    assert [a['id'] for a in j['no_out']] == [90]                           # N(91)·장부 시작 전(92)은 뺌
    assert j['total']['no_approval'] == 1 and j['total']['no_out'] == 1
    with Flask('t').test_request_context('/api/finance/bank?start=2026-11-20&end=2026-11-26&appr=none'):
        assert [r['id'] for r in view().get_json()['rows']] == [2]
