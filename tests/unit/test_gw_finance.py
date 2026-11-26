# -*- coding: utf-8 -*-
"""재무 > 입출금 화면 API — 농협(N통장)은 관리하지 않는 계좌라 화면·합계에서 빠진다(11-26 부장님). 장부엔 남음."""
import inspect
import sys

sys.path.insert(0, '.')

from flask import Flask

from dashboard.blueprints import gw_finance as g


def test_nonghyup_hidden_from_screen(monkeypatch):
    rows = [{'id': 1, 'acct_code': '농협', 'direction': 'out', 'amount': 900_000, 'memo': ''},
            {'id': 2, 'acct_code': '1011', 'direction': 'in', 'amount': 5_000, 'memo': ''}]
    monkeypatch.setattr(g.bl, 'list_tx', lambda *a, **k: list(rows))
    monkeypatch.setattr(g.bl, 'summary', lambda day: [{'code': '1011'}, {'code': '농협'}, {'code': '하도급'}])
    view = inspect.unwrap(g.api_list)          # 로그인·권한 데코레이터 없이 본문만
    with Flask('t').test_request_context('/api/finance/bank'):
        j = view().get_json()
    assert [r['acct_code'] for r in j['rows']] == ['1011']
    assert j['total'] == {'in': 5_000, 'out': 0, 'unassigned': 1}         # 농협 출금은 합계·적요 없음 건수에도 안 들어감
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
