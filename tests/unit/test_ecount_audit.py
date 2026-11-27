"""이카운트 대조(2026-11-27) — 대조 결과를 넣고, 줄마다 처리 상태를 바꾼다. 임시 DB만 씀, 코드·금액은 가짜."""
import sys

sys.path.insert(0, '.')

import pytest
from flask import Flask

from dashboard.services import ecount_audit as ea


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(ea, '_db', lambda: d)
    return d


def _items():
    return [
        {'category': 'dup', 'code': 'G9001-AA', 'site': '가나시 다라로 1', 'eco_supply': 42_000_000, 'eco_vat': 4_200_000,
         'sheet_supply': 21_000_000, 'sheet_vat': 2_100_000, 'note': '8/3 전표 2장'},
        {'category': 'vat_eco', 'code': 'G9002-BB', 'eco_supply': 1_000_000, 'eco_vat': 100_000,
         'sheet_supply': 1_000_000, 'sheet_vat': 0, 'evidence': '계산서: - / - / N입금'},
        {'category': 'code', 'code': 'R9003-CC', 'ecount_code': 'G9003-CC', 'note': '입금이 R 통장'},
    ]


def test_import_and_list_groups_in_order(db):
    assert ea.import_items('2026-11-26', _items()) == 3
    assert ea.import_items('2026-11-26', _items()) == 0          # 두 번 넣어도 그대로
    out = ea.list_items()
    assert out['batch'] == '2026-11-26' and out['counts'] == {'open': 3, 'done': 0, 'total': 3}
    assert [g['key'] for g in out['groups']] == ['dup', 'code', 'vat_eco']   # 화면 순서: 중복 → 코드 → 부가세
    code = out['groups'][1]['items'][0]
    assert code['ecount_code'] == 'G9003-CC' and code['status_label'] == '처리 전'


def test_set_status_moves_done_down_and_can_undo(db):
    ea.import_items('2026-11-26', _items() + [{'category': 'dup', 'code': 'G9000-ZZ'}])
    g = ea.list_items()['groups'][0]
    first = g['items'][0]                       # G9000-ZZ (코드 순)
    ea.set_status(first['id'], 'ecount_fixed', '  한 장 지움 ', by='황샛별')
    out = ea.list_items()
    g = out['groups'][0]
    assert [i['code'] for i in g['items']] == ['G9001-AA', 'G9000-ZZ']      # 처리한 줄은 아래로
    done = g['items'][1]
    assert done['status'] == 'ecount_fixed' and done['memo'] == '한 장 지움' and done['resolved_by'] == '황샛별'
    assert out['counts'] == {'open': 3, 'done': 1, 'total': 4} and g['open'] == 1
    ea.set_status(first['id'], 'open', '', by='황샛별')                       # 되돌리기
    back = [i for i in ea.list_items()['groups'][0]['items'] if i['id'] == first['id']][0]
    assert back['status'] == 'open' and back['resolved_by'] is None and back['memo'] is None
    # 되돌려도 앞 처리·메모는 근거 칸 아래 기록으로 남음, 다시 대조 결과도 한 줄 더함(11-27)
    assert back['history'].endswith('황샛별 이카운트 고침 — 한 장 지움')
    ea.add_history(first['id'], '11/27 다시 대조 — 아직 3장')
    ea.set_status(first['id'], 'keep', '', by='황샛별')
    again = [i for i in ea.list_items()['groups'][0]['items'] if i['id'] == first['id']][0]
    assert again['history'].split('\n')[1] == '11/27 다시 대조 — 아직 3장' and again['memo'] is None
    with pytest.raises(LookupError):
        ea.add_history(999, 'x')


def test_bad_input(db):
    with pytest.raises(ValueError):
        ea.import_items('b', [{'category': 'nope', 'code': 'X'}])
    ea.import_items('b', _items())
    with pytest.raises(ValueError):
        ea.set_status(1, 'deleted', '', by='x')
    with pytest.raises(LookupError):
        ea.set_status(999, 'keep', '', by='x')


def test_empty_db(db):
    assert ea.list_items() == {'batch': None, 'sides': [], 'groups': [], 'counts': {'open': 0, 'done': 0, 'total': 0}}


def test_api(db, monkeypatch):
    import inspect
    from dashboard.blueprints import gw_finance as g
    ea.import_items('2026-11-26', _items())
    monkeypatch.setattr(g, '_by', lambda: '황샛별')
    with Flask('t').test_request_context('/api/finance/ecount/audit'):
        j = inspect.unwrap(g.api_ecount_audit)().get_json()
    assert j['success'] and j['counts']['open'] == 3
    item = j['groups'][0]['items'][0]
    with Flask('t').test_request_context(f"/api/finance/ecount/audit/{item['id']}", method='POST',
                                         json={'status': 'keep', 'memo': '나눠 냄'}):
        r = inspect.unwrap(g.api_ecount_audit_set)(item['id']).get_json()
    assert r['success'] and r['status_label'] == '그대로 둠'
    with Flask('t').test_request_context(f"/api/finance/ecount/audit/{item['id']}", method='POST', json={'status': 'bad'}):
        resp = inspect.unwrap(g.api_ecount_audit_set)(item['id'])
    assert resp[1] == 400
