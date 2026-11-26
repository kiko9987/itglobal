"""재무 > 이카운트 올리기(2026-11-26) — 넣을 전표를 이카운트 웹자료올리기 칸 순서로, [넣음 표시]는 장부에. 이름·코드는 가짜."""
import inspect
import sys
from datetime import date

sys.path.insert(0, '.')

import pytest
from flask import Flask

from dashboard.blueprints import gw_finance as g
from dashboard.services import ecount_sales as es


def _item(code, kind, supply, vat, partner='홈페이지', date_='2026-11-25', addr='서울 가나구 다라로 1', sent=(0, 0),
          target=None, note='', split=False):
    target = target or (sent[0] + supply, sent[1] + vat)
    return {'code': code, 'kind': kind, 'supply': supply, 'vat': vat, 'sent': sent, 'target': target, 'date': date_,
            'partner': partner, 'addr': addr, 'note': note, 'split': split, 'key': es.entry_key(code, sent, target)}


def test_upload_rows_cells_in_ecount_order():
    items = [_item('R1-AA', '공사 확정', 2_200_000, 220_000),
             _item('G2-BB', '공사 확정', 3_000_000, 0, partner='거래처'),                  # 부가세 없음 → 1Z
             _item('R3-CC', '금액 변경', -60_000, -6_000, partner='전화', sent=(4_300_000, 430_000)),
             _item('R4-DD', '공사 확정', 1_000_000, 100_000, partner='중고판매', date_='')]
    out = es.upload_rows(items, today=date(2026, 11, 26))
    s = out['sales']
    assert len(s[0]['cells']) == len(es.SALES_COLS)
    assert s[0]['cells'] == ['20261125', '11', '99655', '홈페이지', '2200000', '', '', '220000', '서울 가나구 다라로 1',
                             '4049', '1089', '', 'R1-AA']
    assert s[1]['cells'][1] == '1Z' and s[1]['cells'][7] == '' and s[1]['cells'][2] == '99237'
    assert s[2]['memo'] == '수정 사유 : ' and s[2]['cells'][4] == '-60000' and s[2]['cells'][1] == '11'
    assert s[3]['cells'][0] == '20261126' and s[3]['cells'][2] == '' and len(s[3]['warn']) == 2   # 거래처 없음·날짜 없음
    assert [p['code'] for p in out['projects']] == ['R1-AA', 'G2-BB', 'R4-DD']                  # 새 공사만
    assert out['projects'][0]['cells'] == ['R1-AA', '서울 가나구 다라로 1', '', '', '', '']


def test_cancel_memo_and_vat_only_warning():
    cancel = _item('R5-EE', '공사 취소', -2_850_000, -285_000, sent=(2_850_000, 285_000), target=(0, 0),
                   note='공사 취소 - 고객 변심')
    vat_only = _item('G6-FF', '부가세 변경', 0, -420_000, partner='거래처', sent=(4_200_000, 420_000), target=(4_200_000, 0))
    s = es.upload_rows([cancel, vat_only])['sales']
    assert s[0]['memo'] == '취소 사유 : 고객 변심'
    assert s[1]['cells'][1] == '11' and '부가세 별도 → 없음' in s[1]['warn']   # 세금계산서 부가세를 빼는 전표


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(es, '_db', lambda: d)
    monkeypatch.setattr(es, '_table_ready', False)
    return d


def test_mark_done_records_only_current_items(db):
    items = [_item('R1-AA', '공사 확정', 2_200_000, 220_000)]
    res = es.mark_done([items[0]['key'], 'R9-ZZ|0,0>1,1'], by='gw:가나', items=items)
    assert res == {'done': 1, 'skipped': 1}
    assert es.ledger_totals()['R1-AA'] == (2_200_000, 220_000)
    assert es.mark_done([items[0]['key']], by='gw:가나', items=items) == {'done': 0, 'skipped': 1}   # 두 번 눌러도 한 번만


def test_api_ecount_list_and_done(monkeypatch, db):
    items = [_item('R1-AA', '공사 확정', 2_200_000, 220_000)]
    monkeypatch.setattr(es, 'collect_items', lambda force_refresh=False: list(items))
    with Flask('t').test_request_context('/api/finance/ecount'):
        j = inspect.unwrap(g.api_ecount)().get_json()
    assert j['success'] and j['sales'][0]['code'] == 'R1-AA' and j['sales_cols'][0] == '일자'
    with Flask('t').test_request_context('/api/finance/ecount/done', method='POST', json={'keys': [items[0]['key']]}):
        j = inspect.unwrap(g.api_ecount_done)().get_json()
    assert j['done'] == 1
    monkeypatch.setattr(es, 'collect_items', lambda force_refresh=False: None)
    with Flask('t').test_request_context('/api/finance/ecount'):
        resp = inspect.unwrap(g.api_ecount)()
    assert resp[1] == 503


def test_exclude_records_as_excluded(db):
    items = [_item('G7-GG', '부가세 변경', 0, -420_000, partner='거래처', sent=(4_200_000, 420_000), target=(4_200_000, 0))]
    assert es.mark_done([items[0]['key']], by='gw:가나', items=items, exclude=True) == {'done': 1, 'skipped': 0}
    assert es.ledger_totals()['G7-GG'] == (0, -420_000)                    # 차액이 장부에 → 다음 목록에서 빠짐
    with es._db()._get_connection() as conn:
        assert conn.execute("SELECT kind, note FROM ecount_sales_entries WHERE project_code='G7-GG' AND kind='제외'").fetchone()[1] \
            == '제외 — 원래 부가세 변경'
