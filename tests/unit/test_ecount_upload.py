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
    # 사유가 없으면 바뀐 금액으로 채움(11-26 부장님 '수정 사유 : 만 있으면 이상')
    assert s[2]['memo'] == '수정 사유 : 공급가액 4,300,000 → 4,240,000 하향'
    assert s[2]['cells'][8] == s[2]['memo'] and s[2]['cells'][4] == '-60000' and s[2]['cells'][1] == '11'
    assert s[3]['cells'][0] == '20261126' and s[3]['cells'][2] == '' and len(s[3]['warn']) == 2   # 거래처 없음·날짜 없음
    assert [p['code'] for p in out['projects']] == ['R1-AA', 'G2-BB', 'R4-DD']                  # 새 공사만
    assert out['projects'][0]['cells'] == ['R1-AA', '서울 가나구 다라로 1', '', '', '', '']


def test_cancel_memo_and_vat_only_warning():
    cancel = _item('R5-EE', '공사 취소', -2_850_000, -285_000, sent=(2_850_000, 285_000), target=(0, 0),
                   note='공사 취소 - 고객 변심')
    vat_only = _item('G6-FF', '부가세 변경', 0, -420_000, partner='거래처', sent=(4_200_000, 420_000), target=(4_200_000, 0))
    s = es.upload_rows([cancel, vat_only])['sales']
    assert s[0]['memo'] == '취소 사유 : 고객 변심'
    assert s[1]['cells'] is None and s[1]['memo'] == '수정 사유 : 부가세 별도 → 없음'   # 원래 전표 고치기(붙여넣을 줄 없음)
    assert '부가세 별도 → 없음' not in s[1]['warn']        # 적요에 있으니 확인할 점에 또 쓰지 않음
    back = _item('G7-GG', '부가세 변경', 0, 300_000, sent=(3_000_000, 0), target=(3_000_000, 300_000))
    assert es.upload_rows([back])['sales'][0]['memo'] == '수정 사유 : 부가세 없음 → 별도'
    split = _item('G9-II', '부가세 변경', 0, -161_350, sent=(3_527_000, 352_700), target=(3_527_000, 191_350))
    assert es.upload_rows([split])['sales'][0]['memo'] == '수정 사유 : 부가세 352,700 → 191,350'   # 나눠 내는 공사
    both = _item('G10-JJ', '금액 변경', -1_000_000, -300_000, sent=(4_000_000, 300_000), target=(3_000_000, 0))
    assert es.upload_rows([both])['sales'][0]['memo'] == '수정 사유 : 공급가액 4,000,000 → 3,000,000 하향, 부가세 별도 → 없음'


def test_request_reason_used_until_recorded(db):
    """금액 수정 요청의 사유가 적요로 — 장부에 적힌 뒤 들어온 사유만(이미 넣은 변경의 사유는 안 씀)."""
    es.record_entry('R8-HH', 3_200_000, 320_000, '이카운트 대조', 'R8-HH|old', by='t')
    with db._get_connection() as c:     # 장부 줄을 예전 시각으로
        c.execute("UPDATE ecount_sales_entries SET recorded_at='2026-11-01 00:00:00' WHERE project_code='R8-HH'")
        es._ensure_table(c)
        c.execute("INSERT INTO ecount_change_reasons (project_code, reason, created_at) VALUES ('R8-HH','옛 사유','2026-10-01 00:00:00')")
        c.commit()
    es.remember_reason('R8-HH', '배관  5M 추가')
    es.remember_reason('R8-HH', '배관  5M 추가')            # 같은 사유 두 번이면 한 번만
    assert es.pending_reasons(['R8-HH', 'R9-ZZ']) == {'R8-HH': '배관 5M 추가'}
    e = _item('R8-HH', '금액 변경', 96_000, 9_600, sent=(3_200_000, 320_000))
    e['reason'] = '배관 5M 추가'
    assert es.memo_of(e) == '수정 사유 : 배관 5M 추가'
    es.record_entry('R8-HH', 96_000, 9_600, '금액 변경', 'R8-HH|new', by='t')     # 넣음 표시 뒤엔 사유 안 씀
    assert es.pending_reasons(['R8-HH']) == {}


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


def test_one_won_rounding_not_pending():
    """G4160-JW(11-26): 공급가 −1·부가세 +1 — 넣을 전표 아님."""
    recs = [{'프로젝트 코드': 'G1-AA', '공사 확정': '2026-10-01', '총액 1': 10876090, '총액 2': 11963700, '부가세': True}]
    assert es.compute_pending(recs, {'G1-AA': (10876091, 1087609)}) == []


def test_trace_and_sh_flag():
    e = _item('G3288-SH', '부가세 변경', 0, -900_000, partner='거래처', sent=(9_000_000, 900_000), target=(9_000_000, 0))
    s = es.upload_rows([e, _item('R1-AA', '공사 확정', 100, 10)])['sales']
    assert s[0]['trace'] == '이카운트 9,900,000원 → 시트 9,000,000원' and 'SH 정산 공사' in s[0]['warn']
    assert s[1]['trace'] == ''


def test_vat_edit_plan_like_sb():
    """SB 규칙(11-26 이카운트 13건): 부가세 있음↔없음은 원래 전표를 고침(새 전표 없음), 금액도 바뀌면 차액만 새 줄,
    나눠 내는 공사의 일부 현금 몫은 원래 전표를 계산서 몫으로 고치고 부가세 없는 몫 한 장 추가."""
    whole = _item('G3288-SH', '부가세 변경', 0, -900_000, partner='거래처', sent=(9_000_000, 900_000), target=(9_000_000, 0))
    r = es.upload_rows([whole])['sales'][0]
    assert r['cells'] is None and '부가세를 없음으로 고쳐' in r['guide'] and '공급가액 합계 9,000,000원 그대로' in r['guide']         and '새 전표는 넣지 않아요' in r['guide']
    part = _item('G4156-JW', '부가세 변경', 0, -161_350, partner='거래처', sent=(3_527_000, 352_700), target=(3_527_000, 191_350))
    r = es.upload_rows([part])['sales'][0]
    assert (r['cells'][1], r['cells'][4], r['cells'][7]) == ('1Z', '1613500', '')
    assert '계산서 매출 합계가 공급가액 1,913,500원·부가세 191,350원이' in r['guide'] and '1,613,500원' in r['guide']
    on = _item('G4208-MJ', '부가세 변경', 0, 280_000, sent=(2_800_000, 0), target=(2_800_000, 280_000))
    r = es.upload_rows([on])['sales'][0]
    assert r['cells'] is None and '합계 280,000원이 되게' in r['guide']
    both = _item('G4083-YG', '금액 변경', 400_000, -550_000, sent=(5_500_000, 550_000), target=(5_900_000, 0))
    r = es.upload_rows([both])['sales'][0]                  # G4083: 원래 전표 vat 제외 + 40만원 새 줄
    assert (r['cells'][1], r['cells'][4], r['cells'][7]) == ('1Z', '400000', '') and '부가세를 없음으로' in r['guide']
    amt = _item('R3-CC', '금액 변경', -60_000, -6_000, sent=(4_300_000, 430_000))
    r = es.upload_rows([amt])['sales'][0]
    assert r['guide'] == '' and (r['cells'][1], r['cells'][4], r['cells'][7]) == ('11', '-60000', '-6000')


def test_mark_done_edit_only_item(db):
    whole = _item('G3288-SH', '부가세 변경', 0, -900_000, sent=(9_000_000, 900_000), target=(9_000_000, 0))
    es.record_entry('G3288-SH', 9_000_000, 900_000, '기준선', 'base', by='t')
    assert es.mark_done([whole['key']], by='gw:가나', items=[whole]) == {'done': 1, 'skipped': 0}
    assert es.ledger_totals()['G3288-SH'] == (9_000_000, 0)            # 고친 뒤 이카운트 = 시트
