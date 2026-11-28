"""이카운트 밤 자동 대조(2026-11-28) — 장부는 이카운트 = 시트일 때만 맞추고, ③ 대조는 매출·매입으로 나눠
지난밤 알린 건을 자동으로 완료/되돌림. 임시 DB 만 씀, 코드·금액·주소는 가짜."""
import inspect
import json
import sys
from datetime import date

sys.path.insert(0, '.')

import pytest
from flask import Flask

from dashboard.services import ecount_audit as ea
from dashboard.services import ecount_recon as er
from dashboard.services import ecount_sales as es


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for m in (ea, er, es):
        monkeypatch.setattr(m, '_db', lambda: d)
    monkeypatch.setattr(es, '_table_ready', False)
    monkeypatch.setattr(er, 'OUT_DIR', tmp_path / 'recon')
    return d


def _rec(code, t1, vat=True, confirmed='2026-10-01', cost=None, **kw):
    r = {'프로젝트 코드': code, '총액 1': t1, '총액 2': round(t1 * 1.1) if vat else t1, '부가세': vat,
         '공사 확정': confirmed, '유입 구분': '거래처', '현장 주소': '가나시 다라로 1'}
    for k, v in (cost or {}).items():
        r[k] = v
    r.update(kw)
    return r


def _row(d, code, s, v, typ='매출전표 I', partner='거래처'):
    return {'d': d, 'code': code, 'typ': typ, 's': s, 'v': v, 'amt': s + v, 'partner': partner, 'site': '가나시 다라로 1'}


# ── 합산·장부 맞춤 ─────────────────────────────────────────────────────

def test_aggregate_sums_by_code_and_skips_sales_in_buy():
    rows = [_row('2026/10/02 -1', 'G9001-AA', 1_000_000, 100_000), _row('2026/10/09 -4', 'G9001-AA', -200_000, -20_000),
            _row('2026/10/03 -2', 'G9001-AA', 500_000, 0, typ='매입전표 I'),
            _row('2026/10/04 -1', 'G9001-AA', 9_000_000, 0, typ='입금보고서')]
    a = er.aggregate(rows)
    assert (a['G9001-AA']['s'], a['G9001-AA']['v'], a['G9001-AA']['n'], a['G9001-AA']['first']) == (10_300_000, 80_000, 4, '2026-10-02')
    b = er.aggregate(rows, buy=True)
    assert b['G9001-AA']['s'] == 500_000 and b['G9001-AA']['n'] == 1


def test_ledger_fixed_only_when_ecount_equals_sheet():
    recs = {
        'G9001-AA': _rec('G9001-AA', 1_000_000),            # 이카운트 = 시트, 장부 없음([넣음 표시] 없이 넣음) → 맞춤
        'G9002-BB': _rec('G9002-BB', 2_000_000),            # 이카운트 ≠ 시트 → 장부 그대로
        'G9003-CC': _rec('G9003-CC', 3_000_000),            # 셋 다 같음 → 할 일 없음
        'G0004-DD': _rec('G0004-DD', 0, confirmed=''),      # 2025년 전 공사의 취소 전표(−)만 보임 → 범위 밖
        'G9005-EE': _rec('G9005-EE', 4_000_000, confirmed='2024-12-20'),   # 2025-02 전 확정 → 범위 밖
    }
    eco = er.aggregate([_row('2026/10/02 -1', 'G9001-AA', 1_000_000, 100_000),
                        _row('2026/10/02 -2', 'G9002-BB', 1_500_000, 150_000),
                        _row('2026/10/02 -3', 'G9003-CC', 3_000_000, 300_000),
                        _row('2025/08/01 -1', 'G0004-DD', -800_000, 0),
                        _row('2025/01/05 -1', 'G9005-EE', 1_000, 0)])
    ledger = {'G9002-BB': (2_000_000, 200_000), 'G9003-CC': (3_000_000, 300_000), 'G0004-DD': (0, 0),
              'G9005-EE': (4_000_000, 400_000)}
    fixes = er.ledger_fixes(eco, recs, ledger)
    assert [(f['code'], f['supply'], f['vat']) for f in fixes] == [('G9001-AA', 1_000_000, 100_000)]


# ── 매출 대조 ──────────────────────────────────────────────────────────

def test_sales_diffs_categories_and_pending_skip():
    recs = {
        'G9001-AA': _rec('G9001-AA', 1_000_000),                   # 이카운트 = 장부 ≠ 시트 → ② 대기, ③ 아님
        'G9002-BB': _rec('G9002-BB', 1_000_000, vat=False),        # 부가세만 이카운트에 있음 → vat_eco
        'G9003-CC': _rec('G9003-CC', 2_000_000),                   # 같은 날 같은 금액 두 장 → dup
        'R9004-DD': _rec('R9004-DD', 500_000),                     # 이카운트엔 G9004-DD → code
        'R9006-FF': _rec('R9006-FF', 300_000),                     # 같은 번호 옛 코드가 넣었다 취소(합 0) → 짝 아님, 맞음
    }
    eco = er.aggregate([
        _row('2026/10/02 -1', 'G9001-AA', 800_000, 80_000),
        _row('2026/10/02 -2', 'G9002-BB', 1_000_000, 100_000),
        _row('2026/10/03 -1', 'G9003-CC', 2_000_000, 200_000), _row('2026/10/03 -5', 'G9003-CC', 2_000_000, 200_000),
        _row('2026/10/04 -1', 'G9004-DD', 500_000, 50_000),
        _row('2026/11/12 -1', 'R9006-ZZ', 39_000_000, 0), _row('2026/11/13 -1', 'R9006-ZZ', -39_000_000, 0),
        _row('2026/11/13 -2', 'R9006-FF', 300_000, 30_000),
    ])
    ledger = {'G9001-AA': (800_000, 80_000), 'G9002-BB': (1_000_000, 0), 'G9003-CC': (2_000_000, 200_000),
              'R9006-FF': (300_000, 30_000)}
    found, mismatch = er.sales_diffs(eco, recs, ledger)
    assert {c: i['category'] for c, i in found.items()} == {'G9002-BB': 'vat_eco', 'G9003-CC': 'dup', 'R9004-DD': 'code'}
    assert found['R9004-DD']['ecount_code'] == 'G9004-DD' and 'G9004-DD' in found['R9004-DD']['note']
    assert '이카운트 공급가 전표: 26.10.03 2,000,000 · 26.10.03 2,000,000' in found['G9003-CC']['note']
    assert mismatch == {'G9002-BB': [1_000_000, 100_000], 'G9003-CC': [4_000_000, 400_000]}
    # 이미 ③에서 보고 있는 공사는 ② 대기여도 계속 보임(근거와 함께)
    found2, _ = er.sales_diffs(eco, recs, ledger, watch={'G9001-AA'})
    assert found2['G9001-AA']['category'] == 'supply' and '② 매출전표 목록에도 있어요' in found2['G9001-AA']['note']


# ── 매입 대조 ──────────────────────────────────────────────────────────

def test_buy_diffs_grace_and_categories():
    cost = lambda a: {'제품대': a, '도급비': 0, '자재비': 0, '기타비': 0}
    recs = {
        'G9101-AA': _rec('G9101-AA', 5_000_000, confirmed='2026-08-01', cost=cost(2_000_000)),   # 같음
        'G9102-BB': _rec('G9102-BB', 5_000_000, confirmed='2026-08-01', cost=cost(2_000_000)),   # 다름
        'G9103-CC': _rec('G9103-CC', 5_000_000, confirmed='2026-08-01'),                         # 원가 빈칸
        'G9104-DD': _rec('G9104-DD', 5_000_000, confirmed='2026-11-20', cost=cost(9)),           # 확정 45일 전 → 아직
        'G9105-EE': _rec('G9105-EE', 0, confirmed='', **{'수금 관련 특이사항': '공사 취소 (사유: 고객변심)'}),
        'R9106-FF': _rec('R9106-FF', 1_000_000, confirmed='2026-08-01', cost=cost(80_000)),
    }
    rows = [_row('2026/08/10 -1', 'G9101-AA', 2_000_000, 200_000, typ='매입전표 I', partner='가나공조'),
            _row('2026/08/10 -2', 'G9102-BB', 2_500_000, 250_000, typ='매입전표 I', partner='가나공조'),
            _row('2026/08/10 -3', 'G9103-CC', 700_000, 70_000, typ='매입전표 I'),
            _row('2026/11/21 -1', 'G9104-DD', 100, 10, typ='매입전표 I'),
            _row('2026/09/01 -1', 'G9105-EE', 300_000, 0, typ='지출결의서'),
            _row('2026/11/20 -1', 'R9106-XX', 80_000, 8_000, typ='매입전표 I', partner='지급운반비'),
            _row('2026/08/10 -9', 'G9103-CC', 999, 0)]                                  # 매출전표는 빠짐
    found = er.buy_diffs(er.aggregate(rows, buy=True), recs, date(2026, 11, 28))
    assert {c: i['category'] for c, i in found.items()} == {
        'G9102-BB': 'buy_diff', 'G9103-CC': 'buy_nocost', 'G9105-EE': 'buy_cancel', 'R9106-FF': 'buy_code'}
    assert found['G9102-BB']['eco_supply'] == 2_500_000 and found['G9102-BB']['sheet_supply'] == 2_000_000
    assert '26.08.10 2,500,000 가나공조' in found['G9102-BB']['note'] and found['G9102-BB']['evidence'].startswith('원가: 제품대 2,000,000')
    assert found['R9106-FF']['ecount_code'] == 'R9106-XX'


# ── ③ 반영: 자동 완료·되돌리기·이어받기 ────────────────────────────────────

def _it(cat, code, es_, ev, ss, sv, note=''):
    return {'category': cat, 'code': code, 'eco_supply': es_, 'eco_vat': ev, 'sheet_supply': ss, 'sheet_vat': sv, 'note': note}


def _row_of(code, side='sales'):
    with ea._conn() as conn:
        return ea._rows(conn, 'batch = ? AND side = ? AND code = ?', (ea.AUTO_BATCH, side, code))[0]


def test_apply_nightly_resolves_reopens_and_carries_over(db):
    # 손 대조(11-26): A 는 그대로 둠, B 는 이카운트 고침(그런데 아직 다름)
    ea.import_items('2026-11-26', [_it('split', 'G9201-AA', 100, 10, 100, 5), _it('supply', 'G9202-BB', 100, 10, 200, 20)])
    with ea._conn() as conn:
        ids = {r['code']: r['id'] for r in ea._rows(conn, 'batch = ?', ('2026-11-26',))}
    ea.set_status(ids['G9201-AA'], 'keep', '나눠 냄', by='황샛별')
    ea.set_status(ids['G9202-BB'], 'ecount_fixed', '', by='황샛별')

    found = {'G9201-AA': _it('split', 'G9201-AA', 100, 10, 100, 5), 'G9202-BB': _it('supply', 'G9202-BB', 100, 10, 200, 20),
             'G9203-CC': _it('supply', 'G9203-CC', 1, 0, 2, 0)}
    r1 = ea.apply_nightly('sales', found, '11/28')
    assert r1['baseline'] and sorted(r1['new']) == ['G9202-BB', 'G9203-CC']
    a = _row_of('G9201-AA')
    assert a['status'] == 'keep' and a['memo'] == '나눠 냄' and "'그대로 둠'을 이어받음" in a['history']
    b = _row_of('G9202-BB')
    assert b['status'] == 'open' and "11/26 '이카운트 고침' 뒤에도 아직 달라 처리 전으로" in b['history']

    # SB 가 C 를 [이카운트 고침] → 다음 밤에도 다르면 처리 전으로, B 는 이번엔 맞음 → 완료
    ea.set_status(_row_of('G9203-CC')['id'], 'ecount_fixed', '고침', by='황샛별')
    found2 = {'G9201-AA': found['G9201-AA'], 'G9203-CC': found['G9203-CC']}
    r2 = ea.apply_nightly('sales', found2, '11/29')
    assert r2['reopened'] == ['G9203-CC'] and r2['resolved'] == ['G9202-BB'] and not r2['baseline']
    c = _row_of('G9203-CC')
    assert c['status'] == 'open' and c['resolved_by'] is None
    assert c['history'].split('\n')[-2].endswith('황샛별 이카운트 고침 — 고침')
    assert c['history'].split('\n')[-1] == "11/29 자동 대조 — '이카운트 고침' 뒤에도 아직 달라 처리 전으로"
    b = _row_of('G9202-BB')
    assert b['status'] == 'resolved' and b['resolved_by'] == '자동 대조' and b['history'].endswith('같아져 맞음')

    # 그대로 둔 A 의 금액이 바뀌면 다시 처리 전, 맞음이던 B 가 다시 달라지면 처리 전
    r3 = ea.apply_nightly('sales', {'G9201-AA': _it('split', 'G9201-AA', 100, 10, 100, 0),
                                    'G9202-BB': _it('supply', 'G9202-BB', 5, 0, 6, 0)}, '11/30')
    assert sorted(r3['reopened']) == ['G9201-AA', 'G9202-BB'] and r3['resolved'] == ['G9203-CC']
    assert _row_of('G9201-AA')['status'] == 'open' and "'그대로 둠' 뒤 금액이 바뀌어 처리 전으로" in _row_of('G9201-AA')['history']


def test_list_items_prefers_auto_and_splits_sides(db):
    ea.import_items('2026-11-26', [_it('supply', 'G9301-AA', 1, 0, 2, 0)])
    assert ea.list_items()['batch'] == '2026-11-26'
    ea.apply_nightly('sales', {'G9301-AA': _it('supply', 'G9301-AA', 1, 0, 2, 0)}, '11/28')
    ea.apply_nightly('buy', {'G9304-DD': _it('buy_code', 'G9304-DD', 80, 0, 0, 0)}, '11/28')
    ea.apply_nightly('cost', {'G9302-BB': _it('buy_nocost', 'G9302-BB', 700, 0, 0, 0),
                              'G9303-CC': _it('buy_diff', 'G9303-CC', 700, 0, 600, 0)}, '11/28')
    out = ea.list_items()
    assert out['batch'] == 'auto'
    assert {s['key']: (s['open'], s['total']) for s in out['sides']} == {'sales': (1, 1), 'buy': (1, 1), 'cost': (2, 2)}
    assert [(g['key'], g['side']) for g in out['groups']] == [
        ('supply', 'sales'), ('buy_code', 'buy'), ('buy_nocost', 'cost'), ('buy_diff', 'cost')]
    # 오래된 '맞음'은 화면에서 빠짐
    ea.apply_nightly('cost', {'G9302-BB': _it('buy_nocost', 'G9302-BB', 700, 0, 0, 0)}, '11/29')
    with ea._conn() as conn:
        conn.execute("UPDATE ecount_audit_items SET resolved_at = datetime('now', '-30 days') WHERE code = 'G9303-CC'")
        conn.commit()
    assert {s['key']: s['total'] for s in ea.list_items()['sides']}['cost'] == 1
    with pytest.raises(ValueError):
        ea.set_status(_row_of('G9301-AA')['id'], 'resolved', '', by='황샛별')   # '맞음'은 사람이 못 누름


# ── 밤 실행 전체 ────────────────────────────────────────────────────────

def _read_file(tmp_path, sales, buy, ok=True):
    p = tmp_path / 'read.json'
    p.write_text(json.dumps({'ok': ok, 'error': '' if ok else '로그인 실패', 'sales': {'rows': sales}, 'buy': {'rows': buy}},
                            ensure_ascii=False), encoding='utf-8')
    return str(p)


def test_run_nightly_end_to_end(db, tmp_path, monkeypatch):
    import dashboard.services.project_service as ps
    recs = [_rec('G9401-AA', 1_000_000), _rec('G9402-BB', 2_000_000), _rec('G9403-CC', 3_000_000)]
    monkeypatch.setattr(ps, 'get_project_records', lambda force_refresh=False: recs)
    monkeypatch.setattr(er, 'MIN_SALES_ROWS', 1)
    monkeypatch.setattr(er, 'MIN_SHEET_ROWS', 1)
    es.record_entry('G9402-BB', 2_000_000, 200_000, '공사 확정', 'k2', 'gw:가나')      # [넣음 표시]는 했는데
    sales = [_row('2026/10/02 -1', 'G9401-AA', 1_000_000, 100_000),                   # 넣음 표시 없이 이카운트에 들어감
             _row('2026/10/02 -2', 'G9402-BB', 1_800_000, 180_000),                    # 이카운트엔 다른 금액
             _row('2026/10/03 -1', 'G9403-CC', 500_000, 50_000)]                       # 새 확정인데 일부만 들어감
    res = er.run_nightly(read_path=_read_file(tmp_path, sales, []), today=date(2026, 11, 28))
    assert res['ok'], res
    led = es.ledger_totals()
    assert led['G9401-AA'] == (1_000_000, 100_000)            # 이카운트 = 시트 → 장부 맞춤
    assert led['G9402-BB'] == (2_000_000, 200_000)            # 버튼만으로는 그대로
    assert 'G9403-CC' not in led                               # 이카운트 ≠ 시트 → 장부에 안 적음
    assert list((tmp_path / 'recon').glob('ledger_backup_*.json'))
    mism = {'G9402-BB': [1_800_000, 180_000], 'G9403-CC': [500_000, 50_000]}
    assert set(res['sales_found']) == {'G9402-BB', 'G9403-CC'} and res['mismatch'] == mism
    assert '장부에는 넣은 것으로 돼 있는데' in res['sales_found']['G9402-BB']['note']
    run = er.last_run()
    assert run['status'] == 'ok' and run['summary']['ledger_fixed'] == 1 and er.eco_mismatch() == mism
    # ② 목록: G9403 은 전액 '공사 확정'으로 뜨지만, 이카운트에 이미 일부가 있으니 확인 필요(바로 복사 안 됨)·① 등록에서는 빠짐
    items = es.compute_pending(recs, es.ledger_totals())
    assert [e['code'] for e in items] == ['G9403-CC']
    rows = es.upload_rows(items, eco_seen=er.eco_mismatch())
    assert rows['sales'][0]['warn'] == ['이카운트에 이미 550,000원 있음(③ 대조 확인)'] and rows['projects'] == []
    # 다음 밤 — SB 가 이카운트를 고쳐 같아짐 → ③ 완료, 장부는 이미 시트와 같음
    sales2 = sales + [_row('2026/11/28 -1', 'G9402-BB', 200_000, 20_000)]
    res2 = er.run_nightly(read_path=_read_file(tmp_path, sales2, []), today=date(2026, 11, 29))
    assert res2['ok'] and res2['summary']['sales']['resolved'] == ['G9402-BB']
    assert _row_of('G9402-BB')['status'] == 'resolved' and _row_of('G9403-CC')['status'] == 'open'
    # 그다음 밤 — G9403 나머지도 들어가 이카운트 = 시트 → 장부에 적히고 ③ 완료, ② 에서도 빠짐
    sales3 = sales2 + [_row('2026/11/29 -1', 'G9403-CC', 2_500_000, 250_000)]
    res3 = er.run_nightly(read_path=_read_file(tmp_path, sales3, []), today=date(2026, 11, 30))
    assert res3['summary']['ledger_fixed'] == 1 and res3['summary']['sales']['resolved'] == ['G9403-CC']
    assert es.ledger_totals()['G9403-CC'] == (3_000_000, 300_000) and es.compute_pending(recs, es.ledger_totals()) == []


def test_run_nightly_refuses_bad_reads(db, tmp_path, monkeypatch):
    import dashboard.services.project_service as ps
    monkeypatch.setattr(er, 'MIN_SALES_ROWS', 1)
    monkeypatch.setattr(ps, 'get_project_records', lambda force_refresh=False: [_rec('G9501-AA', 1)])   # 시트 1줄(백업 파일)
    sales = [_row('2026/10/02 -1', 'G9501-AA', 1, 0)]
    res = er.run_nightly(read_path=_read_file(tmp_path, sales, []))
    assert not res['ok'] and '공사 현황이 1줄뿐' in res['error']
    res = er.run_nightly(read_path=_read_file(tmp_path, sales, [], ok=False))
    assert not res['ok'] and '로그인 실패' in res['error']
    assert er.last_run()['status'] == 'failed' and er.last_run(ok_only=True) is None
    assert es.ledger_totals() == {}                                       # 실패하면 장부·대조 그대로
    # 지난번보다 전표가 크게 줄면 실패
    monkeypatch.setattr(er, 'MIN_SHEET_ROWS', 1)
    monkeypatch.setattr(ps, 'get_project_records', lambda force_refresh=False: [_rec('G9501-AA', 1)])
    assert er.run_nightly(read_path=_read_file(tmp_path, sales * 100, []))['ok']
    res = er.run_nightly(read_path=_read_file(tmp_path, sales * 50, []))
    assert not res['ok'] and '줄었어요' in res['error']


def test_dry_run_writes_nothing(db, tmp_path, monkeypatch):
    import dashboard.services.project_service as ps
    monkeypatch.setattr(er, 'MIN_SALES_ROWS', 1)
    monkeypatch.setattr(er, 'MIN_SHEET_ROWS', 1)
    monkeypatch.setattr(ps, 'get_project_records', lambda force_refresh=False: [_rec('G9601-AA', 1_000_000)])
    res = er.run_nightly(read_path=_read_file(tmp_path, [_row('2026/10/02 -1', 'G9601-AA', 1_000_000, 100_000)], []), apply=False)
    assert res['ok'] and len(res['fixes']) == 1
    assert es.ledger_totals() == {} and er.last_run() is None


# ── ①·②: 오늘 건은 내일로 ───────────────────────────────────────────────

def test_split_today_and_api_later(db, monkeypatch):
    from dashboard.blueprints import gw_finance as g
    today = date.today().isoformat()
    items = [{'code': 'G9701-AA', 'kind': '공사 확정', 'supply': 1, 'vat': 0, 'sent': (0, 0), 'target': (1, 0), 'date': today,
              'partner': '거래처', 'addr': '가나', 'note': '', 'split': False, 'key': 'a'},
             {'code': 'G9702-BB', 'kind': '공사 확정', 'supply': 2, 'vat': 0, 'sent': (0, 0), 'target': (2, 0), 'date': '2026-11-01',
              'partner': '거래처', 'addr': '다라', 'note': '', 'split': False, 'key': 'b'}]
    shown, later = es.split_today(items)
    assert [e['code'] for e in shown] == ['G9702-BB'] and [e['code'] for e in later] == ['G9701-AA']
    monkeypatch.setattr(es, 'collect_items', lambda force_refresh=False: list(items))
    with Flask('t').test_request_context('/api/finance/ecount'):
        j = inspect.unwrap(g.api_ecount)().get_json()
    assert j['later'] == 1 and [r['code'] for r in j['sales']] == ['G9702-BB'] and [p['code'] for p in j['projects']] == ['G9702-BB']


# ── 아침 DM ────────────────────────────────────────────────────────────

def test_build_dm_variants():
    ok = {'status': 'ok', 'summary': {'md': '11/29', 'sales': {'resolved': ['G1']}, 'buy': {'resolved': []}}}
    text, blocks = er.build_dm(ok, {'projects': 0, 'sales': 0}, [], [])
    assert text == '수정할 것 없음 — 이카운트·공사 현황·장부가 모두 맞아요.'
    flat = json.dumps(blocks, ensure_ascii=False)
    assert '1건은 이카운트와 공사 현황이 같아져 완료로' in flat
    sales_open = [{'code': 'G3', 'side': 'sales', 'eco_supply': 100, 'eco_vat': 10, 'sheet_supply': 200, 'sheet_vat': 20,
                   'ecount_code': 'G3', '_fresh': True}]
    buy = [{'code': 'G4', 'side': 'buy', 'category': 'buy_code', 'eco_supply': 700, 'sheet_supply': 0, 'ecount_code': 'G4', '_fresh': False},
           {'code': 'G5', 'side': 'buy', 'category': 'buy_cancel', 'eco_supply': 980, 'ecount_code': 'G5'},
           {'code': 'G6', 'side': 'buy', 'category': 'buy_noeco', 'eco_supply': 0, 'sheet_supply': 300, 'ecount_code': 'G6'}]
    text, blocks = er.build_dm(ok, {'projects': 1, 'sales': 2}, sales_open, buy)
    flat = json.dumps(blocks, ensure_ascii=False)
    assert text == '이카운트 대조 — 고칠 것 7건'
    assert '① 프로젝트 등록 1건 · ② 매출전표 2건' in flat and '`G3` 이카운트 110원 · 공사 현황 220원' in flat
    assert '`G4` 이카운트 매입 700원 · 공사 현황 원가 0원' in flat and '`G5` 취소된 공사에 이카운트 매입 980원 남음' in flat
    assert '`G6` 이카운트 매입 없음 · 공사 현황 원가 300원' in flat
    text, _ = er.build_dm({'status': 'failed'}, {}, [], [])
    assert '읽지 못해 쉬어요' in text


def test_run_nightly_splits_buy_work_from_sheet_cost(db, tmp_path, monkeypatch):
    """매입은 SB 가 이카운트를 고칠 것(매입)과 시트 원가 확인(원가 빈칸·다름)으로 나뉨 — 11-28 부장님 'SB 할 일은 30건만'."""
    import dashboard.services.project_service as ps
    monkeypatch.setattr(er, 'MIN_SALES_ROWS', 1)
    monkeypatch.setattr(er, 'MIN_SHEET_ROWS', 1)
    recs = [_rec('G9801-AA', 1_000_000, confirmed='2026-08-01'), _rec('R9802-BB', 1_000_000, confirmed='2026-08-01')]
    monkeypatch.setattr(ps, 'get_project_records', lambda force_refresh=False: recs)
    sales = [_row('2026/08/01 -1', 'G9801-AA', 1_000_000, 100_000), _row('2026/08/01 -2', 'R9802-BB', 1_000_000, 100_000)]
    buy = [_row('2026/08/05 -1', 'G9801-AA', 400_000, 40_000, typ='매입전표 I'),       # 원가 빈칸 → 시트 원가
           _row('2026/08/05 -2', 'G9802-BB', 300_000, 30_000, typ='매입전표 I')]       # 코드 다름 → 매입(SB)
    res = er.run_nightly(read_path=_read_file(tmp_path, sales, buy), today=date(2026, 11, 28))
    assert set(res['buy_found']) == {'R9802-BB'} and set(res['cost_found']) == {'G9801-AA'}
    sides = {s['key']: s['open'] for s in ea.list_items()['sides']}
    assert sides == {'sales': 0, 'buy': 1, 'cost': 1}
    assert [r['code'] for r in ea.open_items('buy')] == ['R9802-BB']


def test_weekend_skip_monday_uses_friday_night(monkeypatch):
    """주말엔 대조를 안 돌림(화~토 00:10) — 월요일 아침 DM 은 토요일 00:10(금요일 밤) 대조를 씀(11-28 부장님)."""
    monkeypatch.setattr(er, '_is_bd', lambda d: d.weekday() < 5)
    run = lambda kst: {'started_at': (kst - __import__('datetime').timedelta(hours=9)).strftime('%Y-%m-%d %H:%M:%S')}
    from datetime import datetime as dt
    mon, tue = date(2026, 11, 30), date(2026, 12, 1)
    assert er.fresh_for(run(dt(2026, 11, 28, 0, 10)), mon)          # 토 00:10 → 월요일에 씀
    assert not er.fresh_for(run(dt(2026, 11, 27, 0, 10)), mon)      # 금 00:10(목요일 밤) → 낡음
    assert er.fresh_for(run(dt(2026, 12, 1, 0, 10)), tue)
    assert not er.fresh_for(run(dt(2026, 11, 28, 0, 10)), tue)      # 화요일인데 토요일 것 → 낡음(월요일 밤 대조 실패)
    assert not er.fresh_for(None, mon)
    assert er.upto_label(date(2026, 11, 28)) == '11/27(금)'
    ok = {'status': 'ok', 'summary': {'md': '11/28', 'upto': '11/27(금)', 'sales': {}, 'buy': {}}}
    assert '11/27(금)까지 기준' in json.dumps(er.build_dm(ok, {}, [], [])[1], ensure_ascii=False)


def test_scheduler_skips_weekend():
    import inspect as _i
    from dashboard.services import sync_scheduler as ss
    src = _i.getsource(ss.start_scheduler)
    assert "_safe_ecount_recon, 'cron', day_of_week='tue-sat', hour=0, minute=10" in src
    assert "_safe_ecount_recon_dm, 'cron', day_of_week='mon-fri', hour=7, minute=50" in src


class _Proc:
    def __init__(self, code=0, out=b''):
        self.returncode, self.stdout, self.stderr = code, out, b''


def test_reader_via_scheduled_task_waits_for_new_result(tmp_path, monkeypatch):
    """서버(LocalSystem)는 Edge 가 안 떠서(11-28) 부장님 계정 예약 작업에 맡기고 새 결과 파일을 기다린다."""
    out = tmp_path / 'recon'
    out.mkdir()
    (out / 'ecount_read_20261127_001000.json').write_text('{}', encoding='utf-8')   # 지난 결과는 무시
    monkeypatch.setattr(er, 'OUT_DIR', out)
    monkeypatch.setenv('ECOUNT_READ_TASK', 'ITG_ecount_read')
    calls = []

    def fake_run(cmd, **kw):
        calls.append(cmd)
        if cmd[:2] == ['schtasks', '/run']:
            (out / 'ecount_read_20261128_001000.json').write_text('{"ok": true}', encoding='utf-8')
        return _Proc()
    monkeypatch.setattr(er.subprocess, 'run', fake_run)
    monkeypatch.setattr(er.time, 'sleep', lambda s: None)
    path, err = er.run_reader()
    assert err == '' and path.name == 'ecount_read_20261128_001000.json'
    assert calls[0] == ['schtasks', '/run', '/tn', 'ITG_ecount_read']


def test_reader_task_ended_without_result_fails_fast(tmp_path, monkeypatch):
    out = tmp_path / 'recon'
    out.mkdir()
    monkeypatch.setattr(er, 'OUT_DIR', out)
    clock = iter(range(0, 100000, 30))
    monkeypatch.setattr(er.time, 'time', lambda: next(clock))
    monkeypatch.setattr(er.time, 'sleep', lambda s: None)

    def fake_run(cmd, **kw):
        if cmd[:2] == ['schtasks', '/run']:
            (out / 'ecount_read_20261128_001000.log').write_text('00:10:01 [실패] .env 빈 칸\n', encoding='utf-8')
            return _Proc()
        return _Proc(out='"\\ITG_ecount_read","N/A","준비"'.encode('cp949'))
    monkeypatch.setattr(er.subprocess, 'run', fake_run)
    path, err = er._run_reader_task('ITG_ecount_read', 45 * 60)
    assert path is None and '결과 없이 끝남' in err and '.env 빈 칸' in err


def test_reader_task_start_failure(tmp_path, monkeypatch):
    monkeypatch.setattr(er, 'OUT_DIR', tmp_path / 'recon')
    monkeypatch.setattr(er.subprocess, 'run', lambda cmd, **kw: _Proc(1, '오류: 작업 없음'.encode('cp949')))
    path, err = er._run_reader_task('ITG_ecount_read', 60)
    assert path is None and '실행하지 못함(코드 1)' in err
