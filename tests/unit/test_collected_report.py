# -*- coding: utf-8 -*-
"""월별 수금 완료 정산 자료 (에어테이블 대체) — 기간 규칙·필터·묶음 합계·엑셀."""
import io
from datetime import date

import pytest

from dashboard.services import collected_report as cr


def test_period_is_prev_26_to_25():
    assert cr.period_range(2026, 9) == (date(2026, 8, 26), date(2026, 9, 25))
    assert cr.period_range(2025, 1) == (date(2024, 12, 26), date(2025, 1, 25))


@pytest.mark.parametrize('d, p', [
    (date(2026, 9, 25), (2026, 9)), (date(2026, 9, 26), (2026, 10)),
    (date(2026, 12, 26), (2027, 1)), (date(2026, 8, 1), (2026, 8)),
])
def test_period_of(d, p):
    assert cr.period_of(d) == p


def test_parse_period_default_and_bad():
    assert cr.parse_period('2026-10') == (2026, 10)
    assert cr.parse_period('', today=date(2026, 11, 8)) == (2026, 11)
    assert cr.parse_period('2026-13', today=date(2026, 11, 27)) == (2026, 12)


def _r(code, owner, biz, d, ok='TRUE', total1='₩1,000,000', profit='₩300,000', **kw):
    base = {'프로젝트 코드': code, '담당자': owner, '사업자': biz, '수금 날짜': d, '수금 확인': ok,
            '총액 1': total1, '총액 2': '₩1,100,000', '순익': profit, '미수금': '₩0', '부가세': 'TRUE',
            '마진율': '30%', '사업자명': '(주)테스트'}
    base.update(kw)
    return base


def test_collect_filters_period_and_check():
    recs = [
        _r('G1-AA', '박정우', '글로벌', '2026-08-25'),          # 8월 정산
        _r('G2-AA', '박정우', '글로벌', '2026-08-26'),          # 9월 시작
        _r('G3-AA', '박정우', '글로벌', '2026-09-25'),          # 9월 끝
        _r('G4-AA', '박정우', '글로벌', '2026-09-26'),          # 10월
        _r('G5-AA', '박정우', '글로벌', '2026-09-10', ok='FALSE'),  # 수금 확인 안 됨
        _r('G6-AA', '박정우', '글로벌', ''),                    # 날짜 없음
        _r('', '박정우', '글로벌', '2026-09-10'),               # 코드 없음
    ]
    rep = cr.collect(recs, 2026, 9)
    assert [x['프로젝트 코드'] for x in rep['rows']] == ['G2-AA', 'G3-AA']
    assert rep['start'] == '2026-08-26' and rep['end'] == '2026-09-25'


def test_groups_and_sums():
    recs = [
        _r('G1-YG', '박용구', '글로벌', '2026-09-01'),
        _r('P2-YG', '박용구', '플렌트', '2026-09-02', total1='₩2,000,000', profit='₩500,000'),
        _r('G3-JW', '박정우', '글로벌', '2026-09-03'),
        _r('G4-XX', '', '글로벌', '2026-09-04'),
    ]
    rep = cr.collect(recs, 2026, 9)
    owners = [g['담당자'] for g in rep['groups']]
    assert owners == sorted(owners) and '(담당자 없음)' in owners
    yg = next(g for g in rep['groups'] if g['담당자'] == '박용구')
    assert [s['사업자'] for s in yg['groups']] == ['글로벌', '플렌트']
    assert yg['sum']['총액 1'] == 3_000_000 and yg['sum']['순익'] == 800_000 and yg['sum']['건수'] == 2
    assert rep['total']['건수'] == 4 and rep['total']['총액 1'] == 5_000_000
    assert rep['rows'][0]['부가세'] is True and rep['rows'][0]['총액 2'] == 1_100_000


def test_bool_and_timestamp_inputs():
    import pandas as pd
    recs = [_r('G1-AA', 'A', '글로벌', pd.Timestamp('2026-09-05'), ok=True),
            _r('G2-AA', 'A', '글로벌', pd.NaT)]
    rep = cr.collect(recs, 2026, 9)
    assert [r['수금 날짜'] for r in rep['rows']] == ['2026-09-05']


def test_xlsx_has_rows_and_summary_sheet():
    from openpyxl import load_workbook
    recs = [_r('G1-YG', '박용구', '글로벌', '2026-09-01'), _r('G3-JW', '박정우', '글로벌', '2026-09-03')]
    rep = cr.collect(recs, 2026, 9)
    wb = load_workbook(io.BytesIO(cr.to_xlsx(rep)))
    ws = wb['수금 완료']
    assert [c.value for c in ws[1]] == cr.COLUMNS
    assert ws.max_row == 3
    assert ws.cell(2, cr.COLUMNS.index('총액 1') + 1).value == 1_000_000
    assert ws.cell(2, cr.COLUMNS.index('수금 확인') + 1).value == '✓'
    s = wb['담당자별 합계']
    assert s.cell(s.max_row, 1).value == '전체 합계' and s.cell(s.max_row, 3).value == 2
    assert cr.xlsx_filename(rep) == '수금완료_26년9월(08.26-09.25).xlsx'


def test_api_json_and_xlsx_routes(monkeypatch):
    from flask import Flask
    from dashboard import auth
    from dashboard.blueprints import collected_report as bp
    monkeypatch.setattr(auth, '_validate_session', lambda *a, **k: None)
    monkeypatch.setattr(bp, '_records', lambda refresh: [_r('G1-YG', '박용구', '글로벌', '2026-09-01')])
    app = Flask(__name__)
    app.register_blueprint(bp.collected_report_bp)
    c = app.test_client()
    j = c.get('/api/settlement/collected?period=2026-09').get_json()
    assert j['success'] and j['data']['total']['건수'] == 1
    r = c.get('/api/settlement/collected.xlsx?period=2026-09')
    assert r.status_code == 200 and r.data[:2] == b'PK'
    assert '%EC%88%98%EA%B8%88%EC%99%84%EB%A3%8C' in r.headers['Content-Disposition']   # '수금완료' 파일 이름


def test_recent_periods():
    ps = cr.recent_periods(today=date(2026, 11, 26), n=3)
    assert [p['value'] for p in ps] == ['2026-12', '2026-11', '2026-10']
    assert ps[0]['label'] == '26년 12월 (11/26~12/25)'


# ── 달별 기간 바꾸기 (2026-11-09 SB 요청: 명절 긴 달은 기준일이 바뀜) ──
def test_resolve_range_default_and_saved():
    assert cr.resolve_range(2026, 10) == (date(2026, 9, 26), date(2026, 10, 25), False)
    saved = {'2026-10': '2026-09-24|2026-10-23'}
    assert cr.resolve_range(2026, 10, saved) == (date(2026, 9, 24), date(2026, 10, 23), True)
    # 다음 달은 바꾼 끝날 다음 날부터 (빠지는 날 없게), 앞 달 끝은 그대로라 앞 달과는 겹침 → 안내
    assert cr.resolve_range(2026, 11, saved) == (date(2026, 10, 24), date(2026, 11, 25), False)
    assert cr.range_notes(2026, 11, saved) == []
    assert any('겹칩니다' in n for n in cr.range_notes(2026, 10, saved))   # 9월(8/26~9/25)과 9/24~9/25
    # 깨진 값은 무시
    assert cr.resolve_range(2026, 10, {'2026-10': 'x'})[2] is False


def test_range_notes_gap_with_next_month():
    saved = {'2026-10': '2026-09-26|2026-10-23', '2026-11': '2026-10-26|2026-11-25'}
    notes = cr.range_notes(2026, 10, saved)
    assert len(notes) == 1 and '10/24~10/25' in notes[0] and '들어가지 않습니다' in notes[0]


def test_collect_uses_saved_range_and_label():
    recs = [_r('G1-AA', '박정우', '글로벌', '2026-10-24'), _r('G2-AA', '박정우', '글로벌', '2026-09-25')]
    saved = {'2026-10': '2026-09-24|2026-10-23'}
    rep = cr.collect(recs, 2026, 10, saved)
    assert [r['프로젝트 코드'] for r in rep['rows']] == ['G2-AA']
    assert rep['custom'] and rep['label'] == '26년 10월 (9/24~10/23) · 기간 바꿈'
    assert (rep['default_start'], rep['default_end']) == ('2026-09-26', '2026-10-25')
    rep11 = cr.collect(recs, 2026, 11, saved)
    assert [r['프로젝트 코드'] for r in rep11['rows']] == ['G1-AA'] and not rep11['custom']
    assert cr.xlsx_filename(rep) == '수금완료_26년10월(09.24-10.23).xlsx'


def test_check_range():
    assert cr.check_range(date(2026, 10, 2), date(2026, 10, 1))
    assert cr.check_range(date(2026, 8, 1), date(2026, 10, 25))
    assert cr.check_range(date(2026, 9, 24), date(2026, 10, 23)) is None


def test_range_api_save_and_reset(monkeypatch):
    from flask import Flask
    from dashboard import auth
    from dashboard.blueprints import collected_report as bp

    class FakeRedis(dict):
        def hgetall(self, k): return dict(self)
        def hset(self, k, f, v): self[f] = v
        def hdel(self, k, f): self.pop(f, None)

    fake = FakeRedis()
    monkeypatch.setattr(auth, '_validate_session', lambda *a, **k: None)
    monkeypatch.setattr(bp, '_redis', lambda: fake)
    monkeypatch.setattr(bp, '_records', lambda refresh: [_r('G1-YG', '박용구', '글로벌', '2026-10-24')])
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.collected_report_bp)
    c = app.test_client()
    url = '/api/settlement/collected/range'
    assert c.post(url, json={'period': '2026-10', 'start': '2026-09-24', 'end': '2026-10-23'}).get_json()['success']
    assert fake == {'2026-10': '2026-09-24|2026-10-23'}
    j = c.get('/api/settlement/collected?period=2026-11').get_json()
    assert j['data']['start'] == '2026-10-24' and j['data']['total']['건수'] == 1
    # 틀린 기간은 거절, 기본과 같으면 저장 안 함, reset 은 지움
    r = c.post(url, json={'period': '2026-10', 'start': '2026-10-24', 'end': '2026-10-23'})
    assert r.status_code == 400 and not r.get_json()['success']
    assert c.post(url, json={'period': '2026-13', 'start': '2026-10-01', 'end': '2026-10-02'}).status_code == 400
    assert c.post(url, json={'period': '2026-10', 'reset': True}).get_json()['success'] and fake == {}
    c.post(url, json={'period': '2026-10', 'start': '2026-09-26', 'end': '2026-10-25'})
    assert fake == {}
