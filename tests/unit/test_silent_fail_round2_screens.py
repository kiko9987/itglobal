"""조용한 실패 점검 2차(11-24) — 실패를 '없음'·'0'·엉뚱한 원인으로 보이던 화면들.

14 인사·토요휴무 API 오류 분기: KeyError·IndexError·UnicodeError·JSONDecodeError 는 코드 버그 → ERROR 로그 + 500
19 빠른 찾기: 프로젝트·리드 한쪽 실패를 failed 로 알림
20 알림: 한 갈래 실패를 partial 로 알리고 캐시하지 않음
21 홈 내 휴가: 요약 실패를 '인사카드 없음'과 구분(leave_error)
22 홈 영업 숫자: 원천을 못 읽으면 0 집계·캐시 대신 500
"""
import json
import os
from datetime import date

import pandas as pd
import pytest
from flask import Flask

from dashboard.services import gw_home
from dashboard.services import gw_notify as gn
from dashboard.services import gw_search

from test_gw_home import _Fake, client, db  # noqa: F401 — 픽스처 재사용

JS = os.path.join(os.path.dirname(__file__), '..', '..', 'dashboard', 'static', 'js')


def _js(name):
    with open(os.path.join(JS, name), encoding='utf-8') as f:
        return f.read()


def _login(client):
    """일반 시험 계정 — 인사카드(로그인 메일)도 만들어 홈의 '내 휴가'가 보이게."""
    from dashboard.services import hr
    hr.upsert_from_flex({'flex_id': 'T', 'name': '시험자', 'status': '재직', 'login_email': 'tester@example.com',
                         'work_email': 'tester@example.com', 'slack_id': 'UTEST'}, [])
    with client.session_transaction() as s:
        s['user'] = {'email': 'tester@example.com', 'name': '시험자', 'permission_level': 'super_admin'}


class _Log:
    def __init__(self):
        self.errors = []

    def error(self, msg, *a, **k):
        self.errors.append((msg, k.get('exc_info')))

    def warning(self, *a, **k):
        pass

    info = debug = warning


# ── 14 ─────────────────────────────────────────────────────────────────
@pytest.mark.parametrize('mod_name, fn_name', [('hr', '_fail'), ('gw_satoff', '_err')])
def test_bug_errors_are_500_and_logged(monkeypatch, mod_name, fn_name):
    import importlib
    mod = importlib.import_module(f'dashboard.blueprints.{mod_name}')
    log = _Log()
    monkeypatch.setattr(mod, 'logger', log)
    fn = getattr(mod, fn_name)
    app = Flask('t')
    bugs = [KeyError('days'), IndexError('list index out of range'),
            json.JSONDecodeError('Expecting value', '', 0), UnicodeDecodeError('utf-8', b'\xff', 0, 1, 'invalid')]
    with app.app_context():
        for e in bugs:
            resp, code = fn(e)
            assert code == 500, e
            assert 'days' not in resp.get_json()['message']                  # 키 이름이 문구로 나가지 않음
        assert len(log.errors) == len(bugs) and all(exc for _, exc in log.errors)
        # 업무 오류는 그대로: 없음 404 · 입력 400 · 권한 403
        assert fn(LookupError('없는 구성원입니다.'))[1] == 404
        assert fn(ValueError('날짜를 다시 골라 주세요.'))[1] == 400
        assert fn(PermissionError('권한 없음'))[1] == 403
        assert fn(LookupError('없는 구성원입니다.'))[0].get_json()['message'] == '없는 구성원입니다.'
    assert len(log.errors) == len(bugs)                                     # 업무 오류는 ERROR 로그 없음


def test_satoff_business_errors_unchanged():
    from dashboard.blueprints.gw_satoff import _err
    from dashboard.services import sat_off as so
    with Flask('t').app_context():
        assert _err(so.SatOffError('토요일만 고를 수 있어요.'))[1] == 400      # SatOffError 는 ValueError 하위
        assert _err(LookupError('없는 변경 신청이에요.'))[1] == 404


def test_hr_route_keyerror_goes_500(client, monkeypatch):
    from dashboard.services import hr
    monkeypatch.setattr(hr, 'document_file', lambda doc_id: (_ for _ in ()).throw(KeyError('person_id')))
    _login(client)
    r = client.delete('/api/hr/documents/1')
    assert r.status_code == 500 and 'person_id' not in r.get_json()['message']
    monkeypatch.setattr(hr, 'document_file', lambda doc_id: (_ for _ in ()).throw(LookupError('없는 서류입니다.')))
    r = client.delete('/api/hr/documents/1')
    assert r.status_code == 404 and r.get_json()['message'] == '없는 서류입니다.'


# ── 19 ─────────────────────────────────────────────────────────────────
_PROJECTS = pd.DataFrame([{'프로젝트 코드': 'G4266-MJ', '거래처명': '호리바코리아', '현장 주소': '안양'}])
_LEADS = [{'리드 No': 'L-04510', '고객명': '호리바', '플랫폼': '네이버'}]


def _search_sources(monkeypatch, projects, leads):
    from dashboard.services import lead_service, project_service
    monkeypatch.setattr(project_service, 'load_data', lambda *a, **k: projects)
    monkeypatch.setattr(lead_service, 'get_lead_records', lambda: leads)


def test_search_marks_failed_side(monkeypatch):
    log = _Log()
    monkeypatch.setattr(gw_search, 'logger', log)
    _search_sources(monkeypatch, _PROJECTS, [dict(r) for r in _LEADS])
    r = gw_search.search('호리바')
    assert r['failed'] == [] and r['projects'] and r['leads']

    _search_sources(monkeypatch, _PROJECTS, [])                            # 리드 시트 못 읽음(get_lead_records 는 [])
    r = gw_search.search('호리바')
    assert r['failed'] == ['leads'] and r['projects'][0]['code'] == 'G4266-MJ' and r['leads'] == []

    _search_sources(monkeypatch, None, [dict(r) for r in _LEADS])          # 공사현황 못 읽음(load_data None)
    r = gw_search.search('호리바')
    assert r['failed'] == ['projects'] and r['leads'][0]['code'] == 'L-04510'
    assert len(log.errors) == 2 and '호리바' not in ''.join(m for m, _ in log.errors)   # 검색어는 로그에 안 남김

    _search_sources(monkeypatch, _PROJECTS, [dict(r) for r in _LEADS])     # 진짜로 안 맞는 건 실패가 아님
    r = gw_search.search('없는이름')
    assert r['failed'] == [] and r['projects'] == [] and r['leads'] == []


def test_search_js_shows_failed_note():
    s = _js('gw-search.js')
    assert 'data.failed' in s and '찾기가 잠시 안 됐어요. 결과가 빠졌을 수 있어요' in s
    assert "if (!shown && !failed.length) addNote(" in s                   # 실패면 '결과가 없어요'를 안 띄움


# ── 20 ─────────────────────────────────────────────────────────────────
def _notify_sources(monkeypatch, leave_fails):
    ev = {'cat': 'approval', 'at': '2026-11-24 09:00:00', 'title': 't', 'actor_id': ''}

    def leave(sid, since):
        if leave_fails:
            raise RuntimeError('db locked')
        return [dict(ev, cat='leave', at='2026-11-24 10:00:00')]
    monkeypatch.setattr(gn, '_approval_events', lambda sid, since: [dict(ev)])
    monkeypatch.setattr(gn, '_leave_events', leave)
    monkeypatch.setattr(gn, '_yearend_events', lambda person, admin, since: [])
    monkeypatch.setattr(gn, 'seen_at', lambda who: '')


def test_notify_partial_not_cached(monkeypatch):
    log = _Log()
    monkeypatch.setattr(gn, 'logger', log)
    gn._cache.clear()
    person = {'id': 1, 'slack_id': 'UMW'}
    _notify_sources(monkeypatch, leave_fails=True)
    out = gn.feed('staff@example.com', person, False, False)
    assert out['partial'] == ['leave'] and len(out['items']) == 1
    assert 'staff@example.com' not in gn._cache                          # 실패 결과는 캐시하지 않음
    assert len(log.errors) == 1 and log.errors[0][1]

    _notify_sources(monkeypatch, leave_fails=False)                       # 복구되면 바로 그 갈래도 보임
    out = gn.feed('staff@example.com', person, False, False)
    assert 'partial' not in out and len(out['items']) == 2
    assert 'staff@example.com' in gn._cache
    gn._cache.clear()


def test_notify_js_skips_seen_on_partial():
    s = _js('gw-notify.js')
    assert 'partial = !!(j.partial && j.partial.length)' in s
    assert 'if (!failed && !partial) markSeen();' in s
    assert '일부 알림을 불러오지 못했어요' in s
    assert 'function jsonOrEnd(r)' in s                                    # 로그인 만료 판정은 그대로


# ── 21 ─────────────────────────────────────────────────────────────────
def test_home_leave_error_separate_from_unlinked(client, monkeypatch):
    from dashboard.services import leave_view
    _login(client)
    j = client.get('/api/home').get_json()
    assert j['leave'] is not None and 'leave_error' not in j

    def boom(p):
        raise RuntimeError('ledger down')
    monkeypatch.setattr(leave_view, 'summary', boom)
    j = client.get('/api/home').get_json()
    assert j['leave'] is None and j['leave_error'] is True


def test_home_js_leave_error_message():
    s = _js('gw-home.js')
    assert 'renderLeave(j.leave, j.leave_error)' in s
    assert '휴가 정보를 불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.' in s


# ── 22 ─────────────────────────────────────────────────────────────────
def _sales_sources(monkeypatch, leads, df):
    from dashboard.services import lead_service, project_service
    monkeypatch.setattr(lead_service, 'get_lead_records', lambda: leads)
    monkeypatch.setattr(project_service, 'load_data', lambda *a, **k: df)
    gw_home._sales_cache.clear()


_DF = pd.DataFrame([{'공사 확정': '2026-11-03', '총액 2': '1000', '수금 관련 특이사항': ''}])


@pytest.mark.parametrize('leads, df', [
    ([], _DF),                                                            # 리드 시트 못 읽음
    ([{'상담 시간': '2026.11.10. 09:00', '상태': ''}], None),             # 공사현황 못 읽음
    ([{'상담 시간': '2026.11.10. 09:00', '상태': ''}], pd.DataFrame([{'총액 2': '1'}])),   # '공사 확정' 열 없음
])
def test_sales_source_failure_raises_and_not_cached(monkeypatch, leads, df):
    _sales_sources(monkeypatch, leads, df)
    with pytest.raises(gw_home.SalesSourceError):
        gw_home.sales(date(2026, 11, 10))
    assert not gw_home._sales_cache


def test_sales_real_zero_is_not_failure(monkeypatch):
    # 시트는 읽었는데 기간 안에 문의·계약이 없음 → 0 으로 보여 줌
    _sales_sources(monkeypatch, [{'상담 시간': '2025.01.02. 09:00', '상태': ''}],
                   pd.DataFrame([{'공사 확정': '2025-01-03', '총액 2': '1000', '수금 관련 특이사항': ''}]))
    s = gw_home.sales(date(2026, 11, 10))
    assert s['this'] == {'inquiry': 0, 'visit': 0, 'contract': 0, 'amount': 0}
    assert gw_home._sales_cache
    gw_home._sales_cache.clear()


def test_sales_api_500_on_source_failure(client, monkeypatch):
    _sales_sources(monkeypatch, [], _DF)
    _login(client)
    r = client.get('/api/home/sales')
    assert r.status_code == 500 and r.get_json()['message'] == '영업 숫자를 불러오지 못했습니다.'
    assert not gw_home._sales_cache
