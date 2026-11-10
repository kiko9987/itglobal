"""빠른 찾기(2026-11-10) — 프로젝트·리드 찾기 순서·전화번호·제외 규칙·권한."""
import pandas as pd
import pytest
from flask import Flask

from dashboard.services import gw_search

PROJECTS = pd.DataFrame([
    {'프로젝트 코드': 'G4266-MJ', '거래처명': '호리바코리아', '사업자명': '', '현장 주소': '안양 만안구 1', '발주처 담당자': '김담당',
     '발주처 연락처': '010-1234-5678', '담당자': '박민우', '공사 시작': pd.Timestamp('2026-11-21'), '수금 관련 특이사항': ''},
    {'프로젝트 코드': 'G4265-MJ', '거래처명': '', '사업자명': '광명상회', '현장 주소': '광명 오리로 870', '발주처 담당자': '',
     '발주처 연락처': '', '담당자': '박민우', '공사 시작': '', '수금 관련 특이사항': '공사 취소'},
    {'프로젝트 코드': 'R4262-SJ', '거래처명': '플레저성형외과', '사업자명': '', '현장 주소': '강남구 강남대로 476', '발주처 담당자': '',
     '발주처 연락처': '02-555-1234', '담당자': '반승정', '공사 시작': float('nan'), '수금 관련 특이사항': ''},
])

LEADS = [
    {'리드 No': 'L-04510', '고객명': '호리바', '고객 연락처': '010-9999-5678', '방문 주소': '안양 만안구 1', '플랫폼': '네이버',
     '상태': '방문 완료', '상담 시간': '2026-11-09 10:00', '영업 담당자': '박민우', '온라인 상담자': ''},
    {'리드 No': 'ETC-000123', '고객명': '호리바', '고객 연락처': '', '방문 주소': '', '플랫폼': '기타',
     '상태': '', '상담 시간': '', '영업 담당자': '', '온라인 상담자': ''},
    {'리드 No': 'L-04400', '고객명': '김고객', '고객 연락처': '01012345678', '방문 주소': '광명 오리로 870', '플랫폼': '당근',
     '상태': '부재중', '상담 시간': '', '영업 담당자': '', '온라인 상담자': ''},
]


@pytest.fixture(autouse=True)
def data(monkeypatch):
    from dashboard.services import lead_service, project_service
    monkeypatch.setattr(project_service, 'load_data', lambda *a, **k: PROJECTS)
    monkeypatch.setattr(lead_service, 'get_lead_records', lambda: [dict(r) for r in LEADS])


def test_code_first_and_links():
    r = gw_search.search('g4266')
    assert [p['code'] for p in r['projects']] == ['G4266-MJ']
    p = r['projects'][0]
    assert p['title'] == '호리바코리아' and p['meta'] == '박민우 · 2026-11-21'
    assert p['href'] == '/projects?search=G4266-MJ&fresh=1' and not p['cancelled']
    assert gw_search.search('L-04510')['leads'][0]['href'] == '/leads?search=L-04510'


def test_name_address_and_exclusions():
    r = gw_search.search('호리바')
    assert [p['code'] for p in r['projects']] == ['G4266-MJ']
    assert [l['code'] for l in r['leads']] == ['L-04510']          # 플랫폼 '기타'(ETC-)는 뺌
    r = gw_search.search('광명 오리로')
    assert r['projects'][0]['cancelled'] and r['projects'][0]['title'] == '광명상회'   # 취소도 보여 주되 표시
    assert gw_search.search('강남대로476')['projects'][0]['code'] == 'R4262-SJ'        # 띄어쓰기 무시


def test_phone_digits():
    r = gw_search.search('5678')
    assert {p['code'] for p in r['projects']} == {'G4266-MJ'}
    assert {l['code'] for l in r['leads']} == {'L-04510', 'L-04400'}
    r = gw_search.search('010-1234-5678')
    assert [p['code'] for p in r['projects']] == ['G4266-MJ'] and [l['code'] for l in r['leads']] == ['L-04400']
    # 글자가 섞인 코드는 전화번호로 보지 않음(G4400 의 '4400' 이 남의 번호에 걸리면 안 됨)
    assert gw_search.search('L-0440')['projects'] == [] and [l['code'] for l in gw_search.search('L-0440')['leads']] == ['L-04400']
    assert gw_search.search('G1234')['leads'] == [] and gw_search.search('G1234')['projects'] == []


def test_ranking_limit_and_short_query():
    r = gw_search.search('박민우')                                   # 담당자 이름만 맞음 → 코드 번호 큰 순
    assert [p['code'] for p in r['projects']] == ['G4266-MJ', 'G4265-MJ']
    assert gw_search.search('박민우', limit=1)['project_total'] == 2
    empty = gw_search.search('a')
    assert empty['projects'] == [] and empty['leads'] == []


def test_one_side_failure_keeps_other(monkeypatch):
    from dashboard.services import project_service

    def boom(*a, **k):
        raise RuntimeError('sheet down')
    monkeypatch.setattr(project_service, 'load_data', boom)
    r = gw_search.search('호리바')
    assert r['projects'] == [] and r['leads'][0]['code'] == 'L-04510'


def test_api_staff_ok_field_blocked(monkeypatch):
    import dashboard.auth as auth
    from dashboard.blueprints.groupware import groupware_bp
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    app = Flask('t')
    app.secret_key = 't'
    app.register_blueprint(groupware_bp)
    c = app.test_client()
    with c.session_transaction() as s:
        s['user'] = {'email': 'mw@itg-aircon.com', 'name': '박민우', 'permission_level': 'editor'}
    j = c.get('/api/gw/search?q=G4266').get_json()
    assert j['success'] and j['projects'][0]['code'] == 'G4266-MJ'
    with c.session_transaction() as s:
        s['user'] = {'email': '', 'name': '김현장', 'permission_level': 'field', 'gw_field': True, 'hr_person_id': 1}
    assert c.get('/api/gw/search?q=G4266').status_code in (302, 401, 403)
