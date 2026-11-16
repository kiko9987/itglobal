# -*- coding: utf-8 -*-
"""리드 하나에 공사 여러 건 — '방문 현장 불러오기' 검색 (2026-11-16).

입찰이 2건으로 쪼개진 롯데홈쇼핑 L-04421 처럼 한 리드에 공사가 여럿 붙는 경우,
전에는 그 리드를 검색에서 아예 빼서 번호를 그대로 쳐도 안 나왔다.
이제는 보여 주되 linked_codes 를 담아 뒤로 민다.
"""
from unittest.mock import patch

import pandas as pd
import pytest
from flask import Flask

from dashboard.blueprints import leads as leads_mod


def _lead(no, name, status='방문 완료', phone='010-1111-2222', addr='영등포구 양평로21길 10'):
    return {'리드 No': no, '고객명': name, '상태': status, '고객 연락처': phone,
            '방문 주소': addr, '플랫폼': '메일', '방문 예정일': '2026-11-03',
            '영업 담당자': '박정우', '문의 내용': '', '상담 내용': '', '폴더 ID': ''}


LEADS = [
    _lead('L-04421', '롯데홈쇼핑 오민석', status='공사 확정'),   # 이미 R4295-JW 로 등록됨
    _lead('L-04500', '롯데홈쇼핑 2차'),                          # 아직 미등록
    _lead('ETC-7119e8', '롯데홈쇼핑'),                           # 리드 번호가 아님
]

PROJECTS = pd.DataFrame([
    {'프로젝트 코드': 'R4295-JW', 'Lead No': 'L-04421'},
    {'프로젝트 코드': 'G1234-YM', 'Lead No': ''},
])


@pytest.fixture
def client():
    app = Flask(__name__)
    app.secret_key = 'test'
    app.register_blueprint(leads_mod.leads_bp)
    with patch.object(leads_mod, 'get_lead_records', return_value=LEADS), \
         patch.object(leads_mod, 'get_user_role', return_value='Admin'), \
         patch.object(leads_mod, '_attach_folder_links', lambda items: None), \
         patch('dashboard.auth._validate_session', return_value=None), \
         patch('dashboard.services.project_service.load_data', return_value=PROJECTS):
        c = app.test_client()
        with c.session_transaction() as s:
            s['user'] = {'email': 'kiko@itg-aircon.com', 'name': '고광일'}
        yield c


def _search(c, q=''):
    r = c.get(f'/leads/api/search-for-project?q={q}&all=1')
    assert r.status_code == 200
    return (r.get_json().get('data') or {}).get('leads') or []


def test_linked_lead_is_searchable_with_codes(client):
    """번호로 찾으면 이미 공사가 붙은 리드도 나오고, 어느 공사인지 알려 준다."""
    found = _search(client, 'L-04421')
    assert [l['lead_no'] for l in found] == ['L-04421']
    assert found[0]['linked_codes'] == ['R4295-JW']


def test_linked_lead_sorts_behind_unlinked(client):
    """같은 검색어에 둘 다 걸리면 미등록 건이 먼저 (등록된 건은 점수를 깎음)."""
    found = [l['lead_no'] for l in _search(client, '롯데홈쇼핑')]
    assert found == ['L-04500', 'L-04421']


def test_unlinked_lead_has_no_badge_field(client):
    """미등록 건에는 linked_codes 를 담지 않는다 (프론트 배지 안 뜸)."""
    found = _search(client, 'L-04500')
    assert found and 'linked_codes' not in found[0]


def test_default_list_still_hides_linked_leads(client):
    """빈 검색어 기본 목록은 그대로 미등록 건만 — 목록이 등록 건으로 불지 않게."""
    assert [l['lead_no'] for l in _search(client)] == ['L-04500']


def test_project_code_map():
    """Lead No → 공사 코드 매핑. 코드 칸이 '코드' 이름이어도 읽는다."""
    with patch('dashboard.services.project_service.load_data', return_value=PROJECTS):
        assert leads_mod._lead_no_project_codes() == {'L-04421': ['R4295-JW']}
        assert leads_mod._registered_lead_nos() == {'L-04421'}

    alt = PROJECTS.rename(columns={'프로젝트 코드': '코드'})
    with patch('dashboard.services.project_service.load_data', return_value=alt):
        assert leads_mod._lead_no_project_codes() == {'L-04421': ['R4295-JW']}


def test_two_projects_on_one_lead():
    """한 리드에 공사가 2건이면 둘 다 돌려준다."""
    df = pd.DataFrame([
        {'프로젝트 코드': 'R4295-JW', 'Lead No': 'L-04421'},
        {'프로젝트 코드': 'R4300-JW', 'Lead No': 'L-04421'},
    ])
    with patch('dashboard.services.project_service.load_data', return_value=df):
        assert leads_mod._lead_no_project_codes() == {'L-04421': ['R4295-JW', 'R4300-JW']}
