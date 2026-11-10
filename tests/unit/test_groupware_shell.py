"""그룹웨어 새 틀(2026-11-10) — 시범 명단·전체 켜기·메뉴·화면 틀 고르기."""
import pytest
from flask import Flask, render_template_string, session

from dashboard.services import groupware as gw


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(gw, '_db', lambda: d)
    monkeypatch.setattr(gw, '_table_ready', False)
    gw._invalidate()
    yield d
    gw._invalidate()


def test_seed_only_kiko_and_others_keep_old(db):
    assert gw.can_use('kiko@itg-aircon.com')
    assert gw.can_use('KiKO@itg-aircon.com')            # 대소문자 무관
    assert not gw.can_use('someone@itg-aircon.com')
    assert not gw.can_use('')
    assert [p['email'] for p in gw.list_pilots()] == ['kiko@itg-aircon.com']


def test_seed_runs_once_even_after_removal(db, monkeypatch):
    gw.remove_pilot('kiko@itg-aircon.com')
    monkeypatch.setattr(gw, '_table_ready', False)      # 서버 재시작 흉내
    gw._invalidate()
    assert gw.list_pilots() == []                      # 지운 사람이 되살아나지 않음


def test_add_remove_and_open_all(db):
    gw.add_pilot(' SB@itg-aircon.com ', by='kiko@itg-aircon.com')
    assert gw.can_use('sb@itg-aircon.com')              # 캐시도 바로 비움 → 재시작 없이 반영
    gw.remove_pilot('sb@itg-aircon.com')
    assert not gw.can_use('sb@itg-aircon.com')
    with pytest.raises(ValueError):
        gw.add_pilot('이메일아님', by='x')
    gw.set_open_all(True, by='kiko@itg-aircon.com')
    assert gw.can_use('anyone@itg-aircon.com') and gw.is_open_all()
    gw.set_open_all(False, by='kiko@itg-aircon.com')
    assert not gw.can_use('anyone@itg-aircon.com')


def test_session_off_returns_old_layout(db):
    assert gw.wants_new_layout('kiko@itg-aircon.com', session_off=False)
    assert not gw.wants_new_layout('kiko@itg-aircon.com', session_off=True)
    assert not gw.wants_new_layout('someone@itg-aircon.com', session_off=False)


def test_db_error_falls_back_to_old(monkeypatch):
    def boom():
        raise RuntimeError('db down')
    monkeypatch.setattr(gw, '_db', boom)
    monkeypatch.setattr(gw, '_table_ready', False)
    gw._invalidate()
    assert not gw.can_use('kiko@itg-aircon.com')      # 화면은 기존 틀로라도 뜬다
    gw._invalidate()


def test_build_nav_active_and_admin_only():
    nav = gw.build_nav('/leads', 'editor')
    assert [a['label'] for a in nav['areas']] == ['홈', '결재', '영업', '경영지원', '휴가', '인사']      # 관리는 관리자만
    assert nav['areas'][2]['active'] and not nav['areas'][3]['active']
    labels = [i['label'] for i in nav['items']]
    assert labels == ['프로젝트', '고객 리드', '매출통계']            # 시공사례는 관리자만
    assert [i['label'] for i in nav['items'] if i['active']] == ['고객 리드']

    nav = gw.build_nav('/settlement/collected', 'super_admin')
    assert nav['areas'][3]['active'] and [i['label'] for i in nav['items']] == ['월별 수금']
    ap = gw.build_nav('/approval/archive', 'editor')
    assert ap['areas'][1]['active'] and [i['label'] for i in ap['items'] if i['active']] == ['플렉스 보관함']
    assert '시공사례' in [i['label'] for i in gw.build_nav('/projects', 'admin')['items']]

    hr_nav = gw.build_nav('/hr/me', 'editor')
    assert [i['label'] for i in hr_nav['items']] == ['내 정보', '연말정산']     # 구성원·연말정산 현황은 HR 관리자만
    hr_nav = gw.build_nav('/hr/people', 'editor', is_hr_admin=True)
    assert [i['label'] for i in hr_nav['items'] if i['active']] == ['구성원']
    admin_nav = gw.build_nav('/gw/admin', 'admin')
    assert [a['label'] for a in admin_nav['areas'] if a['bottom']] == ['관리'] and admin_nav['areas'][-1]['active']

    # 펼친 왼쪽 줄용 하위 메뉴(children) — 권한대로 거르고, 지금 화면만 켜짐
    hr_area = next(a for a in gw.build_nav('/leave/team', 'editor', leave_all=True)['areas'] if a['key'] == 'leave')
    assert [(c['label'], c['active']) for c in hr_area['children']] == [('내 휴가', False), ('전사 현황', True)]
    sales = next(a for a in gw.build_nav('/home', 'editor')['areas'] if a['key'] == 'sales')
    assert [c['label'] for c in sales['children']] == ['프로젝트', '고객 리드', '매출통계'] and not sales['active']
    assert not any(c['active'] for c in sales['children'])

    nav = gw.build_nav('/admin/cache-monitor', 'admin')        # 메뉴에 없는 화면
    assert nav['items'] == [] and not any(a['active'] for a in nav['areas'])
    assert gw.build_nav('/projectsX', 'admin')['items'] == []   # 앞부분만 같은 주소는 아님


def _app():
    from dashboard.blueprints.groupware import groupware_bp
    app = Flask('t')
    app.secret_key = 't'
    app.register_blueprint(groupware_bp)
    return app


def test_context_processor_picks_layout(db):
    app = _app()
    tpl = "{{ gw_layout }}|{{ gw_can_switch|default(False) }}|{{ (gw_nav['items']|map(attribute='label')|join(',')) if gw_layout else '' }}"
    with app.test_request_context('/projects'):
        session['user'] = {'email': 'someone@itg-aircon.com', 'permission_level': 'admin'}
        assert render_template_string(tpl) == 'False|False|'
    with app.test_request_context('/projects'):
        session['user'] = {'email': 'kiko@itg-aircon.com', 'permission_level': 'super_admin'}
        assert render_template_string(tpl) == 'True|False|프로젝트,고객 리드,매출통계,시공사례'
    with app.test_request_context('/projects'):
        session['user'] = {'email': 'kiko@itg-aircon.com', 'permission_level': 'super_admin'}
        session['gw_off'] = True
        assert render_template_string(tpl) == 'False|True|'


def test_layout_api_blocks_non_pilot(db, monkeypatch):
    import dashboard.auth as auth
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    app = _app()
    c = app.test_client()
    with c.session_transaction() as s:
        s['user'] = {'email': 'someone@itg-aircon.com', 'permission_level': 'viewer'}
    assert c.post('/api/gw/layout', json={'new': True}).status_code == 403
    with c.session_transaction() as s:
        s['user'] = {'email': 'kiko@itg-aircon.com', 'permission_level': 'super_admin'}
    assert c.post('/api/gw/layout', json={'new': False}).get_json()['new'] is False
    with c.session_transaction() as s:
        assert s.get('gw_off') is True
    assert c.post('/api/gw/layout', json={'new': True}).status_code == 200
    with c.session_transaction() as s:
        assert 'gw_off' not in s


def test_new_layout_renders_with_file_versioned_assets(db, tmp_path):
    """새 틀 화면이 실제로 그려지고, 그룹웨어 CSS·JS 주소에 파일 시각 버전이 붙는다(재시작 없이 반영, 11-10)."""
    import os
    import re
    from flask import session, render_template
    from dashboard.blueprints.groupware import groupware_bp
    from dashboard.utils import frontend_helpers as fh
    root = os.path.join(os.path.dirname(__file__), '..', '..', 'dashboard')
    app = Flask('t', template_folder=os.path.join(root, 'templates'), static_folder=os.path.join(root, 'static'))
    app.secret_key = 't'
    fh.init_frontend_helpers(app)
    app.register_blueprint(groupware_bp)
    with app.test_request_context('/hr/me'):
        session['user'] = {'email': 'kiko@itg-aircon.com', 'name': '고광일', 'permission_level': 'super_admin'}
        html = render_template('gw/hr_me.html', person_id=None, fields=[])
    assert 'class="gw-shell"' in html and 'gwRail' in html
    # 왼쪽 줄(당근식, 11-10): 하위 메뉴가 여럿인 인사는 펴진 묶음, 하나뿐인 홈은 바로 가는 링크
    assert 'gw-nav-group active open' in html and 'aria-controls="gwSub-hr"' in html
    assert re.search(r'gw-nav-subitem active" href="/hr/me" aria-current="page"', html)
    assert 'aria-controls="gwSub-home"' not in html and '<a class="gw-nav-area" href="/home"' in html
    # 빠른 찾기(11-10): 사무실 직원은 프로젝트·리드까지, 스크립트는 파일 시각 버전
    assert 'id="gwFindBtn"' in html and 'data-scope="all"' in html and '/static/js/gw-search.js?v=' in html
    css = re.search(r'/static/css/gw-pages\.css\?v=(\d+)', html)
    shell = re.search(r'/static/js/groupware-shell\.js\?v=(\d+)', html)
    assert css and shell and int(css.group(1)) == int(os.path.getmtime(os.path.join(root, 'static', 'css', 'gw-pages.css')))


def test_project_filters_chip_labels_only_in_new_layout():
    """프로젝트 필터(당근식 칩, 11-10): 새 화면은 첫 칸이 필터 이름, 예전 화면은 그대로 '전체'."""
    import os
    from jinja2 import Environment, FileSystemLoader
    root = os.path.join(os.path.dirname(__file__), '..', '..', 'dashboard', 'templates')
    env = Environment(loader=FileSystemLoader(root))
    tpl = env.get_template('components/project_filters.html')
    old, new = tpl.render(gw_layout=False), tpl.render(gw_layout=True)
    assert '>담당자</option>' not in old and old.count('<option value="">전체</option>') == 8
    assert '<option value="">전체</option>' not in new
    for label in ('내부 사업자', '담당자', '유입 구분', '거래처·사업자명', '상태', '세금계산서', '수금 상태', '데이터'):
        assert f'<option value="">{label}</option>' in new
    assert 'placeholder="검색..."' in old and 'placeholder="코드·현장·고객 검색"' in new
