"""현장·미화 슬랙 로그인 링크(2026-11-10) — 1회용·만료·현장 세션 문지기·세션 검사."""
from datetime import datetime, timedelta

import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import groupware as gw
from dashboard.services import gw_login
from dashboard.services import hr

ADMIN = 'kiko@itg-aircon.com'
FIELD_SLACK = 'UFIELD1'
FIELD_MAIL = 'field1@gmail.com'


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, gw_login):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.setenv('GW_PUBLIC_URL', 'https://pm.example.test/')
    gw._invalidate()
    hr.upsert_from_flex({'flex_id': 'F1', 'name': '김현장', 'status': '재직', 'access': 'field',
                         'personal_email': FIELD_MAIL, 'slack_id': FIELD_SLACK}, [])
    hr.upsert_from_flex({'flex_id': 'F2', 'name': '박사무', 'status': '재직', 'access': 'staff',
                         'work_email': 'mw@itg-aircon.com', 'login_email': 'mw@itg-aircon.com', 'slack_id': 'USTAFF'}, [])
    hr.upsert_from_flex({'flex_id': 'F3', 'name': '이퇴사', 'status': '퇴사', 'access': 'field', 'slack_id': 'UGONE'}, [])
    yield d
    gw._invalidate()


def _token(url):
    return url.rsplit('/', 1)[1]


def test_issue_only_field_and_pilot(db):
    assert not gw_login.can_issue(FIELD_SLACK)            # 정식 오픈 전 + 시범 명단 밖
    gw.add_pilot(FIELD_MAIL, ADMIN)
    assert gw_login.can_issue(FIELD_SLACK)
    assert not gw_login.can_issue('USTAFF')                # 사무실 직원은 구글 로그인
    assert not gw_login.can_issue('UGONE')                 # 퇴사자
    assert not gw_login.can_issue('UNOBODY')
    url = gw_login.issue(FIELD_SLACK)
    assert url.startswith('https://pm.example.test/gw/login/') and len(_token(url)) >= 40
    with pytest.raises(PermissionError):
        gw_login.issue('USTAFF')


def test_token_single_use_expiry_and_not_stored_plain(db):
    gw.set_open_all(True, ADMIN)
    t0 = datetime(2026, 11, 10, 9, 0, 0)
    tok = _token(gw_login.issue(FIELD_SLACK, now=t0))
    with db._get_connection() as conn:
        assert conn.execute('SELECT COUNT(*) FROM gw_login_tokens WHERE token_hash = ?', (tok,)).fetchone()[0] == 0
    assert gw_login.consume(tok, now=t0 + timedelta(minutes=11)) is None          # 10분 지남
    tok2 = _token(gw_login.issue(FIELD_SLACK, now=t0))
    p = gw_login.consume(tok2, ip='1.2.3.4', now=t0 + timedelta(minutes=9))
    assert p and p['name'] == '김현장'
    assert gw_login.consume(tok2, now=t0 + timedelta(minutes=9)) is None          # 두 번째는 안 됨
    assert gw_login.consume('없는토큰') is None and gw_login.consume('') is None


def test_consume_rejects_if_person_changed(db):
    gw.set_open_all(True, ADMIN)
    tok = _token(gw_login.issue(FIELD_SLACK))
    pid = hr.find_by_slack(FIELD_SLACK)['id']
    hr.update_person(pid, {'access': 'staff'}, ADMIN)                              # 링크 뒤에 접근이 바뀜
    assert gw_login.consume(tok) is None


# ── 웹: 로그인 링크·문지기·세션 검사 ───────────────────────────────────────
@pytest.fixture
def client(db):
    from dashboard.blueprints.groupware import groupware_bp
    from dashboard.blueprints.hr import hr_bp
    from dashboard.auth import admin_required, login_required
    app = Flask('t')
    app.secret_key = 't'
    app.register_blueprint(groupware_bp)
    app.register_blueprint(hr_bp)

    from flask import Blueprint
    auth_bp = Blueprint('auth', __name__)
    auth_bp.add_url_rule('/login', 'login_page', lambda: 'login')
    app.register_blueprint(auth_bp)

    @app.route('/projects')
    @login_required
    def projects():
        return 'PM'

    @app.route('/api/projects/list')
    @login_required
    def plist():
        return {'data': ['고객 정보']}

    @app.route('/api/admin-only')
    @admin_required
    def admin_only():
        return {'ok': True}

    gw.set_open_all(True, ADMIN)
    return app.test_client()


def test_field_login_flow_and_gatekeeper(client):
    tok = _token(gw_login.issue(FIELD_SLACK))
    r = client.get(f'/gw/login/{tok}')
    assert r.status_code == 302 and r.headers['Location'].endswith('/hr/me')
    assert client.get(f'/gw/login/{tok}').status_code == 410                     # 다시 쓰면 거절 화면

    with client.session_transaction() as s:
        u = s['user']
        assert u['gw_field'] and u['permission_level'] == 'field' and u['email'] == FIELD_MAIL
    me_id = hr.find_by_slack(FIELD_SLACK)['id']
    other_id = hr.find_by_slack('USTAFF')['id']

    # 기존 PM 화면·API 는 문지기가 막음
    r = client.get('/projects')
    assert r.status_code == 302 and r.headers['Location'].endswith('/hr/me')
    assert client.get('/api/projects/list').status_code == 403
    assert client.get('/hr/people').status_code == 302
    assert client.get('/gw/admin').status_code == 302
    assert client.post('/api/gw/layout', json={'new': False}).status_code == 403
    assert client.post('/api/gw/pilots', json={'open_all': False}).status_code == 403
    assert client.get('/api/hr/people').status_code == 403
    assert client.put(f'/api/hr/people/{me_id}', json={'memo': 'x'}).status_code == 403
    assert client.post(f'/api/hr/people/{me_id}/ssn').status_code == 403
    assert client.get('/api/admin-only').status_code == 403

    # 자기 카드만
    mine = client.get(f'/api/hr/people/{me_id}').get_json()
    assert mine['success'] and mine['person']['name'] == '김현장'
    assert not mine['person']['can_edit'] and not mine['person']['can_reveal_ssn']
    assert client.get(f'/api/hr/people/{other_id}').status_code == 403


def test_field_session_ends_when_resigned(client):
    tok = _token(gw_login.issue(FIELD_SLACK))
    client.get(f'/gw/login/{tok}')
    pid = hr.find_by_slack(FIELD_SLACK)['id']
    assert client.get(f'/api/hr/people/{pid}').status_code == 200
    hr.update_person(pid, {'status': '퇴사'}, ADMIN)
    r = client.get(f'/api/hr/people/{pid}')
    assert r.status_code == 302 or r.status_code == 401                          # 세션 종료 → 로그인
    with client.session_transaction() as s:
        assert 'user' not in s


def test_nav_for_field():
    nav = gw.build_nav('/hr/me', 'field', field=True)
    assert [a['label'] for a in nav['areas']] == ['휴가', '인사']
    assert [i['label'] for i in nav['items']] == ['내 정보']
    assert [i['label'] for i in gw.build_nav('/leave', 'field', field=True)['items']] == ['내 휴가']


def test_slack_home_button_only_for_field_pilot(db, monkeypatch):
    from dashboard.blueprints import slack_leave as sl
    monkeypatch.setattr(sl.lv, 'staff_by_slack', lambda uid: None)
    monkeypatch.setattr(sl.lv, 'is_open', lambda uid: False)
    ids = lambda v: [b.get('accessory', {}).get('action_id') for b in v['blocks']]
    assert sl.A_GW_LINK not in ids(sl.home_view(FIELD_SLACK))                     # 시범 명단 밖
    gw.add_pilot(FIELD_MAIL, ADMIN)
    assert sl.A_GW_LINK in ids(sl.home_view(FIELD_SLACK))
    assert sl.A_GW_LINK not in ids(sl.home_view('USTAFF'))
    view = sl.gw_link_view('https://pm.example.test/gw/login/abc')
    assert view['blocks'][1]['elements'][0]['url'].endswith('/gw/login/abc')
