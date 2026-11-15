"""현장 직원 휴대폰 등록(2026-11-15) — 관리자 숫자 → 휴대폰 등록 → 계속 로그인, 끊기·퇴사·로그아웃·시도 제한."""
from datetime import date, datetime, timedelta

import pytest
from cryptography.fernet import Fernet
from flask import Blueprint, Flask, redirect, session

from dashboard.services import groupware as gw
from dashboard.services import gw_device
from dashboard.services import gw_login
from dashboard.services import hr
from dashboard.services import leave as lv
from dashboard.services import leave_view
from dashboard.services import leave_web

ADMIN = hr._DEFAULT_HR_ADMINS[0]     # 기본 인사카드 관리자
STAFF = 'mw@itg-aircon.com'
UA_IPHONE = 'Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 Version/18.0 Mobile/15E148 Safari/604.1'


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, gw_login, gw_device, lv, leave_view):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.setenv('LEAVE_APPROVER_IDS', 'UJW,UYG')
    monkeypatch.setenv('LEAVE_PILOT_IDS', 'UF')
    gw._invalidate()
    hr.upsert_from_flex({'flex_id': 'P1', 'name': '박민우', 'status': '재직', 'work_email': STAFF,
                         'login_email': STAFF, 'slack_id': 'UMW'}, [])
    hr.upsert_from_flex({'flex_id': 'P3', 'name': '권미화', 'status': '재직', 'access': 'field', 'slack_id': 'UF'}, [])
    lv.upsert_staff('권미화', '2023-07-24', 'UF')
    lv.upsert_staff('박민우', '2022-02-03', 'UMW')
    lv.add_grant('권미화', 12, date.today() - timedelta(days=5), date.today() + timedelta(days=200), source='t')
    monkeypatch.setattr(leave_web, '_run_bg', lambda fn: None)        # 슬랙 알림은 이 시험에서 안 봄
    yield d
    gw._invalidate()


@pytest.fixture
def client(db, monkeypatch):
    import dashboard.auth as auth
    real = auth._validate_session
    monkeypatch.setattr(auth, '_validate_session',
                        lambda require_admin=False: real(require_admin) if session.get('user', {}).get('gw_field')
                        else (None if session.get('user') else redirect('/login')))
    from dashboard.blueprints.groupware import groupware_bp
    from dashboard.blueprints.hr import hr_bp
    app = Flask('t')
    app.secret_key = 't'
    app.register_blueprint(groupware_bp)
    app.register_blueprint(hr_bp)
    a = Blueprint('auth', __name__)
    a.add_url_rule('/login', 'login_page', lambda: 'login')

    def _logout():
        session.clear()
        return redirect('/login')
    a.add_url_rule('/logout', 'logout', _logout)
    app.register_blueprint(a)
    gw.set_open_all(True, ADMIN)
    return app.test_client()


def _field_id():
    return next(p['id'] for p in hr.list_people() if p['name'] == '권미화')


def _as(c, email, perm='super_admin'):
    with c.session_transaction() as s:
        s.clear()
        s['user'] = {'email': email, 'name': 'x', 'permission_level': perm}


def _cookie(c):
    ck = c.get_cookie(gw_device.COOKIE)
    return ck.value if ck else None


def _register(c):
    _as(c, ADMIN)
    code = c.post(f'/api/hr/people/{_field_id()}/device-code').get_json()['code']
    with c.session_transaction() as s:
        s.clear()
    r = c.post('/gw/device/redeem', json={'code': code}, headers={'User-Agent': UA_IPHONE})
    return code, r


# ── 숫자 만들기 ─────────────────────────────────────────────────────────
def test_issue_code_only_for_field_with_slack(db):
    c = gw_device.issue_code(_field_id(), ADMIN)
    assert len(c['code']) == 6 and c['code'].isdigit() and c['minutes'] == 10
    with pytest.raises(ValueError):
        gw_device.issue_code(hr.find_by_email(STAFF)['id'], ADMIN)    # 사무실 직원(구글 로그인)은 안 됨
    with db._get_connection() as conn:
        assert conn.execute('SELECT COUNT(*) FROM gw_device_codes').fetchone()[0] == 1
        assert c['code'] not in str([tuple(r) for r in conn.execute('SELECT * FROM gw_device_codes').fetchall()])
    c2 = gw_device.issue_code(_field_id(), ADMIN)                     # 새로 만들면 예전 숫자는 못 씀
    assert gw_device.redeem(c['code'])[2] and not gw_device.redeem(c2['code'])[2]


def test_hr_api_permissions(client):
    pid = _field_id()
    _as(client, STAFF, 'editor')                                       # 인사카드 관리자 아님
    assert client.get(f'/api/hr/people/{pid}/devices').status_code == 403
    assert client.post(f'/api/hr/people/{pid}/device-code').status_code == 403
    _as(client, ADMIN)
    j = client.get(f'/api/hr/people/{pid}/devices').get_json()
    assert j['can_register'] and j['devices'] == []
    staff_pid = hr.find_by_email(STAFF)['id']
    assert client.get(f'/api/hr/people/{staff_pid}/devices').get_json()['can_register'] is False
    assert client.post(f'/api/hr/people/{staff_pid}/device-code').status_code == 400


# ── 등록·계속 로그인 ────────────────────────────────────────────────────
def test_register_then_stays_logged_in(client):
    code, r = _register(client)
    assert r.status_code == 200 and r.get_json()['redirect'] == '/leave'
    tok = _cookie(client)
    assert tok and gw_device.check(tok)
    with client.session_transaction() as s:
        u = s['user']
        assert u['gw_field'] and u['gw_person_id'] == _field_id() and u['gw_slack_id'] == 'UF' and u['gw_device_id']
    assert client.post('/gw/device/redeem', json={'code': code}).status_code == 400    # 숫자는 한 번만

    # 세션이 사라져도(4시간 지남·서버 정리) 휴대폰 열쇠로 다시 열림
    with client.session_transaction() as s:
        s.clear()
    r = client.get('/api/leave/me')
    assert r.status_code == 200 and r.get_json()['leave']['name'] == '권미화'
    assert r.get_json()['form']['can_apply']                          # 휴가 신청도 됨(시범 명단 UF)
    # 오래된 로그인 시각은 당겨 둔다(4시간 끝에 안 걸리게)
    with client.session_transaction() as s:
        s['login_time'] = (datetime.now() - timedelta(hours=5)).isoformat()
    client.get('/api/leave/me')
    with client.session_transaction() as s:
        assert datetime.now() - datetime.fromisoformat(s['login_time']) < timedelta(minutes=1)
    # 현장은 바로 내 휴가, 남의 화면은 여전히 막힘
    assert client.get('/home').headers['Location'].endswith('/leave')
    assert client.get('/api/leave/team').status_code == 403
    assert client.get('/hr/people').status_code == 302

    _as(client, ADMIN)
    devs = client.get(f'/api/hr/people/{_field_id()}/devices').get_json()['devices']
    assert [d['label'] for d in devs] == ['아이폰 · 사파리']


def test_revoke_from_hr_card_ends_session(client):
    _register(client)
    tok = _cookie(client)
    did = gw_device.check(tok)['device']['id']
    assert gw_device.revoke(did, ADMIN, person_id=_field_id())       # 관리자가 인사카드에서 [끊기]
    r = client.get('/api/leave/me')
    assert r.status_code in (401, 302)                                # 열쇠로 연 세션도 닫힘
    assert _cookie(client) is None


def test_logout_on_phone_revokes(client):
    _register(client)
    tok = _cookie(client)
    client.get('/logout')
    assert gw_device.check(tok) is None and _cookie(client) is None


@pytest.mark.parametrize('col,val', [('status', '퇴사'), ('access', 'staff'), ('slack_id', 'UOTHER')])
def test_card_change_ends_device(client, col, val):
    _register(client)
    tok = _cookie(client)
    with hr._conn() as conn:
        conn.execute(f'UPDATE hr_people SET {col} = ? WHERE id = ?', (val, _field_id()))
    assert gw_device.check(tok) is None


def test_idle_days(client):
    _register(client)
    tok = _cookie(client)
    assert gw_device.check(tok, now=datetime.now() + timedelta(days=gw_device.IDLE_DAYS + 1)) is None
    assert gw_device.check(tok, now=datetime.now() + timedelta(days=gw_device.IDLE_DAYS - 1))


def test_wrong_codes_are_rate_limited(db):
    now = datetime.now()
    good = gw_device.issue_code(_field_id(), ADMIN, now)['code']
    bad = '000000' if good != '000000' else '111111'
    for _ in range(gw_device.FAIL_PER_IP):
        assert gw_device.redeem(bad, '1.2.3.4', now=now)[2].startswith('번호가 맞지')
    p, tok, why = gw_device.redeem(good, '1.2.3.4', now=now)          # 맞는 숫자도 잠깐 막힘
    assert p is None and '잠시 막혔어요' in why
    p, tok, why = gw_device.redeem(good, '5.6.7.8', now=now)          # 다른 주소는 됨
    assert p and tok and not why
    assert gw_device.redeem(good, '9.9.9.9', now=now + timedelta(minutes=1))[2]   # 이미 씀


def test_expired_code(db):
    now = datetime.now()
    c = gw_device.issue_code(_field_id(), ADMIN, now)['code']
    assert gw_device.redeem(c, now=now + timedelta(minutes=11))[2]


def test_card_slack_id_swapped_to_someone_else_cannot_touch_their_leave(client):
    """인사카드 관리자가 현장 카드에 남(박민우)의 슬랙 ID 를 넣고 등록해도 그 사람 휴가는 못 건드림."""
    with hr._conn() as conn:
        conn.execute('UPDATE hr_people SET slack_id = ? WHERE id = ?', ('UMW', _field_id()))
    _register(client)
    j = client.get('/api/leave/me').get_json()
    assert j['form']['can_apply'] is False and j['form']['why'] == leave_web.MSG_NOT_LISTED


def test_redeem_path_and_device_label():
    from dashboard.blueprints.groupware import field_path_allowed
    assert field_path_allowed('POST', '/gw/device/redeem')
    assert gw_device.device_label(UA_IPHONE) == '아이폰 · 사파리'
    assert gw_device.device_label('Mozilla/5.0 (Linux; Android 14; SM-S921N) AppleWebKit/537.36 '
                                  'Chrome/130.0 Mobile Safari/537.36') == '안드로이드 · 크롬'
    assert gw_device.device_label('Mozilla/5.0 (Linux; Android 14) SamsungBrowser/26.0 Chrome/122.0 '
                                  'Mobile Safari/537.36') == '안드로이드 · 삼성 인터넷'


def test_security_middleware_lets_redeem_through_only_same_origin():
    from dashboard.utils.security_middleware import SecurityMiddleware
    src = open(SecurityMiddleware.process_request.__code__.co_filename, encoding='utf-8').read()
    i = src.index("request.path == '/gw/device/redeem'")
    assert '_same_origin_request()' in src[i:i + 500] and src.index('# 2. CSRF/인증 검사') > i
