"""그룹웨어 휴가 화면(2026-11-10) — 플렉스 기록 보관·요약·달력·권한."""
from datetime import date

import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import groupware as gw
from dashboard.services import gw_login
from dashboard.services import hr
from dashboard.services import leave as lv
from dashboard.services import leave_view

ADMIN = 'kiko@itg-aircon.com'
STAFF = 'mw@itg-aircon.com'
VIEWER = 'approver@itg-aircon.com'
TODAY = date(2026, 11, 10)


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, gw_login, lv, leave_view):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    gw._invalidate()
    hr.upsert_from_flex({'flex_id': 'P1', 'name': '박민우', 'status': '재직', 'work_email': STAFF,
                         'login_email': STAFF, 'slack_id': 'UMW', 'department': '영업', 'hire_date': '2022-02-03'}, [])
    hr.upsert_from_flex({'flex_id': 'P2', 'name': '박민우', 'status': '퇴사'}, [])          # 동명이인 퇴사자
    hr.upsert_from_flex({'flex_id': 'P3', 'name': '김현장', 'status': '재직', 'access': 'field', 'slack_id': 'UF'}, [])
    lv.upsert_staff('박민우', '2022-02-03', 'UMW')
    lv.add_grant('박민우', 8, date(2026, 2, 3), date(2027, 2, 2), source='flex:박민우')
    me = hr.find_by_email(STAFF)['id']
    gone = [p for p in hr.list_people(include_resigned=True) if p['status'] == '퇴사'][0]['id']
    leave_view.upsert_flex_history([
        {'flex_event_id': 'E1', 'person_id': me, 'name': '박민우', 'type_label': '연차', 'start': '2026-11-03',
         'end': '2026-11-04', 'days': 2, 'unit': 'DAY', 'status': '승인', 'reason': '가족 여행'},
        {'flex_event_id': 'E2', 'person_id': me, 'name': '박민우', 'type_label': '연차', 'start': '2026-11-20',
         'end': '2026-11-20', 'days': 0.5, 'unit': 'HALF_DAY_PM', 'status': '승인'},
        {'flex_event_id': 'E3', 'person_id': me, 'name': '박민우', 'type_label': '연차', 'start': '2026-10-01',
         'end': '2026-10-01', 'days': 1, 'unit': 'DAY', 'status': '취소 승인'},
        {'flex_event_id': 'E4', 'person_id': me, 'name': '박민우', 'type_label': '병가', 'start': '2025-05-01',
         'end': '2025-05-02', 'days': 2, 'unit': 'DAY', 'status': '승인'},
        {'flex_event_id': 'E5', 'person_id': gone, 'name': '박민우', 'type_label': '연차', 'start': '2026-11-05',
         'end': '2026-11-05', 'days': 1, 'unit': 'DAY', 'status': '승인'},
    ])
    yield d
    gw._invalidate()


def test_upsert_is_idempotent(db):
    assert leave_view.upsert_flex_history([{'flex_event_id': 'E1', 'person_id': 1, 'name': '박민우', 'start': '2026-11-03',
                                            'status': '승인', 'days': 2}]) == 1
    with db._get_connection() as conn:
        assert conn.execute('SELECT COUNT(*) FROM gw_leave_history').fetchone()[0] == 5


def test_summary_counts_and_separates_namesakes(db):
    me = hr.find_by_email(STAFF)
    s = leave_view.summary(me, TODAY)
    assert s['linked'] and s['balance'] == 8
    assert s['annual_used'] == 2.5                     # 승인만(취소 승인 제외), 올해만
    assert s['sick_used'] == 0                         # 병가는 작년
    assert [e['id'] for e in s['entries']] == ['E2', 'E1', 'E3', 'E4']     # 퇴사 동명이인 기록(E5) 안 섞임
    assert s['upcoming'][0]['id'] == 'E2' and s['upcoming'][0]['half'] == '오후'
    assert s['entries'][2]['status'] == leave_view.ST_CANCEL


def test_slack_requests_merged(db):
    rid = None
    with lv._conn() as conn:
        cur = conn.execute('INSERT INTO leave_requests (name, slack_id, ltype, start, "end", days, status, reason) '
                           "VALUES ('박민우', 'UMW', 'annual', '2026-12-01', '2026-12-01', 1, 'wait_1', '병원')")
        rid = cur.lastrowid
    s = leave_view.summary(hr.find_by_email(STAFF), TODAY)
    slack = [e for e in s['entries'] if e['source'] == 'slack']
    assert slack and slack[0]['id'] == rid and slack[0]['status'] == leave_view.ST_WAIT and s['pending'] == 1


def test_calendar_and_off_on(db):
    cal = leave_view.calendar(2026, 11)
    assert [e['id'] for e in cal] == ['E1', 'E5', 'E2']          # 취소 빠짐, 날짜순
    assert all('reason' not in e for e in cal)                   # 달력엔 사유 없음
    assert [e['id'] for e in leave_view.off_on(date(2026, 11, 4))] == ['E1']
    assert leave_view.calendar(2026, 12) == []


def test_team_rows(db):
    rows = {r['name'] + ('*' if r['access'] == 'field' else ''): r for r in leave_view.team(TODAY)}
    assert rows['박민우']['balance'] == 8 and rows['박민우']['next']['start'] == '2026-11-20'
    assert rows['김현장*']['linked'] is False and rows['김현장*']['balance'] is None


# ── API 권한 ───────────────────────────────────────────────────────────
@pytest.fixture
def client(db, monkeypatch):
    import dashboard.auth as auth
    real = auth._validate_session
    monkeypatch.setattr(auth, '_validate_session',
                        lambda require_admin=False: real(require_admin) if __import__('flask').session.get('user', {}).get('gw_field') else None)
    from dashboard.blueprints.groupware import groupware_bp
    from dashboard.blueprints.hr import hr_bp
    app = Flask('t')
    app.secret_key = 't'
    app.register_blueprint(groupware_bp)
    app.register_blueprint(hr_bp)
    from flask import Blueprint
    a = Blueprint('auth', __name__)
    a.add_url_rule('/login', 'login_page', lambda: 'login')
    app.register_blueprint(a)
    gw.set_open_all(True, ADMIN)
    gw.set_setting('leave_viewers', '["approver@itg-aircon.com"]', ADMIN)
    return app.test_client()


def _as(c, email, perm='editor'):
    with c.session_transaction() as s:
        s.clear()
        s['user'] = {'email': email, 'name': 'x', 'permission_level': perm}


def test_leave_api_permissions(client):
    me = hr.find_by_email(STAFF)['id']
    field = hr.find_by_slack('UF')['id']

    _as(client, STAFF)
    assert client.get('/api/leave/me').get_json()['leave']['balance'] == 8
    assert client.get(f'/api/leave/person/{me}').status_code == 200
    assert client.get(f'/api/leave/person/{field}').status_code == 403
    assert client.get('/api/leave/team').status_code == 403
    assert client.get('/api/leave/calendar').status_code == 403
    assert client.get('/leave/team').status_code == 302

    _as(client, VIEWER, 'viewer')                                   # 휴가 열람자
    assert client.get('/api/leave/team').status_code == 200
    assert client.get(f'/api/leave/person/{me}').get_json()['leave']['entries'][0]['reason'] == ''
    assert client.get('/api/leave/calendar?month=2026-11').get_json()['month'] == '2026-11'
    assert client.get('/api/leave/calendar?month=2026-13').status_code == 400
    assert client.get(f'/api/hr/people/{me}').status_code == 403       # 인사카드는 못 봄

    _as(client, ADMIN, 'super_admin')                               # 인사카드 관리자도 전사 휴가
    assert client.get('/api/leave/team').status_code == 200

    # 현장 세션: 내 휴가만
    tok = gw_login.issue('UF').rsplit('/', 1)[1]
    with client.session_transaction() as s:
        s.clear()
    client.get(f'/gw/login/{tok}')
    mine = client.get('/api/leave/me').get_json()['leave']             # 휴가 장부 미연결 → 기록만, 잔여는 없음
    assert mine['name'] == '김현장' and mine['linked'] is False and mine['balance'] is None
    assert client.get('/api/leave/team').status_code == 403
    assert client.get(f'/api/leave/person/{me}').status_code == 403
    assert client.get('/api/leave/calendar').status_code == 403


def test_nav_leave_menu():
    assert [i['label'] for i in gw.build_nav('/leave', 'editor')['items']] == ['내 휴가']
    nav = gw.build_nav('/leave/team', 'viewer', leave_all=True)
    assert [i['label'] for i in nav['items'] if i['active']] == ['전사 현황']          # /leave 는 안 켜짐
    assert [i['label'] for i in gw.build_nav('/leave', 'field', field=True)['items']] == ['내 휴가']
