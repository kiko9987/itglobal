"""그룹웨어 알림 🔔(2026-11-10) — 결재·휴가·연말정산 기록에서 나와 관련된 일, 읽음 처리, 현장 범위."""
from datetime import date

import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import e_approval as ea
from dashboard.services import groupware as gw
from dashboard.services import gw_login
from dashboard.services import gw_notify as gn
from dashboard.services import hr
from dashboard.services import leave as lv
from dashboard.services import yearend as ye

KIKO, JW, MW = 'U04UL2ZLJAX', 'U051ENK8AAY', 'UMW'


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, gw_login, ea, lv, ye, gn):
        monkeypatch.setattr(mod, '_db', lambda: d)
    for mod, attr in ((gw, '_table_ready'), (hr, '_table_ready'), (gw_login, '_table_ready'), (ea, '_table_ready'),
                      (lv, '_table_ready'), (ye, '_table_ready'), (gn, '_ready')):
        monkeypatch.setattr(mod, attr, False, raising=False)
    for k in ('EAPPROVAL_EXECUTOR_DM_IDS', 'EAPPROVAL_APPROVER_IDS', 'EAPPROVAL_FINAL_APPROVER_IDS', 'LEAVE_APPROVER_IDS',
              'EAPPROVAL_N_EXECUTOR_IDS', 'SLACK_SETTLEMENT_CHECKER_ID', 'SLACK_SETTLEMENT_CHECKER_EMAIL', 'EAPPROVAL_FINAL_ENABLED'):
        monkeypatch.delenv(k, raising=False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.setattr(ye, 'current_year', lambda today=None: 2026)
    monkeypatch.setattr(ye, '_dm', lambda sid, text: True)
    gn._cache.clear()
    gw._invalidate()
    for fid, name, email, sid in (('K', '고광일', 'kiko@itg-aircon.com', KIKO), ('J', '박정우', 'jw@itg-aircon.com', JW),
                                  ('M', '박민우', 'mw@itg-aircon.com', MW)):
        hr.upsert_from_flex({'flex_id': fid, 'name': name, 'status': '재직', 'login_email': email, 'work_email': email,
                             'slack_id': sid}, [])
    yield d
    gn._cache.clear()
    gw._invalidate()


def _feed(email, field=False, admin=False):
    gn._cache.clear()
    return gn.feed(email, hr.find_by_email(email), field, admin)


def test_approval_turn_and_progress(db):
    row = ea.create(ea.FORM_FUND, MW, '박민우', {'amount': 330000, 'partner': '알만상사', 'grn': 'G'})
    k = _feed('kiko@itg-aircon.com')
    assert any('승인 차례' in i['title'] and i['href'] == f"/approval?box=todo&id={row['id']}" for i in k['items'])
    assert k['unread'] >= 1
    ea.transition(row['id'], 'approve', KIKO, '고광일')
    assert not any('승인 차례' in i['title'] for i in _feed('kiko@itg-aircon.com')['items'])   # 처리하면 사라짐
    mw = _feed('mw@itg-aircon.com')['items']
    assert any(i['title'].endswith('고광일님이 승인했어요') and i['cat'] == 'approval' for i in mw)
    assert any('최종 승인 차례' in i['title'] for i in _feed('jw@itg-aircon.com')['items'])
    ea.transition(row['id'], 'reject', JW, '박정우', '증빙 없음')
    rej = [i for i in _feed('mw@itg-aircon.com')['items'] if '반려' in i['title']]
    assert rej and rej[0]['body'] == '증빙 없음' and rej[0]['icon'] == '❌'
    assert not any('박정우님이 반려' in i['title'] for i in _feed('jw@itg-aircon.com')['items'])     # 내가 한 일은 안 옴


def test_leave_turn_and_result(db):
    lv.upsert_staff('박민우', '2022-02-03', MW)
    lv.add_grant('박민우', 8, date(2026, 2, 3), date(2027, 2, 2), source='t')
    r = lv.create('박민우', MW, MW, '박민우', lv.ANNUAL, '', date(2026, 11, 20), date(2026, 11, 20))
    jw = _feed('jw@itg-aircon.com')['items']
    assert any(i['title'] == '박민우님 휴가 승인 차례예요' and i['cat'] == 'leave' for i in jw)
    lv.transition(r['id'], 'approve', JW, '박정우')
    mine = _feed('mw@itg-aircon.com')['items']
    assert any(i['title'] == '박정우님이 휴가를 승인했어요' and i['href'] == '/leave' for i in mine)


def test_yearend_and_admin(db):
    me = hr.find_by_email('mw@itg-aircon.com')
    ye.set_season(2026, 'kiko@itg-aircon.com', status='open', due_date='2027-01-31')
    assert any(i['title'] == '2026년 귀속 연말정산이 시작됐어요' for i in _feed('mw@itg-aircon.com')['items'])
    assert any(i['title'] == '2026년 귀속 연말정산이 시작됐어요' for i in _feed('', field=True)['items']) is False  # 인사카드 없으면 없음
    ye.save_answers(me['id'], 2026, {'method': 'bulk', 'family_status': 'none', 'bulk_consent': True})
    ye.submit(me['id'], 2026, me)
    adm = _feed('kiko@itg-aircon.com', admin=True)['items']
    assert any(i['title'] == '박민우님이 연말정산 자료를 냈어요' and i['href'] == '/yearend/admin' for i in adm)
    assert not any('자료를 냈어요' in i['title'] for i in _feed('kiko@itg-aircon.com', admin=False)['items'])
    ye.set_status(me['id'], 2026, ye.RETURNED, 'kiko@itg-aircon.com', '등본 누락')
    ret = [i for i in _feed('mw@itg-aircon.com')['items'] if i['title'] == '연말정산 보완 요청이 왔어요']
    assert ret and ret[0]['body'] == '등본 누락'


def test_seen_and_field_scope(db):
    ea.create(ea.FORM_FUND, MW, '박민우', {'amount': 1000, 'partner': 'x', 'grn': 'G'})
    who = gn.who_key('kiko@itg-aircon.com', None, False)
    assert _feed('kiko@itg-aircon.com')['unread'] >= 1
    gn.mark_seen(who)
    f = _feed('kiko@itg-aircon.com')
    assert f['unread'] == 0 and all(not i['unread'] for i in f['items'])
    with db._get_connection() as conn:                          # 본 뒤에 생긴 일만 다시 빨간 점
        conn.execute("UPDATE gw_notify_seen SET seen_at='2000-01-01 00:00:00' WHERE who=?", (who,))
        conn.commit()
    assert _feed('kiko@itg-aircon.com')['unread'] >= 1
    # 현장 세션은 결재 알림 없음
    assert all(i['cat'] != 'approval' for i in gn.feed('', hr.find_by_email('kiko@itg-aircon.com'), True, False)['items'])
    assert gn.who_key('', {'id': 7}, True) == 'field:7'


@pytest.fixture
def client(db, monkeypatch):
    import dashboard.auth as auth
    from flask import Blueprint, session
    real = auth._validate_session
    monkeypatch.setattr(auth, '_validate_session',
                        lambda require_admin=False: real(require_admin) if session.get('user', {}).get('gw_field') else None)
    from dashboard.blueprints.groupware import groupware_bp
    from dashboard.blueprints.hr import hr_bp
    app = Flask('t')
    app.secret_key = 't'
    app.register_blueprint(groupware_bp)
    app.register_blueprint(hr_bp)
    a = Blueprint('auth', __name__)
    a.add_url_rule('/login', 'login_page', lambda: 'login')
    app.register_blueprint(a)
    gw.set_open_all(True, 'kiko@itg-aircon.com')
    hr.upsert_from_flex({'flex_id': 'F', 'name': '김현장', 'status': '재직', 'access': 'field', 'slack_id': 'UF'}, [])
    return app.test_client()


def test_api_staff_and_field(client):
    ea.create(ea.FORM_FUND, MW, '박민우', {'amount': 1000, 'partner': 'x', 'grn': 'G'})
    with client.session_transaction() as s:
        s['user'] = {'email': 'kiko@itg-aircon.com', 'name': '고광일', 'permission_level': 'super_admin'}
    gn._cache.clear()
    j = client.get('/api/gw/notify').get_json()
    assert j['success'] and j['unread'] >= 1
    assert client.post('/api/gw/notify/seen').get_json()['success']
    assert client.get('/api/gw/notify').get_json()['unread'] == 0

    tok = gw_login.issue('UF').rsplit('/', 1)[1]
    with client.session_transaction() as s:
        s.clear()
    client.get(f'/gw/login/{tok}')
    r = client.get('/api/gw/notify')                                     # 현장도 알림은 봄(휴가·연말정산)
    assert r.status_code == 200 and all(i['cat'] != 'approval' for i in r.get_json()['items'])
    assert client.post('/api/gw/notify/seen').status_code == 200
