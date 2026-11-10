"""그룹웨어 홈(2026-11-10) — 결재 차례 판단·휴가자·메일 칸·현장 제한."""
from datetime import date

import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import e_approval as ea
from dashboard.services import groupware as gw
from dashboard.services import gw_home
from dashboard.services import gw_login
from dashboard.services import hr
from dashboard.services import leave as lv
from dashboard.services import leave_view

KIKO, JW, SB, MW = 'U04UL2ZLJAX', 'U051ENK8AAY', 'U0BHC2JV7U5', 'UMW'


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, gw_login, lv, leave_view, ea):
        monkeypatch.setattr(mod, '_db', lambda: d)
        if hasattr(mod, '_table_ready'):
            monkeypatch.setattr(mod, '_table_ready', False)
    for k in ('EAPPROVAL_EXECUTOR_DM_IDS', 'EAPPROVAL_APPROVER_IDS', 'EAPPROVAL_FINAL_APPROVER_IDS',
              'EAPPROVAL_N_EXECUTOR_IDS', 'SLACK_SETTLEMENT_CHECKER_ID', 'SLACK_SETTLEMENT_CHECKER_EMAIL',
              'EAPPROVAL_FINAL_ENABLED'):
        monkeypatch.delenv(k, raising=False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    gw_home._mail_cache.clear()
    gw._invalidate()
    yield d
    gw._invalidate()


def _approval(d, status, requester, form='fund', grn='G'):
    with d._get_connection() as conn:
        ea._ensure_tables(conn)
        cur = conn.execute('INSERT INTO approvals (form, requester_id, requester_name, status, amount, grn, card_channel, card_ts) '
                           "VALUES (?, ?, '요청자', ?, 1000, ?, 'C1', '1700000000.123456')", (form, requester, status, grn))
        conn.commit()
        return cur.lastrowid


def test_whose_turn(db):
    p = _approval(db, ea.PENDING, MW)
    f = _approval(db, ea.FINAL, MW)
    a = _approval(db, ea.APPROVED, MW)
    r = _approval(db, ea.REVISE, MW)
    _approval(db, ea.DONE, MW)
    own = _approval(db, ea.PENDING, KIKO)
    k = gw_home.approvals(KIKO)
    assert [x['id'] for x in k['todo']] == [p]                      # 승인 대기(본인 요청 제외)
    assert k['mine_count'] == 1 and k['mine'][0]['id'] == own
    assert [x['id'] for x in gw_home.approvals(JW)['todo']] == [f]  # 최종 승인 대기
    assert [x['id'] for x in gw_home.approvals(SB)['todo']] == [a]  # 이체 대기(G 통장 → 경영지원)
    m = gw_home.approvals(MW)
    assert [x['id'] for x in m['todo']] == [r] and m['mine_count'] == 4    # 수정 요청은 요청자 차례
    assert m['todo'][0]['link'].endswith('/archives/C1/p1700000000123456')
    assert gw_home.approvals('')['linked'] is False


def test_people_off(db):
    leave_view.upsert_flex_history([
        {'flex_event_id': 'A', 'name': '오늘', 'start': '2026-11-10', 'end': '2026-11-10', 'status': '승인', 'unit': 'HALF_DAY_PM'},
        {'flex_event_id': 'B', 'name': '다음주', 'start': '2026-11-16', 'end': '2026-11-17', 'status': '승인'},
        {'flex_event_id': 'C', 'name': '대기', 'start': '2026-11-11', 'end': '2026-11-11', 'status': '승인 대기'},
        {'flex_event_id': 'D', 'name': '멀리', 'start': '2026-11-30', 'end': '2026-11-30', 'status': '승인'},
        {'flex_event_id': 'E', 'name': '월넘김', 'start': '2026-12-01', 'end': '2026-12-01', 'status': '승인'},
    ])
    o = gw_home.people_off(date(2026, 11, 10))
    assert [(x['name'], x['half']) for x in o['today']] == [('오늘', '오후')]
    assert [x['name'] for x in o['soon']] == ['다음주']                   # 대기·7일 밖 제외
    o2 = gw_home.people_off(date(2026, 11, 28))
    assert [x['name'] for x in o2['soon']] == ['멀리', '월넘김']           # 달을 넘겨도


class _Fake:
    calls = 0

    def users(self): return self
    def labels(self): return self
    def messages(self): return self

    def get(self, userId=None, id=None, **kw):
        self._k = ('get', id)
        return self

    def list(self, **kw):
        self._k = ('list', None)
        return self

    def execute(self):
        _Fake.calls += 1
        kind, mid = self._k
        if kind == 'list':
            return {'messages': [{'id': '1'}]}
        if mid == 'INBOX':
            return {'messagesUnread': 3}
        return {'threadId': 'T1', 'snippet': 's', 'labelIds': ['UNREAD'],
                'payload': {'headers': [{'name': 'From', 'value': '김철수 <a@b.c>'}, {'name': 'Subject', 'value': '제목'},
                                        {'name': 'Date', 'value': 'Tue, 10 Nov 2026 09:12:00 +0900'}]}}


def test_mail_only_company_own_and_cached(db, monkeypatch):
    seen = []
    monkeypatch.setattr(gw_home, '_gmail', lambda email: seen.append(email) or _Fake())
    r = gw_home.mail_summary('KiKO@itg-aircon.com')
    assert r['ok'] and r['unread'] == 3 and r['items'][0]['from'] == '김철수' and r['items'][0]['unread']
    assert 'authuser=kiko@itg-aircon.com#inbox/T1' in r['items'][0]['link']
    assert seen == ['kiko@itg-aircon.com']
    gw_home.mail_summary('kiko@itg-aircon.com')
    assert seen == ['kiko@itg-aircon.com']                                 # 60초 기억
    assert gw_home.mail_summary('me@gmail.com') == {'ok': False, 'reason': 'no_company_mail'}

    def boom(email):
        raise RuntimeError('google down')
    monkeypatch.setattr(gw_home, '_gmail', boom)
    assert gw_home.mail_summary('jw@itg-aircon.com')['ok'] is False        # 실패해도 홈은 뜸


@pytest.fixture
def client(db, monkeypatch):
    import dashboard.auth as auth
    from flask import Blueprint, session
    real = auth._validate_session
    monkeypatch.setattr(auth, '_validate_session',
                        lambda require_admin=False: real(require_admin) if session.get('user', {}).get('gw_field') else None)
    monkeypatch.setattr(gw_home, '_gmail', lambda email: _Fake())
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
    hr.upsert_from_flex({'flex_id': 'K', 'name': '고광일', 'status': '재직', 'login_email': 'kiko@itg-aircon.com',
                         'work_email': 'kiko@itg-aircon.com', 'slack_id': KIKO}, [])
    hr.upsert_from_flex({'flex_id': 'F', 'name': '김현장', 'status': '재직', 'access': 'field', 'slack_id': 'UF'}, [])
    return app.test_client()


def test_home_api_staff_and_field(client, db):
    _approval(db, ea.PENDING, MW)
    with client.session_transaction() as s:
        s['user'] = {'email': 'kiko@itg-aircon.com', 'name': '고광일', 'permission_level': 'super_admin'}
    j = client.get('/api/home').get_json()
    assert j['approvals']['todo_count'] == 1 and 'off' in j and j['leave'] is not None
    assert client.get('/api/home/mail').get_json()['mail']['unread'] == 3

    tok = gw_login.issue('UF').rsplit('/', 1)[1]
    with client.session_transaction() as s:
        s.clear()
    r = client.get(f'/gw/login/{tok}')
    assert r.headers['Location'].endswith('/home')
    j = client.get('/api/home').get_json()
    assert j['field'] and 'approvals' not in j and 'off' not in j       # 현장: 내 휴가만
    assert client.get('/api/home/mail').status_code == 403
    assert client.get('/projects').headers['Location'].endswith('/home')


def test_nav_home_first():
    assert gw.build_nav('/home', 'editor')['areas'][0]['label'] == '홈'
    assert [a['label'] for a in gw.build_nav('/home', 'field', field=True)['areas']] == ['홈', '휴가', '인사']


class _FakeBatchSvc(_Fake):
    """묶음 요청(batch)을 흉내 — 구글 왕복 수를 센다."""
    trips = 0

    def list(self, **kw):
        _FakeBatchSvc.trips += 1
        return super().list(**kw)

    def new_batch_http_request(self, callback):
        svc = self

        class _B:
            def __init__(self):
                self.reqs = []

            def add(self, req, request_id):
                self.reqs.append((request_id, req._k))

            def execute(self):
                _FakeBatchSvc.trips += 1
                for rid, k in self.reqs:
                    svc._k = k
                    callback(rid, svc.execute(), None)
        return _B()

    def get(self, userId=None, id=None, **kw):
        req = _Fake()
        req._k = ('get', id)
        return req

    def labels(self):
        return self


def test_mail_batch_two_trips_and_stale_while_revalidate(db, monkeypatch):
    _FakeBatchSvc.trips = 0
    r = gw_home.mail_summary('kiko@itg-aircon.com', service=_FakeBatchSvc())
    assert r['ok'] and r['unread'] == 3 and len(r['items']) == 1
    assert _FakeBatchSvc.trips == 2                                   # 목록 1번 + 묶음 1번

    calls = []
    monkeypatch.setattr(gw_home, '_refresh_in_background', lambda email, limit: calls.append(email))
    gw_home._mail_cache['kiko@itg-aircon.com'] = (gw_home.time.time() - 120, r)      # 2분 지난 값
    assert gw_home.mail_summary('kiko@itg-aircon.com') is r and calls == ['kiko@itg-aircon.com']
    gw_home._mail_cache['kiko@itg-aircon.com'] = (gw_home.time.time() - 3600, r)     # 1시간 지난 값 → 새로 받음
    monkeypatch.setattr(gw_home, '_gmail', lambda email: _Fake())
    assert gw_home.mail_summary('kiko@itg-aircon.com')['unread'] == 3 and calls == ['kiko@itg-aircon.com']
