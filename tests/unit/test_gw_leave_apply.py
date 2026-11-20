"""그룹웨어 본인 휴가 신청·취소(2026-11-15) + 홈 화면 앱 설치 정보.

슬랙 휴가 앱과 같은 장부(leave.create·transition)·같은 알림(slack_leave.start·apply_action)을 쓰는지,
현장 세션은 본인 것만 되는지, 열기 규칙(is_open)을 따르는지 본다. 슬랙은 가짜 클라이언트로 대신한다.
"""
from datetime import date, timedelta

import pytest
from cryptography.fernet import Fernet
from flask import Blueprint, Flask

from dashboard.services import groupware as gw
from dashboard.services import gw_login
from dashboard.services import hr
from dashboard.services import leave as lv
from dashboard.services import leave_view
from dashboard.services import leave_web

STAFF = 'mw@itg-aircon.com'
OTHER = 'other@itg-aircon.com'
ADMIN = 'kiko@itg-aircon.com'
JW, YG = 'UJW', 'UYG'


class FakeSlack:
    """슬랙 WebClient 대신 — 보낸 것만 적어 둔다(운영 슬랙으로 아무것도 안 나감)."""

    def __init__(self):
        self.posts, self.updates = [], []
        self._n = 0

    def conversations_open(self, users):
        return {'channel': {'id': f'D-{users}'}}

    def chat_postMessage(self, channel, text='', blocks=None):
        self._n += 1
        self.posts.append({'ch': channel, 'text': text, 'blocks': blocks})
        return {'ts': f'1.{self._n}'}

    def chat_update(self, channel, ts, text='', blocks=None):
        self.updates.append({'ch': channel, 'ts': ts, 'text': text, 'blocks': blocks})
        return {'ok': True}

    def to(self, uid):
        return [p for p in self.posts if p['ch'] == f'D-{uid}']


def _next_workday(n=1):
    d = date.today()
    while n:
        d += timedelta(days=1)
        if lv.is_workday(d):
            n -= 1
    return d


@pytest.fixture
def slack(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, gw_login, lv, leave_view):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.setenv('LEAVE_APPROVER_IDS', f'{JW},{YG}')
    monkeypatch.setenv('LEAVE_ANNOUNCE_CHANNEL', 'C-ANNOUNCE')
    monkeypatch.setenv('LEAVE_PILOT_IDS', 'UMW,UF')
    monkeypatch.delenv('LEAVE_ENABLED', raising=False)
    gw._invalidate()
    hr.upsert_from_flex({'flex_id': 'P1', 'name': '박민우', 'status': '재직', 'work_email': STAFF,
                         'login_email': STAFF, 'slack_id': 'UMW', 'hire_date': '2022-02-03'}, [])
    hr.upsert_from_flex({'flex_id': 'P2', 'name': '이직원', 'status': '재직', 'work_email': OTHER,
                         'login_email': OTHER, 'slack_id': 'UOT', 'hire_date': '2022-02-03'}, [])
    hr.upsert_from_flex({'flex_id': 'P3', 'name': '김미화', 'status': '재직', 'access': 'field', 'slack_id': 'UF'}, [])
    for name, sid in (('박민우', 'UMW'), ('이직원', 'UOT'), ('김미화', 'UF')):
        lv.upsert_staff(name, '2022-02-03', sid)
        lv.add_grant(name, 10, date.today() - timedelta(days=30), date.today() + timedelta(days=300), source=f't:{name}')

    fake = FakeSlack()
    from dashboard.blueprints import slack_eapproval as se
    from dashboard.blueprints import slack_leave
    monkeypatch.setattr(slack_leave, '_client', lambda: fake)
    monkeypatch.setattr(se, '_name', lambda uid: {'UMW': '박민우', 'UF': '김미화', JW: '박정우', YG: '박용구'}.get(uid, uid))
    monkeypatch.setattr(se, '_notify', lambda client, uid, text: client.chat_postMessage(channel=f'D-{uid}', text=text))
    monkeypatch.setattr(leave_web, '_run_bg', lambda fn: fn())          # 알림을 바로(시험)
    # 구글 로그인 직원: 로그인 메일로 슬랙에 물어본 ID = 인사카드 ID 라고 친다
    import dashboard.services.gw_approval as ga
    monkeypatch.setattr(ga, 'slack_id_status', lambda email, card: (card, 'ok') if card else ('', 'unlinked'))
    yield fake
    gw._invalidate()


@pytest.fixture
def client(slack, monkeypatch):
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
    a = Blueprint('auth', __name__)
    a.add_url_rule('/login', 'login_page', lambda: 'login')
    app.register_blueprint(a)
    gw.set_open_all(True, ADMIN)
    return app.test_client()


def _as(c, email, perm='editor'):
    with c.session_transaction() as s:
        s.clear()
        s['user'] = {'email': email, 'name': 'x', 'permission_level': perm}


def _field_login(c, sid='UF'):
    tok = gw_login.issue(sid).rsplit('/', 1)[1]
    with c.session_transaction() as s:
        s.clear()
    assert c.get(f'/gw/login/{tok}').status_code == 302


# ── 신청 ───────────────────────────────────────────────────────────────
def test_apply_goes_through_same_ledger_and_slack_flow(client, slack):
    _as(client, STAFF)
    me = client.get('/api/leave/me').get_json()
    assert me['form']['can_apply'] and me['form']['name'] == '박민우' and me['form']['balance'] == 10
    assert [t['value'] for t in me['form']['types']] == [k for k, _ in lv.type_options()]   # 슬랙 창과 같은 종류

    d = _next_workday(3)
    r = client.post('/api/leave/apply', json={'type': 'annual', 'start': d.isoformat(), 'reason': '개인 사정'})
    assert r.status_code == 200, r.get_json()
    rid = r.get_json()['id']
    row = lv.get(rid)
    assert row['status'] == lv.WAIT_1 and row['slack_id'] == 'UMW' and row['submitter_id'] == 'UMW' and row['days'] == 1
    assert lv.balance('박민우', d) == 9                                   # 대기 중에도 떼어 둠
    # 슬랙 신청과 같은 알림: 신청자 카드 + JW 승인 카드(YG 는 아직)
    assert slack.to('UMW') and slack.to(JW) and not slack.to(YG)
    assert any(b.get('action_id') == lv.A_APPROVE for blk in slack.to(JW)[0]['blocks'] if blk['type'] == 'actions'
               for b in blk['elements'])

    # 승인(슬랙 버튼과 같은 함수) → YG 카드 → YG 승인 → 확정 알림 + #휴가_일정
    from dashboard.blueprints import slack_leave
    assert slack_leave.apply_action('approve', rid, JW) is None and slack.to(YG)
    assert slack_leave.apply_action('approve', rid, YG) is None
    assert lv.get(rid)['status'] == lv.APPROVED
    assert [p['text'] for p in slack.posts if p['ch'] == 'C-ANNOUNCE'] == [lv.announce_text(lv.get(rid))]

    # 내 휴가 목록에 [신청 취소] → 웹에서 취소 → 장부·알림이 슬랙 취소와 같게
    entry = [e for e in client.get('/api/leave/me').get_json()['leave']['entries'] if e['source'] == 'slack'][0]
    assert entry['cancel'] == 'all' and entry['cancel_label'] == '신청 취소'
    assert entry['cancel_confirm'] == lv.cancel_confirm_text('all')
    assert client.post(f'/api/leave/{rid}/cancel').get_json()['success']
    assert lv.get(rid)['status'] == lv.CANCELLED and lv.balance('박민우', d) == 10
    assert any(u['ch'] == 'C-ANNOUNCE' and '(취소)' in u['text'] for u in slack.updates)
    assert any('취소됐어요' in p['text'] for p in slack.to(JW))           # 확정 휴가 취소 → 승인자 알림
    entry = [e for e in client.get('/api/leave/me').get_json()['leave']['entries'] if e['id'] == rid][0]
    assert entry['cancel'] == '' and entry['status'] == leave_view.ST_CANCEL


def test_apply_validation_same_as_slack(client):
    _as(client, STAFF)
    d = _next_workday(2)
    r = client.post('/api/leave/apply', json={'type': 'annual', 'start': d.isoformat(), 'reason': ''})
    assert r.status_code == 400 and r.get_json()['errors'] == {'reason': '사유를 적어 주세요.'}
    r = client.post('/api/leave/apply', json={'type': 'event', 'start': d.isoformat(), 'reason': '결혼'})
    assert r.get_json()['errors'] == {'sub': '경조사 종류를 골라 주세요.'}
    r = client.post('/api/leave/apply', json={'type': 'annual', 'start': d.isoformat(),
                                             'end': (d - timedelta(days=1)).isoformat(), 'reason': 'x'})
    assert r.get_json()['errors'] == {'end': '마지막 날이 시작일보다 빨라요.'}
    r = client.post('/api/leave/apply', json={'type': 'reward', 'start': d.isoformat(), 'reason': 'x'})   # 관리자 부여 휴가
    assert 'type' in r.get_json()['errors']
    r = client.post('/api/leave/apply', json={'type': 'annual', 'start': d.isoformat(),
                                             'end': _next_workday(30).isoformat(), 'reason': '길게'})
    assert '남은 연차가' in r.get_json()['errors']['start']
    # 반차는 마지막 날을 보내도 그날 하루(0.5)
    r = client.post('/api/leave/apply', json={'type': 'half_pm', 'start': d.isoformat(),
                                             'end': _next_workday(5).isoformat(), 'reason': '병원'})
    assert r.status_code == 200 and lv.get(r.get_json()['id'])['days'] == 0.5
    # 같은 날 오후 반차 또 → 겹침
    r = client.post('/api/leave/apply', json={'type': 'half_pm', 'start': d.isoformat(), 'reason': '병원'})
    assert '겹쳐요' in r.get_json()['errors']['start']


def test_is_open_rule(client, monkeypatch):
    monkeypatch.setenv('LEAVE_PILOT_IDS', 'USOMEONE')
    _as(client, STAFF)
    form = client.get('/api/leave/me').get_json()['form']
    assert form == {'can_apply': False, 'why': leave_web.MSG_NOT_OPEN}
    r = client.post('/api/leave/apply', json={'type': 'annual', 'start': _next_workday().isoformat(), 'reason': 'x'})
    assert r.status_code == 400 and r.get_json()['message'] == leave_web.MSG_NOT_OPEN
    monkeypatch.setenv('LEAVE_ENABLED', 'true')                         # 플렉스에서 옮긴 뒤 = 전원
    assert client.get('/api/leave/me').get_json()['form']['can_apply']


def test_unverified_slack_id_cannot_apply(client, monkeypatch):
    import dashboard.services.gw_approval as ga
    monkeypatch.setattr(ga, 'slack_id_status', lambda email, card: ('', 'unlinked'))   # 카드 ID ≠ 슬랙이 아는 ID
    _as(client, STAFF)
    assert client.get('/api/leave/me').get_json()['form']['why'] == leave_web.MSG_NO_SLACK
    monkeypatch.setattr(ga, 'slack_id_status', lambda email, card: ('', 'unavailable'))
    assert client.get('/api/leave/me').get_json()['form']['why'] == leave_web.MSG_SLACK_DOWN


def test_cannot_cancel_someone_elses_leave(client):
    d = _next_workday(4)
    other = lv.create('이직원', 'UOT', 'UOT', '이직원', 'annual', '', d, d, 'x')
    _as(client, STAFF)
    r = client.post(f"/api/leave/{other['id']}/cancel")
    assert r.status_code == 400 and r.get_json()['message'] == '본인 휴가만 취소할 수 있어요.'
    assert lv.get(other['id'])['status'] == lv.WAIT_1


def test_rest_cancel_for_leave_in_progress(slack):
    """쓰는 중인 휴가 → [남은 날 취소]: 어제까지 남기고 오늘부터 취소(슬랙과 같은 cut_rest)."""
    today = date(2026, 11, 18)        # 수
    start, end = date(2026, 11, 16), date(2026, 11, 20)
    row = lv.create('박민우', 'UMW', 'UMW', '박민우', 'annual', '', start, end, 'x')
    with lv._conn() as conn:
        conn.execute('UPDATE leave_requests SET status=? WHERE id=?', (lv.APPROVED, row['id']))
    row = lv.get(row['id'])
    assert lv.own_cancel_mode(row, 'UMW', today) == 'rest'
    assert lv.own_cancel_mode(row, 'UOT', today) == ''                # 남의 휴가엔 단추 없음
    assert lv.own_cancel_mode(row, 'UMW', date(2026, 11, 21)) == ''   # 끝난 휴가 → 경영지원만
    b = lv.cancel_button(row, 'UMW', today)
    assert b['text']['text'] == '남은 날 취소' and b['confirm']['text']['text'] == lv.cancel_confirm_text('rest')
    assert b['confirm']['confirm']['text'] == '취소하기' and b['confirm']['deny']['text'] == '그대로 두기'
    s = {'entries': [{'source': 'slack', 'id': row['id']}, {'source': 'flex', 'id': 'E1'}]}
    out = leave_web.annotate(s, {'email': STAFF}, today)['entries']
    assert out[0]['cancel'] == 'rest' and out[0]['cancel_label'] == '남은 날 취소' and out[1]['cancel'] == ''


def test_requester_card_button_unchanged(slack):
    """요청자 DM 카드의 취소 단추는 예전과 같다(대기 = 신청 취소). 11-18 부터 그 앞에 [✏️ 고치기]."""
    d = _next_workday(2)
    row = lv.create('박민우', 'UMW', 'UMW', '박민우', 'annual', '', d, d, 'x')
    acts = [b for blk in lv.card_blocks(row, [], 'requester') if blk['type'] == 'actions' for b in blk['elements']]
    assert [b['text']['text'] for b in acts] == ['✏️ 수정', '신청 취소']
    assert [b['action_id'] for b in acts] == [lv.A_EDIT, lv.A_CANCEL]
    assert acts[1]['confirm']['text']['text'] == '이 휴가를 취소할까요? 쓴 연차는 돌아와요.'


def test_slack_mine_view_has_cancel_buttons(slack):
    from dashboard.blueprints import slack_leave
    d = _next_workday(2)
    row = lv.create('박민우', 'UMW', 'UMW', '박민우', 'annual', '', d, d, 'x')
    v = slack_leave.my_mine_view('UMW')
    assert v['callback_id'] == slack_leave.CB_MINE
    btns = [b['accessory'] for b in v['blocks'] if b.get('accessory')]
    assert [(b['action_id'], b['value'], b['text']['text']) for b in btns] == [(lv.A_CANCEL, str(row['id']), '신청 취소')]
    assert slack_leave.mine_cancel_blocks('UOT') == []                 # 취소할 휴가 없음


# ── 현장 세션 ──────────────────────────────────────────────────────────
def test_field_session_can_apply_and_cancel_own_only(client, slack):
    _field_login(client)
    with client.session_transaction() as s:
        assert s['user']['gw_slack_id'] == 'UF'                     # 링크를 누른 슬랙 계정
    me = client.get('/api/leave/me').get_json()
    assert me['form']['can_apply'] and me['form']['name'] == '김미화'
    d = _next_workday(2)
    r = client.post('/api/leave/apply', json={'type': 'annual', 'start': d.isoformat(), 'reason': '집안일'})
    assert r.status_code == 200
    rid = r.get_json()['id']
    assert lv.get(rid)['slack_id'] == 'UF' and slack.to(JW)
    # 남의 휴가·전사 현황은 여전히 막힘
    other = lv.create('이직원', 'UOT', 'UOT', '이직원', 'annual', '', d, d, 'x')
    assert client.post(f"/api/leave/{other['id']}/cancel").status_code == 400
    assert client.get('/api/leave/team').status_code == 403
    assert client.get('/api/leave/calendar').status_code == 403
    assert client.get(f"/api/leave/person/{hr.find_by_email(STAFF)['id']}").status_code == 403
    assert client.post(f'/api/leave/{rid}/cancel').get_json()['success']
    assert lv.get(rid)['status'] == lv.CANCELLED


def test_field_session_without_link_slack_id_must_relogin(client):
    p = hr.find_by_slack('UF')
    with client.session_transaction() as s:
        s.clear()
        s['user'] = dict(gw_login.session_user(p), gw_slack_id='')    # 이 기능 전에 만든 현장 세션
    assert client.get('/api/leave/me').get_json()['form']['why'] == leave_web.MSG_RELOGIN
    # 링크로 들어온 뒤 인사카드 슬랙 ID 가 바뀌면 → 다시 로그인
    with client.session_transaction() as s:
        s['user'] = dict(gw_login.session_user(p), gw_slack_id='UOT')
    r = client.post('/api/leave/apply', json={'type': 'annual', 'start': _next_workday().isoformat(), 'reason': 'x'})
    assert r.status_code == 400 and r.get_json()['message'] == leave_web.MSG_RELOGIN


def test_field_gatekeeper_paths():
    from dashboard.blueprints.groupware import field_path_allowed
    assert field_path_allowed('POST', '/api/leave/apply')
    assert field_path_allowed('POST', '/api/leave/12/cancel')
    assert field_path_allowed('GET', '/manifest.webmanifest')
    assert not field_path_allowed('GET', '/api/leave/apply')
    assert not field_path_allowed('POST', '/api/leave/12/approve')
    assert not field_path_allowed('GET', '/api/leave/team')
    assert not field_path_allowed('POST', '/api/leave/person/3')


# ── 홈 화면 앱 ─────────────────────────────────────────────────────────
def test_manifest_without_login(client):
    with client.session_transaction() as s:
        s.clear()
    r = client.get('/manifest.webmanifest')
    assert r.status_code == 200 and r.mimetype == 'application/manifest+json'
    m = r.get_json(force=True)
    assert m['name'] == '아이티글로벌' and m['start_url'] == '/home' and m['display'] == 'standalone'
    sizes = {i['sizes']: i['src'] for i in m['icons']}
    assert set(sizes) == {'192x192', '512x512', '180x180'}
    assert sizes['512x512'].startswith('/static/images/icon-512.png?v=')
    import os
    root = os.path.join(os.path.dirname(__file__), '..', '..', 'dashboard', 'static', 'images')
    import struct
    with open(os.path.join(root, 'icon-512.png'), 'rb') as fh:     # PNG 머리(IHDR)의 가로·세로
        head = fh.read(24)
    assert head[:8] == b'\x89PNG\r\n\x1a\n' and struct.unpack('>II', head[16:24]) == (512, 512)


def test_web_app_head_only_in_groupware_layout_and_login():
    import os
    root = os.path.join(os.path.dirname(__file__), '..', '..', 'dashboard', 'templates')
    base = open(os.path.join(root, 'modern_base.html'), encoding='utf-8').read()
    head = base.split('</head>')[0]
    inc = "{% include 'components/_web_app_head.html' %}"
    assert inc in head and head.index('{% if gw_layout %}', head.index('Favicon')) < head.index(inc)
    assert inc in open(os.path.join(root, 'login.html'), encoding='utf-8').read()
    part = open(os.path.join(root, 'components', '_web_app_head.html'), encoding='utf-8').read()
    assert 'rel="manifest"' in part and 'apple-mobile-web-app-capable' in part
    assert 'serviceWorker' not in part                                  # 서비스 워커는 다시 켜지 않는다


def test_install_menu_in_groupware_layout_only(tmp_path):
    """사용자 메뉴 [앱으로 설치](다우오피스처럼, 11-15) — 새 틀에만, 안내 창·스크립트·설치 신호 받기가 같이 있다."""
    import os
    root = os.path.join(os.path.dirname(__file__), '..', '..', 'dashboard')
    menu = open(os.path.join(root, 'templates', 'components', '_user_menu.html'), encoding='utf-8').read()
    assert '{% if gw_layout %}' in menu and menu.index('{% if gw_layout %}') < menu.index('id="gwInstallBtn"')
    shell = open(os.path.join(root, 'templates', 'components', '_gw_shell.html'), encoding='utf-8').read()
    assert 'id="gwInstallModal"' in shell and "gw_asset('js/gw-install.js')" in shell
    for osname in ('ios', 'android', 'pc'):
        assert f'data-os="{osname}"' in shell
    head = open(os.path.join(root, 'templates', 'components', '_web_app_head.html'), encoding='utf-8').read()
    assert 'beforeinstallprompt' in head and '__gwInstallEvt' in head
    js = open(os.path.join(root, 'static', 'js', 'gw-install.js'), encoding='utf-8').read()
    assert '__gwInstallEvt' in js and 'display-mode: standalone' in js and 'serviceWorker' not in js
