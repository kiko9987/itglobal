# -*- coding: utf-8 -*-
"""보안 점검(2026-11-13) 낮음 우선순위 수정 회귀 테스트. create_app 미사용."""
import hashlib
import hmac
import inspect
import io
import json
import logging
import time
from unittest.mock import patch

import pytest
from flask import Flask, session

from dashboard.blueprints import slack_bot as sb


# ── 슬랙 관리자 정비 링크 ─────────────────────────────────────────────
@pytest.fixture
def app():
    a = Flask(__name__)
    a.secret_key = 'test'
    return a


@pytest.mark.parametrize('sfs', ['cross-site', 'same-site'])
def test_manual_trigger_rejects_cross_site(app, sfs):
    with app.test_request_context('/slack/sync-karrot', headers={'Sec-Fetch-Site': sfs}), \
         patch('dashboard.auth._validate_session', return_value=None):
        r = sb._manual_trigger_denied()
    assert r is not None and r[1] == 403


@pytest.mark.parametrize('headers', [{'Sec-Fetch-Site': 'none'}, {'Sec-Fetch-Site': 'same-origin'}, {}])
def test_manual_trigger_allows_typed_or_same_origin_admin(app, headers):
    with app.test_request_context('/slack/sync-karrot', headers=headers), \
         patch('dashboard.auth._validate_session', return_value=None):
        session['user'] = {'email': 'a@x', 'permission_level': 'admin'}
        assert sb._manual_trigger_denied() is None


def test_manual_trigger_no_session_denied(app):
    with app.test_request_context('/slack/sync-karrot', headers={'Sec-Fetch-Site': 'none'}):
        assert sb._manual_trigger_denied()[1] == 403


def test_manual_trigger_uses_full_session_check(app):
    with app.test_request_context('/slack/sync-karrot', headers={'Sec-Fetch-Site': 'none'}), \
         patch('dashboard.auth._validate_session', return_value=('expired', 401)) as vs:
        session['user'] = {'email': 'a@x', 'permission_level': 'admin'}
        r = sb._manual_trigger_denied()
    vs.assert_called_once_with(require_admin=True)
    assert r[1] == 403


def test_migrate_visit_buttons_skips_cancelled_cards(monkeypatch):
    cancelled = {
        'ts': str(time.time()),
        'blocks': [
            {'type': 'section', 'text': {'type': 'mrkdwn', 'text': '🚫 방문 취소\n```🔔 새 방문 일정 `L-00001````'}},
            {'type': 'actions', 'elements': [{'action_id': 'visit_uncancel', 'type': 'button'}]},
        ],
    }

    class _Client:
        def __init__(self):
            self.updates = []

        def conversations_history(self, **kw):
            return {'messages': [cancelled], 'response_metadata': {}, 'has_more': False}

        def chat_update(self, **kw):
            self.updates.append(kw)
            return {'ok': True}

    c = _Client()

    class _App:
        client = c

    monkeypatch.setenv('SLACK_VISIT_CHANNEL', 'C123')
    monkeypatch.setattr(sb, '_visit_slack_app', _App())
    stats = sb._migrate_visit_card_buttons(days=3, dry_run=False)
    assert c.updates == []
    assert stats.get('cancelled_skip') == 1


# ── 슬랙 서명 확인 전 중복 표시 ───────────────────────────────────────
class _Redis:
    def __init__(self):
        self.keys = {}

    def set(self, k, v, nx=False, ex=None):
        if nx and k in self.keys:
            return None
        self.keys[k] = v
        return True


class _RC:
    def __init__(self, r):
        self.redis = r


def _signed(secret, body, ts=None):
    ts = ts or str(int(time.time()))
    sig = 'v0=' + hmac.new(secret.encode(), f'v0:{ts}:{body}'.encode(), hashlib.sha256).hexdigest()
    return {'X-Slack-Request-Timestamp': ts, 'X-Slack-Signature': sig, 'Content-Type': 'application/json'}


def test_unsigned_event_does_not_mark_dedup(app, monkeypatch):
    r = _Redis()
    monkeypatch.setattr(sb, '_SIGNING_SECRET', 'sekret')
    body = json.dumps({'type': 'event_callback', 'api_app_id': 'A1', 'event_id': 'Ev1'})
    with app.test_request_context('/slack/events', method='POST', data=body,
                                  headers={'Content-Type': 'application/json'}), \
         patch('dashboard.utils.redis_client.get_redis_client', return_value=_RC(r)):
        assert sb._is_slack_retry_duplicate() is False
    assert r.keys == {}


def test_signed_event_marks_and_then_dedups(app, monkeypatch):
    r = _Redis()
    monkeypatch.setattr(sb, '_SIGNING_SECRET', 'sekret')
    body = json.dumps({'type': 'event_callback', 'api_app_id': 'A1', 'event_id': 'Ev1'})
    with patch('dashboard.utils.redis_client.get_redis_client', return_value=_RC(r)):
        with app.test_request_context('/slack/events', method='POST', data=body, headers=_signed('sekret', body)):
            assert sb._is_slack_retry_duplicate() is False
        with app.test_request_context('/slack/events', method='POST', data=body, headers=_signed('sekret', body)):
            assert sb._is_slack_retry_duplicate() is True
    assert list(r.keys) == ['slack_event_seen:A1:Ev1']


def test_wrong_secret_signature_rejected(app, monkeypatch):
    monkeypatch.setattr(sb, '_SIGNING_SECRET', 'sekret')
    body = '{"event_id":"Ev1"}'
    with app.test_request_context('/slack/events', method='POST', data=body, headers=_signed('other', body)):
        assert sb._slack_signature_ok() is False


# ── 구글 로그인 확인 ─────────────────────────────────────────────────
def _mgr():
    from dashboard.google_oauth import GoogleOAuthManager
    m = GoogleOAuthManager.__new__(GoogleOAuthManager)
    m.allowed_domain = 'itg-aircon.com'
    return m


@pytest.mark.parametrize('info,ok', [
    ({'email': 'a@itg-aircon.com', 'verified_email': True, 'hd': 'itg-aircon.com'}, True),
    ({'email': 'a@itg-aircon.com', 'verified_email': True}, True),            # hd 없음 — 지금은 경고만
    ({'email': 'a@itg-aircon.com', 'verified_email': False, 'hd': 'itg-aircon.com'}, False),
    ({'email': 'a@itg-aircon.com', 'verified_email': True, 'hd': 'other.com'}, False),
    ({'email': 'a@gmail.com', 'verified_email': True}, False),
])
def test_oauth_validate_user(info, ok, caplog):
    caplog.set_level(logging.WARNING)
    assert (_mgr()._validate_user(info) is not None) is ok


def test_oauth_missing_hd_logged(caplog):
    caplog.set_level(logging.WARNING)
    _mgr()._validate_user({'email': 'a@itg-aircon.com', 'verified_email': True})
    assert any('회사 도메인(hd) 정보 없는 로그인' in r.getMessage() for r in caplog.records)


# ── 세션 ID 재발급 ───────────────────────────────────────────────────
def test_rotate_session_id_calls_regenerate(app):
    from dashboard.auth import rotate_session_id

    class _SI:
        called = []

        def regenerate(self, s):
            self.called.append(s)

    class _Sess(dict):
        sid = 'old'

    si = _SI()
    with app.test_request_context('/'):
        app.session_interface = si
        s = _Sess(user={'email': 'a'})
        with patch('dashboard.auth.session', s):
            rotate_session_id()
    assert si.called and si.called[0] is s


def test_login_paths_rotate_session():
    from dashboard.blueprints import auth as auth_bp, groupware
    assert 'rotate_session_id()' in inspect.getsource(auth_bp.google_callback)
    gsrc = inspect.getsource(groupware.field_login)
    assert 'rotate_session_id()' in gsrc and "headers.get('X-Forwarded-For')" not in gsrc


# ── /api 출처 확인 · 소켓 CORS ───────────────────────────────────────
@pytest.mark.parametrize('headers,ok', [
    ({'Sec-Fetch-Site': 'same-origin'}, True),
    ({'Sec-Fetch-Site': 'none'}, True),
    ({'Sec-Fetch-Site': 'same-site'}, False),
    ({'Sec-Fetch-Site': 'cross-site'}, False),
    ({'Origin': 'https://pm.itg-aircon.com'}, True),
    ({'Origin': 'https://www.itg-aircon.com'}, False),
    ({}, True),                                     # 브라우저 아닌 호출
])
def test_same_origin_request(app, headers, ok):
    from dashboard.utils.security_middleware import SecurityMiddleware
    with app.test_request_context('/api/x', method='POST', base_url='https://pm.itg-aircon.com', headers=headers):
        assert SecurityMiddleware._same_origin_request() is ok


def test_socketio_cors_not_wildcard():
    import dashboard
    src = inspect.getsource(dashboard)
    assert 'cors_allowed_origins="*"' not in src and 'cors_allowed_origins=None' in src


# ── 엑셀·CSV·OCR 수식 ────────────────────────────────────────────────
@pytest.mark.parametrize('v,expected', [
    ('=1+1', "'=1+1"), ('+82', "'+82"), ('-x', "'-x"), ('@a', "'@a"), ('-', '-'), ('홍길동', '홍길동'), (3, 3),
])
def test_csv_safe(v, expected):
    from dashboard.utils.spreadsheet_safe import csv_safe
    assert csv_safe(v) == expected


def test_xlsx_append_text_keeps_value_as_text():
    from openpyxl import Workbook, load_workbook
    from dashboard.utils.spreadsheet_safe import xlsx_append_text
    wb = Workbook()
    ws = wb.active
    xlsx_append_text(ws, ['=HYPERLINK("https://x","y")', 1000, '보통 글'])
    buf = io.BytesIO()
    wb.save(buf)
    ws2 = load_workbook(io.BytesIO(buf.getvalue())).active
    assert ws2['A1'].data_type == 's' and ws2['A1'].value == '=HYPERLINK("https://x","y")'
    assert ws2['B1'].value == 1000 and ws2['C1'].value == '보통 글'


def test_ocr_business_name_neutralised():
    from dashboard.services import business_license_handler as blh
    src = inspect.getsource(blh._maybe_update_business_name)
    assert "sheet_safe_text(ocr_name)" in src
