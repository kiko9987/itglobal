"""구글 로그인 코드 만료(invalid_grant)는 오류 알림 없이 로그인 재시작 (2026-11-12)."""
import logging
from unittest.mock import MagicMock, patch

import pytest
from flask import Blueprint, Flask

from dashboard import google_oauth
from dashboard.google_oauth import GoogleOAuthManager, OAuthCodeExpired


def _manager_with_error(msg):
    mgr = GoogleOAuthManager.__new__(GoogleOAuthManager)
    flow = MagicMock()
    flow.fetch_token.side_effect = Exception(msg)
    mgr.create_flow = MagicMock(return_value=flow)
    return mgr


def test_invalid_grant_raises_and_logs_warning_only(caplog):
    mgr = _manager_with_error('(invalid_grant) Bad Request')
    with caplog.at_level(logging.WARNING, logger=google_oauth.__name__):
        with pytest.raises(OAuthCodeExpired):
            mgr.get_user_info('code', 'state')
    assert not [r for r in caplog.records if r.levelno >= logging.ERROR]
    assert any(r.levelno == logging.WARNING for r in caplog.records)


def test_other_errors_still_logged_as_error(caplog):
    mgr = _manager_with_error('(invalid_client) Unauthorized')
    with caplog.at_level(logging.WARNING, logger=google_oauth.__name__):
        assert mgr.get_user_info('code', 'state') == (None, None)
    assert any(r.levelno == logging.ERROR for r in caplog.records)


@pytest.fixture
def client():
    from dashboard.blueprints.auth import auth_bp
    app = Flask(__name__)
    app.secret_key = 'test'
    projects = Blueprint('projects', __name__)
    projects.add_url_rule('/projects', 'project_list', lambda: 'ok')
    app.register_blueprint(projects)
    app.register_blueprint(auth_bp)
    mgr = MagicMock()
    mgr.get_user_info.side_effect = OAuthCodeExpired('invalid_grant')
    with patch('dashboard.blueprints.auth.is_oauth_configured', return_value=True), \
         patch('dashboard.blueprints.auth.get_google_oauth', return_value=mgr):
        yield app.test_client()


def _callback(c):
    with c.session_transaction() as s:
        s['oauth_state'] = 'st'
    return c.get('/auth/callback?state=st&code=x')


def test_callback_goes_to_login_page_without_message(client):
    r = _callback(client)
    assert r.status_code == 302 and r.headers['Location'].endswith('/login')
    # 세션을 바꾸지 않아야 동시 콜백의 로그인 쿠키를 덮지 않는다
    assert 'Set-Cookie' not in r.headers


def test_login_page_sends_logged_in_user_to_projects(client):
    with client.session_transaction() as s:
        s['user'] = {'email': 'a@itg-aircon.com'}
    r = client.get('/login')
    assert r.status_code == 302 and r.headers['Location'].endswith('/projects')
