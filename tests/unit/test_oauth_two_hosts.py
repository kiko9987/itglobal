"""구글 로그인 복귀 주소 — pm·gw 두 주소(2026-11-19). 들어온 주소에 맞는 복귀 주소를 고른다."""
from flask import Flask

from dashboard.blueprints import auth


def _app():
    app = Flask(__name__)
    app.register_blueprint(auth.auth_bp)
    return app


def test_callback_uri_follows_request_host(monkeypatch):
    monkeypatch.setenv('GOOGLE_OAUTH_REDIRECT_URIS',
                       'https://pm.itg-aircon.com/auth/callback,https://gw.itg-aircon.com/auth/callback')
    monkeypatch.setenv('GOOGLE_OAUTH_DEFAULT_REDIRECT_URI', 'https://pm.itg-aircon.com/auth/callback')
    app = _app()
    with app.test_request_context('/auth/google', base_url='https://gw.itg-aircon.com'):
        assert auth._callback_uri() == 'https://gw.itg-aircon.com/auth/callback'
    with app.test_request_context('/auth/google', base_url='https://pm.itg-aircon.com'):
        assert auth._callback_uri() == 'https://pm.itg-aircon.com/auth/callback'


def test_callback_uri_falls_back_to_default(monkeypatch):
    # 구글 콘솔에 gw 를 아직 안 넣었으면(목록에 없음) 예전처럼 기본값
    monkeypatch.setenv('GOOGLE_OAUTH_REDIRECT_URIS', 'https://pm.itg-aircon.com/auth/callback')
    monkeypatch.setenv('GOOGLE_OAUTH_DEFAULT_REDIRECT_URI', 'https://pm.itg-aircon.com/auth/callback')
    with _app().test_request_context('/auth/google', base_url='https://gw.itg-aircon.com'):
        assert auth._callback_uri() == 'https://pm.itg-aircon.com/auth/callback'


def test_allowed_origins_has_gw(monkeypatch):
    from dashboard.utils.allowed_origins import allowed_hosts
    monkeypatch.delenv('ALLOWED_ORIGINS', raising=False)
    assert {'pm.itg-aircon.com', 'gw.itg-aircon.com'} <= allowed_hosts()
