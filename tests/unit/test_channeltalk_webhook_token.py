"""채널톡 웹훅 토큰 — 값만 넣으면 기록만, ENFORCE=1 일 때만 거부 (2026-10-28)."""
import pytest
from flask import Flask

from dashboard.blueprints.channeltalk import channeltalk_bp


@pytest.fixture
def client():
    app = Flask(__name__)
    app.register_blueprint(channeltalk_bp)
    return app.test_client()


def _post(client, qs=''):
    return client.post('/channeltalk/events' + qs, json={})    # 빈 페이로드 → 토큰 확인 뒤 200


def test_no_token_configured_accepts_all(client, monkeypatch):
    monkeypatch.delenv('CHANNELTALK_WEBHOOK_TOKEN', raising=False)
    assert _post(client).status_code == 200


def test_observe_mode_accepts_mismatch(client, monkeypatch):
    monkeypatch.setenv('CHANNELTALK_WEBHOOK_TOKEN', 'abc')
    monkeypatch.delenv('CHANNELTALK_WEBHOOK_TOKEN_ENFORCE', raising=False)
    assert _post(client).status_code == 200
    assert _post(client, '?token=zzz').status_code == 200
    assert _post(client, '?token=abc').status_code == 200


def test_enforce_rejects_mismatch(client, monkeypatch):
    monkeypatch.setenv('CHANNELTALK_WEBHOOK_TOKEN', 'abc')
    monkeypatch.setenv('CHANNELTALK_WEBHOOK_TOKEN_ENFORCE', '1')
    assert _post(client).status_code == 403
    assert _post(client, '?token=zzz').status_code == 403
    assert _post(client, '?token=abc').status_code == 200
