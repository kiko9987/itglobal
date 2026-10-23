# -*- coding: utf-8 -*-
"""/sms/inbound 본문 형식 회귀 테스트 (2026-10-23 갤럭시 SmsForwarder 세팅).

핵심 보장:
  1. 폼(x-www-form-urlencoded) 본문 — 여러 줄 은행 문자가 줄바꿈 그대로 인입 코어로 전달
  2. 기존 JSON 본문 경로 무변경
  3. 인증됐는데 본문이 비면 경고 로그(조용히 무시하지 않음)
create_app 미사용 — 블루프린트만 단 최소 Flask 앱.
"""
import sys
sys.path.insert(0, '.')

import logging
import pytest
from flask import Flask

import dashboard.blueprints.sms_inbound as si

IBK = "[Web발신]\n2026/10/23 10:00\n입금 1,000원\n홍길동\n452***38801011\n기업"


@pytest.fixture
def client(monkeypatch):
    monkeypatch.setenv('SMS_INBOUND_TOKENS', 'sb:secretSB')
    monkeypatch.delenv('SMS_BANK_SENDERS', raising=False)
    calls = []

    def fake_ingest(text, source='sms', cash_receiver=''):
        calls.append((text, source))
        return {'status': 'ignored', 'reason': 'test'}

    monkeypatch.setattr(si, 'ingest_deposit', fake_ingest)
    app = Flask(__name__)
    app.register_blueprint(si.sms_bp)
    c = app.test_client()
    c.calls = calls
    return c


def test_form_body_multiline_preserved(client):
    r = client.post('/sms/inbound', data={'token': 'secretSB', 'sender': '15662566', 'text': IBK})
    assert r.status_code == 200
    assert client.calls == [(IBK, 'sms:sb')]


def test_json_body_unchanged(client):
    r = client.post('/sms/inbound', json={'token': 'secretSB', 'sender': 'x', 'text': IBK})
    assert r.status_code == 200
    assert client.calls == [(IBK, 'sms:sb')]


def test_bad_token_form_rejected(client):
    r = client.post('/sms/inbound', data={'token': 'nope', 'text': IBK})
    assert r.status_code == 401
    assert client.calls == []


def test_empty_text_warns(client, caplog):
    # 템플릿 필드명 오타(text 대신 content) 같은 설정 오류 → 경고 로그로 드러나야 함
    caplog.set_level(logging.WARNING)
    r = client.post('/sms/inbound', data={'token': 'secretSB', 'content': IBK})
    assert r.get_json() == {'status': 'empty'}
    assert client.calls == []
    assert any('본문 비어 무시' in rec.getMessage() for rec in caplog.records)


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
