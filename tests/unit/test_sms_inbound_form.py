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


NH_IN = "[Web발신]\n농협 입금7,450,600원\n10/26 09:15 352-****-1682-33 이재철(푸른 잔액82,517,408원"
NH_OUT = "[Web발신]\n농협 출금750,000원\n10/23 14:44 352-****-1682-33 주영민 잔액12,345,678원"


def test_nh_inflow_server_only(client):
    # 농협 입금(폰) — 서버에만 보관, 슬랙 카드 코어(ingest_deposit)는 안 탐
    import json as _json
    from dashboard.utils.redis_client import get_redis_client
    rc = get_redis_client().redis
    r = client.post('/sms/inbound', json={'token': 'secretSB', 'sender': '15882100', 'text': NH_IN})
    body = r.get_json()
    assert body['reason'] == 'nh_inflow_recorded'
    assert client.calls == []
    oid = si.dedup_hash('', NH_IN)
    try:
        rec = _json.loads(rc.get(f'sms_nh_inflow:{oid}'))
        assert '82,517,408' not in rec['text'] and '잔액' not in rec['text']   # 잔액 제거본만
        assert rec['partner'] == '이재철(푸른' and rec['amount'] == 7_450_600
        # 인입 카드 dedup 키는 건드리지 않음 → 나중에 같은 문자 붙여넣으면 카드 정상 생성
        assert rc.get(f'sms_intake:seen:{oid}') is None
        # 같은 문자 재수신(폰 여러 대) → 중복 처리
        r2 = client.post('/sms/inbound', json={'token': 'secretSB', 'sender': 'x', 'text': NH_IN})
        assert r2.get_json()['status'] == 'duplicate'
    finally:
        rc.delete(f'sms_nh_inflow:{oid}', f'sms_nh_inflow:seen:{oid}')
        rc.zrem('sms_nh_inflow:index', oid)


def test_nh_outflow_goes_to_common_path(client):
    # 농협 출금은 기존처럼 공통 코어 → 출금 서버 보관 경로
    client.post('/sms/inbound', json={'token': 'secretSB', 'sender': 'x', 'text': NH_OUT})
    assert client.calls == [(NH_OUT, 'sms:sb')]


def test_ibk_with_nonghyup_named_payer_not_blocked(client):
    # 기업 계좌 입금인데 입금자명에 '농협'이 들어간 건은 막지 않음 (계좌로 판정)
    txt = "[Web발신]\n2026/10/26 10:00\n입금 1,000,000원\n(주)농협유통\n452***38801011\n기업"
    client.post('/sms/inbound', json={'token': 'secretSB', 'sender': 'x', 'text': txt})
    assert client.calls == [(txt, 'sms:sb')]


def test_nh_channel_paste_path_unaffected():
    # 붙여넣기(채널) 경로는 ingest_deposit 를 직접 타므로 농협도 정상 인입 대상
    from dashboard.services.sms_intake import has_business_account, looks_like_payment
    assert looks_like_payment(NH_IN) and has_business_account(NH_IN)


def test_empty_text_warns(client, caplog):
    # 템플릿 필드명 오타(text 대신 content) 같은 설정 오류 → 경고 로그로 드러나야 함
    caplog.set_level(logging.WARNING)
    r = client.post('/sms/inbound', data={'token': 'secretSB', 'content': IBK})
    assert r.get_json() == {'status': 'empty'}
    assert client.calls == []
    assert any('본문 비어 무시' in rec.getMessage() for rec in caplog.records)


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
