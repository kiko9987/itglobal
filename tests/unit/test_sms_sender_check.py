# -*- coding: utf-8 -*-
"""SMS 입금 인입 — 발신번호 확인·표시, 폰 경로 현금 차단, 카드 표기 이스케이프 (2026-11-13 보안 점검).

예전: 은행 발신번호 허용 목록이 비어 있고('포함' 비교), 실제 발신번호가 어디에도 안 남았으며,
'현금 100만원' 같은 문자를 폰으로 보내면 계좌 확인 없이 입금 카드가 생겼다.
create_app 미사용 — 블루프린트만 단 최소 Flask 앱.
"""
import pytest
from flask import Flask

import dashboard.blueprints.sms_inbound as si
from dashboard.services.sms_intake import quoted_body, active_display, intake_text_from_card, INTAKE_SEP

IBK = "[Web발신]\n2026/10/23 10:00\n입금 1,000원\n홍길동\n452***38801011\n기업"


@pytest.mark.parametrize('raw,expected', [
    ('1566-2566', '15662566'),
    ('15662566', '15662566'),
    ('+82 10-1234-5678', '01012345678'),
    ('010-1234-5678', '01012345678'),
    ('기업은행', ''),
    ('', ''),
    (None, ''),
])
def test_normalize_sender(raw, expected):
    assert si._normalize_sender(raw) == expected


@pytest.mark.parametrize('d,ok', [
    ('15662566', True), ('15991111', True), ('18882100', True),
    ('01012345678', False), ('0212345678', False), ('1234', False), ('', False),
])
def test_looks_like_bank_rep(d, ok):
    assert si.looks_like_bank_rep(d) is ok


@pytest.fixture
def client(monkeypatch):
    monkeypatch.setenv('SMS_INBOUND_TOKENS', 'sb:secretSB')
    calls = []

    def fake_ingest(text, source='sms', cash_receiver='', **kw):
        calls.append(dict(text=text, source=source, **kw))
        return {'status': 'ignored', 'reason': 'test'}

    monkeypatch.setattr(si, 'ingest_deposit', fake_ingest)
    app = Flask(__name__)
    app.register_blueprint(si.sms_bp)
    c = app.test_client()
    c.calls = calls
    return c


@pytest.mark.parametrize('sender', ['0101566256699', '15662566999', '1566256', '01012345678', ''])
def test_allowlist_exact_match_rejects_lookalikes(client, monkeypatch, sender):
    monkeypatch.setenv('SMS_BANK_SENDERS', '15662566, 1599-1111')
    r = client.post('/sms/inbound', data={'token': 'secretSB', 'sender': sender, 'text': IBK})
    assert r.get_json() == {'status': 'ignored', 'reason': 'sender'}
    assert client.calls == []


@pytest.mark.parametrize('sender', ['15662566', '1566-2566', '15991111'])
def test_allowlist_accepts_listed(client, monkeypatch, sender):
    monkeypatch.setenv('SMS_BANK_SENDERS', '15662566, 1599-1111')
    client.post('/sms/inbound', data={'token': 'secretSB', 'sender': sender, 'text': IBK})
    assert len(client.calls) == 1
    call = client.calls[0]
    assert call['origin'] == si._normalize_sender(sender)
    assert call['origin_checked'] is True
    assert call['allow_cash'] is False


def test_no_allowlist_still_passes_origin_unchecked(client, monkeypatch):
    monkeypatch.delenv('SMS_BANK_SENDERS', raising=False)
    client.post('/sms/inbound', data={'token': 'secretSB', 'sender': '010-1234-5678', 'text': IBK})
    call = client.calls[0]
    assert call['origin'] == '01012345678'
    assert call['origin_checked'] is False
    assert call['allow_cash'] is False


def test_sms_path_cash_text_ignored():
    # 폰 경로(allow_cash=False)에서는 '현금' 자유문장이 입금 카드가 되지 않음 (Redis 전 단계에서 끝)
    r = si.ingest_deposit('현금 100만원 YG 수령', source='sms:test', allow_cash=False, origin='01012345678')
    assert r['status'] == 'ignored'


@pytest.mark.parametrize('preview,expect_warn', [
    ({'origin': '15662566', 'origin_checked': False}, False),      # 은행 대표번호 — 줄 없음
    ({'origin': '15662566', 'origin_checked': True}, False),
    ({'origin': '01012345678', 'origin_checked': False}, True),    # 휴대폰 — 경고
    ({'origin': '01012345678', 'origin_checked': True}, False),    # 허용 목록으로 확인됨
    ({'origin': '-', 'origin_checked': False}, False),             # 번호 모름 — 소음 막으려 줄 없음
    ({}, False),                                                    # 채널 붙여넣기
    ({'cash': True}, False),
])
def test_origin_line_warning_only(preview, expect_warn):
    line = si._origin_line(preview)
    assert bool(line) is expect_warn
    assert '발신번호:' not in line and '01012345678' not in line   # 번호는 카드에 안 보임


def test_card_hides_number_but_warns():
    blocks = si._build_intake_blocks('id1', IBK, {'origin': '01012345678', 'origin_checked': False})
    text = blocks[0]['text']['text']
    assert '01012345678' not in text and '발신번호' not in text
    assert '은행 대표번호가 아닌 번호에서 온 문자입니다' in text
    blocks = si._build_intake_blocks('id2', IBK, {'origin': '15662566', 'origin_checked': False})
    assert '⚠️' not in blocks[0]['text']['text'] and '1566' not in blocks[0]['text']['text']


def test_sms_card_text_escaped_and_round_trips():
    raw = "[Web발신]\n입금 1,000원\n<!channel> <https://evil.example|입금확인> A&B"
    q = quoted_body(raw)
    joined = '\n'.join(q)
    assert '<!channel>' not in joined and '<https://' not in joined
    assert '&lt;!channel&gt;' in joined and 'A&amp;B' in joined
    assert '<' not in active_display(raw)[4:]
    blocks = [{'type': 'section', 'text': {'type': 'mrkdwn',
               'text': '\n'.join(['>head', f'>{INTAKE_SEP}', *q, f'>{INTAKE_SEP}'])}}]
    assert intake_text_from_card(blocks) == raw.replace('입금 1,000원', '입금 1,000원')
