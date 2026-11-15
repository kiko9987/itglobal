# -*- coding: utf-8 -*-
"""5차 보안 점검 #6 — #영업_관리 카드·공사 취소 DM 의 현장 주소(리드 원문일 수 있음) 슬랙 표기 이스케이프.

홈페이지 폼 주소가 카카오·juso 로 안 풀리면 원문 그대로 방문 주소 → 현장 주소로 들어간다.
'<!here>'·'<https://…|견적서 확인>' 이 카드에서 멘션·위장 링크로 동작하지 않아야 한다(표시만 이스케이프)."""
import inspect

import pytest

EVIL = '서울 <!here> <https://example.invalid|견적서 확인> & 2층'
SAFE = '서울 &lt;!here&gt; &lt;https://example.invalid|견적서 확인&gt; &amp; 2층'

PROJECT = {
    '프로젝트 코드': 'G9999', '사업자명': '테스트상사', '거래처명': '', '현장 주소': EVIL,
    '총액 1': '1,000,000', '부가세': 'TRUE', '공사 확정': '2026-11-01', '공사 내용': '<!channel> 교체',
}


class _Client:
    def __init__(self):
        self.texts = []

    def conversations_join(self, **kw):
        return {'ok': True}

    def chat_postMessage(self, **kw):
        self.texts.append(kw.get('text', ''))
        for b in kw.get('blocks') or []:
            t = (b.get('text') or {}).get('text')
            if t:
                self.texts.append(t)
        return {'ok': False}          # 뒤 단계(Redis 대기 저장 등)로 안 넘어가게


@pytest.fixture
def sb(monkeypatch):
    from dashboard.blueprints import slack_bot
    monkeypatch.setenv('SLACK_INVOICE_CHANNEL_ID', 'C_TEST')
    return slack_bot


def _check(texts):
    assert texts
    joined = '\n'.join(texts)
    assert '<!here>' not in joined and '<https://example.invalid' not in joined
    assert '<!channel>' not in joined
    assert SAFE in joined


def test_edit_notice_card(sb):
    c = _Client()
    sb._post_project_edit_notice_card(c, 'G9999', PROJECT, {'공사 내용': '<!channel> 수정'}, '오타', 'KK')
    _check(c.texts)
    assert '&lt;!channel&gt; 교체 → &lt;!channel&gt; 수정' in '\n'.join(c.texts)


def test_amount_request_card(sb, monkeypatch):
    c = _Client()
    monkeypatch.setattr(sb, '_project_client', lambda: c)
    sb._post_amount_edit_request_card(PROJECT, {'총액 1': '2000000'}, '변경', 'U_TEST', 'KK')
    _check(c.texts)


def test_cancel_money_request_card(sb, monkeypatch):
    c = _Client()
    monkeypatch.setattr(sb, '_project_client', lambda: c)
    sb._send_cancel_money_request_card(PROJECT, '고객 취소', 'KK', 'U_TEST')
    _check(c.texts)


def test_cancel_notice_card(sb):
    c = _Client()
    sb._post_project_cancel_notice_card(c, 'G9999', PROJECT, 'KK')
    _check(c.texts)


def test_resume_card(sb, monkeypatch):
    c = _Client()
    monkeypatch.setattr(sb, '_project_client', lambda: c)
    sb._send_project_resume_card('G9999', PROJECT, 'KK', '재개', '1000000')
    _check(c.texts)


def test_cancel_dm_and_modals_escape_in_source(sb):
    src = inspect.getsource(sb)
    assert '📍 {_slack_esc_txt(addr)}\\n💲' in src                     # 공사 취소 경영지원 DM
    assert "사업자등록증 — `{code}` · {_slack_esc_txt(biz)}" in src    # 계산서 스레드 등록증 댓글
    assert src.count('info = _slack_esc_txt(info)') == 2                # 입금 창·A/S 창 공사 정보


def test_payment_alert_address_escaped():
    from dashboard.services import payment_alert as pa
    src = inspect.getsource(pa)
    assert 'addr = slack_esc(addr)' in src
    assert "slack_esc((it.get('address') or '')[:50])" in src
