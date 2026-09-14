# -*- coding: utf-8 -*-
"""공사 금액 PM 직접수정 → 경영지원(황샛별) 별도 DM 알림 회귀 테스트 (2026-09-14).

관리자 금액 직접 편집은 영업사원 요청 게이트를 우회하므로 샛별에게 FYI DM.
네트워크(Slack)·Redis 는 monkeypatch. 트리거/제외 판정만 검증.
"""
import sys
sys.path.insert(0, '.')

import types
import slack_sdk

from dashboard.services import project_slack_notifier as psn


class _FakeWeb:
    posted = []
    opened = []

    def __init__(self, token=None):
        self.token = token

    def conversations_open(self, users=None):
        _FakeWeb.opened.append(users)
        return {'channel': {'id': 'D_SB'}}

    def chat_postMessage(self, **kw):
        _FakeWeb.posted.append(kw)
        return {'ok': True}


def _setup(monkeypatch):
    _FakeWeb.posted = []
    _FakeWeb.opened = []
    monkeypatch.setenv('SLACK_BOT_TOKEN', 'xoxb-test')
    monkeypatch.setenv('SLACK_SETTLEMENT_CHECKER_ID', 'U_SB')
    monkeypatch.setenv('SLACK_SETTLEMENT_CHECKER_EMAIL', 'sb@itg-aircon.com')
    monkeypatch.setattr(slack_sdk, 'WebClient', _FakeWeb)
    # redis (permalink 조회) — 없음 처리
    fake_redis = types.SimpleNamespace(redis=types.SimpleNamespace(get=lambda k: None))
    import dashboard.utils.redis_client as rcmod
    monkeypatch.setattr(rcmod, 'get_redis_client', lambda: fake_redis)


def _amt_change():
    return [{'field_name': '총액 1', 'old_value': 35500000, 'new_value': 36500000}]


def test_amount_edit_by_admin_sends_dm(monkeypatch):
    _setup(monkeypatch)
    ok = psn.notify_amount_edit_to_settlement(
        'G3805-YG', _amt_change(), editor_email='kiko@itg-aircon.com',
        latest_data={'사업자명': '닥터탁 성형외과의원'})
    assert ok is True
    assert _FakeWeb.opened == ['U_SB']
    assert len(_FakeWeb.posted) == 1
    body = _FakeWeb.posted[0]['text']
    assert 'G3805-YG' in body and '총액 1' in body and 'KIKO' in body


def test_settlement_own_edit_skipped(monkeypatch):
    _setup(monkeypatch)
    ok = psn.notify_amount_edit_to_settlement(
        'G3805-YG', _amt_change(), editor_email='sb@itg-aircon.com')
    assert ok is False
    assert _FakeWeb.posted == []


def test_non_amount_change_skipped(monkeypatch):
    _setup(monkeypatch)
    ok = psn.notify_amount_edit_to_settlement(
        'G3805-YG',
        [{'field_name': '방문 주소', 'old_value': 'A', 'new_value': 'B'}],
        editor_email='kiko@itg-aircon.com')
    assert ok is False
    assert _FakeWeb.posted == []


def test_vat_change_triggers(monkeypatch):
    _setup(monkeypatch)
    ok = psn.notify_amount_edit_to_settlement(
        'G3805-YG',
        [{'field_name': '부가세', 'old_value': False, 'new_value': True}],
        editor_email='kiko@itg-aircon.com')
    assert ok is True
    assert len(_FakeWeb.posted) == 1


def test_amount_unchanged_value_skipped(monkeypatch):
    _setup(monkeypatch)
    # 표시 동일(콤마 유무 차이만) → 실질 변화 없음으로 skip
    ok = psn.notify_amount_edit_to_settlement(
        'G3805-YG',
        [{'field_name': '총액 1', 'old_value': 35500000, 'new_value': '35500000'}],
        editor_email='kiko@itg-aircon.com')
    assert ok is False
    assert _FakeWeb.posted == []
