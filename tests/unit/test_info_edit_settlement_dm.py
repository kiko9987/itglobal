# -*- coding: utf-8 -*-
"""공사 현장 주소·유입 구분 변경 → 경영지원(SB) 알림 DM (2026-11-09 SB 요청, 이카운트 반영용)."""
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
        return {'channel': {'id': 'D_' + users}}

    def chat_postMessage(self, **kw):
        _FakeWeb.posted.append(kw)
        return {'ok': True}


def _setup(monkeypatch):
    _FakeWeb.posted = []
    _FakeWeb.opened = []
    monkeypatch.setenv('SLACK_BOT_TOKEN', 'xoxb-test')
    monkeypatch.delenv('SLACK_SETTLEMENT_CHECKER_ID', raising=False)
    monkeypatch.delenv('SLACK_SETTLEMENT_CHECKER_EMAIL', raising=False)
    monkeypatch.setattr(slack_sdk, 'WebClient', _FakeWeb)
    fake_redis = types.SimpleNamespace(redis=types.SimpleNamespace(get=lambda k: None))
    import dashboard.utils.redis_client as rcmod
    monkeypatch.setattr(rcmod, 'get_redis_client', lambda: fake_redis)


SB = 'U0BHC2JV7U5'


def test_address_change_by_manager_goes_to_sb_only(monkeypatch):
    _setup(monkeypatch)
    ok = psn.notify_info_edit_to_settlement(
        'R1-TH', [{'field_name': '현장 주소', 'old_value': '파주 A동 1', 'new_value': '파주 B동 2'}],
        editor_email='th@itg-aircon.com', latest_data={'사업자명': '로보온(주)'})
    assert ok is True
    assert _FakeWeb.opened == [SB]          # 승인자 전원이 아니라 SB 만
    body = _FakeWeb.posted[0]['text']
    assert '공사 정보 변경 알림' in body and '파주 A동 1 → 파주 B동 2' in body
    assert 'TH (PM 직접 편집)' in body and '이카운트' in body


def test_kiko_edit_still_notifies_sb(monkeypatch):
    """금액 알림과 달리 KiKO 편집도 보냄 — 이카운트는 SB 가 고치므로."""
    _setup(monkeypatch)
    ok = psn.notify_info_edit_to_settlement(
        'R1-TH', [{'field_name': '유입 구분', 'old_value': '전화', 'new_value': '거래처'}],
        editor_email='kiko@itg-aircon.com')
    assert ok is True and _FakeWeb.opened == [SB]


def test_sb_own_edit_skipped(monkeypatch):
    _setup(monkeypatch)
    assert psn.notify_info_edit_to_settlement(
        'R1-TH', [{'field_name': '유입 구분', 'old_value': '전화', 'new_value': '거래처'}],
        editor_email='sb@itg-aircon.com') is False
    assert _FakeWeb.posted == []


def test_blank_fill_notifies_but_whitespace_only_does_not(monkeypatch):
    _setup(monkeypatch)
    assert psn.notify_info_edit_to_settlement(
        'R1-TH', [{'field_name': '현장 주소', 'old_value': '-', 'new_value': '파주 A동 1'}],
        editor_email='th@itg-aircon.com') is True
    _FakeWeb.posted = []
    assert psn.notify_info_edit_to_settlement(
        'R1-TH', [{'field_name': '현장 주소', 'old_value': '파주  A동 1', 'new_value': '파주 A동 1 '}],
        editor_email='th@itg-aircon.com') is False
    assert _FakeWeb.posted == []


def test_other_fields_and_system_label_skipped(monkeypatch):
    _setup(monkeypatch)
    assert psn.notify_info_edit_to_settlement(
        'R1-TH', [{'field_name': '총액 1', 'old_value': 1, 'new_value': 2}],
        editor_email='th@itg-aircon.com') is False
    assert psn.notify_info_edit_to_settlement(
        'R1-TH', [{'field_name': '현장 주소', 'old_value': 'A', 'new_value': 'B'}],
        editor_label='시스템 자동기록') is False
    assert _FakeWeb.posted == []


def test_sheet_direct_edit_notifies(monkeypatch):
    _setup(monkeypatch)
    assert psn.notify_info_edit_to_settlement(
        'R1-TH', [{'field_name': '유입 구분', 'old_value': '네이버', 'new_value': '당근'}],
        editor_label='시트 직접수정') is True
    assert '변경 경로 : 시트 직접수정' in _FakeWeb.posted[0]['text']
