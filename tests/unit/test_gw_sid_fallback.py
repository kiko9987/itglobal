# -*- coding: utf-8 -*-
"""결재함 본인 확인 — 슬랙 조회가 잠깐 안 될 때 (2026-11-15).

예전: 조회 실패면 5분 동안 '인사카드에 슬랙 계정이 연결되지 않아'로 보이고 웹 승인이 막혔다. 연결이 안 되면
봇 셋을 5초씩 기다려 최대 15초 멈췄다.
지금: 하루 안에 확인해 둔 ID 를 쓰고(카드와 다시 맞춰 봄), 그것도 없으면 '잠시 뒤 다시' 상태로 구분한다.
"""
from unittest.mock import patch

import pytest

from dashboard.services import gw_approval as ga

WHO = 'staff1@test.invalid'


@pytest.fixture(autouse=True)
def _clear():
    ga._SID_CACHE.clear()
    yield
    ga._SID_CACHE.clear()


def _expire(email, age_last=0.0):
    real, _exp, state, last_id, last_at = ga._SID_CACHE[email]
    ga._SID_CACHE[email] = (real, 0, state, last_id, last_at - age_last)


def test_slack_blip_uses_last_confirmed_id():
    with patch.object(ga, '_lookup_slack_id', return_value='U_SELF'):
        assert ga.slack_id_status(WHO, 'U_SELF') == ('U_SELF', 'ok')
    _expire(WHO)
    with patch.object(ga, '_lookup_slack_id', return_value=None) as lk:
        assert ga.slack_id_status(WHO, 'U_SELF') == ('U_SELF', 'ok')
        assert ga.slack_id_status(WHO, 'U_OTHER') == ('', 'unlinked')   # 카드 바꿔치기는 그대로 막힘
        assert lk.call_count == 1                                         # 실패도 5분 기억


def test_slack_down_without_history_is_unavailable_not_unlinked():
    with patch.object(ga, '_lookup_slack_id', return_value=None):
        assert ga.slack_id_status(WHO, 'U_SELF') == ('', 'unavailable')
    ga._SID_CACHE.clear()
    with patch.object(ga, '_lookup_slack_id', return_value=''):           # 슬랙에 없는 메일
        assert ga.slack_id_status(WHO, 'U_SELF') == ('', 'unlinked')


def test_last_confirmed_id_expires_after_a_day():
    with patch.object(ga, '_lookup_slack_id', return_value='U_SELF'):
        ga.slack_id_status(WHO, 'U_SELF')
    _expire(WHO, age_last=ga._SID_LAST_OK_TTL + 1)
    with patch.object(ga, '_lookup_slack_id', return_value=None):
        assert ga.slack_id_status(WHO, 'U_SELF') == ('', 'unavailable')


def test_slack_says_not_found_drops_old_id():
    with patch.object(ga, '_lookup_slack_id', return_value='U_SELF'):
        ga.slack_id_status(WHO, 'U_SELF')
    _expire(WHO)
    with patch.object(ga, '_lookup_slack_id', return_value=''):           # 슬랙에서 사라진 계정
        assert ga.slack_id_status(WHO, 'U_SELF') == ('', 'unlinked')


def test_verified_slack_id_keeps_old_contract():
    with patch.object(ga, '_lookup_slack_id', return_value=None):
        assert ga.verified_slack_id(WHO, 'U_SELF') == ''


_ENVS = ('SLACK_BOT_TOKEN', 'SLACK_PROJECT_BOT_TOKEN', 'SLACK_EAPPROVAL_BOT_TOKEN')


def test_lookup_stops_on_connection_error(monkeypatch):
    for k in _ENVS:
        monkeypatch.setenv(k, 'tok-' + k)
    calls = []

    class _WC:
        def __init__(self, token, timeout):
            self.token = token

        def users_lookupByEmail(self, email):
            calls.append(self.token)
            raise ConnectionError('down')
    monkeypatch.setattr('slack_sdk.WebClient', _WC)
    assert ga._lookup_slack_id(WHO) is None and len(calls) == 1          # 봇 셋을 5초씩 기다리지 않음


def test_lookup_tries_next_bot_on_api_error(monkeypatch):
    from slack_sdk.errors import SlackApiError
    for k in _ENVS:
        monkeypatch.setenv(k, 'tok-' + k)
    calls = []

    class _WC:
        def __init__(self, token, timeout):
            self.token = token

        def users_lookupByEmail(self, email):
            calls.append(self.token)
            if self.token == 'tok-SLACK_BOT_TOKEN':
                raise SlackApiError('missing_scope', {'ok': False, 'error': 'missing_scope'})
            return {'user': {'id': 'U_SELF'}}
    monkeypatch.setattr('slack_sdk.WebClient', _WC)
    assert ga._lookup_slack_id(WHO) == 'U_SELF' and len(calls) == 2


def test_viewer_flags_unavailable():
    card = {'slack_id': 'U_SELF', 'name': '홍길동'}
    with patch.object(ga.hr, 'find_by_email', return_value=card), \
         patch.object(ga.hr, 'is_hr_admin', return_value=False), \
         patch.object(ga, '_archive_windows', return_value=None), \
         patch.object(ga, '_lookup_slack_id', return_value=None):
        v = ga.viewer(WHO)
    assert v['linked'] is False and v['sid_unavailable'] is True
