# -*- coding: utf-8 -*-
"""safe_slack_call 재시도 분류 회귀 테스트.

핵심 보장: 영구(fatal) 슬랙 오류는 재시도하지 않고 즉시 raise → 호출자 fallback 이 처리.
특히 message_not_found(삭제된 카드 chat.update)는 자기치유 패턴이라 3회 재시도·ERROR
알림(오탐)을 유발하면 안 된다 (2026-09-29 G4109-YG grace update 계기).
"""
import sys
sys.path.insert(0, '.')

import pytest
from slack_sdk.errors import SlackApiError
from dashboard.blueprints.slack_helpers import safe_slack_call


class _Resp(dict):
    headers: dict = {}


def _raiser(err_code, counter):
    def _fn(**kwargs):
        counter['n'] += 1
        raise SlackApiError('x', _Resp({'ok': False, 'error': err_code}))
    return _fn


@pytest.mark.parametrize('err', [
    'message_not_found', 'cant_update_message', 'edit_window_closed',
    'cant_delete_message', 'channel_not_found', 'not_in_channel',
    'is_archived', 'not_authed', 'invalid_auth', 'account_inactive',
])
def test_fatal_errors_no_retry(err):
    # 영구 오류는 단 1회 시도 후 즉시 raise (재시도·지연 없음)
    counter = {'n': 0}
    with pytest.raises(SlackApiError):
        safe_slack_call(_raiser(err, counter), channel='c', ts='1', text='x')
    assert counter['n'] == 1


def test_success_returns_response():
    def _ok(**kwargs):
        return _Resp({'ok': True})
    assert safe_slack_call(_ok, channel='c')['ok'] is True


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
