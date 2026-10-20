# -*- coding: utf-8 -*-
"""Slack 이벤트 중복 전달 차단(_is_slack_retry_duplicate) — 원본 + 1차 재전송 이중 처리 방지 (2026-10-20 R4187-MJ)."""
import json
from unittest.mock import patch

import pytest
from flask import Flask

from dashboard.blueprints import slack_bot as sb


class _FakeRedis:
    def __init__(self):
        self.kv = {}

    def set(self, k, v, nx=False, ex=None):
        if nx and k in self.kv:
            return None
        self.kv[k] = v
        return True


@pytest.fixture
def env():
    app = Flask(__name__)
    fake = _FakeRedis()
    with patch('dashboard.utils.redis_client.get_redis_client') as g:
        g.return_value.redis = fake
        yield app, fake


def _dup(app, body, retry=None, form=False):
    headers = {'X-Slack-Retry-Num': str(retry), 'X-Slack-Retry-Reason': 'http_timeout'} if retry else {}
    if form:
        ctx = app.test_request_context('/slack/project-events', method='POST', data=body, headers=headers)
    else:
        ctx = app.test_request_context('/slack/project-events', method='POST', data=json.dumps(body),
                                       content_type='application/json', headers=headers)
    with ctx:
        return sb._is_slack_retry_duplicate()


EV = {'type': 'event_callback', 'api_app_id': 'A_PROJECT', 'event_id': 'Ev0001', 'event': {'type': 'message'}}


def test_original_then_first_retry_processed_once(env):
    app, _ = env
    assert _dup(app, EV) is False            # 원본 — 처리
    assert _dup(app, EV, retry=1) is True    # 1차 재전송 — 건너뜀 (예전엔 여기서 또 처리)
    assert _dup(app, EV, retry=2) is True


def test_retry_alone_is_processed_when_original_never_arrived(env):
    app, _ = env
    assert _dup(app, EV, retry=1) is False
    assert _dup(app, EV, retry=2) is True


def test_same_event_id_different_app_not_blocked(env):
    app, _ = env
    assert _dup(app, EV) is False
    assert _dup(app, dict(EV, api_app_id='A_INVOICE')) is False


def test_no_event_id_always_processed(env):
    app, fake = env
    assert _dup(app, {'command': '/공사확정'}, form=True) is False          # 슬래시 명령(form)
    assert _dup(app, {'type': 'url_verification', 'challenge': 'x'}) is False
    assert _dup(app, {'type': 'url_verification', 'challenge': 'x'}, retry=1) is False
    assert fake.kv == {}


def test_redis_error_does_not_block(env):
    app, _ = env
    with patch('dashboard.utils.redis_client.get_redis_client', side_effect=RuntimeError('down')):
        assert _dup(app, EV) is False
        assert _dup(app, EV, retry=1) is False
