# -*- coding: utf-8 -*-
"""/slack/list-assignee — 비밀값이 반드시 있어야 하고 X-Auth 가 맞아야 쓴다 (2026-11-13 보안 점검).

예전엔 SLACK_LIST_WEBHOOK_SECRET 이 없으면 검사 없이 받아, 주소만 알면 누구나
리드 영업 담당자를 바꾸거나 지울 수 있었다.
"""
from unittest.mock import patch

import pytest
from flask import Flask

from dashboard.blueprints import slack_bot as sb

SECRET = 'test-secret-value-123'


class _UserDB:
    def __init__(self, users=None, boom=False):
        self.users, self.boom = users or [], boom

    def get_all_users(self):
        if self.boom:
            raise RuntimeError('db down')
        return self.users


STAFF = [{'name': '테스트', 'is_active': True, 'is_resigned': False}]


@pytest.fixture()
def client():
    app = Flask(__name__)
    app.register_blueprint(sb.slack_bp)
    return app.test_client()


def _post(client, body, auth=None):
    headers = {'X-Auth': auth} if auth is not None else {}
    return client.post('/slack/list-assignee', json=body, headers=headers)


@pytest.mark.parametrize('assignee', ['테스트', '', '미정'])
def test_rejected_when_secret_not_configured(client, monkeypatch, assignee):
    monkeypatch.delenv('SLACK_LIST_WEBHOOK_SECRET', raising=False)
    with patch.object(sb, '_update_lead_dispatch') as upd:
        r = _post(client, {'lead_no': 'L-00001', 'assignee': assignee})
    assert r.status_code == 503
    upd.assert_not_called()


@pytest.mark.parametrize('auth', [None, '', 'wrong', SECRET + 'x', '한글'])
def test_rejected_with_wrong_or_missing_header(client, monkeypatch, auth):
    monkeypatch.setenv('SLACK_LIST_WEBHOOK_SECRET', SECRET)
    with patch.object(sb, '_update_lead_dispatch') as upd:
        r = _post(client, {'lead_no': 'L-00001', 'assignee': '테스트'}, auth=auth)
    assert r.status_code == 401
    upd.assert_not_called()


def test_accepts_with_correct_header(client, monkeypatch):
    monkeypatch.setenv('SLACK_LIST_WEBHOOK_SECRET', SECRET)
    with patch('dashboard.utils.user_database.get_user_database', return_value=_UserDB(STAFF)), \
         patch.object(sb, '_update_lead_dispatch') as upd:
        r = _post(client, {'lead_no': 'L-00001', 'assignee': '테스트'}, auth=SECRET)
    assert r.status_code == 200
    upd.assert_called_once_with('L-00001', {'영업 담당자': '테스트'})


def test_clear_with_correct_header(client, monkeypatch):
    monkeypatch.setenv('SLACK_LIST_WEBHOOK_SECRET', SECRET)
    with patch.object(sb, '_update_lead_dispatch') as upd:
        r = _post(client, {'lead_no': 'L-00001', 'assignee': '미정'}, auth=SECRET)
    assert r.status_code == 200
    upd.assert_called_once_with('L-00001', {'영업 담당자': '-'})


def test_staff_list_failure_does_not_write(client, monkeypatch):
    monkeypatch.setenv('SLACK_LIST_WEBHOOK_SECRET', SECRET)
    with patch('dashboard.utils.user_database.get_user_database', return_value=_UserDB(boom=True)), \
         patch.object(sb, '_update_lead_dispatch') as upd:
        r = _post(client, {'lead_no': 'L-00001', 'assignee': '=1+1'}, auth=SECRET)
    assert r.status_code == 503
    upd.assert_not_called()


def test_unknown_name_rejected(client, monkeypatch):
    monkeypatch.setenv('SLACK_LIST_WEBHOOK_SECRET', SECRET)
    with patch('dashboard.utils.user_database.get_user_database', return_value=_UserDB(STAFF)), \
         patch.object(sb, '_update_lead_dispatch') as upd:
        r = _post(client, {'lead_no': 'L-00001', 'assignee': '외부인'}, auth=SECRET)
    assert r.status_code == 400
    upd.assert_not_called()
