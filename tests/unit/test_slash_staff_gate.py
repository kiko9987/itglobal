# -*- coding: utf-8 -*-
"""/수금·/as 사람 확인 (2026-11-13 보안 점검 2차 #7).

PM 이 현장·미화 직원(인사카드 access=field)에게 막는 미수금·고객 연락처·공사 금액이 슬랙 명령으로는 보였다.
재직 사무실 직원(access=staff)과 경영지원 승인자만 허용, 조회 실패는 막음.
"""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb
from dashboard.services import hr


@pytest.fixture
def people(monkeypatch):
    cards = {'U_STAFF': {'access': 'staff'}, 'U_FIELD': {'access': 'field'}}
    monkeypatch.setattr(hr, 'find_by_slack', lambda uid: cards.get(uid))
    monkeypatch.setattr(sb, '_settlement_checker_ids', lambda: ('U_CHECK',))
    return cards


def test_staff_and_checker_allowed(people):
    assert sb._slack_staff_denied('U_STAFF') is None
    assert sb._slack_staff_denied('U_CHECK') is None


@pytest.mark.parametrize('uid', ['U_FIELD', 'U_GUEST', ''])
def test_field_unknown_denied(people, uid):
    assert ':lock:' in sb._slack_staff_denied(uid)


def test_lookup_error_denied(monkeypatch):
    monkeypatch.setattr(sb, '_settlement_checker_ids', lambda: ())

    def boom(uid):
        raise RuntimeError('db down')
    monkeypatch.setattr(hr, 'find_by_slack', boom)
    assert '실패' in sb._slack_staff_denied('U_STAFF')


# ── /as 봇 핸들러가 실제로 거르는지 ─────────────────────────────

class _FakeApp:
    def __init__(self):
        self.h = {}

    def _reg(self, kind):
        def deco(name, *a, **k):
            def wrap(fn):
                self.h[(kind, str(name))] = fn
                return fn
            return wrap
        return deco

    def __getattr__(self, kind):
        if kind in ('command', 'view', 'action', 'options', 'event', 'shortcut', 'message'):
            return self._reg(kind)
        raise AttributeError(kind)


@pytest.fixture
def as_app(people):
    app = _FakeApp()
    sb._register_as_handlers(app)
    return app


def test_as_command_refuses_field(as_app, monkeypatch):
    opened, said = [], []
    monkeypatch.setattr(sb, '_open_as_request_modal', lambda *a, **k: opened.append(a))
    fn = as_app.h[('command', '/as')]
    fn(ack=lambda *a, **k: None, command={'trigger_id': 't', 'user_id': 'U_FIELD'}, client=None,
       respond=lambda **k: said.append(k))
    assert not opened and ':lock:' in said[0]['text']
    fn(ack=lambda *a, **k: None, command={'trigger_id': 't', 'user_id': 'U_STAFF'}, client=None,
       respond=lambda **k: said.append(k))
    assert len(opened) == 1


def test_as_project_search_empty_for_field(as_app, monkeypatch):
    import dashboard.services.as_service as asv
    monkeypatch.setattr(asv, 'search_confirmed_projects', lambda q, limit=100: [{'code': 'G1-AA', 'biz': 'x', 'address': 'y'}])
    got = []
    fn = as_app.h[('options', 'value')]
    fn(ack=lambda **k: got.append(k), body={'block_id': 'as_project_code', 'value': 'G', 'user': {'id': 'U_FIELD'}})
    fn(ack=lambda **k: got.append(k), body={'block_id': 'as_project_code', 'value': 'G', 'user': {'id': 'U_STAFF'}})
    assert got[0] == {'options': []} and got[1]['options'][0]['value'] == 'G1-AA'
