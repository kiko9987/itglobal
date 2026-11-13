# -*- coding: utf-8 -*-
"""보안 점검(2026-11-13) 중간 우선순위 수정 회귀 테스트.

1) 웹 결재 신원 — 인사카드 slack_id 는 슬랙이 아는 본인 ID 와 같을 때만 인정
2) 입금 메모 API — 관리자 전용
3) 캘린더 — 로그인 때 토큰 저장 안 함, 수동 동기화는 관리자·연동 켜짐일 때만
4) SMS 수신마다 발신번호 로그
"""
import inspect
import logging
import re
from unittest.mock import patch

import pytest

from dashboard.services import gw_approval as ga


# ── 1) 웹 결재 신원 ──────────────────────────────────────────────────
@pytest.fixture(autouse=True)
def _clear_sid_cache():
    ga._SID_CACHE.clear()
    yield
    ga._SID_CACHE.clear()


def test_verified_slack_id_match():
    with patch.object(ga, '_lookup_slack_id', return_value='U_SELF'):
        assert ga.verified_slack_id('a@x.com', 'U_SELF') == 'U_SELF'


def test_verified_slack_id_swapped_card_rejected():
    # 인사 관리자가 자기 카드 slack_id 를 결재자 ID 로 바꿔도 슬랙이 아는 본인 ID 와 달라 인정 안 됨
    with patch.object(ga, '_lookup_slack_id', return_value='U_SELF'):
        assert ga.verified_slack_id('a@x.com', 'U_APPROVER') == ''


@pytest.mark.parametrize('lookup', [None, ''])
def test_verified_slack_id_lookup_fail_or_unknown(lookup):
    with patch.object(ga, '_lookup_slack_id', return_value=lookup):
        assert ga.verified_slack_id('a@x.com', 'U_SELF') == ''


def test_verified_slack_id_empty_inputs():
    with patch.object(ga, '_lookup_slack_id') as lk:
        assert ga.verified_slack_id('', 'U1') == ''
        assert ga.verified_slack_id('a@x.com', '') == ''
        lk.assert_not_called()


def test_verified_slack_id_cached():
    with patch.object(ga, '_lookup_slack_id', return_value='U_SELF') as lk:
        ga.verified_slack_id('A@X.com', 'U_SELF')
        ga.verified_slack_id('a@x.com', 'U_SELF')
        assert lk.call_count == 1


def test_viewer_uses_verified_id_for_privileges():
    card = {'slack_id': 'U_APPROVER', 'name': '경영지원'}
    with patch.object(ga.hr, 'find_by_email', return_value=card), \
         patch.object(ga.hr, 'is_hr_admin', return_value=False), \
         patch.object(ga.ea, 'approver_ids', return_value=('U_APPROVER',)), \
         patch.object(ga.ea, 'final_approver_ids', return_value=()), \
         patch.object(ga.ea, 'executor_ids', return_value=()), \
         patch.object(ga.ea, 'n_executor_ids', return_value=()), \
         patch.object(ga, '_lookup_slack_id', return_value='U_SELF'):
        v = ga.viewer('sb@example.com')
    assert v['slack_id'] == '' and v['see_all'] is False
    assert ga.act(1, 'approve', v) is not None   # 처리 거부 문구


def test_archive_detail_requires_name():
    class _Conn:
        def execute(self, *a):
            class _R:
                def fetchone(self_inner):
                    return ('D1', 't', 'f', 's', '', 'c', 'u', 0, '', '{}')
            return _R()

    from contextlib import contextmanager

    @contextmanager
    def _fake():
        yield _Conn()

    with patch.object(ga, '_aconn', _fake):
        assert ga.archive_detail('D1', {'see_all': False, 'name': ''}) is None
        assert ga.archive_detail('D1', {'see_all': True, 'name': ''}) is not None


# ── 2) 입금 메모 API 관리자 전용 ─────────────────────────────────────
def _decorators_of(src, route):
    i = src.index(route)
    head = src[i:i + 400].split('def ', 1)[0]
    return head


def test_field_memo_endpoints_admin_only():
    from dashboard.blueprints import projects
    src = inspect.getsource(projects)
    for route in ("'/api/projects/field-memo'", "'/api/projects/field-memos/batch'"):
        head = _decorators_of(src, route)
        assert '@admin_required' in head and '@editor_required' not in head, route


# ── 3) 캘린더 ────────────────────────────────────────────────────────
def test_login_callback_does_not_save_calendar_token():
    from dashboard.blueprints import auth as auth_bp
    src = inspect.getsource(auth_bp.google_callback)
    assert 'save_calendar_credentials(' not in src


def test_calendar_routes_admin_only():
    from dashboard.blueprints import projects, auth as auth_bp
    assert '@admin_required' in _decorators_of(inspect.getsource(projects), "'/api/projects/calendar/sync'")
    asrc = inspect.getsource(auth_bp)
    for route in ("'/auth/google/calendar')", "'/auth/google/calendar/callback')"):
        assert '@admin_required' in _decorators_of(asrc, route), route


def test_run_sync_once_skips_when_disabled():
    from dashboard.services import calendar_sync_scheduler as cs
    sch = cs.CalendarSyncScheduler.__new__(cs.CalendarSyncScheduler)
    with patch('dashboard.services.calendar_service.is_calendar_enabled', return_value=False), \
         patch.object(cs.CalendarSyncScheduler, '_sync_deleted_events') as run:
        sch.run_sync_once()
    run.assert_not_called()


# ── 4) SMS 수신 로그 ─────────────────────────────────────────────────
def test_sms_inbound_logs_sender_for_every_message(monkeypatch, caplog):
    from flask import Flask
    import dashboard.blueprints.sms_inbound as si
    monkeypatch.setenv('SMS_INBOUND_TOKENS', 'sb:secretSB')
    monkeypatch.delenv('SMS_BANK_SENDERS', raising=False)
    monkeypatch.setattr(si, 'ingest_deposit', lambda *a, **k: {'status': 'ignored', 'reason': 'test'})
    app = Flask(__name__)
    app.register_blueprint(si.sms_bp)
    caplog.set_level(logging.INFO)
    app.test_client().post('/sms/inbound', data={'token': 'secretSB', 'sender': '010-1234-5678', 'text': '출금 1원'})
    msgs = [r.getMessage() for r in caplog.records]
    assert any(re.search(r"수신 device=sb sender='010\*+5678'", m) for m in msgs), msgs
