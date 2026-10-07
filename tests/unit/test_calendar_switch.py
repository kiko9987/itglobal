# -*- coding: utf-8 -*-
"""캘린더 연동 운영 스위치 GOOGLE_CALENDAR_ENABLED (2026-10-06 '당분간 안 함')."""
import sys

sys.path.insert(0, '.')

from dashboard.services import calendar_service as cs


def _configured(monkeypatch):
    monkeypatch.setattr(cs, 'get_client_secret_file', lambda: 'secret.json')
    monkeypatch.setenv('GOOGLE_CALENDAR_ID', 'c_test@group.calendar.google.com')


def test_switch_off_disables_even_when_configured(monkeypatch):
    _configured(monkeypatch)
    for v in ('false', 'FALSE', '0', 'off', 'no', ' false '):
        monkeypatch.setenv('GOOGLE_CALENDAR_ENABLED', v)
        assert cs.is_calendar_enabled() is False


def test_default_and_true_keep_existing_behavior(monkeypatch):
    _configured(monkeypatch)
    monkeypatch.delenv('GOOGLE_CALENDAR_ENABLED', raising=False)
    assert cs.is_calendar_enabled() is True
    monkeypatch.setenv('GOOGLE_CALENDAR_ENABLED', 'true')
    assert cs.is_calendar_enabled() is True


def test_disabled_calls_are_noop(monkeypatch):
    _configured(monkeypatch)
    monkeypatch.setenv('GOOGLE_CALENDAR_ENABLED', 'false')
    monkeypatch.setattr(cs, 'get_calendar_service',
                        lambda: (_ for _ in ()).throw(AssertionError('API 호출되면 안 됨')), raising=False)
    assert cs.create_project_calendar_event({'프로젝트 코드': 'R4015-SJ'}) is None
    cs.update_project_calendar_event({'프로젝트 코드': 'R4015-SJ'})
    cs.delete_project_calendar_event('R4015-SJ')
