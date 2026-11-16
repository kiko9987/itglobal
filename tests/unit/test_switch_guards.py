# -*- coding: utf-8 -*-
"""꺼 둔 기능을 켤 때의 안전장치 (2026-11-16 상태 머신 점검 묶음 H).

① 캘린더 동기화: 한 회차에 캘린더에서 사라진 이벤트가 많으면(다시 켜는 날·연결 문제) 시트 공사 날짜를 비우지 않고 알림.
② 전자결재 주말 시험 모드(EAPPROVAL_EXECUTOR_DM_IDS): 카드에 표시 + 평일 아침이면 관리자 알림.
"""
import sys
from datetime import date

sys.path.insert(0, '.')

import pytest

from dashboard.services import calendar_sync_scheduler as cs


@pytest.fixture
def cal(monkeypatch):
    sch = object.__new__(cs.CalendarSyncScheduler)
    cleared, deleted = [], []

    class _Repo:
        def delete_event(self, code):
            deleted.append(code)
    monkeypatch.setattr(cs, 'get_calendar_service', lambda: object())
    monkeypatch.setattr(cs, 'get_calendar_event_repository', lambda: _Repo())
    import dashboard.services.project_service as ps
    monkeypatch.setattr(ps, 'get_project_records', lambda: [])
    monkeypatch.setattr(sch, '_clear_project_dates', lambda code: cleared.append(code) or True, raising=False)
    monkeypatch.setattr(sch, '_get_calendar_event', lambda svc, cid, eid: None, raising=False)
    monkeypatch.delenv('CALENDAR_SYNC_MAX_DELETES', raising=False)

    def _run(n):
        events = [{'project_code': f'G{4000 + i}-AA', 'calendar_id': 'c', 'event_id': f'e{i}'} for i in range(n)]
        monkeypatch.setattr(sch, '_get_all_calendar_events', lambda repo: events, raising=False)
        sch._sync_deleted_events()
        return cleared, deleted
    return _run


def test_few_deletions_still_clear_dates(cal):
    cleared, deleted = cal(2)
    assert cleared == ['G4000-AA', 'G4001-AA'] and deleted == cleared


def test_many_deletions_held_and_alerted(cal, caplog):
    with caplog.at_level('ERROR'):
        cleared, deleted = cal(10)
    assert cleared == [] and deleted == []                               # 시트 날짜 그대로
    assert any('보류' in r.getMessage() and '10건' in r.getMessage() for r in caplog.records)


def test_max_deletes_env(monkeypatch):
    monkeypatch.setenv('CALENDAR_SYNC_MAX_DELETES', '20')
    assert cs._max_deletes_per_run() == 20
    monkeypatch.setenv('CALENDAR_SYNC_MAX_DELETES', 'abc')
    assert cs._max_deletes_per_run() == 3


# ── 전자결재 주말 시험 모드 ──

def _row(status):
    return {'id': 1, 'status': status, 'form': 'fund', 'amount': 1000, 'requester_id': 'UMGR',
            'requester_name': '강성환', 'partner': '거래처', 'content': '자재'}


def test_card_shows_test_mode(monkeypatch):
    from dashboard.services import e_approval as ea
    monkeypatch.setattr(ea, 'card_text', lambda *a, **k: '본문')
    monkeypatch.setattr(ea, 'buttons', lambda row, role='all': [])
    monkeypatch.delenv('EAPPROVAL_EXECUTOR_DM_IDS', raising=False)
    ctx = lambda bl: [e['text'] for b in bl if b['type'] == 'context' for e in b['elements']]
    assert not any('시험 모드' in t for t in ctx(ea.card_blocks(_row(ea.APPROVED), [], '현장')))
    monkeypatch.setenv('EAPPROVAL_EXECUTOR_DM_IDS', 'U04UL2ZLJAX')
    assert any('시험 모드' in t for t in ctx(ea.card_blocks(_row(ea.APPROVED), [], '현장')))
    assert not any('시험 모드' in t for t in ctx(ea.card_blocks(_row(ea.DONE), [], '현장')))   # 끝난 결재엔 없음


def test_weekday_alert(monkeypatch):
    from dashboard.services import sync_scheduler as ss
    sent = []
    monkeypatch.setattr(ss, '_notify_admin', lambda key, text: sent.append(key))
    monkeypatch.setenv('EAPPROVAL_EXECUTOR_DM_IDS', 'U04UL2ZLJAX')
    assert ss._eap_test_mode_weekday_alert(date(2026, 11, 15)) is False      # 일요일 — 시험 중
    assert ss._eap_test_mode_weekday_alert(date(2026, 11, 16)) is True       # 월요일
    monkeypatch.delenv('EAPPROVAL_EXECUTOR_DM_IDS')
    assert ss._eap_test_mode_weekday_alert(date(2026, 11, 17)) is False
    assert sent == ['eap_test_mode']


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
