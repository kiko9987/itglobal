# -*- coding: utf-8 -*-
"""공사 취소·재개 → 연결 리드 상태 맞추기 (2026-11-16)."""
import sys

sys.path.insert(0, '.')

import pandas as pd
import pytest

from dashboard.services import lead_project_link as lpl


@pytest.fixture
def env(monkeypatch):
    leads = {}
    writes = []
    projects = []

    monkeypatch.setattr(lpl, '_lead_status', lambda no: leads.get(no))

    def _set(no, st):
        writes.append((no, st))
        leads[no] = st
        return True
    monkeypatch.setattr(lpl, '_set_status', _set)
    monkeypatch.setattr(lpl, '_projects_df', lambda: pd.DataFrame(
        projects, columns=['프로젝트 코드', 'Lead No', '수금 관련 특이사항']))
    return type('E', (), {'leads': leads, 'writes': writes, 'projects': projects})


def test_lead_nos_of():
    assert lpl.lead_nos_of('L-04216') == ['L-04216']
    assert lpl.lead_nos_of(' L-1, L-2 ') == ['L-1', 'L-2']
    assert lpl.lead_nos_of('ETC-0012') == ['ETC-0012']
    for v in ('', '-', None, 'nan'):
        assert lpl.lead_nos_of(v) == []


def test_cancel_moves_confirmed_lead(env):
    env.leads['L-1'] = '공사 확정'
    env.projects.append(['G1-AA', 'L-1', '공사 취소 (사유: 고객 변심)'])
    assert lpl.sync_lead_status('G1-AA', {'Lead No': 'L-1'}, cancelled=True) == ['L-1']
    assert env.writes == [('L-1', '공사 취소')]


def test_cancel_keeps_lead_when_other_project_live(env):
    env.leads['L-1'] = '공사 확정'
    env.projects += [['G1-AA', 'L-1', '공사 취소'], ['G2-AA', 'L-1', '-']]
    assert lpl.sync_lead_status('G1-AA', {'Lead No': 'L-1'}, cancelled=True) == []
    assert env.writes == []


def test_cancel_when_other_project_also_cancelled(env):
    env.leads['L-1'] = '공사 확정'
    env.projects += [['G1-AA', 'L-1', '-'], ['G2-AA', 'L-1', '공사 취소 (사유: x)']]
    # G1 은 지금 막 취소(캐시엔 아직 '-') — 자기 자신은 빼고 본다
    assert lpl.sync_lead_status('G1-AA', {'Lead No': 'L-1'}, cancelled=True) == ['L-1']


@pytest.mark.parametrize('status', ['방문 완료', '문의 드랍', '공사 취소', '견적 제출'])
def test_cancel_leaves_other_statuses(env, status):
    env.leads['L-1'] = status
    assert lpl.sync_lead_status('G1-AA', {'Lead No': 'L-1'}, cancelled=True) == []
    assert env.writes == []


def test_cancel_stops_when_projects_unreadable(env, monkeypatch):
    env.leads['L-1'] = '공사 확정'
    monkeypatch.setattr(lpl, '_projects_df', lambda: None)
    assert lpl.sync_lead_status('G1-AA', {'Lead No': 'L-1'}, cancelled=True) == []
    assert env.writes == []


def test_resume_restores_cancelled_lead(env):
    env.leads['L-1'] = '공사 취소'
    assert lpl.sync_lead_status('G1-AA', {'Lead No': 'L-1'}, cancelled=False) == ['L-1']
    assert env.writes == [('L-1', '공사 확정')]


@pytest.mark.parametrize('status', ['방문 완료', '공사 확정', '유선 상담'])
def test_resume_leaves_other_statuses(env, status):
    env.leads['L-1'] = status
    assert lpl.sync_lead_status('G1-AA', {'Lead No': 'L-1'}, cancelled=False) == []


def test_missing_lead_no_looks_up_projects(env):
    env.leads['L-9'] = '공사 취소'
    env.projects.append(['G9-AA', 'L-9', '-'])
    assert lpl.sync_lead_status('G9-AA', {}, cancelled=False) == ['L-9']


def test_no_lead_or_lead_not_found(env):
    assert lpl.sync_lead_status('G1-AA', {'Lead No': '-'}, cancelled=True) == []
    assert lpl.sync_lead_status('G1-AA', {'Lead No': 'L-404'}, cancelled=True) == []
    assert env.writes == []


def test_write_failure_reported_not_raised(env, monkeypatch):
    env.leads['L-1'] = '공사 취소'
    monkeypatch.setattr(lpl, '_set_status', lambda no, st: False)
    assert lpl.sync_lead_status('G1-AA', {'Lead No': 'L-1'}, cancelled=False) == []


def test_async_skipped_under_pytest(monkeypatch):
    called = []
    monkeypatch.setattr(lpl, 'sync_lead_status', lambda *a: called.append(a))
    monkeypatch.delenv('LEAD_LINK_TEST', raising=False)
    lpl.sync_lead_status_async('G1-AA', {}, True)
    assert called == []


def test_cancel_and_resume_hooks_call_sync(monkeypatch):
    """공용 취소·재개 함수가 리드 맞추기를 부른다 (슬랙 쪽 동작은 막고 시작 부분만 본다)."""
    import dashboard.blueprints.slack_bot as sb
    calls = []
    monkeypatch.setattr(lpl, 'sync_lead_status_async', lambda code, p, cancelled: calls.append((code, cancelled)))
    monkeypatch.setattr(sb, '_project_client', lambda: None)
    # _dm_client 는 메인봇 토큰으로 따로 만든다 — 안 막으면 경영지원 4명에게 진짜 취소 DM 이
    #   나간다(2026-11-17 가짜 DM 12건). 루트 conftest 가 한 겹 더 막지만 여기서도 끊는다.
    monkeypatch.setattr(sb, '_dm_client', lambda: None)
    for fn in (sb.apply_project_cancel_to_slack, sb.apply_project_uncancel_to_slack):
        try:
            fn('G1-AA', {'Lead No': 'L-1'}, 'JW')
        except Exception:
            pass                                       # 카드·댓글 뒷단계는 범위 밖
    assert calls == [('G1-AA', True), ('G1-AA', False)]


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
