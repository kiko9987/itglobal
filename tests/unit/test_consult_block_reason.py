# -*- coding: utf-8 -*-
"""재상담 제출 차단 판정 (2026-11-06).

방문을 마친 리드도 리드 채널 카드 [재상담]으로 다시 방문 예약을 잡게 하면서,
공사 확정 리드의 상태 덮어쓰기와 아직 방문 전인 리드의 방문 카드 중복만 막는다.
"""
import sys
sys.path.insert(0, '.')

from datetime import date

import pytest
from dashboard.blueprints import slack_bot
from dashboard.blueprints.slack_bot import _consult_block_reason

TODAY = date(2026, 11, 6)


@pytest.fixture(autouse=True)
def _not_completed(monkeypatch):
    monkeypatch.setattr(slack_bot, '_visit_is_completed', lambda lead_no, consult='': False)


def _lead(status, visit_date='', consult=''):
    return {'상태': status, '방문 예정일': visit_date, '상담 내용': consult}


def test_revisit_after_past_visit_allowed():
    """한참 전에 방문한 리드(상태가 방문 예약으로 남음) — 다시 예약 가능."""
    assert _consult_block_reason(_lead('방문 예약', '2026-08-20'), 'L-03700', '방문 예약', TODAY) == ''


def test_revisit_after_visit_done_status_allowed():
    assert _consult_block_reason(_lead('방문 완료', '2026-09-01'), 'L-03800', '방문 예약', TODAY) == ''


def test_upcoming_visit_blocks_second_booking():
    msg = _consult_block_reason(_lead('방문 예약', "'2026-11-08"), 'L-04490', '방문 예약', TODAY)
    assert '11/8' in msg and '정보 수정' in msg


def test_today_visit_counts_as_upcoming():
    assert _consult_block_reason(_lead('방문 예약', '2026-11-06'), 'L-04491', '방문 예약', TODAY)


def test_range_uses_end_date():
    assert _consult_block_reason(_lead('방문 예약', '2026-11-05~07'), 'L-04492', '방문 예약', TODAY)


def test_upcoming_but_completed_allowed(monkeypatch):
    """사진 자동완료 등으로 이미 완료 표시가 있으면 날짜가 남아 있어도 다시 예약 가능."""
    monkeypatch.setattr(slack_bot, '_visit_is_completed', lambda lead_no, consult='': True)
    assert _consult_block_reason(_lead('방문 예약', '2026-11-08'), 'L-04493', '방문 예약', TODAY) == ''


def test_upcoming_visit_other_status_allowed():
    """방문 전이라도 견적 제출·문의 드랍 등으로 바꾸는 재상담은 지금처럼 허용."""
    assert _consult_block_reason(_lead('방문 예약', '2026-11-08'), 'L-04494', '문의 드랍', TODAY) == ''


def test_no_date_allowed():
    assert _consult_block_reason(_lead('방문 예약', '-'), 'L-04495', '방문 예약', TODAY) == ''


def test_confirmed_project_blocked_any_status():
    for st in ('방문 예약', '유선 상담', '문의 드랍'):
        assert '공사가 확정' in _consult_block_reason(_lead('공사 확정'), 'L-04000', st, TODAY)


def test_first_consult_and_phone_lead_allowed():
    assert _consult_block_reason(_lead(''), 'L-04496', '방문 예약', TODAY) == ''
    assert _consult_block_reason(_lead('유선 상담'), 'L-04497', '방문 예약', TODAY) == ''
    assert _consult_block_reason({}, 'L-04498', '방문 예약', TODAY) == ''
