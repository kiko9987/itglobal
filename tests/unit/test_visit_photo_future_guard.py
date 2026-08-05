# -*- coding: utf-8 -*-
"""방문일 미래 시 사진 첨부 자동완료 skip 게이트 회귀 테스트 (2026-08-05).

L-03575 사고: 08-06 방문 거래처 건에 도면(참고)을 등록 시 첨부했는데, 사진 첨부
자동완료가 방문일과 무관하게 발동해 '방문 완료' 처리됨. 방문일이 미래면 첨부는
참고자료로 보고 완료 skip (사진 저장은 유지). 파싱 실패·빈값은 기존 동작(허용).
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.blueprints.slack_bot as sb


def test_future_date_blocks():
    """먼 미래 방문일 → True (자동완료 skip)."""
    assert sb._visit_not_yet_reached('2999-01-01') is True


def test_future_date_range_uses_start():
    """범위 방문일도 시작일 기준 — 미래면 True."""
    assert sb._visit_not_yet_reached('2999-01-01~2999-01-03') is True


def test_past_date_allows():
    """지난 방문일 → False (자동완료 허용)."""
    assert sb._visit_not_yet_reached('2000-01-01') is False


def test_empty_or_garbage_allows():
    """빈값·파싱불가 → False (기존 동작 유지, 현장 사진 UX 보존)."""
    assert sb._visit_not_yet_reached('') is False
    assert sb._visit_not_yet_reached(None) is False
    assert sb._visit_not_yet_reached('미정') is False


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
