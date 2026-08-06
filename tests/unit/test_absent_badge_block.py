# -*- coding: utf-8 -*-
"""부재중 배지 블록 판정 회귀 테스트 (2026-08-05, L-03527 배지 스택 사고).

모달 부재중 경로가 '이미 배지 있나?'를 '마지막 시도'(배지에 없는 문구)로 오판정 →
매번 새 배지 삽입 → 카드에 부재중 배지가 여러 개 쌓임. _is_absent_badge_block 으로
기존 배지를 전부 식별·제거 후 최신 1개만 prepend 하도록 통일.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.blueprints.slack_bot as sb


def _sec(text):
    return {'type': 'section', 'text': {'type': 'mrkdwn', 'text': text}}


def test_detects_section_badge():
    b = _sec('⠀\n:arrows_counterclockwise: *부재중* (총 *2회*)\n처리자 : JK\n처리 시간 : 08.04 11:17')
    assert sb._is_absent_badge_block(b) is True


def test_detects_context_badge():
    b = {'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': '부재중 (총 1회)'}]}
    assert sb._is_absent_badge_block(b) is True


def test_ignores_inquiry_section():
    b = _sec('⠀\n>:bell: *새 문의 접수 알림 - 온라인 (당근)*  `L-03527`')
    assert sb._is_absent_badge_block(b) is False


def test_ignores_actions_and_nondict():
    assert sb._is_absent_badge_block({'type': 'actions', 'elements': []}) is False
    assert sb._is_absent_badge_block(None) is False
    assert sb._is_absent_badge_block('x') is False


def test_partial_text_not_badge():
    # '*부재중*' 만 있고 '처리 시간' 없으면 배지 아님 (원문 본문 등 오검출 방지)
    assert sb._is_absent_badge_block(_sec('부재중 관련 문의입니다')) is False


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
