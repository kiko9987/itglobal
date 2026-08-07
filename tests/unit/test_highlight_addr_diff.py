# -*- coding: utf-8 -*-
"""방문 카드 원본↔변환 diff 강조 (_highlight_addr_diff) 테스트.

2026-08-06: 1자 차이도 볼드(*)로 통일 (기존 홑따옴표 '..' → 볼드, 사용자 요청).
blockquote 컨텍스트 전용, 청크가 alnum 사이면 Word Joiner(⁠) 로 경계 확보.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.blueprints.slack_bot import _highlight_addr_diff

WJ = '⁠'


def test_single_char_now_bold_not_quote():
    """1자 차이(베↔벤)는 볼드 — 더 이상 홑따옴표 아님 (L-03596)."""
    o, c = _highlight_addr_diff('두산베처다임', '두산벤처다임')
    assert '*베*' in o and '*벤*' in c
    assert "'베'" not in o and "'벤'" not in c


def test_single_char_delete_bold():
    """1자 삭제(경기도 광주'시'→광주)도 볼드."""
    o, c = _highlight_addr_diff('경기도 광주시 목동길45번길 1-1',
                                '광주 목동길45번길 1-1')
    assert '*시*' in o
    assert "'시'" not in o


def test_multichar_still_bold():
    """2자↑(제이비 추가)는 기존대로 볼드 (회귀)."""
    o, c = _highlight_addr_diff('강남구 논현로 841 미소빌딩',
                                '강남구 논현로 841 제이비미소빌딩')
    assert '*제이비*' in c


def test_word_joiner_boundary():
    """청크가 한글 사이에 끼면 Word Joiner 로 mrkdwn 경계 확보."""
    o, c = _highlight_addr_diff('두산베처다임', '두산벤처다임')
    # 두산 *벤* 처다임 — 앞뒤 한글이라 WJ 삽입
    assert f'{WJ}*벤*{WJ}' in c


def test_identical_no_highlight():
    """동일하면 원본 그대로 (강조 없음)."""
    assert _highlight_addr_diff('서울 강남구 A', '서울 강남구 A') == ('서울 강남구 A', '서울 강남구 A')


def test_noise_only_diff_skipped():
    """공백/문장부호만 다르면 하이라이트 스킵 (의미 청크 아님)."""
    o, c = _highlight_addr_diff('서울 강남구 415, 두산', '서울 강남구 415 두산')
    assert '*' not in o and '*' not in c


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
