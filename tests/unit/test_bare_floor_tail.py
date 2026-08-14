# -*- coding: utf-8 -*-
"""숫자 없는 층 표기(지층·지하·반지하·옥탑) tail 보존 (L-03696).

'선부광장남로109 지층 한국유통'에서 '지층'이 숫자 없는 층이라 _TAIL_SIGNAL 이 신호로
인정 못 해 tail(상호까지) 통째 유실되던 버그. 복합어(지하철·지하상가)는 뒤 한글
lookahead 로 제외. 순수함수 — 네트워크 없음.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import (
    _extract_building_tail as T, _TAIL_SIGNAL,
)


@pytest.mark.parametrize('text, expected_tail', [
    ('선부광장남로109 지층 한국유통', '지층 한국유통'),
    ('선부광장남로109 지하 한국유통', '지하 한국유통'),
    ('선부광장남로109 반지하 상회', '반지하 상회'),
    ('선부광장남로109 옥탑 카페', '옥탑 카페'),
    # 숫자형은 기존대로 유지
    ('선부광장남로109 지하1층 한국유통', '지하1층 한국유통'),
])
def test_bare_floor_tail_preserved(text, expected_tail):
    assert T(text) == expected_tail


@pytest.mark.parametrize('signal_text', ['지층 한국유통', '지하 상회', '반지하 카페', '옥탑 사무실'])
def test_tail_signal_matches_bare_floor(signal_text):
    assert _TAIL_SIGNAL.search(signal_text)


@pytest.mark.parametrize('compound', ['지하철역 3번출구', '지하도 통로', '지하수 개발'])
def test_tail_signal_excludes_compound(compound):
    # 지하+한글 복합어는 층 신호로 오인되면 안 됨 (독립 토큰만)
    # (다른 신호가 없는 순수 복합어에 한해)
    assert not _TAIL_SIGNAL.search(compound)


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
