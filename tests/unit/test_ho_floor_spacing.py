# -*- coding: utf-8 -*-
"""한글+숫자 층/호/관/동 붙임 분리 — building_tail 보강 (L-03754).

'상가2238호'→'상가 2238호'. building_tail 은 base normalize(호/층 분리)를 안 거쳐
_post_normalize_display 최종 단계에서 보강. 행정동 접미(N동주민센터)·이미 공백·비한글 앞
(B1층)은 미변경. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _post_normalize_display as P


@pytest.mark.parametrize('addr, expected', [
    ('성남 위례광장로 104 센트럴스퀘어 상가2238호', '성남 위례광장로 104 센트럴스퀘어 상가 2238호'),
    ('강남구 X로 1 래미안101동 502호', '강남구 X로 1 래미안 101동 502호'),
    ('서초구 Y로 2 지하1층 카페', '서초구 Y로 2 지하 1층 카페'),
    ('중구 Z로 3 빌딩3관', '중구 Z로 3 빌딩 3관'),
])
def test_split_glued_floor_ho(addr, expected):
    assert P(addr) == expected


@pytest.mark.parametrize('addr', [
    '노원구 마들로3길 37 행당제1동주민센터',   # 행정동 접미 — 분리 금지
    '강남구 X로 1 상가 2238호',                # 이미 공백
    '서초구 Y로 2 B1층',                       # 비한글 앞(B)
])
def test_preserved(addr):
    assert P(addr) == addr


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
