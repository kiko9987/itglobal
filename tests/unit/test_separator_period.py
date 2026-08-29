# -*- coding: utf-8 -*-
"""번지-상세 사이 구분 마침표 제거 (L-03843).

고객이 번지와 상세 사이에 마침표를 구분자로 넣은 것('42 . 지하1층'·'42.지하1층') → 공백.
소수점(뒤가 숫자)은 미매치. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _post_normalize_display as P


@pytest.mark.parametrize('addr, expected', [
    ('중랑구 면목로29길 42 . 지하 1층', '중랑구 면목로29길 42 지하 1층'),
    ('중랑구 면목로29길 42.지하1층', '중랑구 면목로29길 42 지하 1층'),
    ('X구 Y로 1-5 . 3층', 'X구 Y로 1-5 3층'),
])
def test_separator_period_removed(addr, expected):
    assert P(addr) == expected


@pytest.mark.parametrize('addr', [
    'X구 Y로 152 3층',       # 마침표 없음 — 무변
    'X구 Y로 152 3층 상가',
])
def test_untouched(addr):
    assert P(addr) == addr


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
