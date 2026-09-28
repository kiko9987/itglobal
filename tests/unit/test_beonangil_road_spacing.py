# -*- coding: utf-8 -*-
"""번안길·안길 도로 접미 스페이싱 정규화 (L-04089).

'성현로 135번안길 80-7'(성현로 뒤 공백)은 도로명 '성현로135번안길'이라 붙여야 검증됨.
기존 정규식은 '번?길'만 인식 → '번안길'을 도로명이 아닌 번지+건물로 오분리하고, 이미 붙은
'성현로135번안길'을 '성현로 135 번안길'로 쪼갬. '번?안?길' 가족(길·안길·번길·번안길) 인식.
순수함수(네트워크 미사용).
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.lead_helpers import _normalize_road_spacing as N
from dashboard.services.address_resolver import _join_road_gil as J


@pytest.mark.parametrize('raw, expected', [
    # 로 뒤 공백 + 숫자+번안길 → 붙임
    ('일산동구 성현로 135번안길 80-7', '일산동구 성현로135번안길 80-7'),
    ('성현로 135안길 12', '성현로135안길 12'),
    # 이미 붙은 것은 그대로(쪼개지 않음)
    ('성현로135번안길 80-7', '성현로135번안길 80-7'),
    # 기존 번길·길 회귀 없음
    ('상도로 13길 4', '상도로13길 4'),
    ('부천로 431번길 16', '부천로431번길 16'),
])
def test_normalize_road_spacing(raw, expected):
    assert N(raw) == expected


@pytest.mark.parametrize('raw, expected', [
    # 공백 두 개(성현로 135 번안길) → 붙임
    ('일산동구 성현로 135 번안길 80-7', '일산동구 성현로135번안길 80-7'),
    ('언주로 107 길 27', '언주로107길 27'),   # 기존 케이스 회귀 없음
])
def test_join_road_gil(raw, expected):
    assert J(raw) == expected


def test_plain_beonji_not_glued():
    # '성현로 135'(번지, 뒤에 길 접미 없음)는 붙이지 않음(실제 번지 보존)
    assert N('성현로 135 80-7') == '성현로 135 80-7'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
