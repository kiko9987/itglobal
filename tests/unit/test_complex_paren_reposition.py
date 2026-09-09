# -*- coding: utf-8 -*-
"""끝 괄호 단지명 → 괄호 벗기고 상가/동 앞(도로+번지 뒤)으로 (L-03962).

신도시 상가주소 '도로 번지 상가 N동 N층 N호 (XX도시 N단지)' → '도로 번지 XX도시 N단지
상가 …'. 단지 접미(단지/아파트/마을/타운)로 끝나는 끝 괄호 + 도로+번지 존재 시만. 노트·
법정동·단지접미 없는 괄호는 보존. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _post_normalize_display as P


@pytest.mark.parametrize('addr, expected', [
    ('하남 미사강변서로 167-1 상가 1동 1층 103호 (미사강변도시 17단지)',
     '하남 미사강변서로 167-1 미사강변도시 17단지 상가 1동 1층 103호'),
    ('위례광장로 200 상가 2층 205호 (위례24단지아파트)',
     '위례광장로 200 위례24단지아파트 상가 2층 205호'),
    ('하남 미사강변서로 167-1 (미사강변도시 17단지)',
     '하남 미사강변서로 167-1 미사강변도시 17단지'),
])
def test_complex_paren_repositioned(addr, expected):
    assert P(addr) == expected


@pytest.mark.parametrize('addr', [
    '고잔동 위례광장로 30 (비밀번호 7080)',   # 노트 — 단지 접미 아님, 보존
    '서울 강남구 테헤란로 152 3층',           # 괄호 없음
])
def test_preserved(addr):
    assert P(addr) == addr


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
