# -*- coding: utf-8 -*-
"""기관·건물명 괄호 벗기기 (L-03738).

고객이 '(경희대학교)' 처럼 기관/건물명을 괄호로 감싸면 괄호 제거. 단일 토큰 + 강한
건물/기관 접미(대학교·병원·빌딩 등)만 → 노트·(예정)·법정동·콤마형·(주)는 보존. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _post_normalize_display as P


@pytest.mark.parametrize('addr, expected', [
    ('동대문구 경희대로 26 (경희대학교) 푸른솔문화관 115호',
     '동대문구 경희대로 26 경희대학교 푸른솔문화관 115호'),
    ('강남구 일원로 81 (삼성서울병원) 본관', '강남구 일원로 81 삼성서울병원 본관'),
    ('강남구 테헤란로 1 (강남빌딩)', '강남구 테헤란로 1 강남빌딩'),
    ('성동구 왕십리로 222 (한양대학교) 공대', '성동구 왕십리로 222 한양대학교 공대'),
])
def test_institution_paren_unwrapped(addr, expected):
    assert P(addr) == expected


@pytest.mark.parametrize('addr', [
    '강남구 X로 1 (예정) 2층',            # 계획 마커
    '강남구 X로 1 (주) 2층',              # 법인 표기
    '강남구 X로 1 (101동) 5층',           # 동 번호
    '강남구 X로 1 (주차는 경희대학교 정문) 2층',  # 노트(공백 포함)
])
def test_preserved(addr):
    assert P(addr) == addr


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
