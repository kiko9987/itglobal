# -*- coding: utf-8 -*-
"""평수(면적) 이후 절단 — 주소칸에 섞인 문의/상세 제거 (L-03774).

'…이마트안산고잔점2층 25평 인테리어공사시작단계' → '…이마트안산고잔점2층'. 공백+숫자+평
(뒤 한글 아님) 부터 줄 끝까지 제거. '평택로'·'평화빌딩'(앞 공백+숫자 없음)은 보존. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.lead_helpers import (
    _normalize_road_spacing as N, extract_address_overflow as OV,
)


@pytest.mark.parametrize('raw, expected', [
    ('경기도안산시 단원구 원포공원1로 46 이마트안산고잔점2층 25평 인테리어공사시작단계',
     '25평 인테리어공사시작단계'),
    ('강남구 X로 1 40평 2대 견적문의', '40평 2대 견적문의'),
    ('강남구 테헤란로 152 삼성빌딩 3층', ''),   # 평 없음
    ('평택시 평택로 25 3층', ''),               # 평택(오검출 방지)
])
def test_address_overflow(raw, expected):
    assert OV(raw) == expected


@pytest.mark.parametrize('raw, expected', [
    ('경기도안산시 단원구 원포공원1로 46 이마트안산고잔점2층 25평 인테리어공사시작단계',
     '경기도안산시 단원구 원포공원1로 46 이마트안산고잔점2층'),
    ('강남구 테헤란로 152 삼성빌딩 3층 40평 2대 견적',
     '강남구 테헤란로 152 삼성빌딩 3층'),
])
def test_pyeong_and_after_cut(raw, expected):
    assert N(raw) == expected


@pytest.mark.parametrize('text', [
    '평택시 평택로 25 평화빌딩 3층',   # 평택/평화 — 앞 공백+숫자 없음, 보존
    '강남구 테헤란로 152 3층',          # 평 없음
])
def test_no_false_cut(text):
    assert N(text) == text


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
