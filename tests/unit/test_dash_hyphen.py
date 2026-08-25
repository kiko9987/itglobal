# -*- coding: utf-8 -*-
"""숫자 사이 대시류 → 하이픈 + 끝 마침표 제거 (L-03769).

갤럭시 등에서 '-' 대신 'ㅡ'(U+3161 한글 모음)·en/em대시·마이너스를 입력하면 번지(2ㅡ9)가
파싱 실패. 숫자 사이(공백 허용)에서만 '-'로 치환 → 일반 텍스트 대시는 보존. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.lead_helpers import _normalize_road_spacing as N
from dashboard.services.address_resolver import _post_normalize_display as P


@pytest.mark.parametrize('raw, expected', [
    ('2ㅡ9', '2-9'),
    ('2 ㅡ 9', '2-9'),
    ('산57ㅡ22', '산57-22'),
    ('101ㅡ3호', '101-3호'),
    ('2–9', '2-9'),   # en dash
    ('2—9', '2-9'),   # em dash
    ('2−9', '2-9'),   # minus
    ('2－9', '2-9'),   # fullwidth hyphen
])
def test_dash_between_digits(raw, expected):
    assert N(raw) == expected


@pytest.mark.parametrize('text', [
    '테헤란로ㅡ길',      # 숫자 사이 아님 — 보존
    'ㅡ자형 건물',       # 보존
    '가나ㅡ다',          # 보존
])
def test_non_digit_dash_preserved(text):
    assert N(text) == text


@pytest.mark.parametrize('addr, expected', [
    ('용인 포곡로234번길 2-9 .', '용인 포곡로234번길 2-9'),
    ('강남구 테헤란로 152.', '강남구 테헤란로 152'),
])
def test_trailing_period_stripped(addr, expected):
    assert P(addr) == expected


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
