# -*- coding: utf-8 -*-
"""카카오 건물명 끝의 영문 로마자 괄호 표기 제거 (L-03693).

'브릿지타워(BRIDGE TOWER)' 처럼 카카오 등록 건물명에 국문+영문이 병기돼 방문 주소에
노이즈로 붙던 것 제거. 국문·순수숫자 괄호는 보존(오제거 방지). 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _strip_roman_paren as S


@pytest.mark.parametrize('raw, expected', [
    # 영문 로마자 괄호 — 제거
    ('브릿지타워(BRIDGE TOWER)', '브릿지타워'),
    ('아이파크(IPARK)', '아이파크'),
    ('롯데캐슬(LOTTE CASTLE)', '롯데캐슬'),
    ('브릿지타워(BRIDGE TOWER) ', '브릿지타워'),
    ('SK V1(S K V1)', 'SK V1'),  # 앞 로마자 본문은 유지, 끝 괄호만
])
def test_roman_paren_stripped(raw, expected):
    assert S(raw) == expected


@pytest.mark.parametrize('raw', [
    '현대(주)',          # 국문 — 보존
    '래미안(101동)',      # 국문 동 포함 — 보존
    '파크뷰(A동)',        # 동(국문) 포함 — 보존
    '타워(1234)',        # 순수 숫자(로마자 없음) — 보존
    '브릿지타워',         # 괄호 없음 — 불변
    '봇들마을2단지이지더원아파트',
])
def test_preserved(raw):
    assert S(raw) == raw


def test_none_and_empty():
    assert S(None) is None
    assert S('') == ''


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
