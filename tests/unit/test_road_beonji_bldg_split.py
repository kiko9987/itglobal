# -*- coding: utf-8 -*-
"""도로명+번지+건물명 완전 붙여쓰기 분할 (L-03806).

'가람로124봉명타워' → '가람로 124 봉명타워'. 도로연속(로124번길·언주로107길)·법정동
(세종로1가)·번지단위(124동/호/층)는 보존. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.lead_helpers import _normalize_road_spacing as N


@pytest.mark.parametrize('raw, expected', [
    ('가람로124봉명타워', '가람로 124 봉명타워'),
    ('파주시 가람로124봉명타워', '파주시 가람로 124 봉명타워'),
    ('테헤란로152빌딩', '테헤란로 152 빌딩'),
    ('가람로124-5봉명타워', '가람로 124-5 봉명타워'),
])
def test_split(raw, expected):
    assert N(raw) == expected


@pytest.mark.parametrize('raw', [
    '언주로107길',        # 로+숫자+길 도로연속 — 보존
    '가람로124번길',      # 로+숫자+번길 도로연속 — 보존
    '판교로241번길',      # 보존
    '세종로1가',          # 법정동 번호 — 보존
    '가람로 124동',       # 번지단위(동) — 보존
])
def test_preserved(raw):
    assert N(raw) == raw


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
