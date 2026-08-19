# -*- coding: utf-8 -*-
"""_ROAD_PATTERN — '공단1대로195번길'(한글+숫자+한글+로) 형태 지원 (L-03716).

기존 패턴이 숫자 뒤 한글이 오는 도로명(공단1대로)을 못 잡아 행안부 도로확인이 불발돼
[추정]으로 빠졌음. \\d* 뒤에 [가-힣]* 추가. 기존 도로명은 불변. 순수 정규식.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _ROAD_PATTERN as RP


@pytest.mark.parametrize('text, expected', [
    ('공단1대로195번길 38', '공단1대로195번길 38'),   # 신규 지원
    ('공단1대로 38', '공단1대로 38'),
    ('테헤란로 152', '테헤란로 152'),                 # 회귀 방지
    ('봉은사로26길 12', '봉은사로26길 12'),
    ('판교로 393', '판교로 393'),
    ('산업2로 5', '산업2로 5'),
    ('백제고분로19길 13', '백제고분로19길 13'),
])
def test_road_pattern_matches(text, expected):
    m = RP.search(text)
    assert m and m.group(1) == expected


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
