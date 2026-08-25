# -*- coding: utf-8 -*-
"""단지명 반복 접두 병합 + 인접 동일 토큰 제거 (L-03783).

다음 위젯 building_name '종암2차 아이파크' unwrap 후 고객이 상세에 '아이파크상가동'을
다시 붙여 아이파크 중복 → 접두만 제거하고 구역(상가동)은 보존. 숫자 동은 base 분리로
'자이 자이 101동'이 되므로 인접 동일 토큰 제거로 커버. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _post_normalize_display as P


@pytest.mark.parametrize('addr, expected', [
    # 접두 병합 — 나머지 '…동' 구역 보존
    ('성북구 종암로9길 86 종암2차 아이파크 아이파크상가동 1층104호',
     '성북구 종암로9길 86 종암2차 아이파크 상가동 1층 104호'),
    ('X구 로 1 하늘채 하늘채관리동 2층', 'X구 로 1 하늘채 관리동 2층'),
    # 숫자 동 — base 분리 후 인접 동일 토큰 제거
    ('X구 로 1 자이 자이101동 302호', 'X구 로 1 자이 101동 302호'),
])
def test_prefix_dedup(addr, expected):
    assert P(addr) == expected


@pytest.mark.parametrize('addr', [
    'X구 성수동 성수동물병원',        # 나머지 '물병원' 동 아님 — 보존
    'X구 로 1 래미안 래미안아파트',   # 나머지 '아파트' 동 아님 — 보존
    'X구 로 1 상가동 사무실',         # 접두관계 아님 — 보존
    'X구 종로 1가 2가',              # 동일 아님 — 보존
])
def test_preserved(addr):
    assert P(addr) == addr


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
