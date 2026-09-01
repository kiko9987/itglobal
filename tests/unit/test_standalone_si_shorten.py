# -*- coding: utf-8 -*-
"""단독 'XX시'(도 접두 없음) 축약 — 뒤가 도로명·법정동·번지여도 시 제거 (L-03867).

'동두천시 아차노리로…'·'수원시 매영로 …'처럼 구 없는 시나 구 생략 입력에서 시가 남던 갭.
도 접두 있는 비수도권(전북 김제)·광역시(인천)·서울은 기존 규칙 유지. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import normalize_display as N


@pytest.mark.parametrize('addr, expected', [
    ('동두천시 아차노리로52번길 59-7 단독주택', '동두천 아차노리로52번길 59-7 단독주택'),
    ('수원시 매영로 95', '수원 매영로 95'),
    ('포천시 내촌면 금강로 2947', '포천 내촌면 금강로 2947'),
    ('고양시 덕양구 향기2로 21', '고양 덕양구 향기2로 21'),
    ('제주시 노형동 925', '제주 노형동 925'),
])
def test_standalone_si_shortened(addr, expected):
    assert N(addr) == expected


@pytest.mark.parametrize('addr, expected', [
    ('전북 김제시 남북로 218', '전북 김제 남북로 218'),   # 비수도권 도 접두 유지
    ('인천 연수구 갯벌로 36', '인천 연수구 갯벌로 36'),   # 광역시 유지
    ('서울 강남구 테헤란로 152', '강남구 테헤란로 152'),  # 서울 제거
])
def test_existing_rules_intact(addr, expected):
    assert N(addr) == expected


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
