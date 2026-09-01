# -*- coding: utf-8 -*-
"""옛 '군' → '시' 승격 지명 후보 생성 (L-03863).

포천군→포천시처럼 시로 승격된 지명은 카카오/행안부가 현재명(시)만 인식 → 옛 '군' 지번은
변환 실패. 'XX군'(행정 토큰) → 'XX시' 후보를 만들고, 실제 채택은 resolve_address 가
재귀 verify 로 게이팅(아직 군인 가평군 등은 verify 실패로 원본 유지). 여기선 순수 후보
생성만 검증. '군포시'(군이 접두)·'국군병원'(경계 아님)은 미치환.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _promote_gun_to_si as G


@pytest.mark.parametrize('raw, expected', [
    ('경기도 포천군 내촌면 소학리 2-6', '경기도 포천시 내촌면 소학리 2-6'),
    ('가평군 청평면 청평리 100', '가평시 청평면 청평리 100'),   # 후보만(verify-gate가 거부)
    ('경기도 여주군 여주읍 5', '경기도 여주시 여주읍 5'),
])
def test_promote(raw, expected):
    assert G(raw) == expected


@pytest.mark.parametrize('raw', [
    '군포시 산본동 100',     # '군'이 접두 — 미치환
    '국군병원 앞 5',         # 경계(공백/끝) 아님 — 미치환
    '서울 강남구 테헤란로 1',  # 군 없음
])
def test_untouched(raw):
    assert G(raw) == raw


def test_none_safe():
    assert G(None) is None
    assert G('') == ''


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
