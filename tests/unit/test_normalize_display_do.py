# -*- coding: utf-8 -*-
"""normalize_display 도(道) 접두 처리 — 비수도권 도 유지 / 경기 제거 (L-03646).

2026-08-06 사용자 결정: '전북 김제'처럼 비수도권 도는 도 접두 유지(담당자 먼 지역
인지용), 경기(수도권)는 기존대로 제거. 순수함수 — 네트워크 없음.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import normalize_display as N


@pytest.mark.parametrize('addr, expected', [
    # 비수도권 도 — 도 유지 + 시/군 제거
    ('전북특별자치도 김제시 남북로 218', '전북 김제 남북로 218'),
    ('전북 김제시 남북로 218', '전북 김제 남북로 218'),
    ('전라남도 순천시 중앙로 100', '전남 순천 중앙로 100'),
    ('경북 김천시 시청로 1', '경북 김천 시청로 1'),
    ('경남 창원시 의창구 중앙대로 1', '경남 창원 의창구 중앙대로 1'),
    ('강원특별자치도 춘천시 중앙로 1', '강원 춘천 중앙로 1'),
    ('충남 아산시 배방읍 배방로 25', '충남 아산 배방읍 배방로 25'),
    ('전남 완주군 소양면 화심로 1', '전남 완주 소양면 화심로 1'),
])
def test_non_metro_do_kept(addr, expected):
    assert N(addr) == expected


@pytest.mark.parametrize('addr, expected', [
    # 경기(수도권) — 도 제거 (기존 유지)
    ('경기 수원시 영통구 광교로 145', '수원 영통구 광교로 145'),
    ('경기도 성남시 분당구 판교로 1', '성남 분당구 판교로 1'),
    ('경기 광주시 경안로 100', '광주 경안로 100'),
])
def test_gyeonggi_dropped(addr, expected):
    assert N(addr) == expected


@pytest.mark.parametrize('addr, expected', [
    # 회귀 — 서울/광역시/광주광역시 규칙 불변
    ('서울 강남구 테헤란로 152', '강남구 테헤란로 152'),
    ('인천 연수구 갯벌로 36', '인천 연수구 갯벌로 36'),
    ('광주 동구 충장로 1', '전남 광주 동구 충장로 1'),
])
def test_others_unchanged(addr, expected):
    assert N(addr) == expected


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
