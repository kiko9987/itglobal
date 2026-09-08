# -*- coding: utf-8 -*-
"""법정동+건물명만 POI 구제 가드 (L-03956).

'감이동 벨솔레파크' → '하남 감일중앙로 60 감일벨솔레파크'. 강가드: ①도로·번지·지역
토큰 없어야 진입 ②법정동이 전국 단일 시/구(행안부) ③POI 지역 만장일치 ④건물명이
place_name 꼬리. 네트워크는 monkeypatch 로 차단(hermetic).
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


def _mk_kakao(docs):
    return lambda url: {'documents': docs}


@pytest.fixture(autouse=True)
def _no_net(monkeypatch):
    # 기본: 법정동 단일(1), kakao 빈 응답 — 개별 테스트가 덮어씀
    monkeypatch.setattr(ar, '_dong_region_count', lambda d: 1)
    monkeypatch.setattr(ar, '_kakao_get_json', _mk_kakao([]))


def test_resolve_dong_building(monkeypatch):
    monkeypatch.setattr(ar, '_kakao_get_json', _mk_kakao([
        {'place_name': '감일벨솔레파크', 'address_name': '경기 하남시 감이동 441-1',
         'road_address_name': '경기 하남시 감일중앙로 60'},
    ]))
    assert ar._dong_building_poi_fallback('감이동 벨솔레파크') == \
        '하남 감일중앙로 60 감일벨솔레파크'


def test_reject_multi_city_dong(monkeypatch):
    # 법정동이 여러 시/구(신정동=5) → 거부 (kakao 결과 있어도)
    monkeypatch.setattr(ar, '_dong_region_count', lambda d: 5)
    monkeypatch.setattr(ar, '_kakao_get_json', _mk_kakao([
        {'place_name': '목련아파트', 'address_name': '울산 남구 신정동 100',
         'road_address_name': '울산 남구 거마로98번길 16'},
    ]))
    assert ar._dong_building_poi_fallback('신정동 목련아파트') is None


def test_reject_building_not_tail(monkeypatch):
    # place_name 꼬리가 건물명이 아니면 거부(생활상호 오매칭 방지)
    monkeypatch.setattr(ar, '_kakao_get_json', _mk_kakao([
        {'place_name': '메가커피강남점', 'address_name': '경기 하남시 감이동 5',
         'road_address_name': '경기 하남시 감일중앙로 60'},
    ]))
    assert ar._dong_building_poi_fallback('감이동 메가커피') is None


def test_reject_when_road_or_region_present():
    # 도로·번지·지역 토큰 있으면 진입 안 함 (다른 경로 소관)
    assert ar._dong_building_poi_fallback('하남시 감이동 벨솔레파크') is None
    assert ar._dong_building_poi_fallback('감이동 123 벨솔레파크') is None
    assert ar._dong_building_poi_fallback('감이로 60 벨솔레파크') is None


def test_reject_no_building_or_no_dong():
    assert ar._dong_building_poi_fallback('감이동') is None
    assert ar._dong_building_poi_fallback('벨솔레파크') is None


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
