# -*- coding: utf-8 -*-
"""'김포구래동'을 '김포구 래동'이 아니라 '김포 구래동'으로 나눔 + 동+건물명 POI 구제 보강 (L-04376).

김포시엔 구가 없는데 붙여쓴 행정구역 분리가 '김포구'(구) + '래동'으로 잘랐음. 또 '동+건물명'
POI 구제가 시 약칭('김포')을 건물명에 섞고, 표기 흔들림(카림에비뉴/카림애비뉴)을 못 맞춤.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar
from dashboard.services.lead_helpers import _normalize_road_spacing as N


@pytest.mark.parametrize('raw, expected', [
    ('김포구래동 카림에비뉴', '김포 구래동 카림에비뉴'),
    ('경기도고양시덕양구도내동', '경기도 고양시 덕양구 도내동'),   # 기존 동작
    ('부산해운대구우동', '부산 해운대구 우동'),
    ('서구가좌동', '서구 가좌동'),
])
def test_unglue_city_short(raw, expected):
    assert N(raw) == expected


def test_ae_fold():
    assert ar._ae_fold('카림애비뉴') == ar._ae_fold('카림에비뉴')


def test_dong_building_poi_with_city_short(monkeypatch):
    monkeypatch.setattr(ar, '_dong_region_count', lambda d: 1)
    monkeypatch.setattr(ar, '_kakao_get_json', lambda url: {'documents': [
        {'place_name': '카림애비뉴김포', 'address_name': '경기 김포시 구래동 6880-1',
         'road_address_name': '경기 김포시 김포한강4로 487'}]})
    assert ar._dong_building_poi_fallback('김포 구래동 카림에비뉴') == '김포 김포한강4로 487 카림애비뉴김포'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
