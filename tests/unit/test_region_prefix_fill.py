# -*- coding: utf-8 -*-
"""juso 지역 접두 전체 비교 — 시 누락 보충(건물 유지) vs 지역 오기 재구성 (L-04114).

'권선구 오목천로152번길 40 첨단벤처밸리'(수원 누락)는 juso '수원 권선구 …' 와 비교해 앞
시만 붙이고 건물 유지. '남구 …'(강남구 오기)는 지역 토큰 자체가 달라 juso base 로 재구성.
_region_prefix 는 순수함수. resolve 는 네트워크 monkeypatch.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


def test_region_prefix_pure():
    assert ar._region_prefix('권선구 오목천로152번길 40 첨단벤처밸리') == '권선구'
    assert ar._region_prefix('수원 권선구 오목천로152번길 40') == '수원 권선구'
    assert ar._region_prefix('남구 삼성로 155 대치퍼스트') == '남구'
    assert ar._region_prefix('오목천로152번길 40') == ''   # 지역 없음


@pytest.fixture
def _stub(monkeypatch):
    monkeypatch.setattr(ar, 'verify_address', lambda *a, **k: None)
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda addr, text: addr)
    for fn in ['_try_poi_fallback', '_road_poi_fallback', '_dong_building_poi_fallback',
               '_jibun_road_fallback', '_partner_alias_lookup']:
        monkeypatch.setattr(ar, fn, lambda *a, **k: None)


def test_missing_city_prepended_building_kept(_stub, monkeypatch):
    # 시 누락 — 앞 시만 붙이고 건물(첨단벤처밸리) 유지
    monkeypatch.setattr(ar, '_juso_fallback',
                        lambda *a, **k: ('수원 권선구 오목천로152번길 40', 'road'))
    raw = '권선구 오목천로152번길 40 첨단벤처밸리 101호'
    rx = '권선구 오목천로152번길 40 첨단벤처밸리'
    addr, lv = ar.resolve_address(raw, rx, 'level3')
    assert lv == 'verified'
    assert addr.startswith('수원 권선구 오목천로152번길 40')
    assert '첨단벤처밸리' in addr   # 건물 보존


def test_wrong_region_rebuilds(_stub, monkeypatch):
    # 지역 토큰 자체가 다름 — juso base 로 재구성(건물명 juso 것 채택)
    monkeypatch.setattr(ar, '_juso_fallback',
                        lambda *a, **k: ('강남구 삼성로 155 대치퍼스트빌딩', 'road'))
    raw = '남구 삼성로 155 대치퍼스트 2층'
    rx = '남구 삼성로 155 대치퍼스트'
    addr, lv = ar.resolve_address(raw, rx, 'level3')
    assert lv == 'verified'
    assert addr.startswith('강남구 삼성로 155 대치퍼스트빌딩')


def test_same_region_unchanged(_stub, monkeypatch):
    monkeypatch.setattr(ar, '_juso_fallback',
                        lambda *a, **k: ('수원 권선구 오목천로152번길 40', 'road'))
    raw = '수원 권선구 오목천로152번길 40 무슨빌딩'
    rx = '수원 권선구 오목천로152번길 40 무슨빌딩'
    addr, lv = ar.resolve_address(raw, rx, 'level3')
    assert lv == 'verified'
    assert addr.startswith('수원 권선구') and addr.count('수원') == 1


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
