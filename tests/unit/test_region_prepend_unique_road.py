# -*- coding: utf-8 -*-
"""지역 없는 입력 + 전국 유일 도로 → 지역 접두 부착 (L-03988).

'풍초로 115'(하남 유일) → '하남 풍초로 115'. 다도시 도로('중앙로 100')는 지역 미부착
(오확정 방지). 네트워크(juso/kakao)는 monkeypatch 로 차단(hermetic).
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.fixture
def _stub(monkeypatch):
    # raw fallback 까지 도달하도록 앞단계 모두 None, juso 는 도로 base 반환
    monkeypatch.setattr(ar, 'verify_address', lambda t, r=None: None)
    for fn in ['_try_poi_fallback', '_jibun_road_fallback', '_road_poi_fallback',
               '_dong_building_poi_fallback']:
        monkeypatch.setattr(ar, fn, lambda *a, **k: None)
    monkeypatch.setattr(ar, '_juso_fallback',
                        lambda t, r: ('하남 풍초로 115', 'road'))


def test_unique_road_gets_region(_stub, monkeypatch):
    monkeypatch.setattr(ar, '_road_region_count', lambda c: 1)   # 유일
    assert ar.resolve_address('풍초로 115', None, '') == ('하남 풍초로 115', 'verified')


def test_ambiguous_road_no_region(_stub, monkeypatch):
    monkeypatch.setattr(ar, '_road_region_count', lambda c: 3)   # 다도시
    # 지역 미부착 — 문자열 그대로(오확정 방지)
    addr, lvl = ar.resolve_address('풍초로 115', None, '')
    assert addr == '풍초로 115' and lvl == 'verified'


def test_region_already_present_not_doubled(_stub, monkeypatch):
    monkeypatch.setattr(ar, '_road_region_count', lambda c: 1)
    addr, lvl = ar.resolve_address('하남 풍초로 115', None, '')
    assert addr.count('하남') == 1


def test_admin_present_no_duplication(monkeypatch):
    # L-04030: 입력에 시/면 admin 이 있어도 base 로 교체(중복 방지). '화성 정남면' +
    #   base '화성 효행구 정남면' 이 '화성 효행구 정남면 화성 정남면' 로 겹치면 안 됨.
    monkeypatch.setattr(ar, 'verify_address', lambda t, r=None: None)
    for fn in ['_try_poi_fallback', '_jibun_road_fallback', '_road_poi_fallback',
               '_dong_building_poi_fallback']:
        monkeypatch.setattr(ar, fn, lambda *a, **k: None)
    monkeypatch.setattr(ar, '_juso_fallback',
                        lambda t, r: ('화성 효행구 정남면 신리길 43-25', 'road'))
    monkeypatch.setattr(ar, '_road_region_count', lambda c: 1)
    addr, lvl = ar.resolve_address('화성 정남면 신리길 43-25', None, '')
    assert addr == '화성 효행구 정남면 신리길 43-25'
    assert addr.count('정남면') == 1 and addr.count('화성') == 1


def test_admin_present_building_tail_kept(monkeypatch):
    monkeypatch.setattr(ar, 'verify_address', lambda t, r=None: None)
    for fn in ['_try_poi_fallback', '_jibun_road_fallback', '_road_poi_fallback',
               '_dong_building_poi_fallback']:
        monkeypatch.setattr(ar, fn, lambda *a, **k: None)
    monkeypatch.setattr(ar, '_juso_fallback',
                        lambda t, r: ('화성 효행구 정남면 신리길 43-25', 'road'))
    monkeypatch.setattr(ar, '_road_region_count', lambda c: 1)
    addr, lvl = ar.resolve_address('화성 정남면 신리길 43-25 무슨빌딩 3층', None, '')
    assert addr == '화성 효행구 정남면 신리길 43-25 무슨빌딩 3층'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
