# -*- coding: utf-8 -*-
"""juso 도로 확인 시 지역이 다르면 juso base 채택 (L-04091).

고객이 지역을 틀리거나 빠뜨리면('남구 삼성로 155'←강남구) step-2 는 기존엔 문자열을 그대로
두고 level 만 올렸음(L-03671). 이제 juso base(정부 DB)의 지역이 정규식과 다르면 juso base
로 재구성(지역·공식 건물명) + 상호 부착 후 _enrich_with_poi 재실행(공식 상호명). 지역이
같으면 기존대로 유지. 네트워크는 monkeypatch 로 차단(hermetic).
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.fixture
def _stub(monkeypatch):
    monkeypatch.setattr(ar, 'verify_address', lambda *a, **k: None)
    for fn in ['_try_poi_fallback', '_road_poi_fallback', '_dong_building_poi_fallback',
               '_jibun_road_fallback', '_partner_alias_lookup']:
        monkeypatch.setattr(ar, fn, lambda *a, **k: None)
    # _enrich_with_poi: 상호가 있으면 공식명으로 치환하는 것처럼 흉내
    monkeypatch.setattr(ar, '_enrich_with_poi',
                        lambda addr, text: addr.replace('웰라스 피부과', '웰라스피부과의원'))


def test_region_differs_adopts_juso_base(_stub, monkeypatch):
    monkeypatch.setattr(ar, '_juso_fallback',
                        lambda *a, **k: ('강남구 삼성로 155 대치퍼스트빌딩', 'road'))
    raw = '남구 삼성로 155 대치퍼스트 상가 2층 웰라스 피부과'
    rx = '남구 삼성로 155 대치퍼스트 상가 2층 웰라스'
    addr, lv = ar.resolve_address(raw, rx, 'level3')
    assert lv == 'verified'
    assert addr == '강남구 삼성로 155 대치퍼스트빌딩 2층 웰라스피부과의원'


def test_region_same_keeps_regex_string(_stub, monkeypatch):
    # 지역 동일 → juso base 채택 안 함(문자열 유지 + level 만 승격, L-03671)
    monkeypatch.setattr(ar, '_juso_fallback',
                        lambda *a, **k: ('강남구 삼성로 155', 'road'))
    raw = '강남구 삼성로 155 3층 무슨상호'
    rx = '강남구 삼성로 155 3층 무슨상호'
    addr, lv = ar.resolve_address(raw, rx, 'level3')
    assert lv == 'verified'
    assert '무슨상호' in addr and addr.startswith('강남구 삼성로 155')


def test_no_juso_no_change(_stub, monkeypatch):
    # juso 미확인 → 기존 level 유지
    monkeypatch.setattr(ar, '_juso_fallback', lambda *a, **k: None)
    addr, lv = ar.resolve_address('강남구 삼성로 155 3층 카페', '강남구 삼성로 155 3층 카페', 'level3')
    assert lv == 'level3'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
