# -*- coding: utf-8 -*-
"""정규식 경로도 _post_normalize_display(dedup) 적용 (L-04113).

정규식 경로(step-2)만 _post_normalize_display 후처리를 빠뜨려, extract 가 건물명 꼬리('차')를
떨군 뒤 enrich 가 원문 tail 을 재부착하면 '대륭테크노타운12 대륭테크노타운12차' 같은 인접
유사 토큰 중복이 남던 갭. 네트워크는 monkeypatch 로 차단(hermetic).
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.fixture
def _stub(monkeypatch):
    monkeypatch.setattr(ar, 'verify_address', lambda *a, **k: None)
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda addr, text: addr)  # 네트워크 차단
    for fn in ['_try_poi_fallback', '_road_poi_fallback', '_dong_building_poi_fallback',
               '_jibun_road_fallback', '_juso_fallback', '_partner_alias_lookup']:
        monkeypatch.setattr(ar, fn, lambda *a, **k: None)


def test_adjacent_near_dup_deduped(_stub):
    # 정규식 경로가 _post_normalize_display 를 타 '대륭테크노타운12 대륭테크노타운12차'
    #   인접 유사 토큰 중복을 dedup (뒤 '차' 포함형 유지).
    dup = '금천구 가산디지털2로 14 대륭테크노타운12 대륭테크노타운12차 1216호'
    addr, lv = ar.resolve_address(dup, dup, 'level3')
    assert addr == '금천구 가산디지털2로 14 대륭테크노타운12차 1216호'
    assert addr.count('대륭테크노타운') == 1


def test_clean_regex_unchanged(_stub):
    # 중복 없는 정규식 결과는 그대로(정규화만)
    raw = '강남구 테헤란로 152 3층'
    addr, lv = ar.resolve_address(raw, '강남구 테헤란로 152 3층', 'level3')
    assert addr == '강남구 테헤란로 152 3층'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
