# -*- coding: utf-8 -*-
"""시/도 있으나 지번↔도로명 불일치 → 구 기준 POI 구제 회귀 테스트 (2026-07-31 L-03473).

배경: '인천 부평구 일신동 25 송암노인요양원' — 매니저가 지번(일신동)으로 입력했으나
실제 도로명은 일신로 25. 카카오 주소 API 는 지번을 0건 반환 → verify 실패 → raw 배지.
건물명 POI(송암노인요양원)는 '인천 부평구 일신로 25' 를 돌려주므로, 지번↔도로명 공통
안정 단위인 '구'(부평구)로 검증해 구제. _kakao_search_poi 를 monkeypatch (네트워크 없음).
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services import address_resolver as ar


def _patch_poi(monkeypatch, results):
    monkeypatch.setattr(ar, '_kakao_search_poi', lambda q: tuple(results))


class TestPoiFallbackByGu:
    def test_sido_present_gu_match_upgrades(self, monkeypatch):
        """시/도+구 + 건물명 POI 정확매치 + POI 도로명에 구 포함 → 도로명 채택."""
        _patch_poi(monkeypatch, [('송암노인요양원', '인천 부평구 일신로 25')])
        out = ar._try_poi_fallback('인천 부평구 일신동 25 송암노인요양원')
        assert out == '인천 부평구 일신로 25 송암노인요양원'

    def test_no_gu_skips(self, monkeypatch):
        """시/도는 있으나 구가 없으면 힌트 약함 → skip (오탐 방지)."""
        _patch_poi(monkeypatch, [('송암노인요양원', '인천 부평구 일신로 25')])
        # 구 없는 입력 (시/군만) — POI 가 매치돼도 채택 안 함
        assert ar._try_poi_fallback('인천 일신동 25 송암노인요양원') is None

    def test_place_name_mismatch_skips(self, monkeypatch):
        """POI place_name 이 상호 후보와 정확 매치 안 되면 skip."""
        _patch_poi(monkeypatch, [('전혀다른건물', '인천 부평구 일신로 25')])
        assert ar._try_poi_fallback('인천 부평구 일신동 25 송암노인요양원') is None

    def test_gu_not_in_poi_road_skips(self, monkeypatch):
        """POI 도로명에 원본 구가 없으면(다른 지역) skip."""
        _patch_poi(monkeypatch, [('송암노인요양원', '서울 강남구 테헤란로 1')])
        assert ar._try_poi_fallback('인천 부평구 일신동 25 송암노인요양원') is None

    def test_empty_poi_skips(self, monkeypatch):
        """POI 무결과 → None."""
        _patch_poi(monkeypatch, [])
        assert ar._try_poi_fallback('인천 부평구 일신동 25 송암노인요양원') is None


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
