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

    def test_road_present_skips_shop_override(self, monkeypatch):
        """입력에 도로명+번지 이미 있으면 상호 POI 로 다른 지점 도로 갈아치우기 금지
        (L-03679: '분당내곡로 131 … 포케올데이 판교점' → 서현역점 황새울로 오매칭 방지).

        같은 브랜드 다른 지점(서현역점)을 POI 가 먼저 줘도, 도로+번지가 있으면 채택 안 함
        → 도로 기반 fallback(_road_poi_fallback·Juso)이 입력 도로를 검증하도록 위임."""
        _patch_poi(monkeypatch, [('포케올데이 서현역점', '경기 성남시 분당구 황새울로360번길 28')])
        out = ar._try_poi_fallback(
            '경기 성남시 분당구 분당내곡로 131 판교 테크원타워 지하 1층 25호 포케올데이 판교점')
        assert out is None

    def test_jibun_only_still_rescued(self, monkeypatch):
        """도로명 없는 순수 지번 입력은 기존대로 상호 POI 구제 유지 (회귀 방어)."""
        _patch_poi(monkeypatch, [('송암노인요양원', '인천 부평구 일신로 25')])
        out = ar._try_poi_fallback('인천 부평구 일신동 25 송암노인요양원')
        assert out == '인천 부평구 일신로 25 송암노인요양원'

    def test_road_present_no_sido_skips_shop(self, monkeypatch):
        """시/도 생략된 '강남구 논현로159길 10 신사빌딩' → A분기로 가서 다른 신사빌딩
        (언주로 817)로 도로 갈아치우던 버그 방지 (L-03644). 도로+번지 있으면 skip."""
        _patch_poi(monkeypatch, [('신사빌딩', '서울 강남구 언주로 817')])
        assert ar._try_poi_fallback('강남구 논현로159길 10 신사빌딩 3층') is None


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
