# -*- coding: utf-8 -*-
"""순수 지번 주소: 행안부(juso) 권위 우선 · POI 퍼지 폴백 오매칭 차단 회귀 (L-03875).

배경 (2026-09-01 L-03875 정유진): '서울 서초구 반포동 18-3' 처럼 카카오 미인덱싱
순수 지번을, POI 폴백이 같은 구(서초구)의 엉뚱한 장소(내곡동 '헌릉과 인릉' 왕릉)로
정확매치해 **완전히 다른 주소를 confident verified** 로 반환 → 매니저 오방문 직결.

근본 원인 2겹:
  (a) '서울'이 '시'로 안 끝나 _ADMIN_SUFFIX_RE 를 통과 → 순수 지번의 유일한 상호
      후보로 뽑혀 POI query('서울 서초구')가 'cand + 공백'으로 시작하는 장소를 매치.
  (b) POI 퍼지 폴백(1c-poi)이 행안부 juso(1b-juso, 권위 소스)보다 먼저 실행돼
      juso 정답('반포대로 287')이 도달조차 못 함.

수정: (a) 시/도 이름을 상호 후보에서 제외, (b) 순수 지번은 juso 를 POI 보다 먼저 시도.
네트워크 없이 경계 함수(verify_address/_juso_fallback/_try_poi_fallback) monkeypatch.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services import address_resolver as ar


class TestSidoNotShopCandidate:
    def test_pure_jibun_with_sido_yields_no_shop_candidate(self):
        """'서울 서초구 반포동 18-3' → 상호 후보 없음 (시/도·구·동 전부 제외)."""
        assert ar._extract_shop_candidates('서울 서초구 반포동 18-3') == []

    def test_sido_short_and_suffixed_excluded(self):
        assert ar._extract_shop_candidates('서울 경기도 강원특별자치도') == []

    def test_real_shop_still_extracted(self):
        """시/도로 시작하는 실제 상호(서울대입구·세종병원)는 후보 유지 (fullmatch)."""
        out = ar._extract_shop_candidates('서초구 반포동 세종병원 서울대입구역')
        assert '세종병원' in out
        assert '서울대입구역' in out

    def test_named_shop_jibun_still_candidate(self):
        """지번 + 실제 건물명 입력은 건물명이 후보로 남아 POI 구제 유지 (L-03473)."""
        out = ar._extract_shop_candidates('인천 부평구 일신동 25 송암노인요양원')
        assert out == ['송암노인요양원']


class TestJusoPriorityOverPoi:
    def _wire(self, monkeypatch, *, juso, poi_road):
        # 카카오 주소검색(step1) 미인덱싱 → verify 실패 시뮬
        monkeypatch.setattr(ar, 'verify_address', lambda t, r=None: None)
        # 하위 keyword/road POI 폴백은 이 테스트 범위 밖 → 무력화
        monkeypatch.setattr(ar, '_jibun_road_fallback', lambda t: None)
        monkeypatch.setattr(ar, '_road_poi_fallback', lambda t: None)
        # 행안부 juso 결과
        monkeypatch.setattr(ar, '_juso_fallback', lambda t, r=None: juso)
        # POI 퍼지 폴백 결과(엉뚱한 장소)
        monkeypatch.setattr(ar, '_try_poi_fallback', lambda t: poi_road)
        # enrichment 는 실제 값 유지(부작용 최소) — 그대로 둠

    def test_juso_wins_over_wrong_poi(self, monkeypatch):
        """juso 정답이 있으면 POI 가 엉뚱한 장소를 줘도 juso 채택 (오방문 차단)."""
        self._wire(monkeypatch,
                   juso=('서초구 반포대로 287', 'jibun'),
                   poi_road='서초구 헌인릉길 36-10 서울')  # 왕릉 (오매칭)
        addr, level = ar.resolve_address('서울 서초구 반포동 18-3')
        assert level == 'verified'
        assert addr.startswith('서초구 반포대로 287')
        assert '헌인릉' not in addr and '헌릉' not in addr

    def test_poi_used_when_juso_misses(self, monkeypatch):
        """juso 미매치면 POI 폴백은 기존대로 동작 (fallback 유지)."""
        self._wire(monkeypatch,
                   juso=None,
                   poi_road='인천 부평구 일신로 25 송암노인요양원')
        addr, level = ar.resolve_address('인천 부평구 일신동 25 송암노인요양원')
        assert level == 'verified'
        assert '일신로 25' in addr


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
