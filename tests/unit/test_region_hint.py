# -*- coding: utf-8 -*-
"""_extract_region_hint 시/도 축약형 fallback 회귀 테스트 (2026-07-31 L-03254).

배경: 정규화 주소 첫 토큰은 시 접미 없는 축약형(양주·시흥·화성) → `시|군|구|도$`
매칭 실패 → region 빈값 → _enrich_with_poi 의 POI 쿼리에 지역 없어 전국 동명 상호가
나와 도로 불일치 → 지점명(더마트 송추점 등) 복원 실패. 구/동 우선, 없으면(시+면/리)
첫 토큰으로 fallback.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services import address_resolver as ar


class TestExtractRegionHint:
    def test_gu_preferred(self):
        assert ar._extract_region_hint('인천 부평구 일신로 25') == '부평구'
        assert ar._extract_region_hint('서울 구로구 경인로65길 44') == '구로구'

    def test_si_suffix_word(self):
        # '남양주시' 처럼 시 접미 붙은 토큰은 그대로
        assert ar._extract_region_hint('남양주시 와부읍 율석리 723') == '남양주시'

    def test_rural_si_short_fallback(self):
        # 시+면/리 rural, 구/동 없음 → 첫 토큰(축약형 시) fallback (기존엔 '')
        assert ar._extract_region_hint('양주 장흥면 가마골로 3 더마트') == '양주'
        assert ar._extract_region_hint('화성 팔탄면 서근내길 81-10') == '화성'

    def test_metro_first_token(self):
        # 광역시 축약형도 구가 없으면 첫 토큰
        assert ar._extract_region_hint('시흥 은계중앙로 247 럭스나인') == '시흥'

    def test_empty(self):
        assert ar._extract_region_hint('') == ''


class TestBuildCandidatesRegionPrefix:
    """도로+번지 후보에 지역 접두 부착 — 도시 모호 도로 오매칭 방지 (L-03659)."""

    def test_region_prefixed_road_first(self):
        # '경인로 789'는 서울 영등포·인천 부평 양쪽 존재 → 지역 포함 후보가 먼저
        cands = ar._build_candidates('인천 부평구 경인로 789 서광빌딩 1층 카센타', None)
        assert cands[0] == '인천 부평구 경인로 789'
        assert '경인로 789' in cands  # 지역 없는 fallback 도 존재

    def test_gu_only_region(self):
        cands = ar._build_candidates('영등포구 경인로 789 카센타', None)
        # 시/도 없이 구만 있으면 지역접두 미부착(정규식 시/도 시작 요구) → 도로만
        assert '경인로 789' in cands

    def test_no_region_no_prefix(self):
        cands = ar._build_candidates('테헤란로 152 삼성빌딩', None)
        assert '테헤란로 152' in cands


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
