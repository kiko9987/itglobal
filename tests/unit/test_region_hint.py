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


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
