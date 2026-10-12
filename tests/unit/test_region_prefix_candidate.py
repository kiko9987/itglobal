# -*- coding: utf-8 -*-
"""시/도 없이 시·구로 시작하는 주소도 지역+도로+번지 후보를 먼저 조회 (L-04219).

'수원 장안구 조원로 97-1'이 지역 없는 '조원로 97-1'로 먼저 조회돼 카카오 첫 결과
'서울 관악구 조원로 97-1'로 오매칭되던 버그. 기존 지역 접두(L-03659)는 시/도로 시작할 때만.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.mark.parametrize('text, first', [
    ('수원 장안구 조원로 97-1 1층 조원한의원', '수원 장안구 조원로 97-1'),
    ('성남시 분당구 정자일로 95', '성남시 분당구 정자일로 95'),
    ('강남구 테헤란로 152 3층', '강남구 테헤란로 152'),
    ('화성 효행구 봉담읍 상리2길 55', '화성 효행구 봉담읍 상리2길 55'),
    ('인천 부평구 경인로 789', '인천 부평구 경인로 789'),       # 기존 시/도 경로 유지
])
def test_region_road_candidate_first(text, first):
    assert ar._build_candidates(text, None)[0] == first


def test_no_region_unchanged():
    # 지역 없이 도로부터 시작하면 기존대로 도로+번지가 첫 후보
    assert ar._build_candidates('조원로 97-1 조원한의원', None)[0] == '조원로 97-1'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
