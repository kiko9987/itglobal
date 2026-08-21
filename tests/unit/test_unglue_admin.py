# -*- coding: utf-8 -*-
"""붙여쓴 행정구역 접두 분리 (L-03741).

'경기도고양시덕양구도내동 964' 처럼 시/도+시+구+동을 붙여 입력하면 ADDRESS_PATTERNS
(공백 요구)를 못 타 [추정]으로 빠졌음. 시/도+시?+구/군?+동/읍/면/리를 분리. 어간 1~4자로
'서구'·'우동'·'구로구' 커버. 이미 띄어쓴 것·도로명·단일 동은 미변경. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.lead_helpers import _unglue_admin_prefix as U


@pytest.mark.parametrize('glued, expected', [
    ('경기도고양시덕양구도내동 964', '경기도 고양시 덕양구 도내동 964'),
    ('서울특별시구로구고척동 123', '서울특별시 구로구 고척동 123'),   # 구로구(내부 구) 보존
    ('부산해운대구우동 1', '부산 해운대구 우동 1'),                    # 우동(1자 어간)
    ('인천서구가정동 산1', '인천 서구 가정동 산1'),                    # 서구(1자 어간)
    ('수원시팔달구인계동 111', '수원시 팔달구 인계동 111'),
    ('경기광주시오포동 5', '경기 광주시 오포동 5'),
    ('인천중구운서동 2840', '인천 중구 운서동 2840'),
])
def test_glued_admin_split(glued, expected):
    assert U(glued) == expected


@pytest.mark.parametrize('unchanged', [
    '경기도 고양시 덕양구 도내동 964',   # 이미 띄어쓴 것
    '고양 덕양구 도내동 964',
    '중동 큰길',                         # 단일 동
    '강남구청 앞',                       # 동 없음(구+청)
    '서울대로 123',                      # 도로명
    '서울숲it밸리',                      # 상호
    '서울특별시 강남구 역삼동 823 서울숲빌딩',
    '역삼동825번지 스타',
])
def test_not_touched(unchanged):
    assert U(unchanged) == unchanged


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
