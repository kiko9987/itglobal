# -*- coding: utf-8 -*-
"""괄호 '(법정동 건물명)' 무-콤마 정리 — 법정동 제거·건물명 unwrap (L-03702).

매니저가 '(갈현동 벧엘)' 처럼 법정동+건물을 괄호로 적으면 법정동 빼고 건물명만 밖으로.
카카오/관행 정식 괄호는 '(법정동, 건물)' 콤마형이라 미매치(괄호 유지). 번지형은 건물부가
숫자로 시작해 미매치. 순수함수 _post_normalize_display.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _post_normalize_display as P


@pytest.mark.parametrize('addr, expected', [
    # 무-콤마 (법정동 건물) → unwrap
    ('은평구 통일로 875 (갈현동 벧엘) 1층 중앙약국', '은평구 통일로 875 벧엘 1층 중앙약국'),
    ('수원 팔달로 1 (매탄동 삼성전자) 2층', '수원 팔달로 1 삼성전자 2층'),
    ('서초구 강남대로 1 (서초동 우성빌딩)', '서초구 강남대로 1 우성빌딩'),
])
def test_no_comma_dong_building_unwrapped(addr, expected):
    assert P(addr) == expected


@pytest.mark.parametrize('addr', [
    '김포 사우중로 1 (걸포동 172-1)',   # 번지형 — 건물부 숫자 시작, 미매치(보존)
    '김포 사우중로 1 (걸포동)',          # 동 단독 — 공백+건물 없음, 미매치(보존)
])
def test_preserved(addr):
    assert P(addr) == addr


def test_comma_form_paren_kept():
    # 콤마형 '(법정동, 건물)' 은 unwrap 안 함 — 콤마만 정리되고 괄호 유지
    assert P('강남구 테헤란로 152 (역삼동, 강남빌딩) 3층') \
        == '강남구 테헤란로 152 (역삼동 강남빌딩) 3층'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
