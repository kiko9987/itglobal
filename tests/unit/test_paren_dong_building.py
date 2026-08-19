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
    # 동N가·로N가 법정동 (L-03707)
    ('성동구 아차산로 103 (성수동2가 영동테크노타워) 10층 1004호',
     '성동구 아차산로 103 영동테크노타워 10층 1004호'),
    ('중구 세종대로 1 (을지로3가 삼성빌딩) 5층', '중구 세종대로 1 삼성빌딩 5층'),
    # 중첩 괄호 (법인 (주)/㈜) — 내부 괄호 보존 unwrap (L-03716)
    ('시흥 공단1대로195번길 38 (정왕동 (주)삼인)', '시흥 공단1대로195번길 38 (주)삼인'),
    ('부천 부천로 1 (심곡동 ㈜대성) 2층', '부천 부천로 1 ㈜대성 2층'),
])
def test_no_comma_dong_building_unwrapped(addr, expected):
    assert P(addr) == expected


@pytest.mark.parametrize('addr', [
    '김포 사우중로 1 (걸포동 172-1)',   # 번지형 — 건물부 숫자 시작, 미매치(보존)
    '김포 사우중로 1 (걸포동)',          # 동 단독 — 공백+건물 없음, 미매치(보존)
    '서초구 방배로 20 (서초대로 삼성빌딩) 3층',  # 도로명(숫자+가 없음) — 법정동 오인 방지, 보존
])
def test_preserved(addr):
    assert P(addr) == addr


def test_comma_form_paren_kept():
    # 콤마형 '(법정동, 건물)' 은 unwrap 안 함 — 콤마만 정리되고 괄호 유지
    assert P('강남구 테헤란로 152 (역삼동, 강남빌딩) 3층') \
        == '강남구 테헤란로 152 (역삼동 강남빌딩) 3층'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
