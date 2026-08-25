# -*- coding: utf-8 -*-
"""법정동 단독 괄호 '(원창동)' 제거 — 홈페이지 다음 우편번호 위젯 표기 통일 (L-03779).

'도로명 번지 (법정동) 층' → '도로명 번지 층'. 번지(숫자) 직후 괄호만, 건물 구역동
(관리동·상가동)·아파트 동(101동)·라틴(A동)은 보존. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _post_normalize_display as P


@pytest.mark.parametrize('addr, expected', [
    ('인천 서해구 원창로64번길 16-8 (원창동) 2층',
     '인천 서해구 원창로64번길 16-8 2층'),
    ('서울 강남구 테헤란로 152 (역삼동)', '서울 강남구 테헤란로 152'),
    ('인천 서구 경명대로 100 (원당동) 3층', '인천 서구 경명대로 100 3층'),
    ('서울 중구 세종대로 110 (을지로3가) 5층', '서울 중구 세종대로 110 5층'),
])
def test_bare_dong_paren_removed(addr, expected):
    assert P(addr) == expected


@pytest.mark.parametrize('addr', [
    '시흥 공단1대로 38 (관리동) 2층',        # 건물 구역동 — 보존
    '안산 어쩌고로 12 (상가동)',              # 건물 구역동 — 보존
    '서울 강남구 X로 1 롯데캐슬 (101동) 302호',  # 아파트 동(숫자) — 보존
    '서울 강남구 X로 1 무슨빌딩 (A동)',       # 라틴 동 — 보존
])
def test_preserved(addr):
    assert P(addr) == addr


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
