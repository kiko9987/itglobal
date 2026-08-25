# -*- coding: utf-8 -*-
"""건물명 뒤 중복어 '건물' 제거 (L-03778).

'프라임하우스 건물 1층' → '프라임하우스 1층'. 강한 건물 접미(빌딩/하우스/타워 등) 뒤
단독 '건물'만. '건물주'·'건물관리'(뒤 한글)·접미 없는 상호(관리동 건물)는 보존. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _post_normalize_display as P


@pytest.mark.parametrize('addr, expected', [
    ('강서구 강서로16길 43 프라임하우스 건물 1층 체리쁘띠푸 카페베이커리',
     '강서구 강서로16길 43 프라임하우스 1층 체리쁘띠푸 카페베이커리'),
    ('강남구 X로 1 삼성빌딩 건물 3층', '강남구 X로 1 삼성빌딩 3층'),
    ('강남구 X로 1 롯데타워 건물', '강남구 X로 1 롯데타워'),
])
def test_redundant_building_removed(addr, expected):
    assert P(addr) == expected


@pytest.mark.parametrize('addr', [
    '강남구 X로 1 건물주 김씨',        # 건물주 — 뒤 한글, 보존
    '강남구 X로 1 관리동 건물 2층',    # 강한 접미 아님(동), 보존
])
def test_preserved(addr):
    assert P(addr) == addr


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
