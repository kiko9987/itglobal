# -*- coding: utf-8 -*-
"""시/군/구가 도로명·하위 행정구역에 붙은 글루 분리 후보 생성 (L-03839).

'수원시매영로' → '수원시 매영로'. 연속 글루(수원시영통구매영로)는 비탐욕으로 각 경계
분리. '청로'·'민로'(1자+로)는 lookahead 미매치라 시청로·시민로 보존. 이 함수는 **후보
생성만** 담당하고, 실제 채택은 resolve_address 가 재귀 verify 로 게이팅한다(순수함수).
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _unglue_si_before_road as U


@pytest.mark.parametrize('raw, expected', [
    ('수원시매영로345번길95', '수원시 매영로345번길95'),      # 시+도로명
    ('수원시영통구매영로11', '수원시 영통구 매영로11'),        # 시+구+도로명 (연속)
    ('의정부시신흥로50', '의정부시 신흥로50'),                # 4자 시 어간
    ('안양시만안구예술로9', '안양시 만안구 예술로9'),
])
def test_unglue(raw, expected):
    assert U(raw) == expected


@pytest.mark.parametrize('raw', [
    '수원시청로',    # 시청로 — '청로'(1자+로) lookahead {2,} 미매치, 보존
    '수원시민로',    # 시민로 — 동일
    '서울시청',      # 도로 아님
])
def test_protected(raw):
    assert U(raw) == raw


def test_none_safe():
    assert U(None) is None
    assert U('') == ''


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
