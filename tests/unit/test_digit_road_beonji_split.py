# -*- coding: utf-8 -*-
"""도로명 내부 숫자 + 번지 붙여쓰기 분할 (L-03828).

'향기2로21' → '향기2로 21'. 기존 BEONJI_SPLIT 도로명 패턴([가-힣]{2,}대?로)이 도로명
내부 숫자(향기'2'로)를 못 받아 미분할 → 파싱 실패([확인 필요]). 도로명 패턴에 \\d* 허용.
대로·법정동(세종로1가)·도로연속(공단1대로195번길)은 보존. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.lead_helpers import _normalize_road_spacing as N


@pytest.mark.parametrize('raw, expected', [
    ('향기2로21', '향기2로 21'),
    ('향기2로21건물', '향기2로 21 건물'),
    ('동탄반석로172', '동탄반석로 172'),
    ('동탄반석대로172', '동탄반석대로 172'),   # 대로 여전히 동작
    ('테헤란로152', '테헤란로 152'),
])
def test_split(raw, expected):
    assert N(raw) == expected


@pytest.mark.parametrize('raw', [
    '세종로1가',            # 법정동 번호 — 보존
    '공단1대로195번길',     # 도로연속(번길) — 보존
])
def test_preserved(raw):
    assert N(raw) == raw


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
