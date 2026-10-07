# -*- coding: utf-8 -*-
"""영문 층 표기 'NF'/'BNF' → 'N층'/'지하N층' 정규화 (L-04193).

'도산대로 428 10F'의 '10F'를 층으로 못 알아봐 '10'만 번지 뒤 숫자로 잡히고 검증에서
떨어져 층 정보가 유실되던 갭. 순수함수(_normalize_road_spacing).
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.lead_helpers import _normalize_road_spacing as N


@pytest.mark.parametrize('raw, expected', [
    ('강남구 도산대로 428 10F', '강남구 도산대로 428 10층'),
    ('강남구 도산대로 428 10f', '강남구 도산대로 428 10층'),
    ('강남구 도산대로 428 B1F', '강남구 도산대로 428 지하1층'),
    ('다산지금로 202 B동 5F 0001호', '다산지금로 202 B동 5층 0001호'),
    ('강남구 도산대로 428 10F 인투익스', '강남구 도산대로 428 10층 인투익스'),
])
def test_f_floor_normalized(raw, expected):
    assert N(raw) == expected


@pytest.mark.parametrize('raw', [
    '강남구 테헤란로 152 3층',          # 이미 한글 층
    '강남구 테헤란로 152 A10F호',      # 앞이 영문(유닛 코드) — 미대상
    '강남구 테헤란로 152 5FLOOR',      # 뒤가 영문 — 미대상
])
def test_untouched(raw):
    assert N(raw) == raw


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
