# -*- coding: utf-8 -*-
"""층 범위 '8~9층' 앞 숫자를 번지로 grab 방지 (L-04115).

'국제금융로2길 36 8~9층'의 '8'(8~9층의 시작)을 번지로 잡아 '36 8'로 만든 뒤 enrich 가 원문
'8~9층'을 재부착해 '36 8 8~9층' 중복이 되던 갭. ADDRESS_PATTERNS 번지 반복 그룹 +
_EXTEND_TOKEN_RE 의 bare 숫자에 '~' negative lookahead. 순수/파이프라인 hermetic.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.lead_helpers import extract_korean_address as E
import dashboard.services.address_resolver as ar


def test_extract_does_not_grab_floor_start():
    # extract 가 '8~9층'의 '8'을 번지로 안 잡음 → '36'까지만
    r = E('영등포구 국제금융로2길 36 8~9층 (주)인젠트')
    assert r is not None
    assert r[0] == '영등포구 국제금융로2길 36'
    assert '36 8' not in r[0]


@pytest.mark.parametrize('raw, expected', [
    ('강남구 테헤란로 152 3층', '강남구 테헤란로 152 3층'),           # 단일 층 회귀 없음
    ('서초구 반포대로 58 101호', '서초구 반포대로 58 101호'),        # 호 회귀 없음
])
def test_no_regression(raw, expected):
    assert E(raw)[0] == expected


def test_pipeline_no_dup_floor(monkeypatch):
    # 전체 파이프라인: '8' 중복 없이 '…36 8~9층 (주)인젠트'
    monkeypatch.setattr(ar, 'verify_address', lambda *a, **k: None)
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda addr, text: addr)
    for fn in ['_try_poi_fallback', '_road_poi_fallback', '_dong_building_poi_fallback',
               '_jibun_road_fallback', '_juso_fallback', '_partner_alias_lookup']:
        monkeypatch.setattr(ar, fn, lambda *a, **k: None)
    raw = '영등포구 국제금융로2길 36 8~9층 (주)인젠트'
    rx = E(raw)
    addr, lv = ar.resolve_address(raw, rx[0], rx[1])
    # (주)인젠트 는 실제 _enrich_with_poi(카카오)가 붙임 — mock 이라 층까지만.
    #   핵심: '8' 중복 없음 + 층 한 번.
    assert addr.startswith('영등포구 국제금융로2길 36 8~9층')
    assert '36 8 8' not in addr and addr.count('8~9층') == 1


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
