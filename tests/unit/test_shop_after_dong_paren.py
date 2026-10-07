# -*- coding: utf-8 -*-
"""번지와 상호 사이 법정동 괄호가 있어도 상호 보존 (L-04191).

홈페이지 주소검색 위젯은 '서부샛길 168 (독산동) 삼부주류판매'처럼 번지와 상호 사이에
'(법정동)'을 끼워 '번지 + 상호' 부착 규칙(1-b)이 깨지고 상호가 유실됐음. 법정동 단독
괄호만 걷어내고 매칭. 건물 구역동('(관리동)')은 미대상. _enrich_with_poi 는 monkeypatch.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.fixture(autouse=True)
def _no_poi(monkeypatch):
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda addr, text: addr)


@pytest.mark.parametrize('raw, expected', [
    ('서울 금천구 서부샛길 168 (독산동) 삼부주류판매', '금천구 서부샛길 168 삼부주류판매'),
    ('서울 금천구 서부샛길 168(독산동) 삼부주류판매', '금천구 서부샛길 168 삼부주류판매'),
    ('서울 금천구 서부샛길 168 삼부주류판매', '금천구 서부샛길 168 삼부주류판매'),  # 회귀 없음
])
def test_shop_kept_across_dong_paren(raw, expected):
    assert ar._enrich_verified_address('금천구 서부샛길 168', raw, None) == expected


def test_dong_only_no_shop():
    # 상호 없이 법정동 괄호만 — 아무것도 붙지 않음
    out = ar._enrich_verified_address('금천구 서부샛길 168', '서울 금천구 서부샛길 168 (독산동)', None)
    assert out == '금천구 서부샛길 168'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
