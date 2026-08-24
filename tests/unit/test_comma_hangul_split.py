# -*- coding: utf-8 -*-
"""콤마+한글(공백 없음) → 공백 (L-03752).

'15층,뉴헤어의원'→'15층 뉴헤어의원'. 기존 규칙은 '콤마+공백'만 처리해 공백 없는 콤마가
남았음. 숫자,숫자(호수 나열 305,306호)·번지 콤마는 보존. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _post_normalize_display as P


@pytest.mark.parametrize('addr, expected', [
    ('강남구 테헤란로 115 서림빌딩 15층,뉴헤어의원', '강남구 테헤란로 115 서림빌딩 15층 뉴헤어의원'),
    ('강남구 X로 1 센트럴스퀘어,관리사무소', '강남구 X로 1 센트럴스퀘어 관리사무소'),
    ('강남구 X로 1 서울숲,강남빌딩', '강남구 X로 1 서울숲 강남빌딩'),
])
def test_comma_hangul_to_space(addr, expected):
    assert P(addr) == expected


@pytest.mark.parametrize('addr', [
    '강남구 X로 1 삼성빌딩 305,306호',   # 숫자 나열 — 보존
    '강남구 X로 1 삼성빌딩 5층,301호',    # 콤마 뒤 숫자 — 규칙(콤마+한글) 미대상, 보존
])
def test_number_comma_preserved(addr):
    assert P(addr) == addr


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
