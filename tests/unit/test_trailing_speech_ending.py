# -*- coding: utf-8 -*-
"""끝 말끝(문장 종결어) 제거 — '…폭스라이브 입니다' → '…폭스라이브' (L-04024).

고객이 주소 끝에 붙인 '입니다/이에요/예요' 등 종결어 제거. 공백형·붙임형 모두. 종결어가
끝이 아니면(예요기획) 보존. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _post_normalize_display as P


@pytest.mark.parametrize('addr, expected', [
    ('중랑구 망우동 용마산로 439 지하 폭스라이브 입니다',
     '중랑구 망우동 용마산로 439 지하 폭스라이브'),
    ('중랑구 용마산로 439 폭스라이브입니다', '중랑구 용마산로 439 폭스라이브'),  # 붙임형
    ('시흥 서울대학로278번길 8 레드동 A201~205호입니다',
     '시흥 서울대학로278번길 8 레드동 A201~205호'),
    ('강남구 테헤란로 152 스타벅스 예요', '강남구 테헤란로 152 스타벅스'),
])
def test_trailing_ending_stripped(addr, expected):
    assert P(addr) == expected


@pytest.mark.parametrize('addr', [
    'X구 Y로 1 예요기획',           # 예요 가 끝 아님(상호 일부) — 보존
    '강남구 테헤란로 152 3층',       # 종결어 없음
])
def test_preserved(addr):
    assert P(addr) == addr


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
