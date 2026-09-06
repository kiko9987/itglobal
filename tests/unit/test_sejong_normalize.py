# -*- coding: utf-8 -*-
"""세종특별자치시 → 세종 축약 (L-03929).

'세종특별자치시'를 일반 '시' 축약(끝 글자만 제거)하면 '세종특별자치'로 깨짐. 세종은 구가
없고 특별자치시 접미라 '세종 [도로/동]'으로 정리. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import normalize_display as N


@pytest.mark.parametrize('addr, expected', [
    ('세종특별자치시 한누리대로 288 세종갤러리밸류시티 1층 110호',
     '세종 한누리대로 288 세종갤러리밸류시티 1층 110호'),
    ('세종특별자치시 나성동 764', '세종 나성동 764'),
    ('세종특별자치시 조치원읍 5', '세종 조치원읍 5'),
    ('세종시 한누리대로 288', '세종 한누리대로 288'),
    ('세종 한누리대로 288', '세종 한누리대로 288'),
])
def test_sejong_shortened(addr, expected):
    assert N(addr) == expected


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
