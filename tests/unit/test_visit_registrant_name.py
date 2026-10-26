# -*- coding: utf-8 -*-
"""방문 카드 등록자 = 영업 담당자 > 온라인 상담자, '-'(미배정)는 건너뜀
(2026-10-26 ETC-a1a726: 기타 '아무나 방문 가능'은 영업 담당자가 '-' → [정보 수정] 뒤 등록자가 '-'로 바뀜)."""
import sys

sys.path.insert(0, '.')

import pandas as pd

from dashboard.services.lead_helpers import registrant_name


def test_dash_sales_falls_back_to_counselor():
    assert registrant_name({'영업 담당자': '-', '온라인 상담자': '고광일'}) == '고광일'
    assert registrant_name({'영업 담당자': '@박용구', '온라인 상담자': '고광일'}) == '박용구'
    assert registrant_name({'영업 담당자': '', '온라인 상담자': '@고광일'}) == '고광일'
    assert registrant_name({'영업 담당자': '-', '온라인 상담자': '-'}) == ''
    assert registrant_name({}) == '' and registrant_name(None) == ''


def test_pandas_row():
    row = pd.Series({'영업 담당자': '-', '온라인 상담자': '@고광일'})
    assert registrant_name(row) == '고광일'
    assert registrant_name(pd.Series({'영업 담당자': float('nan'), '온라인 상담자': '박정우'})) == '박정우'
