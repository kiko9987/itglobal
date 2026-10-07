# -*- coding: utf-8 -*-
"""재문의 칸 '이전 이름 (같은 연락처)' (2026-10-08 L-04199 김장하는날 ↔ L-00520 이창덕).

재문의 판정은 같은 연락처인데, 같은 사람이 이름↔가게명을 바꿔 남기면 다른 사람처럼 보임.
"""
import sys
sys.path.insert(0, '.')

import pandas as pd

from dashboard.services.lead_sync import _build_repeat_section, _get_existing_phone_lookup


def test_phone_lookup_carries_customer_name():
    df = pd.DataFrame([{'리드 No': 'L-00520', '고객 연락처': '010-9145-3180', '고객명': '이창덕',
                        '상담 시간': '2026.01.22 13:57', '상태': '상담 대기', '플랫폼': '당근'}])
    lk = _get_existing_phone_lookup(df)
    assert lk['01091453180'][0]['customer_name'] == '이창덕'


def test_repeat_section_shows_previous_name_with_basis():
    lead = {'_meta_previous_leads': [{
        'lead_no': 'L-00520', 'consult_time': '2026.01.22 13:57', 'customer_name': '이창덕',
        'inquiry': '현 가게 천장형', 'status': '상담 대기', 'address': '-',
        'feedback': '', 'consultant': '', 'sales_rep': ''}]}
    out = _build_repeat_section(lead)
    assert '>*이전 문의* : 2026.01.22 13:57 / L-00520\n>*이전 이름* : 이창덕 (같은 연락처)\n' in out


def test_missing_name_shows_dash():
    out = _build_repeat_section({'_meta_previous_leads': [{'lead_no': 'L-1', 'customer_name': ''}]})
    assert '>*이전 이름* : - (같은 연락처)' in out
    assert _build_repeat_section({}) == ''
