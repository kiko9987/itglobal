# -*- coding: utf-8 -*-
"""당근 sync 초(second)-정밀도 중복 생성 회귀 테스트 (2026-09-14 설희정 무한중복).

버그: 당근 시트 '응답 일시'는 초까지(예 09:15:37) 있어 _meta_consult_dt(new_dt)에
초가 보존되나, 시트 저장값 '상담 시간'은 _format_main_dt('%H:%M')로 초를 버려 되읽으면
항상 초=0. exact-match 를 `==`(초 포함) 로 하면 09:15:37 != 09:15:00 으로 영원히 실패 →
더 최신의 같은번호 리드(거래처 재유입 등)가 생기면 1시간 윈도우도 빗나가 매 sync 신규 생성.

수정: exact-match 를 분(minute) 정밀도로 비교(저장 정밀도와 일치).
네트워크 없이 순수 함수 map_karrot_row_to_lead + _get_existing_phone_lookup 로 왕복 재현.
"""
import sys
sys.path.insert(0, '.')

import pandas as pd

from dashboard.services import lead_sync as ls
from dashboard.services.lead_sync import (
    map_karrot_row_to_lead, _get_existing_phone_lookup,
    KCOL_CONSULT, KCOL_NAME, KCOL_PHONE,
)


def _karrot_row(consult_raw, phone='010-2619-8031', name='설희정'):
    # 주소 비우면 카카오 검증 경로를 타지 않아 네트워크 불필요
    return pd.Series({KCOL_CONSULT: consult_raw, KCOL_NAME: name, KCOL_PHONE: phone})


def test_seconds_bearing_source_still_dedups_by_minute():
    """당근 응답일시에 초가 있어도 시트 저장(분 정밀도)과 분-정밀도로 매치돼 중복 판정."""
    row = _karrot_row('2026-07-11 09:15:37')
    lead = map_karrot_row_to_lead(row)
    new_dt = lead['_meta_consult_dt']
    assert new_dt.second == 37, '당근 소스는 초를 보존해야 한다(버그 재현 조건)'

    # 이 lead 가 시트에 등록됐다고 가정 → 되읽기용 main_df ('상담 시간'은 초 없는 저장형)
    main_df = pd.DataFrame([{
        '리드 No': 'L-03197',
        '고객 연락처': lead['고객 연락처'],
        '상담 시간': lead['상담 시간'],   # '2026.07.11. 09:15' (초 버려짐)
    }])
    lookup = _get_existing_phone_lookup(main_df)
    import re
    digits = re.sub(r'\D', '', lead['고객 연락처'])
    entries = lookup[digits]

    stored_dt = entries[0]['consult_dt']
    # 저장 왕복으로 초는 0 이 됨 → raw == 는 실패(=옛 버그), 분-정밀도는 성립(=수정)
    assert stored_dt == stored_dt.replace(second=0, microsecond=0)
    assert (stored_dt == new_dt) is False, '옛 초-포함 비교는 실패했어야 한다(버그 조건)'
    _new_min = new_dt.replace(second=0, microsecond=0)
    assert any(
        e['consult_dt'].replace(second=0, microsecond=0) == _new_min
        for e in entries if e['consult_dt']
    ), '분-정밀도 exact-match 로는 동일 문의로 잡혀 skip 돼야 한다'


def test_zero_second_source_unaffected():
    """초가 00 인 소스는 기존과 동일하게 정상 매치(회귀 없음)."""
    row = _karrot_row('2026-07-11 09:15:00')
    lead = map_karrot_row_to_lead(row)
    new_dt = lead['_meta_consult_dt']
    main_df = pd.DataFrame([{
        '리드 No': 'L-03197',
        '고객 연락처': lead['고객 연락처'],
        '상담 시간': lead['상담 시간'],
    }])
    lookup = _get_existing_phone_lookup(main_df)
    import re
    entries = lookup[re.sub(r'\D', '', lead['고객 연락처'])]
    _new_min = new_dt.replace(second=0, microsecond=0)
    assert any(
        e['consult_dt'].replace(second=0, microsecond=0) == _new_min
        for e in entries if e['consult_dt']
    )
