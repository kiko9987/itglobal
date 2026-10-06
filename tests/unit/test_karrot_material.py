# -*- coding: utf-8 -*-
"""당근 리드폼 광고 소재 표시 (2026-10-06).

당근 자동연동 시트의 '소재 ID' 칸 → 문의 내용 앞 [유입:당근/리드폼/<ID>] 마커(시트 영속, 소재별 집계용)
→ 카드엔 '광고 소재' 줄(문구 이름)로 표시하고 제목은 '온라인 (당근)' 그대로(중복 '당근 · 당근' 금지).
주소를 비워 카카오 검증 경로(네트워크)를 타지 않게 함.
"""
import sys
sys.path.insert(0, '.')

import pandas as pd

from dashboard.services.lead_sync import (
    map_karrot_row_to_lead, build_inquiry_blocks,
    KCOL_CONSULT, KCOL_NAME, KCOL_PHONE, KCOL_INQUIRY, KCOL_MATERIAL,
)
from dashboard.services.lead_helpers import karrot_material_label


def _row(material='1787591723962696000', inquiry='10평 상가 천장형 1대 견적'):
    return pd.Series({KCOL_CONSULT: '2026-10-06 10:55:19', KCOL_NAME: '테스트', KCOL_PHONE: '010-1234-5678',
                      KCOL_INQUIRY: inquiry, KCOL_MATERIAL: material})


def _text(lead):
    blocks, _ = build_inquiry_blocks(lead, 'L-00001', source='당근')
    return '\n'.join(b.get('text', {}).get('text', '') for b in blocks if b.get('type') == 'section')


def test_marker_saved_in_sheet_content():
    lead = map_karrot_row_to_lead(_row())
    assert lead['문의 내용'] == '[유입:당근/리드폼/1787591723962696000]\n10평 상가 천장형 1대 견적'


def test_card_shows_material_line_and_plain_title():
    t = _text(map_karrot_row_to_lead(_row()))
    assert '>*광고 소재* : 수동1 무료방문견적 (읍면동)\n' in t
    assert '온라인 (당근)*' in t and '당근 · 당근' not in t
    assert '[유입:' not in t
    assert '10평 상가 천장형 1대 견적' in t


def test_no_material_id_unchanged():
    lead = map_karrot_row_to_lead(_row(material=''))
    assert lead['문의 내용'] == '10평 상가 천장형 1대 견적'
    assert '광고 소재' not in _text(lead)


def test_material_only_without_inquiry():
    lead = map_karrot_row_to_lead(_row(inquiry=''))
    assert lead['문의 내용'] == '[유입:당근/리드폼/1787591723962696000]'
    assert '>*광고 소재* :' in _text(lead)


def test_unknown_material_shows_id_tail():
    assert karrot_material_label('1799999999999123000') == 'ID …123000'
    assert karrot_material_label('') == ''
