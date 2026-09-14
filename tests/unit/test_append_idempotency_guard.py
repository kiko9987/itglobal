# -*- coding: utf-8 -*-
"""등록 직전 멱등 가드 회귀 테스트 (2026-09-14 재발 방지 안전망).

_append_leads_to_main_locked 은 dedup 이 뚫려도 '방금 읽은' 시트에 같은
연락처+상담시각(분)이 이미 있으면 생성하지 않고 ''(빈문자)로 정렬 반환한다.
네트워크/Redis 는 monkeypatch 로 제거하고 가드 판정만 검증.
"""
import sys
sys.path.insert(0, '.')

from datetime import datetime
from unittest.mock import MagicMock

import pandas as pd

from dashboard.services import lead_sync as ls
from dashboard.services.lead_sync import (
    _lead_phone_minute_key, _build_phone_minute_index,
)


# ── 순수 헬퍼 ────────────────────────────────────────────────
def test_phone_minute_key_uses_minute_not_seconds():
    lead = {'고객 연락처': '010-1111-2222',
            '_meta_consult_dt': datetime(2026, 7, 11, 9, 15, 37)}
    assert _lead_phone_minute_key(lead) == ('01011112222', '202607110915')


def test_phone_minute_key_falls_back_to_string_field():
    lead = {'고객 연락처': '010-1111-2222', '상담 시간': '2026.07.11. 09:15'}
    assert _lead_phone_minute_key(lead) == ('01011112222', '202607110915')


def test_phone_minute_key_none_when_no_time_or_phone():
    assert _lead_phone_minute_key({'고객 연락처': '010-1111-2222'}) is None
    assert _lead_phone_minute_key({'_meta_consult_dt': datetime(2026, 7, 11, 9, 15)}) is None


def test_build_index_minute_precision():
    df = pd.DataFrame([
        {'리드 No': 'L-1', '고객 연락처': '010-1111-2222', '상담 시간': '2026.07.11. 09:15'},
        {'리드 No': 'L-2', '고객 연락처': '', '상담 시간': '2026.07.11. 09:16'},  # 전화없음 제외
    ])
    idx = _build_phone_minute_index(df)
    assert ('01011112222', '202607110915') in idx
    assert len(idx) == 1


# ── 전체 함수 (가드) ─────────────────────────────────────────
def _patch_append(monkeypatch, existing_df):
    """네트워크/Redis 제거 — append 는 성공 응답, 발번은 순차."""
    monkeypatch.setattr(ls, 'load_leads_data', lambda force_refresh=False: existing_df)

    seq = {'n': 4000}
    def _alloc(sheet_max, count):
        start = max(sheet_max, seq['n']) + 1
        seq['n'] = start + count - 1
        return list(range(start, start + count))
    monkeypatch.setattr(ls, '_allocate_lead_numbers', _alloc)

    fake_mgr = MagicMock()
    (fake_mgr.service.spreadsheets.return_value.values.return_value
     .append.return_value.execute.return_value) = {
        'updates': {'updatedRange': "'고객 리드 관리'!A2:P9",
                    'updatedRows': 99, 'updatedCells': 99}}
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: fake_mgr)
    monkeypatch.setattr(ls, 'invalidate_leads_cache', lambda: None)
    monkeypatch.setattr(ls, '_reset_row_background', lambda *a, **k: None)
    return fake_mgr


CFG = {'sheet_id': 'sid', 'sheet_name': '고객 리드 관리'}


def test_guard_blocks_existing_phone_minute_despite_seconds(monkeypatch):
    df = pd.DataFrame([
        {'리드 No': 'L-03197', '고객 연락처': '010-1111-2222',
         '상담 시간': '2026.07.11. 09:15'},  # 저장형(초 없음)
    ])
    _patch_append(monkeypatch, df)
    leads = [
        {'고객 연락처': '010-1111-2222', '고객명': 'A', '플랫폼': '당근',
         '_meta_consult_dt': datetime(2026, 7, 11, 9, 15, 37)},   # 초 있음 → 분매치로 차단
        {'고객 연락처': '010-9999-8888', '고객명': 'B', '플랫폼': '당근',
         '_meta_consult_dt': datetime(2026, 7, 11, 10, 0, 0)},    # 신규 허용
    ]
    out = ls._append_leads_to_main_locked(leads, CFG)
    assert out[0] == '' and out[1].startswith('L-')
    assert len(out) == 2


def test_guard_blocks_intra_batch_duplicate(monkeypatch):
    _patch_append(monkeypatch, pd.DataFrame(columns=['리드 No', '고객 연락처', '상담 시간']))
    leads = [
        {'고객 연락처': '010-2619-8031', '고객명': '설희정', '플랫폼': '당근',
         '_meta_consult_dt': datetime(2026, 7, 11, 9, 15, 5)},
        {'고객 연락처': '010-2619-8031', '고객명': '설희정', '플랫폼': '당근',
         '_meta_consult_dt': datetime(2026, 7, 11, 9, 15, 50)},  # 같은 분 → 차단
    ]
    out = ls._append_leads_to_main_locked(leads, CFG)
    assert out[0].startswith('L-') and out[1] == ''


def test_guard_allows_different_minute(monkeypatch):
    df = pd.DataFrame([
        {'리드 No': 'L-03197', '고객 연락처': '010-1111-2222',
         '상담 시간': '2026.07.11. 09:15'},
    ])
    _patch_append(monkeypatch, df)
    leads = [{'고객 연락처': '010-1111-2222', '고객명': 'A', '플랫폼': '당근',
              '_meta_consult_dt': datetime(2026, 7, 11, 9, 16, 0)}]  # 재문의(다른 분)
    out = ls._append_leads_to_main_locked(leads, CFG)
    assert out[0].startswith('L-')


def test_guard_all_blocked_returns_empty_no_append(monkeypatch):
    df = pd.DataFrame([
        {'리드 No': 'L-03197', '고객 연락처': '010-1111-2222',
         '상담 시간': '2026.07.11. 09:15'},
    ])
    mgr = _patch_append(monkeypatch, df)
    leads = [{'고객 연락처': '010-1111-2222', '고객명': 'A', '플랫폼': '당근',
              '_meta_consult_dt': datetime(2026, 7, 11, 9, 15, 30)}]
    out = ls._append_leads_to_main_locked(leads, CFG)
    assert out == ['']
    # 전량 차단이면 시트 append 자체를 호출하지 않아야 함
    mgr.service.spreadsheets.return_value.values.return_value.append.assert_not_called()
