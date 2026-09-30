# -*- coding: utf-8 -*-
"""PM 편집 저장은 바뀐 셀만 쓴다 — 전체 행 덮어쓰기 레이스 회귀 테스트 (2026-09-30 G4139-MJ).

배경: PUT 이 PUT 시점 스냅샷(current_values)으로 A:AS 전체 행을 write-behind 큐에 넣어,
큐 처리 전 수 초 사이 다른 경로(수금 인입 확인의 U/V/W 값 기록 등)가 쓴 셀을 옛 값으로
되돌렸다. SB가 계산서만 'N입금'으로 고친 저장이 방금 기록된 잔금 650,000 을 0 으로 덮어
#수금_관리 카드가 유실됨. → 바뀐 필드 + _version 셀만 batch 로 쓴다.
"""
import sys
sys.path.insert(0, '.')

from dashboard.blueprints import projects as P


def test_col_letter():
    assert P._col_letter(0) == 'A'
    assert P._col_letter(22) == 'W'
    assert P._col_letter(25) == 'Z'
    assert P._col_letter(26) == 'AA'
    assert P._col_letter(27) == 'AB'
    assert P._col_letter(P.VERSION_COL_INDEX) == 'AS'


def _row():
    vals = [''] * (P.VERSION_COL_INDEX + 1)
    vals[0] = 'G4139-MJ'
    vals[22] = ''            # W 잔금 — 스냅샷 시점엔 비어 있었음
    vals[24] = '잔금 - N입금'  # Y 계산서 (변경 후)
    vals[27] = 'N입금'        # AB 잔금 계산서 (변경 후)
    vals[P.VERSION_COL_INDEX] = '1'
    return vals


F2I = {'프로젝트 코드': 0, '잔금': 22, '계산서': 24, '잔금 계산서': 27, '_version': 44}


def test_changed_cells_excludes_untouched_payment_cell():
    changes = [
        {'field_name': '잔금 계산서', 'old_value': '', 'new_value': 'N입금'},
        {'field_name': '계산서', 'old_value': '', 'new_value': '잔금 - N입금'},
    ]
    cells = P._changed_cells(changes, F2I, _row())
    assert cells == [
        {'col': 'Y', 'value': '잔금 - N입금'},
        {'col': 'AB', 'value': 'N입금'},
        {'col': 'AS', 'value': '1'},
    ]
    assert all(c['col'] != 'W' for c in cells)   # 잔금 값 셀은 절대 안 씀


def test_changed_cells_includes_code_column_on_rekey():
    cells = P._changed_cells([], F2I, _row(), code_changed=True)
    assert [c['col'] for c in cells] == ['A', 'AS']


def test_changed_cells_short_row_pads_blank():
    cells = P._changed_cells([], F2I, ['G1'])
    assert cells == [{'col': 'AS', 'value': ''}]


class _FakeManager:
    def __init__(self):
        self.batch = None
        self.full = None

    def batch_update_cells(self, sheet_id, updates):
        self.batch = (sheet_id, updates)

    def update_row(self, sheet_id, row_number, values, range_name):
        self.full = (sheet_id, row_number, values, range_name)

    def update_cell_note(self, *a, **k):   # _process_payment_field_comments 경로용
        pass


def _patch_manager(monkeypatch):
    fm = _FakeManager()
    import dashboard.services.project_service as ps
    monkeypatch.setattr(ps, 'get_sheets_manager', lambda: fm)
    return fm


def test_handler_writes_only_given_cells(monkeypatch):
    fm = _patch_manager(monkeypatch)
    P._handle_project_update_sheet({
        'sheet_id': 'SID', 'sheet_name': '공사 현황', 'row_number': 4140,
        'cells': [{'col': 'AB', 'value': 'N입금'}, {'col': 'AS', 'value': '1'}],
        'field_changes': [{'field_name': '잔금 계산서', 'old_value': '', 'new_value': 'N입금'}],
        'project_code': 'G4139-MJ',
    })
    assert fm.full is None
    assert fm.batch == ('SID', [
        {'range': '공사 현황!AB4140', 'values': [['N입금']]},
        {'range': '공사 현황!AS4140', 'values': [['1']]},
    ])


def test_handler_legacy_payload_falls_back_to_full_row(monkeypatch):
    fm = _patch_manager(monkeypatch)
    P._handle_project_update_sheet({
        'sheet_id': 'SID', 'sheet_name': '공사 현황', 'row_number': 7,
        'range_name': '공사 현황!A7:AS7', 'current_values': ['X'],
        'field_changes': [], 'project_code': 'G0006-TH',
    })
    assert fm.batch is None
    assert fm.full == ('SID', 7, ['X'], '공사 현황!A7:AS7')
