# -*- coding: utf-8 -*-
"""공사 쓰기 경로 (2026-11-16 상태 머신 점검 묶음 E).

① 큐 작업은 쓰기 직전에 그 행이 아직 그 공사인지 확인 — 밀렸으면 코드로 다시 찾고, 없으면 쓰지 않음.
② 슬랙 수정·취소·재개·매출0 같은 시스템 쓰기도 _version +1, 데드레터 재시도는 그 사이 다른 저장이면 거절.
③ 다른 화면 저장이 시트 반영 대기 중이면 같은 옛 버전을 든 오래된 화면 저장은 409.
④ '공사 취소' 판정 한 곳, 메모 칸으로 취소 못 함, 시트 직접 취소도 스냅샷.
⑤ PM 새 공사는 지속 큐로 — 다시 돌아도 한 줄, 수식 행 번호는 실제로 쓰는 행.
"""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.projects as P
from dashboard.services.project_state import is_cancelled, spec_cancelled

F2L = {'프로젝트 코드': 'A', '총액 1': 'T', '수금 관련 특이사항': 'AB', '수금 확인': 'AD', '공사 확정': 'S',
       '_version': 'AT'}


class _Mgr:
    def __init__(self, cells=None, rows=None):
        self.cells = dict(cells or {})
        self.rows = dict(rows or {})          # 코드 → 행
        self.batches, self.appended, self.colors = [], [], []

    def get_cell_value(self, sid, sn, cell):
        return self.cells.get(cell, '')

    def find_row_by_project_code(self, sid, code, rng=None, use_cache=True):
        return self.rows.get(code)

    def get_field_to_letter(self):
        return dict(F2L)

    def batch_update_cells(self, sid, updates):
        self.batches.append(updates)
        return True

    def update_row_background_color(self, **kw):
        self.colors.append(kw)

    def find_next_empty_row(self, sid, *a):
        return 4200

    def append_row(self, sid, values, *a):
        self.appended.append(values)
        return {'updatedCells': len(values)}


@pytest.fixture
def mgr(monkeypatch):
    m = _Mgr()
    import dashboard.services.project_service as ps
    monkeypatch.setattr(ps, 'get_sheets_manager', lambda: m)
    monkeypatch.setattr(ps, 'update_project_in_cache', lambda code, payload: True)
    return m


# ── ① 행 확인 ──

def test_queue_row_verifies_and_relocates(mgr):
    base = {'sheet_id': 's', 'sheet_name': '공사 현황', 'row_number': 100, 'expect_code': 'G4100-AA'}
    mgr.cells['A100'] = 'G4100-AA'
    assert P._queue_row(mgr, base) == 100
    mgr.cells['A100'] = 'G4099-ZZ'                                        # 위에 행이 끼어 밀림
    mgr.rows['G4100-AA'] = 101
    assert P._queue_row(mgr, base) == 101
    del mgr.rows['G4100-AA']                                              # 지워짐 → 쓰지 않음
    with pytest.raises(RuntimeError, match='못 찾음'):
        P._queue_row(mgr, base)
    assert P._queue_row(mgr, dict(base, expect_code='')) == 100           # 옛 작업은 예전대로


def test_batch_write_retargets_and_bumps_version(mgr):
    mgr.cells.update({'A100': 'G4099-ZZ', 'AT101': '7'})
    mgr.rows['G4100-AA'] = 101
    P._handle_sheet_batch_write({
        'sheet_id': 's', 'sheet_name': '공사 현황', 'row_number': 100, 'expect_code': 'G4100-AA',
        'bump_version': True, 'tag': 'slack_edit:G4100-AA',
        'updates': [{'range': '공사 현황!L100', 'values': [['천장형 교체']]}]})
    ranges = {u['range']: u['values'][0][0] for u in mgr.batches[-1]}
    assert ranges == {'공사 현황!L101': '천장형 교체', '공사 현황!AT101': '8'}


def test_as_batch_write_without_code_unchanged(mgr):
    P._handle_sheet_batch_write({'sheet_id': 's', 'updates': [{'range': 'A/S 관리!L5', 'values': [['접수 완료']]}],
                                 'tag': 'as_update:AS-0001'})
    assert mgr.batches[-1] == [{'range': 'A/S 관리!L5', 'values': [['접수 완료']]}]


def test_cancel_handler_bumps_version(mgr, monkeypatch):
    monkeypatch.setattr(P, '_update_project_background_color', lambda *a, **k: None)
    mgr.cells.update({'A100': 'G4100-AA', 'AT100': '3'})
    P._handle_cancel_sheet({'sheet_id': 's', 'sheet_name': '공사 현황', 'row_number': 100,
                            'project_code': 'G4100-AA', 'expect_code': 'G4100-AA', 'spec': '공사 취소'})
    assert {'range': '공사 현황!AT100', 'values': [['4']]} in mgr.batches[-1]


# ── ② 데드레터 재시도 ──

def test_retry_guard(mgr):
    mgr.cells['AT100'] = '6'
    P._queue_retry_guard(mgr, {'sheet_id': 's', 'sheet_name': '공사 현황', 'version': 7, '_retried': True}, 100)
    with pytest.raises(RuntimeError, match='다른 수정'):
        P._queue_retry_guard(mgr, {'sheet_id': 's', 'sheet_name': '공사 현황', 'version': 6, '_retried': True}, 100)
    P._queue_retry_guard(mgr, {'sheet_id': 's', 'sheet_name': '공사 현황', 'version': 6}, 100)   # 첫 시도는 그대로


def test_retry_failed_marks_payload(monkeypatch):
    import json
    from dashboard.services import sheet_write_queue as q

    class _R:
        def __init__(self):
            self.failed = [json.dumps({'op_id': 'x1', 'payload': {'project_code': 'G1'}, 'attempts': 3})]
            self.queue = []

        def lrange(self, k, a, b):
            return list(self.failed)

        def lrem(self, k, n, v):
            self.failed.remove(v)

        def rpush(self, k, v):
            self.queue.append(v)
    r = _R()
    monkeypatch.setattr(q, '_rc', lambda: r)
    assert q.retry_failed('x1')
    assert json.loads(r.queue[0])['payload']['_retried'] is True


# ── ③ 대기 중인 다른 저장 ──

def test_optimistic_lock_blocks_stale_screen_while_other_save_queued(monkeypatch):
    from flask import Flask
    monkeypatch.setattr(P, '_version_col_index', lambda m: 0)
    pend = {}
    monkeypatch.setattr(P, '_pending_version_get', lambda code: pend.get(code))

    class _M:
        def get_row_values(self, **k):
            return ['5']

        def get_column_mapping(self):
            return {'AT': '_version'}
    with Flask(__name__).test_request_context():
        assert P._check_optimistic_lock_update(_M(), 's', 'n', 'G1', 100, ['5'], '5') == (6, None)
        pend['G1'] = '6'                                                    # 다른 화면 저장이 큐에 있음
        v, err = P._check_optimistic_lock_update(_M(), 's', 'n', 'G1', 100, ['5'], '5')
        assert v is None and err[1] == 409
        assert P._check_optimistic_lock_update(_M(), 's', 'n', 'G1', 100, ['5'], '6') == (7, None)  # 그 화면 자신


# ── ④ 취소 판정·메모 칸·시트 직접 취소 스냅샷 ──

@pytest.mark.parametrize('spec,yes', [('공사 취소 (사유: 변심)', True), ('공사취소', True), ('공사  취소', True),
                                      ('취소 문의', False), ('', False), (None, False)])
def test_spec_cancelled(spec, yes):
    assert spec_cancelled(spec) is yes and is_cancelled({'수금 관련 특이사항': spec}) is yes


def test_memo_cannot_cancel_project():
    idx = {'수금 관련 특이사항': 0}
    for perm in ('editor', 'admin'):
        msg = P._cancelled_edit_gate([''], idx, {'수금 관련 특이사항': '고객 공사취소 요청'}, perm)
        assert msg and msg[1] == 400 and '[공사 취소] 버튼' in msg[0]
    assert P._cancelled_edit_gate([''], idx, {'수금 관련 특이사항': '취소 문의 있음'}, 'editor') is None


def test_reconcile_sheet_cancel_saves_snapshot(monkeypatch):
    from dashboard.services import cancel_snapshot as cs
    from dashboard.services import project_reconcile as rec
    saved = {}
    monkeypatch.setattr(cs, 'load_raw', lambda code: None)
    monkeypatch.setattr(cs, 'save', lambda code, snap, by='': saved.update({code: (snap, by)}))
    prev = {'총액 1': '3400000', '공사 확정': '2026-10-01', '수금 확인': 'TRUE'}
    cur = {'총액 1': 0, '공사 확정': '2026-10-01', '수금 확인': True, '수금 관련 특이사항': '공사 취소'}
    rec._save_sheet_cancel_snapshot('G4100-AA', prev, cur)
    snap, by = saved['G4100-AA']
    assert snap == {'confirmed_date': '2026-10-01', 'payment_confirmed': True, 'total_1': '3400000'} and by == '시트정리'
    monkeypatch.setattr(cs, 'load_raw', lambda code: '{"total_1": 1}')     # 버튼 취소 스냅샷이 있으면 그대로
    saved.clear()
    rec._save_sheet_cancel_snapshot('G4100-AA', prev, cur)
    assert saved == {}


# ── ⑤ PM 새 공사 지속 큐 ──

def _create_payload():
    return {'sheet_id': 's', 'project_code': 'G4300-AA', 'data': {'프로젝트 코드': 'G4300-AA', '총액 1': '₩1,000'},
            'project_data': {'프로젝트 코드': 'G4300-AA'}, 'user_email': 'mgr@x', 'ip_address': '1.1.1.1',
            'partner_promote': [['G4300-AA'], '현장']}


def test_create_append_uses_real_row_and_finalizes_once(mgr, monkeypatch):
    done, promoted, rows_used = [], [], []
    monkeypatch.setattr(P, '_finalize_project_creation_bg', lambda code, d, pd, user_email='', ip_address=None: done.append(code))
    monkeypatch.setattr(P, '_promote_partner_after_create', lambda code, pp, ed='': promoted.append(pp))
    monkeypatch.setattr(P, '_new_row_formulas', lambda f2l, row: rows_used.append(row) or {'미수금': f'=X{row}'})
    monkeypatch.setattr(P, '_build_row_values', lambda data, m, row: [data.get('프로젝트 코드'), data.get('미수금')])
    P._handle_project_create_append(_create_payload())
    assert rows_used == [4200] and mgr.appended == [['G4300-AA', '=X4200']]
    assert done == ['G4300-AA'] and promoted == [(['G4300-AA'], '현장')]

    # 응답만 늦었던 재실행 — 행이 이미 있고 카드도 있으면 아무것도 안 함
    mgr.rows['G4300-AA'] = 4200

    class _R:
        def exists(self, k):
            return k == 'project_card_msg:G4300-AA'
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': _R()})())
    P._handle_project_create_append(_create_payload())
    assert len(mgr.appended) == 1 and done == ['G4300-AA']


def test_create_append_raises_for_queue_retry(mgr, monkeypatch):
    monkeypatch.setattr(P, '_new_row_formulas', lambda f2l, row: {})
    monkeypatch.setattr(P, '_build_row_values', lambda data, m, row: ['x'])
    monkeypatch.setattr(mgr, 'append_row', lambda *a: None)
    with pytest.raises(RuntimeError):
        P._handle_project_create_append(_create_payload())


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
