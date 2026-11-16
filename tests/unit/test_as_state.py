# -*- coding: utf-8 -*-
"""A/S 상태 한 곳 (2026-11-16 상태 머신 점검 묶음 I).

① 상태 바꾸기는 as_service.transition 하나 — 잠금 + '방금 바꾼 상태'까지 보고 비교(슬랙·PM 동시 처리 한 번만).
② 빈 상태 칸 = 요청됨 — 접수 없이 바로 '처리 완료'가 되지 않게.
③ 메모/이력(M열)은 큐 처리기가 그때의 시트 값에 덧붙임 — 직전 쓰기가 큐에 있어도 줄이 사라지지 않게.
④ '요청 취소' 상태 — 잘못 올린 요청을 가짜 조치 내용 없이 닫고, 같은 공사 새 요청을 막지 않게.
"""
import json
import sys

sys.path.insert(0, '.')

import pytest

from dashboard.services import as_service as svc


class _R:
    def __init__(self):
        self.d = {}

    def get(self, k):
        return self.d.get(k)

    def set(self, k, v, nx=False, ex=None):
        if nx and k in self.d:
            return False
        self.d[k] = v
        return True

    def delete(self, *ks):
        for k in ks:
            self.d.pop(k, None)


@pytest.fixture
def env(monkeypatch):
    r = _R()
    sheet = {'진행 상태': '요청됨', '조치 내용': ''}
    writes, queued = [], []
    monkeypatch.setattr(svc, '_redis', lambda: r)
    monkeypatch.setattr(svc, 'get_as_data', lambda no: dict(sheet))
    monkeypatch.setattr(svc, 'update_as_row', lambda no, upd: writes.append(upd) or True)
    import dashboard.services.sheet_write_queue as q
    monkeypatch.setattr(q, 'enqueue', lambda op, payload, meta=None: queued.append((op, payload)) or 'op')
    return type('E', (), {'r': r, 'sheet': sheet, 'writes': writes, 'queued': queued})


def test_transition_once_even_before_sheet_catches_up(env):
    ok, st = svc.transition('AS-0100', svc.STATUS_REQUESTED, svc.STATUS_ACCEPTED, {'H': 'JW'})
    assert ok and st == svc.STATUS_ACCEPTED and env.writes[-1] == {'H': 'JW', 'L': '접수 완료'}
    # 시트는 아직 '요청됨'(큐가 쓰기 전) — 두 번째 접수(슬랙·PM 동시)는 방금 바꾼 상태를 보고 거절
    ok, st = svc.transition('AS-0100', svc.STATUS_REQUESTED, svc.STATUS_ACCEPTED, {'H': 'SB'})
    assert not ok and st == svc.STATUS_ACCEPTED and len(env.writes) == 1
    assert 'as_state_lock:AS-0100' not in env.r.d                         # 잠금은 풀림


def test_empty_status_is_requested(env):
    env.sheet['진행 상태'] = ''
    ok, st = svc.transition('AS-0101', svc.STATUS_ACCEPTED, svc.STATUS_COMPLETED)
    assert not ok and st == svc.STATUS_REQUESTED                          # 예전엔 빈 칸이면 바로 완료


def test_transition_busy_and_unreadable(env, monkeypatch):
    env.r.d['as_state_lock:AS-0102'] = '1'
    monkeypatch.setattr('time.sleep', lambda s: None)
    ok, why = svc.transition('AS-0102', svc.STATUS_REQUESTED, svc.STATUS_ACCEPTED)
    assert not ok and '다른 처리' in why
    monkeypatch.setattr(svc, 'get_as_data', lambda no: None)
    ok, why = svc.transition('AS-0103', svc.STATUS_REQUESTED, svc.STATUS_ACCEPTED)
    assert not ok and '읽지 못했습니다' in why


def test_transition_without_redis_still_checks_sheet(env, monkeypatch):
    monkeypatch.setattr(svc, '_redis', lambda: None)
    assert svc.transition('AS-0104', svc.STATUS_REQUESTED, svc.STATUS_ACCEPTED)[0]


def test_cancel_from_open_states_only(env):
    assert svc.transition('AS-0105', svc.OPEN_STATUSES, svc.STATUS_CANCELLED)[0]
    env.r.d.clear()
    env.sheet['진행 상태'] = svc.STATUS_COMPLETED
    ok, st = svc.transition('AS-0106', svc.OPEN_STATUSES, svc.STATUS_CANCELLED)
    assert not ok and st == svc.STATUS_COMPLETED


def test_append_log_queues_entry_and_keeps_both_lines(env):
    env.sheet['조치 내용'] = '[11.10 09:00 JW] 접수 메모'
    v1 = svc.append_as_log('AS-0107', '현장 확인', 'SB')
    v2 = svc.append_as_log('AS-0107', '부품 교체 완료', 'SB')            # 앞 줄이 아직 큐에 있어도
    assert v2.startswith('[11.10 09:00 JW] 접수 메모\n') and '현장 확인' in v2 and v2.endswith('부품 교체 완료')
    assert [op for op, _ in env.queued] == ['as_log_append', 'as_log_append'] and v1 in v2
    assert all('entry' in p and 'as_no' in p for _, p in env.queued)


def test_log_handler_appends_to_fresh_cell_once(monkeypatch):
    class _M:
        def __init__(self):
            self.cell = '[11.10 09:00 JW] 접수 메모\n[11.10 09:05 SB] 다른 경로가 방금 쓴 줄'
            self.batches = []

        def get_cell_value(self, sid, sn, cell):
            return self.cell

        def batch_update_cells(self, sid, ups):
            self.batches.append(ups)
            self.cell = ups[0]['values'][0][0]
    m = _M()
    monkeypatch.setattr(svc, '_get_sheet', lambda: (m, 'sid', 'A/S 관리'))
    monkeypatch.setattr(svc, 'find_row_by_as_no', lambda no: 7)
    svc._handle_as_log_append({'as_no': 'AS-0108', 'entry': '[11.10 10:00 SB] 조치 완료'})
    assert m.cell.endswith('다른 경로가 방금 쓴 줄\n[11.10 10:00 SB] 조치 완료')      # 덮지 않고 덧붙임
    assert m.batches[-1][0]['range'] == 'A/S 관리!M7'
    svc._handle_as_log_append({'as_no': 'AS-0108', 'entry': '[11.10 10:00 SB] 조치 완료'})   # 재시도
    assert len(m.batches) == 1
    monkeypatch.setattr(svc, 'find_row_by_as_no', lambda no: None)
    with pytest.raises(RuntimeError):
        svc._handle_as_log_append({'as_no': 'AS-0109', 'entry': 'x'})


def test_log_handler_registered():
    from dashboard.services import sheet_write_queue as q
    assert q._registry.get('as_log_append') is svc._handle_as_log_append


# ── 슬랙 카드·요청 취소 ──

def _btn_ids(blocks):
    return [e['action_id'] for b in blocks if b['type'] == 'actions' for e in b['elements']]


def test_card_buttons_and_cancelled_header(monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    data = {'No': 'AS-0110', '프로젝트 코드': '-', '요청 내용': '소음', '조치 내용': '[11.10 10:00 JW] 요청 취소 — 잘못 등록'}
    assert _btn_ids(sb._build_as_blocks(data, view_state='requested')) == ['as_accept_open', 'as_cancel_open']
    assert _btn_ids(sb._build_as_blocks(data, view_state='accepted')) == ['as_complete_open', 'as_cancel_open']
    blocks = sb._build_as_blocks(data, view_state='cancelled')
    assert _btn_ids(blocks) == []
    assert '[A/S 요청 취소]' in blocks[0]['text']['text'] and '잘못 등록' in blocks[0]['text']['text']


def test_cancel_submission(env, monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    refreshed, told = [], []
    monkeypatch.setattr(sb, 'as_refresh_card', lambda no, **k: refreshed.append((no, k)))
    monkeypatch.setattr(sb, '_slack_user_to_initial', lambda c, u: 'JW')
    monkeypatch.setattr(sb, '_as_tell_actor', lambda c, b, ch, text: told.append(text))
    view = {'private_metadata': json.dumps({'as_no': 'AS-0111', 'channel': 'C_AS'}),
            'state': {'values': {'reason': {'value': {'value': '잘못 등록'}}}}}
    sb._process_as_cancel_submission(object(), {'user': {'id': 'U1'}}, view)
    assert env.writes[-1] == {'L': svc.STATUS_CANCELLED}
    assert refreshed and refreshed[-1][1]['override']['진행 상태'] == svc.STATUS_CANCELLED
    assert '요청 취소 — 잘못 등록' in refreshed[-1][1]['override']['조치 내용']
    sb._process_as_cancel_submission(object(), {'user': {'id': 'U1'}}, view)            # 두 번째는 거절
    assert told and '취소할 수 없는 상태' in told[-1]


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
