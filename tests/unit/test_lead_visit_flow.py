# -*- coding: utf-8 -*-
"""리드 방문 흐름 정리 (2026-11-16 상태 머신 점검 묶음 F 2단계).

① 방문 취소·되돌리기가 '공사 확정'을 덮지 않음. 시트에 못 쓰면 List·카드도 그대로 두고 멈춤.
② 같은 리드를 다시 예약하면 이전 '취소됨' 표시도 지움(재취소가 조용히 무시되지 않게).
③ '방문 예약'에서 다른 상태로 바뀌면 방문 카드를 닫고 List 행을 지움.
④ 전화 병합: 방문 예약 리드에 통화 기록(유선 상담 등)이 합쳐져도 방문 예약은 그대로, '문의 드랍'만 바꿈.
⑤ 채널톡 등록 잠금은 등록할 게 없으면 바로 풂.
"""
import json
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb


class _RC:
    def __init__(self, d=None):
        self.d = dict(d or {})

    def get(self, k):
        return self.d.get(k)

    def set(self, k, v, nx=False, ex=None):
        if nx and k in self.d:
            return False
        self.d[k] = v
        return True

    def delete(self, *ks):
        n = 0
        for k in ks:
            n += int(self.d.pop(k, None) is not None)
        return n


@pytest.fixture
def rc(monkeypatch):
    import dashboard.utils.redis_client as rcm
    r = _RC()
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': r})())
    return r


class _Client:
    def __init__(self, card_blocks=None):
        self.ephemeral, self.updates = [], []
        self.card_blocks = card_blocks or []

    def chat_postEphemeral(self, **k):
        self.ephemeral.append(k['text'])

    def chat_update(self, **k):
        self.updates.append(k)

    def chat_postMessage(self, **k):
        pass

    def conversations_replies(self, **k):
        return {'messages': [{'text': '카드', 'blocks': self.card_blocks}]}


def _cancel_view():
    return {'private_metadata': json.dumps({'lead_no': 'L-04300', 'channel': 'C_VISIT', 'message_ts': '1.2'}),
            'state': {'values': {'reason': {'reason': {'value': '고객 일정 변경'}}}}}


@pytest.fixture
def cancel(monkeypatch, rc):
    wrote, hooks = [], []
    monkeypatch.setattr(sb, '_try_acquire_action_lock', lambda *a: True)
    monkeypatch.setattr(sb, '_slack_user_to_initial', lambda *a: 'JW')
    monkeypatch.setattr(sb, '_update_lead_dispatch', lambda no, upd: wrote.append(upd))
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', lambda *a, **k: hooks.append(a[0]))

    def _run(status='방문 예약'):
        monkeypatch.setattr(sb, '_find_lead_by_no', lambda *a, **k: {'상태': status, '상담 내용': '첫 상담'})
        c = _Client()
        try:
            sb._process_visit_cancel_confirmed(c, {'user': {'id': 'U_JW'}}, _cancel_view())
        except Exception:
            pass                                                   # 카드·DM 뒷단계는 범위 밖
        return c
    return type('F', (), {'run': staticmethod(_run), 'wrote': wrote, 'hooks': hooks})


# ── ① 방문 취소·되돌리기 ──

def test_cancel_keeps_construction_confirmed(cancel):
    cancel.run('공사 확정')
    assert '상태' not in cancel.wrote[0] and '[방문 취소' in cancel.wrote[0]['상담 내용']
    assert cancel.hooks == ['SLACK_VISIT_CANCEL_WEBHOOK_URL']           # 방문 일정은 취소


def test_cancel_normal_sets_cancelled(cancel):
    cancel.run('방문 예약')
    assert cancel.wrote[0]['상태'] == '방문 취소'


def test_cancel_stops_when_sheet_write_fails(cancel, monkeypatch, rc):
    def _boom(no, upd):
        raise RuntimeError('시트 502')
    monkeypatch.setattr(sb, '_update_lead_dispatch', _boom)
    c = cancel.run('방문 예약')
    assert cancel.hooks == [] and c.updates == []                       # List·카드 그대로
    assert 'visit_cancelled:L-04300' not in rc.d and '저장하지 못해' in c.ephemeral[0]


def _uncancel_body():
    return {'actions': [{'value': 'L-04300'}], 'channel': {'id': 'C_VISIT'}, 'message': {'ts': '1.2'},
            'user': {'id': 'U_JW'}}


@pytest.mark.parametrize('status,written', [('방문 취소', [{'상태': '방문 예약'}]), ('공사 확정', [])])
def test_uncancel_restores_unless_confirmed(monkeypatch, rc, status, written):
    wrote, hooks = [], []
    rc.d['visit_cancelled:L-04300'] = '1'
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda *a, **k: {'상태': status, '플랫폼': '당근'})
    monkeypatch.setattr(sb, '_update_lead_dispatch', lambda no, upd: wrote.append(upd))
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', lambda *a, **k: hooks.append(a[0]))
    monkeypatch.setattr(sb, '_slack_user_to_initial', lambda *a: 'JW')
    monkeypatch.setattr(sb, '_build_visit_notice_blocks', lambda **k: ('카드', []))
    c = _Client()
    sb._process_visit_uncancel(c, _uncancel_body())
    assert wrote == written and hooks == ['SLACK_VISIT_RESTORE_WEBHOOK_URL'] and c.updates
    assert 'visit_cancelled:L-04300' not in rc.d


def test_uncancel_stops_when_lead_unreadable(monkeypatch, rc):
    def _boom(*a, **k):
        raise sb.LeadReadError('502')
    monkeypatch.setattr(sb, '_find_lead_by_no', _boom)
    monkeypatch.setattr(sb, '_update_lead_dispatch', lambda *a: (_ for _ in ()).throw(AssertionError('쓰면 안 됨')))
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', lambda *a, **k: (_ for _ in ()).throw(AssertionError('웹훅 안 됨')))
    c = _Client()
    sb._process_visit_uncancel(c, _uncancel_body())
    assert '읽지 못해' in c.ephemeral[0] and c.updates == []


# ── ② 재방문 때 취소 표시 해제 ──

def test_clear_visit_completion_also_clears_cancelled(rc):
    rc.d.update({'visit_cancelled:L-1': '1', 'visit_completed:L-1': '1'})
    sb._clear_visit_completion('L-1')
    assert rc.d == {}


# ── ③ 방문 예약에서 벗어나면 방문 카드·List 정리 ──

_ACTIVE = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': '>*방문일* : 11/20\n>이름 : 홍길동'}},
           {'type': 'actions', 'elements': [{'action_id': 'visit_edit_info'}, {'action_id': 'visit_complete'},
                                            {'action_id': 'visit_cancel'}]}]


def test_close_visit_card_when_leaving_reservation(monkeypatch, rc):
    hooks = []
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', lambda *a, **k: hooks.append(a[:4]))
    monkeypatch.setattr(sb, '_visit_slack_app', None)
    rc.d['visit_notice_msg:L-04300'] = 'C_VISIT|9.9'
    c = _Client(_ACTIVE)
    assert sb._close_visit_card_for_status_change(c, 'L-04300', '문의 드랍', 'JW') is True
    assert hooks == [('SLACK_VISIT_CANCEL_WEBHOOK_URL', 'L-04300', 'C_VISIT', '9.9')]
    up = c.updates[-1]
    assert up['ts'] == '9.9' and "방문 일정 종료 — 상담 결과 '문의 드랍'" in up['text'] and '홍길동' in up['text']
    assert all(b['type'] != 'actions' for b in up['blocks'])              # 버튼 없음


def test_close_visit_card_skips_done_or_missing(monkeypatch, rc):
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', lambda *a, **k: (_ for _ in ()).throw(AssertionError()))
    monkeypatch.setattr(sb, '_visit_slack_app', None)
    assert sb._close_visit_card_for_status_change(_Client(), 'L-1', '유선 상담') is False   # 카드 위치 없음
    rc.d['visit_notice_msg:L-1'] = 'C|1.0'
    done = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': '방문 완료'}}]
    c = _Client(done)
    assert sb._close_visit_card_for_status_change(c, 'L-1', '유선 상담') is False and c.updates == []


# ── ④ 전화 병합 ──

class _Svc:
    def __init__(self):
        self.bodies = []

    def spreadsheets(self):
        return self

    def values(self):
        return self

    def batchUpdate(self, spreadsheetId, body):
        self.bodies.append(body)
        return self

    def execute(self):
        return {}


@pytest.mark.parametrize('wf_status,expect', [('유선 상담', None), ('부재중', None), ('문의 드랍', '문의 드랍'),
                                              ('방문 예약', '방문 예약')])
def test_phone_merge_keeps_reservation(wf_status, expect):
    from dashboard.services import lead_sync as ls
    svc = _Svc()
    mgr = type('M', (), {'service': svc})()
    ls._update_existing_lead_from_workflow(
        mgr, {'sheet_name': '고객 리드 관리', 'sheet_id': 's'}, 50,
        {'상태': wf_status, '상담 내용': '방문 전 문의 통화'},
        existing_row={'상태': '방문 예약', '상담 내용': '첫 상담'})
    cells = {d['range'].split('!')[1]: d['values'][0][0] for d in svc.bodies[0]['data']}
    assert cells.get('D50') == expect
    assert cells['K50'].startswith('첫 상담\n─────────\n') and '방문 전 문의 통화' in cells['K50']


def test_phone_merge_other_status_unchanged():
    from dashboard.services import lead_sync as ls
    svc = _Svc()
    mgr = type('M', (), {'service': svc})()
    ls._update_existing_lead_from_workflow(mgr, {'sheet_name': 'x', 'sheet_id': 's'}, 7, {'상태': '유선 상담'},
                                           existing_row={'상태': '부재중'})
    assert {d['range'].split('!')[1]: d['values'][0][0] for d in svc.bodies[0]['data']}['D7'] == '유선 상담'


# ── ⑤ 채널톡 등록 잠금 ──

def test_channeltalk_register_releases_lock_when_nothing_pending(rc):
    from dashboard.blueprints import channeltalk as ct
    assert ct._register_pending_lead_if_any('chat1') == ''
    assert 'channeltalk_lead_lock:chat1' not in rc.d                      # 예전엔 60초 잡고 있었음


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
