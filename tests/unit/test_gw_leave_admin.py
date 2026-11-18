"""그룹웨어 휴가 고치기·변경 신청·경영지원 기능(2026-11-18) — 웹 API·슬랙 알림까지.

같은 장부·같은 알림을 쓰는지(본인 고치기 → 승인자 다시 알림, 경영지원 고침 → 본인 알림, 등록 → #휴가_일정),
권한(경영지원만)을 본다. 슬랙은 가짜(test_gw_leave_apply.FakeSlack).
"""
import pytest

from dashboard.services import hr
from dashboard.services import leave as lv
from dashboard.services import leave_web
from tests.unit.test_gw_leave_apply import (ADMIN, JW, STAFF, YG, _as, _next_workday,  # noqa: F401
                                            client, slack)


@pytest.fixture
def mgr(slack, monkeypatch):
    """부장님(ADMIN)을 경영지원(휴가 관리자)으로 — 인사카드 슬랙 ID UKIKO."""
    hr.upsert_from_flex({'flex_id': 'P9', 'name': '고광일', 'status': '재직', 'work_email': ADMIN,
                         'login_email': ADMIN, 'slack_id': 'UKIKO', 'hire_date': '2025-01-01'}, [])
    monkeypatch.setattr(lv, 'manager_ids', lambda: ('UKIKO',))
    return slack


def _pid(name):
    return next(p['id'] for p in hr.list_people() if p['name'] == name)


def _body(d, end=None, t='annual', reason='개인 사정'):
    return {'type': t, 'start': d.isoformat(), 'end': (end or d).isoformat(), 'reason': reason}


def test_self_edit_pending_via_web_notifies_approver(client, slack):
    _as(client, STAFF)
    d = _next_workday(3)
    rid = client.post('/api/leave/apply', json=_body(d)).get_json()['id']
    slack.posts.clear()
    r = client.post(f'/api/leave/{rid}/edit', json=_body(d, _next_workday(4)))
    j = r.get_json()
    assert r.status_code == 200 and j['mode'] == 'edit' and j['days'] == 2
    assert any('고쳤어요' in p['text'] for p in slack.to(JW))
    me = client.get('/api/leave/me').get_json()['leave']['entries']
    e = next(x for x in me if x.get('id') == rid)
    assert e['edit'] == 'edit' and e['edit_label'] == '✏️ 고치기' and e['edit_vals']['start'] == d.isoformat()


def test_change_request_via_web_then_approve_cancels_original(client, slack):
    _as(client, STAFF)
    d = _next_workday(5)
    rid = client.post('/api/leave/apply', json=_body(d)).get_json()['id']
    from dashboard.blueprints import slack_leave
    assert slack_leave.apply_action('approve', rid, JW) is None
    assert slack_leave.apply_action('approve', rid, YG) is None          # 확정 → #휴가_일정 글
    assert lv.get(rid)['announce_ts']
    e = next(x for x in client.get('/api/leave/me').get_json()['leave']['entries'] if x.get('id') == rid)
    assert e['edit'] == 'change' and e['edit_label'] == '✏️ 변경 신청'
    j = client.post(f'/api/leave/{rid}/edit', json=_body(_next_workday(6))).get_json()
    assert j['success'] and j['mode'] == 'change'
    new = j['id']
    # 변경 신청 대기 중엔 또 못 함
    again = client.post(f'/api/leave/{rid}/edit', json=_body(_next_workday(7))).get_json()
    assert not again['success'] and '이미 변경 신청' in again['message']
    assert slack_leave.apply_action('approve', new, JW) is None
    assert slack_leave.apply_action('approve', new, YG) is None
    assert lv.get(rid)['status'] == lv.CANCELLED and lv.get(new)['status'] == lv.APPROVED
    # 원래 휴가의 #휴가_일정 글이 '변경 신청으로 바뀜'으로
    assert any('변경 신청으로 바뀜' in u['text'] for u in slack.updates if u['ch'] == 'C-ANNOUNCE')


def test_others_cannot_edit(client, slack):
    _as(client, STAFF)
    rid = client.post('/api/leave/apply', json=_body(_next_workday(3))).get_json()['id']
    from tests.unit.test_gw_leave_apply import OTHER
    _as(client, OTHER)
    r = client.post(f'/api/leave/{rid}/edit', json=_body(_next_workday(4)))
    assert r.status_code == 400 and '본인 휴가만' in r.get_json()['message']
    r = client.post(f'/api/leave/{rid}/admin-edit', json=_body(_next_workday(4)))
    assert r.status_code == 400 and '경영지원만' in r.get_json()['message']


def test_manager_edit_cancel_adjust_register(client, mgr):
    _as(client, STAFF)
    d = _next_workday(3)
    rid = client.post('/api/leave/apply', json=_body(d)).get_json()['id']
    _as(client, ADMIN)
    info = client.get(f'/api/leave/person/{_pid("박민우")}').get_json()['admin']
    assert info['can'] and rid in info['editable'] and info['rows'][str(rid)]['start'] == d.isoformat()
    assert {'value': 'spouse_birth', 'label': '배우자 출산'} in info['admin_types']
    mgr.posts.clear()
    # 고치기 — 상태 그대로, 본인에게 알림
    j = client.post(f'/api/leave/{rid}/admin-edit', json=_body(d, _next_workday(4))).get_json()
    assert j['success'] and j['days'] == 2 and lv.get(rid)['status'] == lv.WAIT_1
    assert any('경영지원(' in p['text'] and '고쳤어요' in p['text'] for p in mgr.to('UMW'))
    # 연차 조정
    before = lv.balance('박민우')
    j = client.post(f'/api/leave/person/{_pid("박민우")}/adjust', json={'days': -1, 'reason': '결근'}).get_json()
    assert j['success'] and lv.balance('박민우') == before - 1
    assert any('연차가 조정됐어요' in p['text'] for p in mgr.to('UMW'))
    bad = client.post(f'/api/leave/person/{_pid("박민우")}/adjust', json={'days': 1, 'reason': ''}).get_json()
    assert not bad['success'] and bad['message'] == '사유를 적어 주세요.'
    # 관리자 부여 휴가 등록 → 확정 + 본인 카드·알림 + #휴가_일정
    d2 = _next_workday(10)
    j = client.post(f'/api/leave/person/{_pid("박민우")}/register',
                    json={'type': 'reward', 'start': d2.isoformat(), 'end': d2.isoformat(), 'reason': '포상'}).get_json()
    assert j['success']
    row = lv.get(j['id'])
    assert row['status'] == lv.APPROVED and row['ltype'] == 'reward'
    assert any('휴가를 등록했어요' in p['text'] for p in mgr.to('UMW'))
    assert any(p['ch'] == 'C-ANNOUNCE' for p in mgr.posts)
    # 본인은 관리자 등록 휴가를 취소·변경 못 함, 경영지원은 취소 가능
    assert lv.own_cancel_mode(row, 'UMW') == '' and lv.own_edit_mode(row, 'UMW') == ''
    r = client.post(f'/api/leave/{row["id"]}/admin-cancel')
    assert r.status_code == 200 and lv.get(row['id'])['status'] == lv.CANCELLED
    assert [a['days'] for a in info['adjustments']] == []
    assert [a['days'] for a in client.get(f'/api/leave/person/{_pid("박민우")}').get_json()['admin']['adjustments']] == [-1]


def test_non_manager_admin_api_refused(client, slack):
    _as(client, STAFF)
    assert client.post(f'/api/leave/person/{_pid("박민우")}/adjust', json={'days': 1, 'reason': 'x'}).status_code == 400
    r = client.post(f'/api/leave/person/{_pid("박민우")}/register', json={'type': 'reward', 'start': '2026-12-21'})
    assert r.status_code == 400 and '경영지원만' in r.get_json()['message']
    assert client.get(f'/api/leave/person/{_pid("박민우")}').get_json()['admin'] == {'can': False}


def test_slack_edit_view_prefilled(slack):
    from dashboard.blueprints import slack_leave
    d = _next_workday(3)
    row = lv.create('박민우', 'UMW', 'UMW', '박민우', 'annual', '', d, d, '개인 사정')
    v = slack_leave.edit_view(row['id'], 'UMW')
    assert v['title']['text'] == '휴가 고치기' and v['submit']['text'] == '고치기'
    assert '"edit": %d' % row['id'] in v['private_metadata'] and '"mode": "edit"' in v['private_metadata']
    starts = [b for b in v['blocks'] if b.get('block_id') == 'lv_start']
    assert starts[0]['element']['initial_date'] == d.isoformat()
    assert slack_leave.edit_view(row['id'], 'UOT') is None             # 남의 것
