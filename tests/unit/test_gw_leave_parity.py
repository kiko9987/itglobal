"""그룹웨어 휴가 관리 플렉스 맞추기(2026-11-18) — 웹 API: 조정 적용일·취소·일괄, 엑셀, 월별 상세, 맞춤 휴가."""
import io

from dashboard.services import leave as lv
from test_gw_leave_admin import _pid, mgr  # noqa: F401
from test_gw_leave_apply import ADMIN, STAFF, _as, _next_workday, client, slack  # noqa: F401


def test_adjust_with_applies_on_and_cancel(client, mgr):
    _as(client, ADMIN)
    pid = _pid('박민우')
    before = lv.balance('박민우')
    j = client.post(f'/api/leave/person/{pid}/adjust', json={'days': -0.5, 'reason': '병원', 'applies_on': 'xx'}).get_json()
    assert not j['success'] and '적용일' in j['message']
    j = client.post(f'/api/leave/person/{pid}/adjust', json={'days': -0.5, 'reason': '병원'}).get_json()
    assert j['success'] and lv.balance('박민우') == before - 0.5
    info = client.get(f'/api/leave/person/{pid}').get_json()['admin']
    gid = info['adjustments'][0]['id']
    assert info['adjustments'][0]['cancelled'] is False
    mgr.posts.clear()
    assert client.post(f'/api/leave/adjust/{gid}/cancel').get_json()['success']
    assert lv.balance('박민우') == before
    assert any('조정이 취소됐어요' in p['text'] for p in mgr.to('UMW'))
    info = client.get(f'/api/leave/person/{pid}').get_json()['admin']
    assert info['adjustments'][0]['cancelled'] is True


def test_adjust_bulk_web(client, mgr):
    _as(client, ADMIN)
    team = client.get('/api/leave/team').get_json()
    assert team['can_manage'] is True
    pid = _pid('박민우')
    before = lv.balance('박민우')
    r = client.post('/api/leave/adjust-bulk', json={'items': [{'person_id': pid, 'days': -99}], 'reason': '정정'})
    j = r.get_json()
    assert r.status_code == 400 and str(pid) in j['errors'] and lv.balance('박민우') == before
    j = client.post('/api/leave/adjust-bulk', json={'items': [{'person_id': pid, 'days': 1}, {'person_id': 'x', 'days': 1}],
                                                    'reason': '연말 보상'}).get_json()
    assert j['success'] and j['count'] == 1 and lv.balance('박민우') == before + 1
    assert any('연차가 조정됐어요' in p['text'] for p in mgr.to('UMW'))


def test_non_manager_cannot_bulk_or_custom(client, slack):
    _as(client, STAFF)
    assert client.post('/api/leave/adjust-bulk', json={'items': []}).status_code == 400
    r = client.post(f'/api/leave/person/{_pid("박민우")}/custom-grant', json={'type': 'health', 'days': 1})
    assert r.status_code == 400 and '경영지원만' in r.get_json()['message']
    assert client.get('/api/leave/export?kind=balance').status_code == 403


def test_custom_grant_flow_web(client, mgr):
    _as(client, ADMIN)
    pid = _pid('박민우')
    mgr.posts.clear()
    j = client.post(f'/api/leave/person/{pid}/custom-grant',
                    json={'type': 'health', 'days': 1, 'valid_to': '', 'note': '보건 휴가'}).get_json()
    assert j['success']
    assert any('보건 1일이 부여됐어요' in p['text'] for p in mgr.to('UMW'))
    info = client.get(f'/api/leave/person/{pid}').get_json()
    g = info['admin']['custom_grants'][0]
    assert g['label'] == '보건' and g['left'] == 1 and g['valid_to'] == ''
    assert info['leave']['custom'][0]['ltype'] == 'health'
    # 직원 창에 보건이 보이고 신청 가능
    _as(client, STAFF)
    form = client.get('/api/leave/me').get_json()['form']
    assert {'value': 'health', 'label': '보건 (남은 1일)'} in form['types']
    d = _next_workday(4)
    r = client.post('/api/leave/apply', json={'type': 'health', 'start': d.isoformat(), 'end': d.isoformat(), 'reason': '보건'})
    assert r.get_json()['success']
    form = client.get('/api/leave/me').get_json()['form']
    assert 'health' not in [t['value'] for t in form['types']]
    # 신청이 걸려 있으면 부여 취소 안 됨
    _as(client, ADMIN)
    r = client.post(f"/api/leave/custom-grant/{g['id']}/cancel")
    assert r.status_code == 400 and '먼저 취소' in r.get_json()['message']


def test_export_three_kinds(client, mgr):
    from openpyxl import load_workbook
    _as(client, ADMIN)
    d = _next_workday(3)
    lv.admin_register('박민우', 'annual', '', d, d, '개인 =SUM(1)', 'UKIKO', '고광일')
    for kind, first in (('balance', '이름'), ('usage', '이름'), ('monthly', '이름')):
        r = client.get(f'/api/leave/export?kind={kind}&year={d.year}')
        assert r.status_code == 200 and 'spreadsheetml' in r.headers['Content-Type'], kind
        ws = load_workbook(io.BytesIO(r.data)).active
        assert ws.cell(1, 1).value == first
        names = [ws.cell(i, 1).value for i in range(2, ws.max_row + 1)]
        assert '박민우' in names, kind
        if kind == 'usage':
            row = next(i for i in range(2, ws.max_row + 1) if ws.cell(i, 1).value == '박민우' and ws.cell(i, 2).value == '연차')
            assert ws.cell(row, 8).data_type == 's'                 # 사유가 수식으로 안 들어감
        if kind == 'monthly':
            row = next(i for i in range(2, ws.max_row + 1) if ws.cell(i, 1).value == '박민우')
            assert ws.cell(row, 2 + d.month + 0).value >= 1 or ws.cell(row, 15).value >= 1
    assert client.get('/api/leave/export?kind=nope').status_code == 400


def test_monthly_api_self_and_others(client, mgr):
    _as(client, STAFF)
    j = client.get('/api/leave/monthly').get_json()
    assert j['success'] and 'months' in j['monthly']
    from test_gw_leave_apply import OTHER  # noqa: F401
    other_pid = _pid('고광일')
    assert client.get(f'/api/leave/monthly?pid={other_pid}').status_code == 403
    _as(client, ADMIN)
    assert client.get(f'/api/leave/monthly?pid={_pid("박민우")}&year=2026').get_json()['success']


# ── 연차 사용 촉진 ───────────────────────────────────────────────────
import pytest  # noqa: E402


@pytest.fixture
def lp(mgr, monkeypatch):
    from dashboard.services import leave_promo
    monkeypatch.setattr(leave_promo, '_ready', False)
    return leave_promo


def test_promo_send_ack_plan_designate_web(client, mgr, lp):
    from dashboard.blueprints import slack_leave
    from test_gw_leave_apply import JW
    to = next(g['valid_to'] for g in lv.grants('박민우'))
    _as(client, STAFF)
    assert client.get('/api/leave/promotions').status_code == 403
    r = client.post('/api/leave/promotions/send', json={'name': '박민우', 'kind': 'annual', 'grant_to': to})
    assert r.status_code == 400 and '경영지원만' in r.get_json()['message']
    _as(client, ADMIN)
    mgr.posts.clear()
    j = client.post('/api/leave/promotions/send', json={'name': '박민우', 'kind': 'annual', 'grant_to': to}).get_json()
    assert j['success']
    pid = j['id']
    dm = mgr.to('UMW')[-1]
    assert '연차 사용 촉진' in dm['text']
    ids = [e['action_id'] for b in dm['blocks'] if b['type'] == 'actions' for e in b['elements']]
    assert slack_leave.A_PROMO_ACK in ids
    again = client.post('/api/leave/promotions/send', json={'name': '박민우', 'kind': 'annual', 'grant_to': to}).get_json()
    assert not again['success'] and '이미' in again['message']
    lst = client.get('/api/leave/promotions').get_json()
    assert lst['can_manage'] and lst['rows'][0]['state'] == 'wait_plan'
    # 직원: 내 휴가에 촉구 문서 → 받았어요 → 계획
    _as(client, STAFF)
    promos = client.get('/api/leave/me').get_json()['leave']['promotions']
    assert promos[0]['id'] == pid and '근로기준법 제61조' in promos[0]['notice_text'] and promos[0]['can_plan']
    assert client.post(f'/api/leave/promotions/{pid}/ack').get_json()['success']
    d1, d2 = _next_workday(5), _next_workday(6)
    mgr.posts.clear()
    j = client.post(f'/api/leave/promotions/{pid}/plan',
                    json={'items': [{'date': d1.isoformat()}, {'date': d2.isoformat(), 'half': 'pm'}]}).get_json()
    assert j['success']
    assert mgr.to(JW)                                                  # 계획 날짜는 보통 신청 → 승인자 카드
    assert any('사용 계획을 냈어요' in p['text'] for p in mgr.to('UKIKO'))
    rec = lp.get(pid)
    assert rec['status'] == lp.PLANNED and rec['ack_at'] and len(rec['plan_rids']) == 2
    # 경영지원 지정은 계획 기한 뒤에만
    _as(client, ADMIN)
    r = client.post(f'/api/leave/promotions/{pid}/designate', json={'items': [{'date': _next_workday(8).isoformat()}]})
    assert r.status_code == 400 and '계획 기한' in r.get_json()['message']
    with lv._conn() as c:
        c.execute("UPDATE leave_promotions SET plan_due='2000-01-01' WHERE id=?", (pid,))
    mgr.posts.clear()
    d3 = _next_workday(9)
    j = client.post(f'/api/leave/promotions/{pid}/designate', json={'items': [{'date': d3.isoformat()}]}).get_json()
    assert j['success']
    assert any('사용 시기 지정 통보' in p['text'] for p in mgr.to('UMW'))
    rec = lp.get(pid)
    assert rec['status'] == lp.DESIGNATED and lv.get(rec['designated_rids'][0])['status'] == lv.APPROVED


def test_promo_daily_auto_and_manager_summary(slack, mgr, lp, monkeypatch):
    from datetime import timedelta
    from dashboard.blueprints import slack_leave
    monkeypatch.setenv('LEAVE_ENABLED', 'true')
    to = lv._d(next(g['valid_to'] for g in lv.grants('박민우')))
    day = lp.schedule(lp.ANNUAL, to)['notice_from'] + timedelta(days=1)
    mgr.posts.clear()
    out = slack_leave.run_promotion(day)
    names = sorted(r['name'] for r in out['sent'])
    assert '박민우' in names
    assert any('연차 사용 촉진 안내' in p['text'] for p in mgr.to('UMW'))
    assert any('오늘 촉구를 보냈어요' in p['text'] for p in mgr.to('UKIKO'))
    monkeypatch.setenv('LEAVE_PROMOTION_AUTO', 'false')
    lp_out = lp.daily(day)
    assert lp_out['sent'] == []


def test_slack_edit_view_keeps_used_up_custom_type(client, mgr):
    from dashboard.blueprints import slack_leave
    lv.custom_grant('박민우', 'health', 1, None, None, '', 'UKIKO', '고광일')
    d = _next_workday(4)
    row = lv.create('박민우', 'UMW', 'UMW', '박민우', 'health', '', d, d, '보건')
    assert lv.custom_left('박민우', 'health') == 0
    v = slack_leave.edit_view(row['id'], 'UMW')
    sel = next(b for b in v['blocks'] if b.get('block_id') == 'lv_type')['element']
    assert sel['initial_option']['value'] == 'health'
    assert 'health' in [o['value'] for o in sel['options']]
    # 고쳐도(같은 1일) 맞춤 휴가 한도 안 — 자기 몫은 빼고 본다
    new, errs = lv.edit_request(row['id'], 'health', '', _next_workday(5), None, '보건', 'UMW', '박민우')
    assert errs == {} and new['start'] == _next_workday(5).isoformat()


def test_settlement_api_manager_only(client, mgr):
    from datetime import date as _d
    _as(client, STAFF)
    assert client.get(f'/api/leave/person/{_pid("박민우")}/settlement?last_day=2027-06-30').status_code in (400, 403)
    _as(client, ADMIN)
    j = client.get(f'/api/leave/person/{_pid("박민우")}/settlement?last_day={_d.today().isoformat()}').get_json()
    assert j['success'] and 'short' in j['settlement'] and 'unused' in j['settlement']
    r = client.get(f'/api/leave/person/{_pid("박민우")}/settlement?last_day=')
    assert r.status_code == 400


def test_fiscal_accrual_dm_goes_to_managers_only(slack, mgr, monkeypatch):
    from datetime import date as _d
    from dashboard.blueprints import slack_leave
    monkeypatch.setenv('LEAVE_ENABLED', 'true')
    monkeypatch.setattr(lv, 'ACCRUE_HOLD_OVER', 99)
    mgr.posts.clear()
    slack_leave.run_accrual(_d(2026, 12, 15))
    assert any(g['source'] == 'fiscal:박민우:2027' for g in lv.grants('박민우'))
    assert not any('2027' in p['text'] and '연차' in p['text'] for p in mgr.to('UMW'))
    assert any('회계연도 연차를 넣었어요' in p['text'] and '박민우' in p['text'] for p in mgr.to('UKIKO'))


def test_team_marks_people_still_on_flex(client, mgr):
    _as(client, ADMIN)
    people = {p['name']: p for p in client.get('/api/leave/team').get_json()['people']}
    assert people['박민우']['on_flex'] is False                 # 신청 열림(시범 명단)
    assert people['이직원']['on_flex'] is True                  # 아직 플렉스로 신청
