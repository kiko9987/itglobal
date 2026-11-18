"""그룹웨어 휴가 관리 플렉스 맞추기(2026-11-18) — 웹 API: 조정 적용일·취소·일괄, 엑셀, 월별 상세, 맞춤 휴가."""
import io

from dashboard.services import leave as lv
from tests.unit.test_gw_leave_admin import _pid, mgr  # noqa: F401
from tests.unit.test_gw_leave_apply import ADMIN, STAFF, _as, _next_workday, client, slack  # noqa: F401


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
    from tests.unit.test_gw_leave_apply import OTHER  # noqa: F401
    other_pid = _pid('고광일')
    assert client.get(f'/api/leave/monthly?pid={other_pid}').status_code == 403
    _as(client, ADMIN)
    assert client.get(f'/api/leave/monthly?pid={_pid("박민우")}&year=2026').get_json()['success']
