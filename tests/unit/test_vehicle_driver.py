"""실제 운전자 받기(2026-11-22) — 확정한 날만 묻기, 혼자 탄 차는 자동, 이름·나눠서 운전 버튼, 다른 사람 DM 갱신,
운행일지가 배정보다 실제 운전자를 먼저 쓰기. 이름은 모두 가짜."""
import json

import pytest

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import vehicle_assign as va
from dashboard.services import vehicle_driver as vd
from dashboard.services import vehicle_log as vl

ASSIGNER = 'sj@example.com'
STAFF = [{'code': c, 'name': n, 'slack_id': 'U' + c, 'dept': ''} for c, n in [
    ('YM', '김가람'), ('LSH', '이나래'), ('TH', '서준'), ('LJW', '한별'), ('SJ', '오늘'), ('YG', '임산'), ('JK', '강해')]]


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, vl, va, hr, vd):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setattr(va, 'staff', lambda: [dict(p) for p in STAFF])
    monkeypatch.setattr(va, 'assigners', lambda: [ASSIGNER])
    monkeypatch.setattr(va, '_dm', lambda sid, text: True)
    calls = []

    def fake_slack(method, payload):
        calls.append((method, payload))
        return {'ok': True, 'channel': 'D' + payload.get('channel', 'X'), 'ts': f'1.{len(calls)}'}
    monkeypatch.setattr(vd, '_slack', fake_slack)
    gw._invalidate()
    yield calls
    gw._invalidate()


def _button(blocks, label, action=None):
    for b in blocks:
        for e in b.get('elements') or []:
            if e.get('type') == 'button' and e['text']['text'] == label and (action is None or e['action_id'] == action):
                return e['action_id'], e['value']
    raise AssertionError(label)


def test_asks_only_confirmed_and_auto_single(db):
    calls = db
    rows = [{'plate': '231허6854', 'people': ['YM', 'LSH']}, {'plate': '53우1783', 'people': ['TH']}]
    va.save_day('2026-11-23', rows, ASSIGNER)                                   # 저장만 — 확정 전
    assert vd.send_asks('2026-11-23') == {'day': '2026-11-23', 'sent': 0, 'auto': 0} and not calls
    va.save_day('2026-11-23', rows, ASSIGNER, confirm=True)
    r = vd.send_asks('2026-11-23')
    assert (r['sent'], r['auto']) == (2, 1)                                      # 레이 두 명에게, 스파크는 혼자
    assert vd.answer('2026-11-23', '53우1783')['am'] == 'TH'
    assert {p['channel'] for m, p in calls if m == 'chat.postMessage'} == {'UYM', 'ULSH'}
    assert vd.send_asks('2026-11-23')['sent'] == 0                               # 다시 안 보냄


def test_one_tap_answer_updates_everyone(db):
    calls = db
    va.save_day('2026-11-23', [{'plate': '231허6854', 'people': ['YM', 'LSH']}], ASSIGNER, confirm=True)
    vd.send_asks('2026-11-23')
    blocks = next(p['blocks'] for m, p in calls if m == 'chat.postMessage')
    calls.clear()
    act, val = _button(blocks, '이나래')
    vd.handle_action(act, val, 'UYM', 'DUYM', '1.1')                              # 김가람이 '이나래' 를 누름
    a = vd.answer('2026-11-23', '231허6854')
    assert (a['am'], a['pm']) == ('LSH', 'LSH')
    ups = [p for m, p in calls if m == 'chat.update']
    assert len(ups) == 2 and all('운전: 이나래 (답: 김가람)' in p['text'] for p in ups)


def test_split_answer_am_pm(db):
    calls = db
    va.save_day('2026-11-23', [{'plate': '16소4041', 'people': ['YG', 'JK'], 'pm': ['SJ', 'JK']}], ASSIGNER, confirm=True)
    vd.send_asks('2026-11-23')
    blocks = next(p['blocks'] for m, p in calls if m == 'chat.postMessage')
    act, val = _button(blocks, '나눠서 운전(오전·오후)')
    calls.clear()
    vd.handle_action(act, val, 'UJK', 'DUJK', '1.1')
    assert calls[-1][0] == 'chat.update' and calls[-1][1]['ts'] == '1.1'        # 누른 사람 DM 만 오전·오후 고르기로
    split_blocks = calls[-1][1]['blocks']
    _, am_val = _button(split_blocks, '임산', vd.A_AM)
    _, pm_val = _button(split_blocks, '오늘', vd.A_PM)
    vd.handle_action(vd.A_AM, am_val, 'UJK', 'DUJK', '1.1')
    assert not [p for m, p in calls if m == 'chat.update' and '답:' in p.get('text', '')]   # 반쪽만 — 아직
    vd.handle_action(vd.A_PM, pm_val, 'UJK', 'DUJK', '1.1')
    a = vd.answer('2026-11-23', '16소4041')
    assert (a['am'], a['pm'], a['answered_by']) == ('YG', 'SJ', 'JK')
    assert any('오전 임산 · 오후 오늘' in p.get('text', '') for m, p in calls if m == 'chat.update')


def test_log_uses_actual_driver(db):
    from tests.unit.test_vehicle_log import _xlsx, LEADS
    trips = [('2026-11-23(월)', '09:10:00', '경기도 김포시 걸포동', '서울특별시 영등포구 영등포동', 30.0),
             ('2026-11-23(월)', '14:00:00', '서울특별시 영등포구 영등포동', '서울특별시 강남구 역삼동', 15.0)]
    vl.import_upload('a.xlsx', _xlsx(trips=trips), 'x')
    with vl._conn() as conn:
        conn.execute("UPDATE vehicle_trips SET plate = '231허6854'")
    va.save_day('2026-11-23', [{'plate': '231허6854', 'people': ['YM', 'LSH']}], ASSIGNER, confirm=True)
    d = vl.vehicle_days('2026-11', '231허6854', LEADS)['days'][0]
    assert d['source'] == 'assigned' and [t['driver'] for t in d['trips']] == ['김가람', '김가람']
    vd._save('2026-11-23', '231허6854', 'LSH', 'YM', 'LSH')                           # 실제: 오전 이나래, 오후 김가람
    d = vl.vehicle_days('2026-11', '231허6854', LEADS)['days'][0]
    assert d['source'] == 'confirmed'
    assert [t['driver'] for t in sorted(d['trips'], key=lambda t: t['start'])] == ['이나래', '김가람']
    assert d['drivers'][:2] == ['이나래', '김가람']


def test_handle_bad_values(db):
    assert vd.handle_action(vd.A_ONE, 'not json', 'UYM') is None
    assert '바뀌어' in vd.handle_action(vd.A_ONE, json.dumps({'day': '2026-11-30', 'plate': 'x', 'code': 'YM'}), 'UYM')
