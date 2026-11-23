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
    monkeypatch.setattr(va, 'CAR_DM_ON', True)         # 운영에선 꺼 둠(11-23)
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


# ── 그룹웨어 [운행 종료](11-22) ────────────────────────────────────────────
from datetime import date as _date  # noqa: E402


def test_end_trip_web_then_no_dm(db):
    calls = db
    va.save_day('2026-11-23', [{'plate': '231허6854', 'people': ['YM', 'LSH']}], ASSIGNER, confirm=True)
    today = _date(2026, 11, 23)
    vd.end_trip('2026-11-23', '231허6854', 'YM', 'LSH', 'LSH', today=today)   # 김가람이 '이나래 운전'으로 적음
    a = vd.answer('2026-11-23', '231허6854')
    assert (a['am'], a['pm'], a['answered_by'], a['via']) == ('LSH', 'LSH', 'YM', 'web') and a['ended_at']
    assert vd.send_asks('2026-11-23')['sent'] == 0 and not [c for c in calls if c[0] == 'chat.postMessage']


def test_end_trip_after_dm_updates_dm(db):
    calls = db
    va.save_day('2026-11-23', [{'plate': '231허6854', 'people': ['YM', 'LSH']}], ASSIGNER, confirm=True)
    vd.send_asks('2026-11-23')
    calls.clear()
    vd.end_trip('2026-11-23', '231허6854', 'LSH', 'YM', 'LSH', today=_date(2026, 11, 24))   # 다음 날 아침에 나눠서
    ups = [p for m, p in calls if m == 'chat.update']
    assert len(ups) == 2 and all('오전 김가람 · 오후 이나래' in p['text'] for p in ups)
    first = vd.answer('2026-11-23', '231허6854')['ended_at']
    vd.end_trip('2026-11-23', '231허6854', 'YM', 'YM', 'YM', today=_date(2026, 11, 24))   # 고치기 — 마지막 답이 맞음
    a = vd.answer('2026-11-23', '231허6854')
    assert (a['am'], a['pm']) == ('YM', 'YM') and a['ended_at'] == first     # 종료 시각은 처음 것


def test_end_trip_rules(db):
    va.save_day('2026-11-23', [{'plate': '231허6854', 'people': ['YM', 'LSH']}, {'plate': '53우1783', 'people': ['TH']}],
                ASSIGNER)                                                    # 저장만 해도 적을 수 있음
    today = _date(2026, 11, 23)
    with pytest.raises(ValueError, match='탄 사람만'):
        vd.end_trip('2026-11-23', '231허6854', 'TH', 'YM', 'YM', today=today)
    with pytest.raises(ValueError, match='중에서'):
        vd.end_trip('2026-11-23', '231허6854', 'YM', 'TH', 'TH', today=today)
    with pytest.raises(ValueError, match='골라'):
        vd.end_trip('2026-11-23', '231허6854', 'YM', '', '', today=today)
    with pytest.raises(ValueError, match='마친 뒤'):
        vd.end_trip('2026-11-23', '231허6854', 'YM', 'YM', 'YM', today=_date(2026, 11, 22))
    with pytest.raises(ValueError, match='경영지원'):
        vd.end_trip('2026-11-23', '231허6854', 'YM', 'YM', 'YM', today=_date(2026, 12, 5))
    vd.end_trip('2026-11-23', '53우1783', 'TH', '', '', today=today)           # 혼자 탄 차는 고르지 않아도
    assert vd.answer('2026-11-23', '53우1783')['am'] == 'TH'


def test_my_trips(db):
    va.save_day('2026-11-23', [{'plate': '16소4041', 'people': ['YG', 'JK'], 'pm': ['SJ', 'JK']},
                               {'plate': '231허6854', 'people': ['YM', 'LSH']}], ASSIGNER)
    t = vd.my_trips('jk', '2026-11-23')
    assert len(t) == 1 and t[0]['plate'] == '16소4041' and t[0]['answer'] is None
    assert [r['code'] for r in t[0]['riders']] == ['YG', 'JK', 'SJ']
    assert t[0]['am_people'] == ['임산', '강해'] and t[0]['pm_people'] == ['오늘', '강해']
    assert vd.my_trips('SJ', '2026-11-23')[0]['plate'] == '16소4041'          # 오후에만 탄 사람도
    assert vd.my_trips('TH', '2026-11-23') == [] and vd.my_trips('', '2026-11-23') == []
    vd.end_trip('2026-11-23', '16소4041', 'JK', 'YG', 'SJ', today=_date(2026, 11, 23))
    a = vd.my_trips('YG', '2026-11-23')[0]['answer']
    assert (a['am_name'], a['pm_name'], a['by_name'], a['via']) == ('임산', '오늘', '강해', 'web') and a['ended_at']
    assert vd.code_of_name('강해') == 'JK' and vd.code_of_name('없는사람') == ''


def test_trip_routes_and_field_paths(db, monkeypatch):
    import dashboard.auth as auth
    from flask import Flask
    from dashboard.blueprints import gw_vehicles as bp
    from dashboard.blueprints.groupware import field_path_allowed
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    monkeypatch.setattr(gw, 'DEV_PREVIEW_USERS', ())
    me = {'name': '강해'}
    monkeypatch.setattr(bp, '_my_person', lambda: me)
    today = _date.today().isoformat()
    va.save_day(today, [{'plate': '16소4041', 'people': ['YG', 'JK']}], ASSIGNER)
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': 'staff@example.com', 'permission_level': 'user'}
        j = cl.get('/api/vehicles/trip').get_json()
        assert j['success'] and [t['plate'] for t in j['trips']] == ['16소4041']
        r = cl.post('/api/vehicles/trip/end', json={'day': today, 'plate': '16소4041', 'am': 'YG', 'pm': 'YG'}).get_json()
        assert r['success'] and r['trips'][0]['answer']['am_name'] == '임산'
        bad = cl.post('/api/vehicles/trip/end', json={'day': today, 'plate': '16소4041', 'am': 'TH', 'pm': 'TH'})
        assert bad.status_code == 400
        me['name'] = '모르는사람'
        assert cl.get('/api/vehicles/trip').get_json()['trips'] == []
        assert cl.post('/api/vehicles/trip/end', json={'day': today, 'plate': '16소4041'}).status_code == 403
    assert field_path_allowed('GET', '/api/vehicles/trip') and field_path_allowed('POST', '/api/vehicles/trip/end')


def test_confirm_dm_mentions_end_button_only_after_open(db, monkeypatch):
    rows = {'231허6854': {'people': ['YM', 'LSH'], 'note': '', 'pm': None},
            '53우1783': {'people': ['TH'], 'note': '', 'pm': None}}
    cars = {c['plate']: c for c in va.fleet(True)}
    who = va._who_map()
    assert '운행 종료' not in va._person_text('2026-11-23', 'YM', rows, cars, who)      # 개발 중엔 안 붙임
    monkeypatch.setattr(gw, 'DEV_PREVIEW_USERS', ())
    assert '[운행 종료]' in va._person_text('2026-11-23', 'YM', rows, cars, who)
    assert '운행 종료' not in va._person_text('2026-11-23', 'TH', rows, cars, who)      # 혼자 탄 차는 자동
