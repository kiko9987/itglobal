"""차량 배정(2026-11-20) — 카톡 글 해석·저장·확정 DM·바뀐 사람만 DM·배정 요청·화면 권한·운행 기록 연결.
이름은 모두 가짜."""
from datetime import date

import pytest
from flask import Flask

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import vehicle_assign as va
from dashboard.services import vehicle_log as vl

ASSIGNER = 'sj@example.com'
STAFF = [{'code': c, 'name': n, 'slack_id': 'U' + c, 'dept': ''} for c, n in [
    ('YM', '김가람'), ('LSH', '이나래'), ('MJ', '박다솜'), ('JH', '최마루'), ('JW', '정바다'), ('LJW', '한별'),
    ('KMW', '윤솔'), ('TH', '서준'), ('JSH', '문채'), ('SJ', '오늘'), ('JK', '강해'), ('YG', '임산')]]



@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, vl, va, hr):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setattr(va, 'staff', lambda: [dict(p) for p in STAFF])
    monkeypatch.setattr(va, 'assigners', lambda: [ASSIGNER])
    sent = []
    monkeypatch.setattr(va, '_dm', lambda sid, text: sent.append((sid, text)) or True)
    gw._invalidate()
    yield sent
    gw._invalidate()


def _rows(view):
    return {c['plate']: c['people'] for c in view['cars'] if c['people']}


def test_seeded_fleet(db):
    cars = {c['plate']: c for c in va.fleet()}
    assert len(cars) == 11 and len(va.fleet(True)) == 14      # 파나메라·벤츠·랜드로버는 배정에서 뺌
    assert cars['237허5358']['label'] == '레이 5358' and cars['237허5358']['usage'] == va.SHARED
    assert cars['143고9687']['usage'] == va.SHARED and cars['53우1783']['usage'] == va.SHARED   # 카니발·스파크·모닝도 매일 배정
    assert '204두8988' not in cars and {c['plate']: c for c in va.fleet(True)}['204두8988']['usage'] == va.OFF
    assert cars['143고9687']['needs_log'] == 0
    va.update_car('133가6633', {'usage': va.FIXED, 'fixed_user': 'MJ'}, ASSIGNER)
    assert va.fixed_users()['133가6633'] == 'MJ'
    with pytest.raises(ValueError):
        va.update_car('53우1783', {'usage': 'x'}, ASSIGNER)


def test_visitors_with_visits_that_day(db):
    from tests.unit.test_vehicle_log import _lead
    leads = [_lead('L-1', '2026-11-21', '용인 수지구 심곡로 11', '김가람, 이나래'),
             _lead('L-2', '2026-11-21', '부천 길주로 1', '김가람'),
             _lead('L-3', '2026-11-22', '부천 길주로 2', '박다솜'),                     # 다른 날
             _lead('L-4', '2026-11-21', '부천 길주로 3', '외부인'),                     # 직원 명단에 없음
             _lead('L-5', '2026-11-21', '부천 길주로 4', '서준', status='유선 상담')]   # 가지 않은 리드
    assert va.visitors('2026-11-21', leads) == [{'code': 'YM', 'name': '김가람', 'visits': 2},
                                                 {'code': 'LSH', 'name': '이나래', 'visits': 1}]


def test_save_confirm_and_change_dms(db):
    sent = db
    with pytest.raises(ValueError):
        va.save_day('2026-11-21', [{'plate': '237허5358', 'people': ['YM']},
                                   {'plate': '231허6854', 'people': ['YM']}], ASSIGNER)
    rows = [{'plate': '237허5358', 'people': ['TH', 'JSH'], 'note': '유리용접'},
            {'plate': '03모8021', 'people': ['JW', 'LJW']}]
    res = va.save_day('2026-11-21', rows, ASSIGNER)
    assert sorted(res['changed']) == ['03모8021', '237허5358'] and res['notified'] == 0 and not sent
    res = va.save_day('2026-11-21', rows, ASSIGNER, confirm=True)
    assert res['notified'] == 4 and res['confirmed']
    th = next(t for sid, t in sent if sid == 'UTH')
    assert '11/21(토)' in th and '레이 5358' in th and '문채' in th and '유리용접' in th
    sent.clear()
    # 확정 뒤 고치기 — LJW 를 레이로 옮기고 JSH 를 뺌 → 바뀐 사람(LJW·JSH·TH 같이 타는 사람 바뀜·JW)만
    res = va.save_day('2026-11-21', [{'plate': '237허5358', 'people': ['TH', 'LJW'], 'note': '유리용접'},
                                     {'plate': '03모8021', 'people': ['JW']}], ASSIGNER)
    who = {sid: t for sid, t in sent}
    assert set(who) == {'UTH', 'ULJW', 'UJSH', 'UJW'}
    assert '빠졌어요' in who['UJSH'] and '차량이 바뀌었어요' in who['ULJW']
    assert va.person_car('2026-11-21', 'LJW')['plate'] == '237허5358'


def test_day_view_draft_from_previous_shared_only(db):
    va.update_car('133가6633', {'usage': va.FIXED}, ASSIGNER)     # 고정 차(빌려 쓰기)는 초안에 안 따라옴
    va.save_day('2026-11-20', [{'plate': '237허5358', 'people': ['TH']},
                               {'plate': '133가6633', 'people': ['MJ']}], ASSIGNER)
    v = va.day_view('2026-11-21')
    assert v['draft_from'] == '2026-11-20' and _rows(v) == {'237허5358': ['TH']}   # 빌려 쓴 고정 차는 안 따라옴
    assert va.day_view('2026-11-20')['draft_from'] is None


def test_request_assignment_once(db, monkeypatch):
    sent = db
    monkeypatch.setattr(hr, 'find_by_email', lambda e: {'slack_id': 'USJ'} if e == ASSIGNER else None)
    assert va.request_assignment('2026-11-21') == 1 and '11/21(토)' in sent[0][1] and '/vehicles/assign?day=2026-11-21' in sent[0][1]
    assert va.request_assignment('2026-11-21') == 0                 # 이미 보냄
    va.save_day('2026-11-23', [{'plate': '237허5358', 'people': ['TH']}], ASSIGNER, confirm=True)
    assert va.request_assignment('2026-11-23') == 0                 # 이미 확정


def test_next_workday():
    assert va.next_workday(date(2026, 11, 20)) == '2026-11-21'     # 금 → 토
    assert va.next_workday(date(2026, 11, 21)) == '2026-11-23'     # 토 → (일 건너뛰고) 월


def test_vehicle_log_prefers_assignment(db):
    from tests.unit.test_vehicle_log import _xlsx, T1, LEADS
    vl.import_upload('a.xlsx', _xlsx(trips=T1), 'x')
    # 운행일지 시험 차량(12가3456)은 차량 목록에 없어 배정할 수 없다 → 실제 차 번호로 바꿔 넣는다
    with vl._conn() as conn:
        conn.execute("UPDATE vehicle_trips SET plate = '237허5358'")
    va.save_day('2026-11-20', [{'plate': '237허5358', 'people': ['TH', 'JSH']}], ASSIGNER)
    d = vl.vehicle_days('2026-11', '237허5358', LEADS)
    assert d['days'][0]['drivers'] == ['서준', '문채'] and d['days'][0]['source'] == 'assigned'
    assert d['days'][1]['source'] == 'guess'
    car = vl.month_summary('2026-11', LEADS)['vehicles'][0]
    assert car['source'] == 'assigned' and car['drivers'] == ['서준'] and car['known_days'] == 1


def test_routes_view_for_all_edit_for_assigner(db, monkeypatch):
    import dashboard.auth as auth
    from dashboard.blueprints import gw_vehicles as bp
    from dashboard.blueprints.groupware import field_path_allowed
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    monkeypatch.setattr(hr, 'hr_admins', lambda: [])
    monkeypatch.setattr(bp, 'render_template', lambda name: name)
    from tests.unit.test_vehicle_log import _lead
    monkeypatch.setattr(bp, '_leads', lambda: [_lead('L-1', '2026-11-21', '부천 길주로 1', '서준')])
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': 'staff@example.com', 'permission_level': 'user'}
        j = cl.get('/api/vehicles/assign?day=2026-11-21').get_json()
        assert j['success'] and not j['can_edit'] and len(j['cars']) == 11 and j['day_label'] == '11/21(토)'
        assert j['visitors'] == [{'code': 'TH', 'name': '서준', 'visits': 1}]
        assert cl.get('/vehicles/assign').get_data(as_text=True) == 'gw/vehicle_assign.html'
        assert cl.post('/api/vehicles/assign', json={'day': '2026-11-21', 'rows': []}).status_code == 403
        with cl.session_transaction() as s:
            s['user'] = {'email': ASSIGNER, 'permission_level': 'user'}
        assert cl.get('/api/vehicles/assign?day=2026-11-21').get_json()['can_edit']
        r = cl.post('/api/vehicles/assign', json={'day': '2026-11-21', 'confirm': True,
                                                  'rows': [{'plate': '237허5358', 'people': ['TH']}]}).get_json()
        assert r['success'] and r['notified'] == 1
        bad = cl.post('/api/vehicles/assign', json={'day': '2026-13-01', 'rows': []})
        assert bad.status_code == 400
        assert cl.post('/api/vehicles/assign/parse', json={'text': 'YM 레이'}).status_code == 404   # 카톡 붙여넣기 없음
        f = cl.post('/api/vehicles/fleet/133가6633', json={'fixed_user': 'JK'}).get_json()
        assert f['car']['fixed_user'] == 'JK'
    assert field_path_allowed('GET', '/vehicles/assign') and field_path_allowed('GET', '/api/vehicles/assign')
    assert not field_path_allowed('POST', '/api/vehicles/assign')


def test_menu():
    items = {it['href']: it for a in gw.AREAS for it in a['items']}
    assert items['/vehicles/assign'].get('field') and not items['/vehicles/assign'].get('hr_admin')
    assert items['/vehicles'].get('hr_admin')
