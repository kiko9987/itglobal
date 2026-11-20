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

DAY1 = 'YM+LSH  레이\nMJ+JH 레이\nJW+LJW+KMW 8021\nTH+JSH 레이 5358\nSJ 전기트럭'


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


def test_parse_kakao_text(db):
    r = va.parse_text(DAY1 + '\n2026년 11월 20일 금요일\nXX 레이', '2026-11-21')
    rows = {x['plate']: x for x in r['rows']}
    assert rows['03모8021']['people'] == ['JW', 'LJW', 'KMW'] and not rows['03모8021']['unsure']
    assert rows['237허5358']['people'] == ['TH', 'JSH'] and not rows['237허5358']['unsure']
    assert rows['90머9131']['people'] == ['SJ']
    # 그냥 '레이' 두 줄 → 남은 레이 두 대에 하나씩, 확인 필요
    rays = [x for x in r['rows'] if x['plate'] in ('231허6854', '231허6873')]
    assert len(rays) == 2 and all(x['unsure'] for x in rays)
    assert sorted(p for x in rays for p in x['people']) == ['JH', 'LSH', 'MJ', 'YM']
    assert set(rays[0]['choices']) == {'237허5358', '231허6854', '231허6873'}
    assert r['skipped'] == ['2026년 11월 20일 금요일', 'XX 레이']   # 이니셜이 아닌 줄·모르는 사람
    assert va.parse_text('JW 파나메라')['skipped'] == ['JW 파나메라']   # 배정에서 뺀 차


def test_parse_later_line_wins_and_merges_same_car(db):
    r = va.parse_text('TH+JSH 레이 5358\nTH 5358(유리용접)\nJK 4041\nJK 쏘렌토 8021', '2026-11-21')
    rows = {x['plate']: x for x in r['rows']}
    assert rows['237허5358']['people'] == ['TH', 'JSH'] and rows['237허5358']['note'] == '유리용접'
    assert '16소4041' not in rows and rows['03모8021']['people'] == ['JK']   # 고쳐 쓴 뒤 줄이 이긴다


def test_parse_prefers_recent_car(db):
    va.save_day('2026-11-20', [{'plate': '231허6873', 'people': ['YM']}], ASSIGNER)
    r = va.parse_text('YM 레이', '2026-11-21')
    assert r['rows'][0]['plate'] == '231허6873' and r['rows'][0]['unsure']


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
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': 'staff@example.com', 'permission_level': 'user'}
        j = cl.get('/api/vehicles/assign?day=2026-11-21').get_json()
        assert j['success'] and not j['can_edit'] and len(j['cars']) == 11 and j['day_label'] == '11/21(토)'
        assert cl.get('/vehicles/assign').get_data(as_text=True) == 'gw/vehicle_assign.html'
        assert cl.post('/api/vehicles/assign', json={'day': '2026-11-21', 'rows': []}).status_code == 403
        assert cl.post('/api/vehicles/assign/parse', json={'text': 'YM 레이'}).status_code == 403
        with cl.session_transaction() as s:
            s['user'] = {'email': ASSIGNER, 'permission_level': 'user'}
        assert cl.get('/api/vehicles/assign?day=2026-11-21').get_json()['can_edit']
        r = cl.post('/api/vehicles/assign', json={'day': '2026-11-21', 'confirm': True,
                                                  'rows': [{'plate': '237허5358', 'people': ['TH']}]}).get_json()
        assert r['success'] and r['notified'] == 1
        bad = cl.post('/api/vehicles/assign', json={'day': '2026-13-01', 'rows': []})
        assert bad.status_code == 400
        p = cl.post('/api/vehicles/assign/parse', json={'day': '2026-11-22', 'text': 'JK 4041'}).get_json()
        assert p['rows'][0]['plate'] == '16소4041'
        f = cl.post('/api/vehicles/fleet/133가6633', json={'fixed_user': 'JK'}).get_json()
        assert f['car']['fixed_user'] == 'JK'
    assert field_path_allowed('GET', '/vehicles/assign') and field_path_allowed('GET', '/api/vehicles/assign')
    assert not field_path_allowed('POST', '/api/vehicles/assign')


def test_menu():
    items = {it['href']: it for a in gw.AREAS for it in a['items']}
    assert items['/vehicles/assign'].get('field') and not items['/vehicles/assign'].get('hr_admin')
    assert items['/vehicles'].get('hr_admin')
