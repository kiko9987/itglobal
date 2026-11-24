"""차량 배정(2026-11-20) — 카톡 글 해석·저장·확정 DM·바뀐 사람만 DM·배정 요청·화면 권한·운행 기록 연결.
이름은 모두 가짜."""
from datetime import date

import pytest
from flask import Flask

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import vehicle_assign as va
from dashboard.services import vehicle_log as vl

ORIG_STAFF = va.staff   # 아래 db 픽스처가 가짜 명단으로 바꾸기 전 원래 함수
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
    monkeypatch.setattr(gw, 'DEV_PREVIEW_USERS', ())   # 개발 중 문은 아래 test_dev_preview_* 에서 따로 본다
    monkeypatch.setattr(va, 'CAR_DM_ON', True)         # 운영에선 꺼 둠(11-23) — 꺼진 때는 test_car_dm_off 에서
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
    # 11-22 부장님: 카니발=YG 차(모자랄 때 빌려 씀), 스파크=MW 개인 차(유류비 지원) → 고정 차. 모닝은 공용
    assert (cars['143고9687']['usage'], cars['143고9687']['fixed_user']) == (va.FIXED, 'YG')
    assert (cars['53우1783']['usage'], cars['53우1783']['fixed_user'], cars['53우1783']['corp']) == (va.FIXED, 'MW', '개인(MW)')
    assert cars['60거5484']['usage'] == va.SHARED
    assert '204두8988' not in cars and {c['plate']: c for c in va.fleet(True)}['204두8988']['usage'] == va.OFF
    assert cars['143고9687']['needs_log'] == 0
    va.update_car('133가6633', {'usage': va.FIXED, 'fixed_user': 'MJ'}, ASSIGNER)
    assert va.fixed_users()['133가6633'] == 'MJ'
    with pytest.raises(ValueError):
        va.update_car('53우1783', {'usage': 'x'}, ASSIGNER)


def test_staff_excludes_office(monkeypatch):
    """사무직(경영지원 부서·비영업 명단)은 차 배정 명단에서 뺀다 — 샛별(11-20 부장님)."""
    from dashboard.services import staff_directory as sd
    monkeypatch.setattr(sd, 'suffix_map', lambda: {'김가람': 'YM', '구름': 'SB', '송이': 'MW', '하늘': 'HN'})
    monkeypatch.setattr(sd, 'non_sales_owners', lambda: ['하늘'])
    monkeypatch.setattr(hr, '_rows', lambda sql, args=(): [
        {'name': '김가람', 'slack_id': 'U1', 'department': '기술영업'},
        {'name': '구름', 'slack_id': 'U2', 'department': '경영지원'},
        {'name': '송이', 'slack_id': 'U3', 'department': '영업'},
        {'name': '하늘', 'slack_id': 'U4', 'department': ''}])
    assert [p['code'] for p in ORIG_STAFF()] == ['MW', 'YM']


def _leads_1121():
    from test_vehicle_log import _lead
    return [_lead('L-1', '2026-11-21', '부천 길주로 1', '오늘, 윤솔'),
            _lead('L-2', '2026-11-21', '부천 길주로 2', '오늘, 윤솔'),
            _lead('L-3', '2026-11-21', '부천 길주로 3', '서준, 한별'),
            _lead('L-4', '2026-11-21', '부천 길주로 4', '김가람, 서준, 한별'),   # 한 번만 같이 간 조합
            _lead('L-5', '2026-11-21', '부천 길주로 5', '강해'),
            _lead('L-6', '2026-11-21', '부천 길주로 6', '외부인')]


def test_teams_from_schedule(db):
    """일정확정 조합(영업 담당자)이 팀 — 겹치면 방문 많은 조합, 같으면 작은 조합이 먼저(11-21 실제 모양)."""
    t = {tuple(x['members']): x['visits'] for x in va.teams('2026-11-21', _leads_1121())}
    assert t == {('SJ', 'KMW'): 2, ('JK',): 1, ('TH', 'LJW'): 1, ('YM',): 1}


def test_visitors_with_visits_that_day(db):
    from test_vehicle_log import _lead
    leads = [_lead('L-1', '2026-11-21', '용인 수지구 심곡로 11', '김가람, 이나래'),
             _lead('L-2', '2026-11-21', '부천 길주로 1', '김가람'),
             _lead('L-3', '2026-11-22', '부천 길주로 2', '박다솜'),                     # 다른 날
             _lead('L-4', '2026-11-21', '부천 길주로 3', '외부인'),                     # 직원 명단에 없음
             _lead('L-5', '2026-11-21', '부천 길주로 4', '서준', status='유선 상담')]   # 가지 않은 리드
    assert va.visitors('2026-11-21', leads) == [
        {'code': 'YM', 'name': '김가람', 'visits': 2,
         'places': [{'label': '용인 수지구', 'addr': '용인 수지구 심곡로 11'}, {'label': '부천', 'addr': '부천 길주로 1'}]},
        {'code': 'LSH', 'name': '이나래', 'visits': 1,
         'places': [{'label': '용인 수지구', 'addr': '용인 수지구 심곡로 11'}]}]


def test_visitor_places_use_road_units(db):
    """갈 곳은 방문 일정과 같은 도로명 기준 — 찾아 둔 지번 동('풍덕천동')이 있어도 '용인 수지구'(11-24 부장님),
    같은 주소 두 번은 한 곳."""
    from test_vehicle_log import _lead
    with vl._conn() as conn:
        conn.execute('INSERT INTO vehicle_addr_dong (addr, dong, checked_at) VALUES (?,?,?)',
                     ('용인 수지구 심곡로 11', '풍덕천동', '2026-11-21 09:00'))
    leads = [_lead('L-1', '2026-11-21', '용인 수지구 심곡로 11', '김가람'),
             _lead('L-2', '2026-11-21', '용인 수지구 심곡로 11', '김가람')]
    v = va.visitors('2026-11-21', leads)
    assert v[0]['visits'] == 2 and v[0]['places'] == [{'label': '용인 수지구', 'addr': '용인 수지구 심곡로 11'}]
    assert vl.unresolved(leads, '2026-11-21', '2026-11-21') == []


def test_day_off_from_groupware_leave_and_saturday(db, monkeypatch):
    """그날 쉬는 사람 = 그룹웨어 휴가(승인) + 토요일 휴무표 — 캔버스 아님(11-21 부장님). 반차는 따로, 명단 밖 사람은 뺌."""
    from dashboard.services import leave_view, sat_off
    monkeypatch.setattr(leave_view, 'off_on', lambda d: [
        {'name': '김가람', 'type': '연차', 'half': ''},
        {'name': '이나래', 'type': '연차', 'half': '오전'},
        {'name': '외부인', 'type': '연차', 'half': ''}])
    monkeypatch.setattr(sat_off, 'picks', lambda m: {'2026-11-21': ['박다솜', '김가람']})
    monkeypatch.setattr(sat_off, 'month_row', lambda m: {'status': sat_off.CONFIRMED})
    sat = va.day_off('2026-11-21')
    assert sat['off'] == [{'code': 'YM', 'name': '김가람', 'why': '연차'},          # 휴가가 토요 휴무보다 먼저
                          {'code': 'MJ', 'name': '박다솜', 'why': '스케줄 근무 쉬는 날', 'soft': False}]
    assert sat['half'] == [{'code': 'LSH', 'name': '이나래', 'why': '오전 반차'}]
    monkeypatch.setattr(sat_off, 'month_row', lambda m: {'status': sat_off.OPEN})
    assert va.day_off('2026-11-21')['off'][1] == {'code': 'MJ', 'name': '박다솜', 'why': '스케줄 근무 쉬는 날(확정 전)', 'soft': True}
    assert [o['code'] for o in va.day_off('2026-11-20')['off']] == ['YM']        # 평일엔 휴무표 안 봄


def test_day_schedule_like_canvas1(db):
    """그날 방문 일정 칸 = 캔버스1과 같은 묶음(기타·거래처·온라인)·표시, 기간 방문은 걸친 날 모두, 가지 않은 리드는 뺌."""
    def lead(no, day, platform, who='', status='방문 예약', **kw):
        d = {'리드 No': no, '방문 예정일': day, '플랫폼': platform, '영업 담당자': who, '상태': status,
             '고객 연락처': '010-1234-5678', '방문 주소': '부천 길주로 1 상가', '상담 내용': '에어컨 2대'}
        d.update(kw)
        return d
    leads = [lead('L-1', '2026-11-21', '당근', '김가람'),
             lead('L-2', '2026-11-20~22', '거래처', '서준, 한별', status='공사 확정'),
             lead('ETC-a1', '2026-11-21', '기타', **{'본인 방문 여부': '본인 방문 필수'}),
             lead('L-3', '2026-11-21', '홈페이지', '오늘', status='유선 상담'),           # 가지 않음
             lead('L-4', '2026-11-22', '홈페이지', '오늘'),                              # 다른 날
             lead('L-5', '2026-11-21', '전화', '-', **{'고객 연락처': '/'})]
    s = va.day_schedule('2026-11-21', leads)
    assert [g['cat'] for g in s] == ['기타', '거래처', '온라인 방문']
    etc, partner, online = (g['items'] for g in s)
    assert etc[0]['lead_no'] == 'ETC-a1' and etc[0]['marks'] == ['본인 방문 필수'] and etc[0]['platform'] == '기타' and etc[0]['who'] == []
    assert partner[0]['marks'] == [] and partner[0]['platform'] == '거래처' and partner[0]['who'] == ['서준', '한별'] and partner[0]['date'] == '11월 20~22일'
    assert [i['lead_no'] for i in online] == ['L-1', 'L-5']
    assert online[0]['marks'] == [] and online[0]['platform'] == '당근' and online[1]['platform'] == '전화' and online[0]['phone'] == '010-1234-5678' and online[0]['text'] == '에어컨 2대'
    assert online[1]['phone'] == '' and online[1]['who'] == []                     # '/'·'-' 는 빈칸


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
    assert '빠졌어요' in who['UJSH'] and '차량 배정이 바뀌었어요' in who['ULJW'] and '레이 5358' in who['ULJW']
    assert va.person_car('2026-11-21', 'LJW')['plate'] == '237허5358'
    # 확정 뒤 [알리지 않고 저장](메모 오타 등) — 저장은 되고 DM 은 안 감(11-21)
    sent.clear()
    res = va.save_day('2026-11-21', [{'plate': '237허5358', 'people': ['TH', 'LJW'], 'note': '유리 용접'},
                                     {'plate': '03모8021', 'people': ['JW']}], ASSIGNER, notify=False)
    assert res['changed'] == ['237허5358'] and res['notified'] == 0 and not sent and res['confirmed']
    assert va.day_view('2026-11-21')['cars'][0]['note'] == '유리 용접'


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
    from test_vehicle_log import _xlsx, T1, LEADS
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
    monkeypatch.setattr(bp, 'render_template', lambda name, **k: name)
    from test_vehicle_log import _lead
    monkeypatch.setattr(bp, '_leads', lambda: [_lead('L-1', '2026-11-21', '부천 길주로 1', '서준')])
    monkeypatch.setattr(va, 'day_off', lambda d: {'off': [{'code': 'JK', 'name': '강해', 'why': '연차'}], 'half': []})   # 휴가 DB 안 읽게
    resolve = []
    monkeypatch.setattr(bp, '_resolve_later', lambda s, e: resolve.append((s, e)))   # 행안부 부르지 않게
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': 'staff@example.com', 'permission_level': 'user'}
        j = cl.get('/api/vehicles/assign?day=2026-11-21').get_json()
        assert j['success'] and not j['can_edit'] and len(j['cars']) == 12 and j['day_label'] == '11/21(토)'   # 11대 + 자차 칸(11-23)
        assert j['visitors'] == [{'code': 'TH', 'name': '서준', 'visits': 1,
                                  'places': [{'label': '부천', 'addr': '부천 길주로 1'}]}]
        # 지난 날(슬랙으로 정한 날)은 시트 담당자 코스로 — 방문 일정 화면과 같게(11-23)
        assert [(t['label'], t['members'], t['visits']) for t in j['teams']] == [('코스 1', ['TH'], 1)]
        assert j['courses']['source'] == 'sheet'
        assert j['resolving'] and resolve == [('2026-11-21', '2026-11-21')]   # 동은 뒤에서 찾고 화면이 다시 받아 감
        assert j['off'] == [{'code': 'JK', 'name': '강해', 'why': '연차'}] and j['half'] == []
        assert [g['cat'] for g in j['schedule']] == ['온라인 방문'] and j['schedule'][0]['items'][0]['who'] == ['서준']
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


def test_dev_preview_only_kiko(db, monkeypatch):
    """차량은 개발 중 — 부장님만 들어가고, 다른 사람은 '아직 개발 중' 안내·API 403(11-21). 스케줄 근무는 11-25 엶.
    차량 배정 화면만 배정 담당(SJ·JW)에게 먼저(11-24)."""
    import dashboard.auth as auth
    from dashboard.blueprints import gw_satoff as so_bp
    from dashboard.blueprints import gw_vehicles as bp
    from dashboard.blueprints import hr as hr_bp
    kiko = 'kiko@itg-aircon.com'
    monkeypatch.setattr(gw, 'DEV_PREVIEW_USERS', (kiko,))
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    monkeypatch.setattr(bp, 'render_template', lambda name, **k: name)
    monkeypatch.setattr(hr_bp, 'render_template', lambda name, **k: name)
    monkeypatch.setattr(bp, '_leads', lambda: [])
    monkeypatch.setattr(va, 'day_off', lambda d: {'off': [], 'half': []})
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    app.register_blueprint(so_bp.gw_satoff_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': ASSIGNER, 'permission_level': 'admin'}
        # 차량 배정 화면만 배정 담당에게 먼저 열림(11-24 — JW 가 방문 일정 확정 뒤 바로 차를 넣음). 다른 차량 화면은 그대로 개발 중
        assert cl.get('/vehicles/assign').get_data(as_text=True) == 'gw/vehicle_assign.html'
        assert cl.get('/api/vehicles/assign?day=2026-11-21').get_json()['success']
        # 스케줄 근무는 11-25 개발 중 문을 뗐다(시공·미화 빼고 모두) — 개발 중 안내가 아님. 볼 수 있는 사람은 test_sat_off 에서
        assert cl.get('/leave/saturday').get_data(as_text=True) != 'gw/dev_preview.html'
        r = cl.get('/api/satoff')
        assert '개발 중' not in (r.get_json() or {}).get('message', '')
        # 운행일지는 경영지원에게 열림(11-24 "메뉴는 다 오픈") — 배정 담당(경영지원 아님)은 경영지원만 안내
        assert cl.get('/vehicles').status_code == 302
        r = cl.get('/api/vehicles')
        assert r.status_code == 403 and '경영지원' in r.get_json()['message']
        assert cl.post('/api/satoff/toggle', json={}).status_code == 403
        with cl.session_transaction() as s:
            s['user'] = {'email': kiko, 'permission_level': 'super_admin'}
        assert cl.get('/vehicles/assign').get_data(as_text=True) == 'gw/vehicle_assign.html'
        assert cl.get('/api/vehicles/assign?day=2026-11-21').get_json()['success']
    # 배정 담당은 차량 배정 화면을 열 수 있으니(11-24) 개발 중 명단과 상관없이 '차량 배정해 주세요' DM 을 받는다
    monkeypatch.setattr(hr, 'find_by_email', lambda e: {'slack_id': 'USJ'})
    assert va.request_assignment('2026-11-21') == 1


def test_menu():
    """차량 배정은 차량 영역에서 빠지고 방문 일정 탭으로만(11-24 부장님) — 그 화면에서는 영업 > 방문 일정이 켜진다."""
    items = {it['href']: it for a in gw.AREAS for it in a['items']}
    assert '/vehicles/assign' not in items and items['/vehicles'].get('hr_admin')
    nav = gw.build_nav('/vehicles/assign', 'admin', is_hr_admin=True, planner=True)
    assert [(a['label'], [c['label'] for c in a['children'] if c['active']]) for a in nav['areas'] if a['active']] == [('영업', ['방문 일정'])]
    assert [a['key'] for a in gw.build_nav('/vehicles', 'admin', is_hr_admin=True, planner=True)['areas'] if a['active']] == ['vehicle']


def test_latest_note_skips_header_only_rounds():
    """상담 내용은 마지막 회차만 — '방문 완료' 머리말만 붙은 회차는 건너뜀(11-21 실제 모양)."""
    assert va._latest_note('데코타일 2장 픽업 / 비용 없음 ──── [11.17 09:18 MJ · 방문 완료]') == '데코타일 2장 픽업 / 비용 없음'
    assert va._latest_note('[11.10 10:00 SJ · 방문 예약] 첫 상담 ─── [11.12 11:00 YM · 재상담] 두번째') == '두번째'
    assert va._latest_note('') == ''


def test_online_duty_saved_at_schedule_confirm(db):
    """/일정확정 때 캔버스2 '온라인' 당번을 그날에 저장 — 배정 명단에서 뺀다(11-21 부장님). 다시 확정하면 덮어씀."""
    assert va.online_duty('2026-11-23') == []                                      # 확정 전
    assert va.record_online_duty('2026-11-23', ['JK', 'SD', 'JK', 'MJ'], {'MJ': '오전'})
    assert va.online_duty('2026-11-23') == [{'code': 'JK', 'name': '강해', 'shift': ''},
                                            {'code': 'MJ', 'name': '박다솜', 'shift': '오전'}]   # 명단 밖(SD)은 뺌
    va.save_day('2026-11-23', [{'plate': '237허5358', 'people': ['TH']}], ASSIGNER, confirm=True)
    assert [o['code'] for o in va.online_duty('2026-11-23')] == ['JK', 'MJ']       # 배정 확정이 당번을 지우지 않음
    va.record_online_duty('2026-11-23', ['YG'])
    assert [o['code'] for o in va.online_duty('2026-11-23')] == ['YG']
    assert va.day_view('2026-11-23')['confirmed_at']                               # 당번 저장이 확정을 지우지 않음


def test_halves_saved_and_validated(db):
    """오전·오후 나누기(11-22) — 같은 사람은 반나절에 한 차, 오전 레이·오후 포터 가능."""
    with pytest.raises(ValueError, match='오전에 두 차'):
        va.save_day('2026-11-23', [{'plate': '231허6854', 'people': ['YM']},
                                   {'plate': '817루4904', 'people': ['YM']}], ASSIGNER)
    rows = [{'plate': '231허6854', 'people': ['YM', 'LJW'], 'pm': ['LJW']},           # 오후엔 LJW 만
            {'plate': '817루4904', 'people': [], 'pm': ['YM', 'TH']},                 # 오후에만 쓰는 차
            {'plate': '16소4041', 'people': ['YG', 'JK'], 'pm': ['SJ']},
            {'plate': '03모8021', 'people': ['JW'], 'pm': ['JW']}]                    # 오전=오후 → 나누지 않음
    va.save_day('2026-11-23', rows, ASSIGNER)
    cars = {c['plate']: c for c in va.day_view('2026-11-23')['cars']}
    assert (cars['231허6854']['people'], cars['231허6854']['pm']) == (['YM', 'LJW'], ['LJW'])
    assert (cars['817루4904']['people'], cars['817루4904']['pm']) == ([], ['YM', 'TH'])
    assert cars['03모8021']['pm'] is None
    with pytest.raises(ValueError, match='오후에 두 차'):
        va.save_day('2026-11-23', rows + [{'plate': '231허6873', 'people': ['TH']}], ASSIGNER)   # TH 종일 = 오후도
    assert va.person_car('2026-11-23', 'TH')['half'] == 'pm'


def test_halves_dm_text_and_changes(db):
    sent = db
    rows = [{'plate': '231허6854', 'people': ['YM', 'LJW'], 'pm': ['LJW']},
            {'plate': '817루4904', 'people': [], 'pm': ['YM', 'TH']},
            {'plate': '03모8021', 'people': ['JW', 'MJ']}]
    va.save_day('2026-11-23', rows, ASSIGNER, confirm=True)
    ym = next(t for sid, t in sent if sid == 'UYM')
    assert '오전·오후가 달라요' in ym and '오전: *레이 6854*' in ym and '오후: *포터 4904*' in ym and '함께 서준' in ym
    jw = next(t for sid, t in sent if sid == 'UJW')
    assert '차량: *쏘렌토 8021*' in jw and '함께 타는 사람: 박다솜' in jw and '운전:' not in jw and '(운전)' not in jw   # 운전자는 배정 때 안 정함 — 저녁 DM·[운행 종료](11-22)
    sent.clear()
    va.save_day('2026-11-23', [rows[0], dict(rows[1], pm=['YM']), rows[2]], ASSIGNER)   # TH 가 오후 포터에서 빠짐
    who = {sid for sid, _ in sent}
    assert who == {'UTH', 'UYM'}                                         # 같이 타는 사람이 바뀐 YM 도
    assert '빠졌어요' in next(t for sid, t in sent if sid == 'UTH')


def test_trip_driver_by_half(db):
    from test_vehicle_log import _xlsx, LEADS
    trips = [('2026-11-23(월)', '09:10:00', '경기도 김포시 걸포동', '서울특별시 영등포구 영등포동', 30.0),
             ('2026-11-23(월)', '14:00:00', '서울특별시 영등포구 영등포동', '서울특별시 강남구 역삼동', 15.0)]
    vl.import_upload('a.xlsx', _xlsx(trips=trips), 'x')
    with vl._conn() as conn:
        conn.execute("UPDATE vehicle_trips SET plate = '16소4041'")
    va.save_day('2026-11-23', [{'plate': '16소4041', 'people': ['YG', 'JK'], 'pm': ['SJ']}], ASSIGNER)
    d = vl.vehicle_days('2026-11', '16소4041', LEADS)['days'][0]
    assert [t['driver'] for t in sorted(d['trips'], key=lambda t: t['start'])] == ['임산', '오늘']
    assert d['drivers'][:2] == ['임산', '오늘'] and d['source'] == 'assigned'


def test_seed_fix_only_untouched_rows(db):
    """11-22 고정 차 바꾸기는 처음 시드 그대로인 줄에만 — 화면에서 고친 줄은 그대로."""
    with va._conn() as conn:
        conn.execute("UPDATE vehicle_fleet SET usage = 'shared', fixed_user = '', updated_by = 'seed' WHERE plate = '53우1783'")
        conn.execute("UPDATE vehicle_fleet SET usage = 'shared', fixed_user = '', updated_by = 'boss' WHERE plate = '143고9687'")
    va._table_ready = False
    cars = {c['plate']: c for c in va.fleet()}
    assert cars['53우1783']['usage'] == va.FIXED and cars['143고9687']['usage'] == va.SHARED


def test_plan_changed_dm_only_when_cars_saved(db, monkeypatch):
    """방문 일정을 다시 확정하면(11-23) — 차량 배정을 저장해 둔 날만 배정 담당에게 DM, 저장 시각도 돌려줌."""
    sent = db
    monkeypatch.setattr(hr, 'find_by_email', lambda e: {'slack_id': 'USJ'} if e == ASSIGNER else None)
    assert va.notify_plan_changed('2026-11-24', '추가 서준') == 0 and not sent and va.last_saved_at('2026-11-24') == ''
    va.save_day('2026-11-24', [{'plate': '237허5358', 'people': ['TH']}], ASSIGNER)
    assert va.last_saved_at('2026-11-24')
    assert va.notify_plan_changed('2026-11-24', '추가 서준') == 1
    assert sent[0][0] == 'USJ' and '추가 서준' in sent[0][1] and '/vehicles/assign?day=2026-11-24' in sent[0][1]
    sent.clear()      # 다시 확정한 사람이 배정 담당이면 본인에게는 안 보냄(바로 차량 배정 탭으로, 11-24)
    assert va.notify_plan_changed('2026-11-24', '추가 서준', ASSIGNER.upper()) == 0 and not sent


def test_half_shared_car_dm(db):
    """차 한 대를 두 코스가 오전·오후로 나눠 씀(11-24) — 확정 뒤 나누면 바뀐 사람에게 'X 오전만/오후만' +
    반대 반나절에 그 차를 쓰는 사람. 같은 차 오전·오후 그대로인 사람(바뀐 게 없는 사람)에게는 안 감."""
    sent = db
    va.save_day('2026-11-25', [{'plate': '03모8021', 'people': ['MJ']}, {'plate': va.OWN, 'people': ['YG', 'TH']},
                               {'plate': '16소4041', 'people': ['JSH']}], ASSIGNER, confirm=True)
    sent.clear()
    res = va.save_day('2026-11-25', [{'plate': '03모8021', 'people': ['MJ'], 'pm': ['YG', 'TH']},
                                     {'plate': '16소4041', 'people': ['JSH']}], ASSIGNER)
    assert res['notified'] == 3
    dm = {sid: t for sid, t in sent}
    assert set(dm) == {'UMJ', 'UYG', 'UTH'}
    assert '쏘렌토 8021' in dm['UMJ'] and '오전만' in dm['UMJ'] and '오후에는' in dm['UMJ'] and '임산' in dm['UMJ']
    assert '오후만' in dm['UTH'] and '함께 타는 사람: 임산' in dm['UTH'] and '오전에는 박다솜' in dm['UTH']
    assert '배정 없음' not in dm['UYG']


def test_own_car_dm_text():
    """본인 차량 DM 은 화면 꼬리표와 같은 말(11-24)."""
    assert va._car_text(va.OWN, {}) == '*본인 차량*으로 이동'


def test_off_person_blocks_confirm(db, monkeypatch):
    """휴무(종일)인 사람을 차에 넣은 채로는 확정·알림이 안 됨(방문 일정과 같게, 11-23). 저장만은 됨, 확정 전 토요 휴무는 막지 않음."""
    off = {'off': [{'code': 'TH', 'name': '서준', 'why': '연차'},
                   {'code': 'YM', 'name': '김가람', 'why': '스케줄 근무 쉬는 날(확정 전)', 'soft': True}], 'half': []}
    monkeypatch.setattr(va, 'day_off', lambda day: off)
    rows = [{'plate': '237허5358', 'people': ['TH', 'JSH']}, {'plate': '03모8021', 'people': ['YM']}]
    assert va.save_day('2026-11-24', rows, ASSIGNER)['changed']                  # 저장만
    with pytest.raises(ValueError, match='휴무인 서준'):
        va.save_day('2026-11-24', rows, ASSIGNER, confirm=True)
    ok = [{'plate': '237허5358', 'people': ['JSH']}, {'plate': '03모8021', 'people': ['YM']}]
    assert va.save_day('2026-11-24', ok, ASSIGNER, confirm=True)['confirmed']
    with pytest.raises(ValueError, match='휴무인 서준'):                          # 확정 뒤 알리며 고칠 때도
        va.save_day('2026-11-24', rows, ASSIGNER)
    assert va.save_day('2026-11-24', rows, ASSIGNER, notify=False)['changed']   # 알리지 않고 저장은 됨


def test_car_dm_off(db, monkeypatch):
    """차량 DM 스위치를 끄면(11-23 부장님) 확정·변경·배정 요청·방문 일정 변경·저녁 운전자 묻기 DM 이 하나도 안 나가고, 저장·확정 기록은 된다."""
    from dashboard.services import vehicle_driver as vd
    sent = db
    monkeypatch.setattr(va, 'CAR_DM_ON', False)
    monkeypatch.setattr(hr, 'find_by_email', lambda e: {'slack_id': 'USJ'})
    rows = [{'plate': '237허5358', 'people': ['TH', 'JSH']}]
    res = va.save_day('2026-11-24', rows, ASSIGNER, confirm=True)
    assert res['confirmed'] and res['notified'] == 0 and res['dm_off']
    assert va.day_view('2026-11-24')['confirmed_at']
    assert va.save_day('2026-11-24', [{'plate': '237허5358', 'people': ['TH']}], ASSIGNER)['notified'] == 0
    assert va.request_assignment('2026-11-25') == 0 and va.notify_plan_changed('2026-11-24', '추가 서준') == 0
    assert vd.send_asks('2026-11-24') == {'skipped': 'car_dm_off'}
    assert not sent


def test_own_car_slot(db):
    """자차 칸(11-23 부장님: 차가 모자라면 자기 차로 이동) — 차처럼 넣고 저장, 한 사람이 두 곳이면 막고,
    운행일지·운전자 묻기·최근 탄 차에서는 빠진다. DM 은 '자차 (자기 차로 이동)', 같이 타는 사람 없음."""
    from dashboard.services import vehicle_driver as vd
    sent = db
    view = va.day_view('2026-11-24')
    assert view['cars'][-1]['plate'] == va.OWN and view['cars'][-1]['usage'] == 'own'
    rows = [{'plate': '03모8021', 'people': ['JW', 'JSH', 'LJW']}, {'plate': va.OWN, 'people': ['JH', 'YM']}]
    with pytest.raises(ValueError):
        va.save_day('2026-11-24', rows + [{'plate': '16소4041', 'people': ['JH']}], ASSIGNER)
    res = va.save_day('2026-11-24', rows, ASSIGNER, confirm=True)
    assert res['confirmed'] and set(res['changed']) == {'03모8021', va.OWN}
    assert _rows(va.day_view('2026-11-24'))[va.OWN] == ['JH', 'YM']
    own_dm = [t for sid, t in sent if sid == 'UJH'][0]
    assert '본인 차량' in own_dm and '함께' not in own_dm and '최마루' not in own_dm.split('\n', 1)[-1]
    assert va.OWN not in va.car_rows('2026-11-24') and '03모8021' in va.car_rows('2026-11-24')
    assert all(k[1] != va.OWN for k in va.assigned_halves('2026-11-24', '2026-11-24'))
    assert all(p != va.OWN for p, _ in vd._cars_to_ask_all('2026-11-24'))


def test_no_draft_on_empty_day(db):
    """빈 날은 비어서 열린다 — 최근 탄 차·전날 배정으로 미리 넣던 초안 없앰(11-23 부장님: 타는 차·인원이 매일 다름)."""
    va.save_day('2026-11-20', [{'plate': '231허6854', 'people': ['SJ', 'KMW']},
                               {'plate': '03모8021', 'people': ['JW', 'TH']}], ASSIGNER)
    teams = [{'members': ['SJ', 'KMW'], 'visits': 2, 'big': False}, {'members': ['JW', 'TH', 'YM'], 'visits': 3, 'big': True}]
    for v in (va.day_view('2026-11-21', teams), va.day_view('2026-11-21')):
        assert not _rows(v) and v['draft_kind'] is None and v['draft_from'] is None
    assert _rows(va.day_view('2026-11-20')) == {'231허6854': ['SJ', 'KMW'], '03모8021': ['JW', 'TH']}


def test_request_skipped_when_confirmer_assigns(db, monkeypatch):
    """방문 일정을 확정한 사람이 배정 담당(JW)이면 '차량 배정해 주세요' DM 안 보냄 — 바로 차량 배정 탭으로(11-24)."""
    monkeypatch.setattr(va, 'CAR_DM_ON', True)
    monkeypatch.setattr(gw, 'DEV_PREVIEW_USERS', ())
    sent = []
    monkeypatch.setattr(va, '_dm', lambda sid, text: sent.append(sid) or True)
    assert set(va._DEFAULT_ASSIGNERS) == {'jw@itg-aircon.com', 'sj@itg-aircon.com'}     # 기본 배정 담당 = SJ + JW
    monkeypatch.setattr(va, 'assigners', lambda: ['sj@example.com', 'jw@example.com'])
    assert va.request_assignment('2026-11-26', 'JW@example.com') == 0 and sent == []      # JW 가 확정 → 안 보냄


def test_day_leads_prefilter():
    """차량 배정 조회는 그날에 걸친 리드만 먼저 고른다(11-24 속도) — 한 날짜·기간·달 넘김·끝이 앞·빈칸."""
    from dashboard.blueprints.gw_vehicles import _day_leads
    L = [{'리드 No': n, '방문 예정일': v} for n, v in [
        ('a', '2026-11-25'), ('b', "'2026-11-25"), ('c', '2026-11-24~26'), ('d', '2026-11-29~12-02'),
        ('e', '2026-11-24'), ('f', '2026-11-25~20'), ('g', ''), ('h', None), ('i', '2026.11.25')]]
    assert [x['리드 No'] for x in _day_leads(L, '2026-11-25')] == ['a', 'b', 'c', 'f', 'i']
    assert [x['리드 No'] for x in _day_leads(L, '2026-12-01')] == ['d']


def test_confirm_reports_people_who_missed_dm(db, monkeypatch):
    """확정 DM 을 못 받은 사람은 이름으로 돌려준다 — 예전엔 '알릴 사람이 없었어요'로 보였음(11-24 조용한 실패 점검 3번)."""
    sent = db
    monkeypatch.setattr(va, '_dm', lambda sid, text: (sid != 'UTH') and (sent.append((sid, text)) or True))
    rows = [{'plate': '237허5358', 'people': ['TH', 'JSH']}, {'plate': '03모8021', 'people': ['JW', 'LJW']}]
    res = va.save_day('2026-11-21', rows, ASSIGNER, confirm=True)
    assert res['notified'] == 3 and len(res['failed']) == 1 and res['confirmed']
    monkeypatch.setattr(va, '_dm', lambda sid, text: False)
    res = va.save_day('2026-11-22', rows, ASSIGNER, confirm=True)
    assert res['notified'] == 0 and len(res['failed']) == 4


def test_resend_to_people_who_missed_dm(db, monkeypatch):
    """보내다 실패한 사람(retry)만 [다시 보내기]로 다시 — 슬랙 계정이 없는 사람은 failed 에만(11-24 추가분)."""
    sent = db
    down = {'UTH'}
    monkeypatch.setattr(va, '_dm', lambda sid, text: bool(sid) and sid not in down and (sent.append((sid, text)) or True))
    rows = [{'plate': '237허5358', 'people': ['TH', 'JSH']}, {'plate': '03모8021', 'people': ['김영숙']}]
    with pytest.raises(ValueError):
        va.resend('2026-11-21', ['TH'])                                   # 확정 전엔 없음
    res = va.save_day('2026-11-21', rows, ASSIGNER, confirm=True)
    assert res['retry'] == ['TH'] and set(res['failed']) == {'서준', '김영숙'}
    sent.clear()
    down.clear()
    again = va.resend('2026-11-21', res['retry'])
    assert again == {'notified': 1, 'failed': [], 'retry': []} and [s for s, _ in sent] == ['UTH']


def test_admins_can_plan_and_assign(db):
    """관리자(admin·super_admin)는 방문 일정 짜기·확정과 차량 배정을 모두(11-25 부장님 '어드민한테는 다 열어줘')."""
    from dashboard.services import visit_plan as vp
    d = hr._db()
    d.create_user('관리자', 'boss2@example.com', 'pw-Test-1234!', 'admin')
    d.create_user('편집자', 'staff2@example.com', 'pw-Test-1234!', 'editor')
    assert hr.is_admin_email('boss2@example.com') and hr.is_admin_email('BOSS2@example.com ')
    assert not hr.is_admin_email('staff2@example.com') and not hr.is_admin_email('nobody@example.com') and not hr.is_admin_email('')
    assert vp.can_plan('boss2@example.com') and va.can_assign('boss2@example.com')
    assert not vp.can_plan('staff2@example.com') and not va.can_assign('staff2@example.com')
    assert vp.can_plan('jw@itg-aircon.com')                       # 일정 담당은 그대로
