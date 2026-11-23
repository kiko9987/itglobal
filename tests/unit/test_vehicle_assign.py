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
    from tests.unit.test_vehicle_log import _lead
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


def test_team_draft_uses_recent_car(db):
    va.save_day('2026-11-20', [{'plate': '231허6854', 'people': ['SJ']},
                               {'plate': '03모8021', 'people': ['TH']}], ASSIGNER)
    teams = va.teams('2026-11-21', _leads_1121())
    v = va.day_view('2026-11-21', teams)
    assert v['draft_kind'] == 'teams'
    assert _rows(v) == {'231허6854': ['SJ', 'KMW'], '03모8021': ['TH', 'LJW']}   # 탄 적 없는 팀(JK·YM)은 비워 둠
    assert va.day_view('2026-11-21')['draft_kind'] == 'prev'                     # 팀 정보 없으면 전날 배정


def test_visitors_with_visits_that_day(db):
    from tests.unit.test_vehicle_log import _lead
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


def test_visitor_places_use_found_dong(db):
    """갈 곳은 동까지 — 찾아 둔 동이 있으면 '용인 풍덕천동', 같은 주소 두 번은 한 곳(11-21 부장님)."""
    from tests.unit.test_vehicle_log import _lead
    with vl._conn() as conn:
        conn.execute('INSERT INTO vehicle_addr_dong (addr, dong, checked_at) VALUES (?,?,?)',
                     ('용인 수지구 심곡로 11', '풍덕천동', '2026-11-21 09:00'))
    leads = [_lead('L-1', '2026-11-21', '용인 수지구 심곡로 11', '김가람'),
             _lead('L-2', '2026-11-21', '용인 수지구 심곡로 11', '김가람')]
    v = va.visitors('2026-11-21', leads)
    assert v[0]['visits'] == 2 and v[0]['places'] == [{'label': '용인 풍덕천동', 'addr': '용인 수지구 심곡로 11'}]
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
                          {'code': 'MJ', 'name': '박다솜', 'why': '토요 휴무', 'soft': False}]
    assert sat['half'] == [{'code': 'LSH', 'name': '이나래', 'why': '오전 반차'}]
    monkeypatch.setattr(sat_off, 'month_row', lambda m: {'status': sat_off.OPEN})
    assert va.day_off('2026-11-21')['off'][1] == {'code': 'MJ', 'name': '박다솜', 'why': '토요 휴무(확정 전)', 'soft': True}
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
    assert partner[0]['marks'] == ['공사확정'] and partner[0]['platform'] == '거래처' and partner[0]['who'] == ['서준', '한별'] and partner[0]['date'] == '11월 20~22일'
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
        assert j['success'] and not j['can_edit'] and len(j['cars']) == 11 and j['day_label'] == '11/21(토)'
        assert j['visitors'] == [{'code': 'TH', 'name': '서준', 'visits': 1,
                                  'places': [{'label': '부천', 'addr': '부천 길주로 1'}]}]
        assert j['teams'] == [{'members': ['TH'], 'visits': 1, 'big': False}]
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
    """토요일 휴무·차량은 개발 중 — 부장님만 들어가고, 다른 사람은 '아직 개발 중' 안내·API 403(11-21)."""
    import dashboard.auth as auth
    from dashboard.blueprints import gw_satoff as so_bp
    from dashboard.blueprints import gw_vehicles as bp
    from dashboard.blueprints import hr as hr_bp
    kiko = 'kiko@itg-aircon.com'
    monkeypatch.setattr(gw, 'DEV_PREVIEW_USERS', (kiko,))
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    monkeypatch.setattr(bp, 'render_template', lambda name: name)
    monkeypatch.setattr(hr_bp, 'render_template', lambda name: name)
    monkeypatch.setattr(bp, '_leads', lambda: [])
    monkeypatch.setattr(va, 'day_off', lambda d: {'off': [], 'half': []})
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    app.register_blueprint(so_bp.gw_satoff_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': ASSIGNER, 'permission_level': 'admin'}
        for path in ('/vehicles/assign', '/vehicles', '/leave/saturday'):
            assert cl.get(path).get_data(as_text=True) == 'gw/dev_preview.html', path
        for path in ('/api/vehicles/assign', '/api/vehicles', '/api/satoff'):
            r = cl.get(path)
            assert r.status_code == 403 and '개발 중' in r.get_json()['message'], path
        assert cl.post('/api/satoff/toggle', json={}).status_code == 403
        with cl.session_transaction() as s:
            s['user'] = {'email': kiko, 'permission_level': 'super_admin'}
        assert cl.get('/vehicles/assign').get_data(as_text=True) == 'gw/vehicle_assign.html'
        assert cl.get('/api/vehicles/assign?day=2026-11-21').get_json()['success']
    # 배정 담당자(SJ)가 못 여는 동안은 '차량 배정해 주세요' DM 도 보내지 않는다
    monkeypatch.setattr(hr, 'find_by_email', lambda e: {'slack_id': 'USJ'})
    assert va.request_assignment('2026-11-21') == 0 and db == []
    monkeypatch.setattr(gw, 'DEV_PREVIEW_USERS', ())                                  # 모두에게 열면 다시 보냄
    assert va.request_assignment('2026-11-21') == 1


def test_menu():
    items = {it['href']: it for a in gw.AREAS for it in a['items']}
    assert items['/vehicles/assign'].get('field') and not items['/vehicles/assign'].get('hr_admin')
    assert items['/vehicles'].get('hr_admin')


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


# ── 11-22 부장님: 큰 팀은 짝 기록대로 나눠 초안, 한 차를 오전·오후 다른 사람이 운전 ──
def test_split_big_team_by_pairs(db):
    pairs = {frozenset(('JW', 'TH')): 3, frozenset(('YM', 'LJW')): 2}
    assert va.split_team(['JW', 'TH', 'YM', 'LJW'], pairs) == [['JW', 'TH'], ['YM', 'LJW']]
    assert va.split_team(['JW', 'TH', 'YM', 'LJW'], {}) is None                # 짝 기록 없으면 몰아넣지 않음
    assert va.split_team(['MJ', 'KMW'], {}) == [['MJ', 'KMW']]                 # 2명은 그대로
    assert va.split_team(['YG', 'JK', 'LSH'], {frozenset(('YG', 'JK')): 1}) == [['YG', 'JK'], ['LSH']]


def test_team_draft_splits_with_history(db):
    va.save_day('2026-11-20', [{'plate': '03모8021', 'people': ['JW', 'TH']},
                               {'plate': '231허6854', 'people': ['YM', 'LJW']}], ASSIGNER)
    teams = [{'members': ['JW', 'TH', 'YM', 'LJW'], 'visits': 4, 'big': True}]
    v = va.day_view('2026-11-23', teams)
    assert _rows(v) == {'03모8021': ['JW', 'TH'], '231허6854': ['YM', 'LJW']}
    assert not _rows(va.day_view('2026-11-23', [{'members': ['MJ', 'KMW', 'SJ'], 'visits': 1, 'big': True}]))


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
    from tests.unit.test_vehicle_log import _xlsx, LEADS
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


def test_off_person_blocks_confirm(db, monkeypatch):
    """휴무(종일)인 사람을 차에 넣은 채로는 확정·알림이 안 됨(방문 일정과 같게, 11-23). 저장만은 됨, 확정 전 토요 휴무는 막지 않음."""
    off = {'off': [{'code': 'TH', 'name': '서준', 'why': '연차'},
                   {'code': 'YM', 'name': '김가람', 'why': '토요 휴무(확정 전)', 'soft': True}], 'half': []}
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
