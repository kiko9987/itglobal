"""U+커넥트 기능 옮기기(2026-11-24) — API 로 운행 받기(출퇴근 구분)·운행기록부 엑셀·운전 습관·가동률·아침 알림·
실시간 위치·가까운 차·사고 기록·차량 상태·화면 권한. U+ API 는 모두 가짜 응답. 이름·주소·좌표는 모두 가짜."""
import io
from datetime import date, datetime, timedelta

import pytest

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import uplus_api as up
from dashboard.services import vehicle_api_sync as vas
from dashboard.services import vehicle_assign as va
from dashboard.services import vehicle_driver as vd
from dashboard.services import vehicle_log as vl
from dashboard.services import vehicle_manage as vm
from dashboard.services import vehicle_trip_watch as tw

ASSIGNER = 'sj@example.com'
STAFF = [{'code': c, 'name': n, 'slack_id': 'U' + c, 'dept': ''} for c, n in [
    ('YM', '김가람'), ('LSH', '이나래'), ('TH', '서준')]]
DAY = '2026-11-23'
RAY, SPARK = '231허6854', '53우1783'
CO = (37.6000, 126.7000)          # 가짜 회사
HOME_YM = (37.5000, 126.8000)     # 가짜 집
SITE = (37.4500, 127.0500)


def T(hm, day=DAY):
    return datetime.strptime(f'{day} {hm}', '%Y-%m-%d %H:%M')


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, vl, va, hr, vd, tw, vas, vm):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setattr(va, 'staff', lambda: [dict(p) for p in STAFF])
    monkeypatch.setattr(va, 'assigners', lambda: [ASSIGNER])
    monkeypatch.setattr(va, '_dm', lambda sid, text: True)
    monkeypatch.setattr(vas, 'company', lambda: CO)
    monkeypatch.setattr(vas, 'home_of', lambda code: {'YM': HOME_YM}.get(code))
    monkeypatch.setattr(vas, 'reverse', lambda lat, lon: {CO: '경기 김포시 걸포동 1', HOME_YM: '경기 고양시 일산동구 장항동 2',
                                                           SITE: '서울 강남구 역삼동 3'}.get((lat, lon), '어딘가'))
    monkeypatch.setattr(vas, 'geocode', lambda addr: SITE if '역삼' in addr else None)
    monkeypatch.setattr(up, 'enabled', lambda name='status': True)
    gw._invalidate()
    yield d
    gw._invalidate()


def _trip(start, end, km, **kw):
    t = {'start': T(start), 'end': T(end), 'odo_start': 1000.0, 'odo_end': 1000.0 + km, 'km': km, 'fuel_ml': 0.0,
         'eco': 90.0, 'safe': 80.0, 'avg_speed': 40.0, 'max_speed': 90.0, 'rapid_acc': 0, 'rapid_dec': 0, 'sudden_stop': 0,
         'rapid_turn': 0, 'rapid_u_turn': 0, 'rapid_lane': 0, 'rapid_pass': 0}
    t.update(kw)
    return t


def _fake_api(monkeypatch):
    trips = {RAY: [_trip('07:40', '08:20', 25.0),                                    # 집 → 회사 = 출퇴근
                   _trip('09:00', '10:00', 30.0, safe=60.0, max_speed=131.0, rapid_acc=2, sudden_stop=1),   # 회사 → 현장
                   _trip('16:00', '17:00', 30.0)]}                                    # 현장 → 회사
    route = {RAY: [(T('07:40'), *HOME_YM), (T('08:20'), *CO), (T('09:00'), *CO), (T('10:00'), *SITE),
                   (T('16:00'), *SITE), (T('17:00'), *CO)]}
    monkeypatch.setattr(up, 'trips', lambda plate, day: trips.get(plate, []) if day == DAY else [])
    monkeypatch.setattr(up, 'route_day', lambda plate, day: route.get(plate, []))


def test_sync_day_commute_and_logbook(db, monkeypatch):
    _fake_api(monkeypatch)
    va.save_day(DAY, [{'plate': RAY, 'people': ['YM', 'LSH']}], ASSIGNER, confirm=True)
    r = vas.sync_day(DAY)
    assert (r['trips'], r['added']) == (3, 3) and not r['errors']
    assert vas.sync_day(DAY)['added'] == 0                                        # 같은 운행 다시 안 넣음
    with vl._conn() as conn:
        rows = [dict(x) for x in conn.execute('SELECT * FROM vehicle_trips ORDER BY start_time')]
    assert [(x['start_time'], x['from_addr'][:5], x['to_addr'][:5], x['commute_km'], x['business_km']) for x in rows] == [
        ('07:40', '경기 고양', '경기 김포', 25.0, 0), ('09:00', '경기 김포', '서울 강남', 0, 30.0), ('16:00', '서울 강남', '경기 김포', 0, 30.0)]
    m = vas.metrics(DAY, DAY, RAY)
    assert m[(RAY, DAY, '09:00')]['driver'] == 'YM' and m[(RAY, DAY, '07:40')]['commute'] == 1
    # 운행기록부 — 성명은 배정 맨 앞, 합계·업무사용비율
    lb = vl.logbook_rows(RAY, '2026-11-01', '2026-11-30')
    assert [x['name'] for x in lb] == ['김가람'] * 3 and lb[0]['day'] == '2026-11-23(월)'
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(vl.logbook_xlsx([RAY], '2026-11-01', '2026-11-30', corp='가짜법인')))
    ws = wb.worksheets[0]
    assert ws['A1'].value == '업무용승용차 운행기록부' and ws['E3'].value == RAY
    last = [c.value for c in ws[ws.max_row]]
    assert last[0] == '합계' and last[5] == 85.0 and last[6] == 25.0 and last[7] == 60.0 and '100.0%' in last[8]
    # 운행일지 펼침창 — 경로 단추·급가속 등
    days = vl.vehicle_days('2026-11', RAY, [])['days']
    t9 = next(t for t in days[0]['trips'] if t['start'] == '09:00')
    assert t9['route'] and t9['max_speed'] == 131 and t9['events']['rapid_acc'] == 2


def test_habits_util_and_alerts(db, monkeypatch):
    _fake_api(monkeypatch)
    va.save_day(DAY, [{'plate': RAY, 'people': ['YM', 'LSH']}], ASSIGNER, confirm=True)
    vas.sync_day(DAY)
    h = vas.habits('2026-11-01', '2026-11-30')
    p = h['people'][0]
    assert (p['name'], p['trips'], p['km'], p['speeding'], p['acc'], p['stop'], p['max_speed']) == ('김가람', 3, 85.0, 1, 2, 1, 131)
    assert p['safe'] == round((80 * 25 + 60 * 30 + 80 * 30) / 85, 1)
    u = vas.utilization('2026-11-01', '2026-11-30')
    ray = next(c for c in u['cars'] if c['plate'] == RAY)
    assert ray['days'] == 1 and ray['trips'] == 3 and ray['hours'] == 2.7 and u['workdays'] > 0
    # 아침 알림 — 과속·배터리·신호 끊김
    vas.save_state([{'plate': SPARK, 'lat': 37.5, 'lon': 126.9, 'at': datetime.now() - timedelta(days=4), 'voltage': 11.2}])
    lines = vas.alert_items(date(2026, 11, 24))
    assert any('과속' in x and '131km/h' in x and '김가람' in x for x in lines)
    assert any('배터리 낮음' in x for x in lines) and any('단말 신호 끊김' in x for x in lines)
    monkeypatch.setattr(vas, 'ALERT_ON', False)
    assert vas.send_alerts(date(2026, 11, 24)) == {'skipped': 'off'}


def test_expiry_alert_once(db):
    vm.update_info(RAY, {'insurance_end': '2026-12-24'}, 'x')
    first = [x for x in vas.alert_items(date(2026, 11, 24)) if '보험 만기' in x]
    assert len(first) == 1 and '30일 남음' in first[0]
    assert not [x for x in vas.alert_items(date(2026, 11, 24)) if '보험 만기' in x]     # 같은 줄은 한 번만
    assert not [x for x in vas.alert_items(date(2026, 11, 25)) if '보험 만기' in x]     # 29일 전은 안 알림


def test_now_view_and_nearest(db, monkeypatch):
    monkeypatch.setattr(vas, 'refresh_state', lambda max_age_sec=0: {'cached': True})
    today = date.today().isoformat()
    va.save_day(today, [{'plate': RAY, 'people': ['YM', 'LSH']}], ASSIGNER)
    vas.save_state([{'plate': RAY, 'lat': SITE[0] + 0.01, 'lon': SITE[1], 'at': datetime.now(), 'speed': 40.0},
                    {'plate': SPARK, 'lat': CO[0], 'lon': CO[1], 'at': datetime.now() - timedelta(minutes=30), 'speed': 0.0}])
    v = vas.now_view()
    ray = next(c for c in v['cars'] if c['plate'] == RAY)
    spark = next(c for c in v['cars'] if c['plate'] == SPARK)
    assert ray['moving'] and ray['riders'] == ['김가람', '이나래'] and not spark['moving'] and spark['mins'] >= 29
    n = vas.nearest('서울 강남구 역삼동')
    assert [c['plate'] for c in n['cars'][:2]] == [RAY, SPARK] and n['cars'][0]['km'] < 2
    with pytest.raises(ValueError):
        vas.nearest('없는 주소')


def test_route_filter(db, monkeypatch):
    pts = [(T('08:59'), 1.5, 1.5), (T('09:00'), 37.1, 127.1), (T('09:30'), 37.2, 127.2), (T('10:02'), 37.3, 127.3)]
    monkeypatch.setattr(up, 'route_day', lambda plate, day: pts)
    vas._route_cache.clear()
    assert vas.route(RAY, DAY, '09:00', '10:00') == [[37.1, 127.1], [37.2, 127.2]]
    assert len(vas.route(RAY, DAY)) == 4


def test_accidents_and_state_in_overview(db):
    with pytest.raises(ValueError):
        vm.save_accident({'plate': RAY, 'day': DAY, 'kind': '모름'}, 'x')
    with pytest.raises(ValueError):
        vm.save_accident({'plate': RAY, 'day': '2099-01-01', 'kind': '차대차'}, 'x')
    a = vm.save_accident({'plate': RAY, 'day': DAY, 'tm': '14:20', 'kind': '차대차', 'driver': '김가람', 'place': '가짜로 1'}, 'x')
    assert vm.accidents(RAY)[0]['status'] == '처리 중'
    vm.save_accident({'plate': RAY, 'day': DAY, 'kind': '차대차', 'status': '처리 끝'}, 'x', a['id'])
    assert vm.accidents(RAY)[0]['status'] == '처리 끝'
    vas.save_state([{'plate': RAY, 'lat': 37.5, 'lon': 126.9, 'at': datetime.now(), 'voltage': 11.0, 'fuel': 40.0, 'odometer': 51837.0}])
    ov = vm.overview()
    car = next(c for c in ov['cars'] if c['plate'] == RAY)
    assert car['state']['issues'] == ['배터리 낮음'] and car['warn'] and car['km'] == 51837.0 and ov['open_accidents'] == 0
    vm.delete_accident(a['id'])
    assert vm.accidents(RAY) == []


def test_uplus_trips_merge(monkeypatch):
    monkeypatch.undo()
    monkeypatch.setenv('UPLUS_SECURE_KEY', 'test-key')
    for n in ('ECO', 'SAFE'):
        monkeypatch.setenv(f'UPLUS_APP_{n}', '7')
    monkeypatch.setattr(up, '_tokens', {})

    def fake_post(path, body):
        if path == '/v1/session':
            return {'extensions': {'code': '20000000', 'data': [{'token': 't'}]}}
        if path == up.APIS['eco']:
            return {'extensions': {'code': '20000000', 'totalRecord': 3, 'data': [
                {'use_start_dt': '2026-11-23T09:00:00.000+00:00', 'use_end_dt': '2026-11-23T10:00:00.000+00:00',
                 'before_odometer': 100.0, 'after_odometer': 130.0, 'mileage': 30.0, 'echo_driving_index': 95.0, 'max_speed': 80},
                {'use_start_dt': '2026-11-23T11:00:00.000+00:00', 'use_end_dt': '2026-11-23T11:00:00.000+00:00',
                 'before_odometer': 130.0, 'after_odometer': 130.0, 'mileage': 0.0},                 # 시동만 — 뺌
                {'use_start_dt': '2026-11-22T23:00:00.000+00:00', 'use_end_dt': '2026-11-23T00:10:00.000+00:00', 'mileage': 5.0}]}}
        if path == up.APIS['safe']:
            return {'extensions': {'code': '20000000', 'totalRecord': 1, 'data': [
                {'use_start_dt': '2026-11-23T09:00:00.000+00:00', 'safe_driving_index': 70.0, 'rapid_acceleration': 3, 'max_speed': 120}]}}
        raise AssertionError(path)
    monkeypatch.setattr(up, '_post', fake_post)
    out = up.trips(RAY, DAY)
    assert len(out) == 1
    t = out[0]
    assert (t['start'], t['km'], t['eco'], t['safe'], t['rapid_acc'], t['max_speed']) == (T('09:00'), 30.0, 95.0, 70.0, 3, 120)
    monkeypatch.setenv('UPLUS_TIME_UTC', '1')
    assert up.parse_time('2026-11-23T00:30:00.000+00:00') == datetime(2026, 11, 23, 9, 30)


def test_now_menu_only_after_keys(monkeypatch):
    monkeypatch.setattr(gw, '_uplus_on', lambda: False)
    assert [a['key'] for a in gw.build_nav('/home', 'field', field=True)['areas']] == ['leave', 'hr']
    monkeypatch.setattr(gw, '_uplus_on', lambda: True)
    nav = gw.build_nav('/vehicles/now', 'field', field=True)
    assert [a['key'] for a in nav['areas']] == ['leave', 'hr', 'vehicle'] and [i['label'] for i in nav['items']] == ['실시간 위치']
    assert [a['label'] for a in gw.build_nav('/home', 'editor')['areas']][-1] == '차량'


def test_now_page_open_to_all_and_field(db, monkeypatch):
    import dashboard.auth as auth
    from flask import Flask
    from dashboard.blueprints import gw_vehicles as bp
    from dashboard.blueprints.groupware import field_path_allowed
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    monkeypatch.setattr(gw, 'DEV_PREVIEW_USERS', ('boss@example.com',))     # 개발 중 명단이 있어도
    monkeypatch.setattr(vas, 'refresh_state', lambda max_age_sec=0: {'cached': True})
    app = Flask(__name__, template_folder='../../dashboard/templates')
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': 'staff@example.com', 'permission_level': 'user'}
        j = cl.get('/api/vehicles/now').get_json()
        assert j['success'] and 'cars' in j
        assert cl.get('/api/vehicles/insights').status_code in (302, 403)      # 운전 습관은 경영지원·개발 중 명단만
        assert cl.get('/api/vehicles/logbook.xlsx').status_code in (302, 403)
    assert field_path_allowed('GET', '/vehicles/now') and field_path_allowed('GET', '/api/vehicles/now/near')
    assert not field_path_allowed('GET', '/api/vehicles/insights')
