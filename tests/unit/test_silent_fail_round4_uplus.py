"""11-24 조용한 실패 점검 4차 — U+커넥트 차량 연동(키 들어온 뒤 조용히 틀어질 곳).

① 실시간 위치: U+ 조회 실패를 지금 위치처럼 보이지 않게(error·last_ok), 실패 뒤 잠깐 다시 안 부름, 연속 실패 ERROR
② 정해진 시각 작업: 결과의 error/errors → ERROR, 연속 실패·마지막 재시도 날이면 force_alert, 단말 조회 예외 ERROR
③ 주소↔좌표: 카카오 오류와 '결과 없음' 구별(오류는 저장·낱말 떼기 안 함, 결과 없음은 7일 기한), 회사 좌표 고정값
④ 운행 표: U+ 로 넣은 줄은 다시 받을 때 주소·출퇴근을 다시 맞춤, 엑셀로 올린 줄은 그대로
U+·카카오 API 는 모두 가짜. 이름·주소·좌표도 가짜."""
import logging
from datetime import date, datetime, timedelta

import pytest

from dashboard.services import address_resolver as ar
from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import sync_scheduler as sched
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
RAY = '231허6854'
CO = (37.6000, 126.7000)
HOME_YM = (37.5000, 126.8000)
SITE = (37.4500, 127.0500)


def T(hm, day=DAY):
    return datetime.strptime(f'{day} {hm}', '%Y-%m-%d %H:%M')


@pytest.fixture
def raw(tmp_path, monkeypatch):
    """DB 만 가짜로 — 카카오·회사·집 좌표 함수는 진짜(그 아래 카카오 호출만 테스트마다 가짜로)."""
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, vl, va, hr, vd, tw, vas, vm):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setattr(va, 'staff', lambda: [dict(p) for p in STAFF])
    monkeypatch.setattr(va, 'assigners', lambda: [ASSIGNER])
    monkeypatch.setattr(va, '_dm', lambda sid, text: True)
    monkeypatch.setattr(up, 'enabled', lambda name='status': True)
    monkeypatch.setattr(vas, '_state_fail', {'at': 0.0, 'count': 0, 'msg': ''})
    monkeypatch.setattr(ar, '_kakao_key', lambda: 'fake-key')
    monkeypatch.setattr(ar, '_kakao_get_json', lambda url: (_ for _ in ()).throw(AssertionError('카카오 가짜 안 씀')))
    gw._invalidate()
    yield d
    gw._invalidate()


@pytest.fixture
def db(raw, monkeypatch):
    """좌표 함수까지 가짜(기존 test_vehicle_api_sync 와 같은 판)."""
    monkeypatch.setattr(vas, 'company', lambda: CO)
    monkeypatch.setattr(vas, 'home_of', lambda code: {'YM': HOME_YM}.get(code))
    return raw


def _errors(caplog, needle=''):
    return [r for r in caplog.records if r.levelno >= logging.ERROR and needle in r.getMessage()]


# ── ① 실시간 위치 ────────────────────────────────────────────────────────
def test_now_view_shows_failure_and_backs_off(db, monkeypatch, caplog):
    va.save_day(date.today().isoformat(), [{'plate': RAY, 'people': ['YM']}], ASSIGNER)
    old = datetime.now() - timedelta(minutes=10)
    vas.save_state([{'plate': RAY, 'lat': SITE[0], 'lon': SITE[1], 'at': old, 'speed': 0.0}], now=old)
    monkeypatch.setattr(vas, 'reverse', lambda lat, lon: '서울 강남구 역삼동 3')
    calls = []

    def boom():
        calls.append(1)
        raise up.UplusError('car_status HTTP 502')
    monkeypatch.setattr(up, 'car_status', boom)
    v = vas.now_view()
    assert v['error'] and 'U+' in v['error'] and v['last_ok'] == old.strftime('%H:%M')
    assert next(c for c in v['cars'] if c['plate'] == RAY)['lat'] == SITE[0]                     # 옛 위치는 그대로 보여 주되, 실패라고 알림
    # 바로 다시 불러도(30초마다 화면) U+ 를 다시 부르지 않는다
    v2 = vas.now_view()
    assert len(calls) == 1 and v2['error']
    # 기다린 뒤엔 다시 부르고, 3번 연속 실패면 ERROR(force_alert)
    for _ in range(2):
        vas._state_fail['at'] -= vas.STATE_RETRY_SEC + 1
        vas.now_view()
    assert len(calls) == 3
    errs = _errors(caplog, '연속 실패')
    assert len(errs) == 1 and getattr(errs[0], 'force_alert', False) is True
    # 다시 되면 실패 표시가 사라지고 횟수도 처음으로
    monkeypatch.setattr(up, 'car_status', lambda: [{'plate': RAY, 'lat': CO[0], 'lon': CO[1], 'at': datetime.now()}])
    vas._state_fail['at'] -= vas.STATE_RETRY_SEC + 1
    v3 = vas.now_view()
    assert v3['error'] == '' and v3['last_ok'] == datetime.now().strftime('%H:%M') and vas._state_fail['count'] == 0


def test_now_view_before_approval_has_no_error(db, monkeypatch):
    monkeypatch.setattr(up, 'enabled', lambda name='status': False)
    v = vas.now_view()
    assert v['enabled'] is False and v['error'] == '' and v['company'] == CO


def test_scheduler_refresh_ignores_backoff(db, monkeypatch):
    """정해진 시각 조회(max_age_sec=0)는 방금 화면에서 실패했어도 부른다."""
    calls = []
    monkeypatch.setattr(up, 'car_status', lambda: calls.append(1) or [])
    vas._state_fail.update(at=__import__('time').time(), count=1, msg='x')
    assert vas.refresh_state() == {'saved': 0} and calls == [1]


# ── ② 정해진 시각 작업 ──────────────────────────────────────────────────────
def test_scheduler_wrapper_errors_and_streak(monkeypatch, caplog):
    monkeypatch.setattr(sched, '_vehicle_api_fails', {})
    results = [{'error': 'HTTP 502'}, {'error': 'HTTP 502'}, {'devices': 3}, {'skipped': 'no_api'}]
    monkeypatch.setattr(vas, 'check_devices', lambda: results.pop(0))
    with caplog.at_level(logging.INFO):
        sched._safe_vehicle_api_devices()
        sched._safe_vehicle_api_devices()
    errs = _errors(caplog, 'U+ 단말 상태 실패')
    assert len(errs) == 2
    assert getattr(errs[0], 'force_alert', False) is False and getattr(errs[1], 'force_alert', False) is True
    sched._safe_vehicle_api_devices()                           # 성공하면 연속 횟수 처음으로
    assert sched._vehicle_api_fails == {}
    caplog.clear()
    sched._safe_vehicle_api_devices()                           # 키 없음은 조용히
    assert not caplog.records


def test_scheduler_wrapper_exception_and_urgent(monkeypatch, caplog):
    monkeypatch.setattr(sched, '_vehicle_api_fails', {})
    monkeypatch.setattr(vas, 'sync_recent', lambda: {'days': [], 'errors': ['2026-11-21 RAY: HTTP 500'],
                                                     'error_days': 1, 'urgent': True})
    sched._safe_vehicle_api_trips()
    e = _errors(caplog, 'U+ 운행 기록 실패')
    assert e and getattr(e[0], 'force_alert', False) is True and '오류 남은 날 1일' in e[0].getMessage()

    def raise_():
        raise RuntimeError('DB 잠김')
    monkeypatch.setattr(vas, 'refresh_state', raise_)
    sched._safe_vehicle_api_state()
    e = _errors(caplog, 'U+ 차량 상태 실패')
    assert e and e[0].exc_info


def test_check_devices_exception_logged(db, monkeypatch, caplog):
    monkeypatch.setattr(up, 'device_status', lambda day: (_ for _ in ()).throw(up.UplusError('인증 실패')))
    assert vas.check_devices('2026-11-23') == {'error': '인증 실패'}
    assert _errors(caplog, '단말 상태 조회 실패')


def test_sync_recent_reports_error_days_and_last_chance(db, monkeypatch):
    today = date(2026, 11, 24)
    oldest = (today - timedelta(days=vas.SYNC_BACK_DAYS)).isoformat()

    def trips(plate, day):
        if day == oldest:
            raise up.UplusError('eco HTTP 503')
        return []
    monkeypatch.setattr(up, 'trips', trips)
    r = vas.sync_recent(today)
    assert r['error_days'] == 1 and r['urgent'] is True
    assert any('eco HTTP 503' in e for e in r['errors']) and any('마지막 자동 재시도' in e for e in r['errors'])
    # 다음 날 다시 받아 되면 그 날은 오류에서 빠짐(창 안의 날만 셈)
    monkeypatch.setattr(up, 'trips', lambda plate, day: [])
    r2 = vas.sync_recent(today)
    assert r2['error_days'] == 0 and 'errors' not in r2 and 'urgent' not in r2


# ── ③ 주소 ↔ 좌표 ──────────────────────────────────────────────────────────
def _kakao(monkeypatch, answer):
    """answer(query) → dict(정상 응답)·None(키/인증 실패)·예외(일시 오류). 부른 검색어를 돌려줌."""
    seen = []

    def fake(url):
        import urllib.parse
        q = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
        key = q.get('query', [''])[0] or f"{q.get('y', [''])[0]},{q.get('x', [''])[0]}"
        seen.append(key)
        return answer(key)
    monkeypatch.setattr(ar, '_kakao_get_json', fake)
    return seen


def _cached(key):
    with vas._conn() as conn:
        r = conn.execute('SELECT v FROM vehicle_geo_cache WHERE k = ?', (key,)).fetchone()
    return r['v'] if r else None


def test_geocode_transient_error_not_saved_no_word_strip(raw, monkeypatch):
    def answer(q):
        raise ar._KakaoTransientError()
    seen = _kakao(monkeypatch, answer)
    assert vas.geocode('경기도 김포시 가짜로 1 101동') is None
    assert seen == ['경기도 김포시 가짜로 1 101동']              # 낱말 떼기 안 함 → '경기도 김포시' 좌표 저장 안 됨
    assert _cached('g:경기도 김포시 가짜로 1 101동') is None
    assert vas._geo_state.err


def test_geocode_auth_fail_and_no_key_not_saved(raw, monkeypatch):
    seen = _kakao(monkeypatch, lambda q: None)                  # 401/403 → None
    assert vas.geocode('경기도 김포시 가짜로 2') is None and len(seen) == 1
    assert _cached('g:경기도 김포시 가짜로 2') is None
    monkeypatch.setattr(ar, '_kakao_key', lambda: '')
    assert vas.geocode('경기도 김포시 가짜로 2') is None and len(seen) == 1     # 키 없으면 부르지도 않음
    assert _cached('g:경기도 김포시 가짜로 2') is None


def test_geocode_no_result_saved_with_ttl_and_strip_still_works(raw, monkeypatch):
    seen = _kakao(monkeypatch, lambda q: {'documents': [{'x': '126.8', 'y': '37.5'}]} if q == '경기도 고양시 가짜로 3'
                  else {'documents': []})
    assert vas.geocode('경기도 고양시 가짜로 3 101동 202호') == (37.5, 126.8)      # 결과 없음이면 끝 낱말 떼기는 그대로
    assert seen[-1] == '경기도 고양시 가짜로 3'
    assert vas.geocode('없는시 없는로') is None
    assert _cached('g:없는시 없는로') == ''
    n = len(seen)
    assert vas.geocode('없는시 없는로') is None and len(seen) == n                 # 기한 안이면 저장된 '없음'
    with vas._conn() as conn:
        conn.execute('UPDATE vehicle_geo_cache SET at = ? WHERE k = ?',
                     ((datetime.now() - timedelta(days=vas.GEO_MISS_TTL_DAYS + 1)).strftime('%Y-%m-%d %H:%M:%S'),
                      'g:없는시 없는로'))
    vas.geocode('없는시 없는로')
    assert len(seen) == n + 1                                                       # 기한 지나면 다시 찾음


def test_reverse_auth_fail_not_saved(raw, monkeypatch):
    seen = _kakao(monkeypatch, lambda q: None)
    assert vas.reverse(37.1, 127.1) == '' and _cached('r:37.1000,127.1000') is None
    _kakao(monkeypatch, lambda q: {'documents': [{'address': {'address_name': '경기 가짜시 가짜동 1'}}]})
    assert vas.reverse(37.1, 127.1) == '경기 가짜시 가짜동 1' and _cached('r:37.1000,127.1000') == '경기 가짜시 가짜동 1'
    _kakao(monkeypatch, lambda q: (_ for _ in ()).throw(ar._KakaoTransientError()))
    assert vas.reverse(37.2, 127.2) == '' and _cached('r:37.2000,127.2000') is None
    assert seen


def test_company_uses_fixed_point_and_errors_when_lookup_fails(raw, monkeypatch, caplog):
    assert vas.company() == vas.COMPANY_PT                      # 카카오를 안 불러도 됨(가짜는 부르면 AssertionError)
    monkeypatch.setattr(vas, 'COMPANY_PT', None)
    _kakao(monkeypatch, lambda q: (_ for _ in ()).throw(ar._KakaoTransientError()))
    assert vas.company() is None
    assert _errors(caplog, '회사 좌표')


def test_home_of_error_not_frozen(raw, monkeypatch):
    monkeypatch.setattr(va, '_who_map', lambda: {'YM': {'name': '김가람'}})
    monkeypatch.setattr(hr, '_rows', lambda q, p=(): [{'address': '경기도 고양시 가짜로 9'}])
    _kakao(monkeypatch, lambda q: (_ for _ in ()).throw(ar._KakaoTransientError()))
    assert vas.home_of('YM') is None
    with vas._conn() as conn:
        assert conn.execute('SELECT COUNT(*) FROM vehicle_home_geo').fetchone()[0] == 0   # '집 없음'으로 굳지 않음
    _kakao(monkeypatch, lambda q: {'documents': [{'x': '126.8', 'y': '37.5'}]})
    assert vas.home_of('YM') == (37.5, 126.8)


def test_nearest_kakao_error_is_not_not_found(raw, monkeypatch):
    _kakao(monkeypatch, lambda q: None)
    with pytest.raises(RuntimeError):
        vas.nearest('서울 강남구 역삼동 1')
    _kakao(monkeypatch, lambda q: {'documents': []})
    with pytest.raises(ValueError):
        vas.nearest('서울 강남구 없는동 1')


# ── ④ 운행 표 다시 맞추기 ────────────────────────────────────────────────────
def _trip(start, end, km):
    return {'start': T(start), 'end': T(end), 'odo_start': 1000.0, 'odo_end': 1000.0 + km, 'km': km, 'fuel_ml': 0.0,
            'eco': 90.0, 'safe': 80.0, 'avg_speed': 40.0, 'max_speed': 90.0, 'rapid_acc': 0, 'rapid_dec': 0,
            'sudden_stop': 0, 'rapid_turn': 0, 'rapid_u_turn': 0, 'rapid_lane': 0, 'rapid_pass': 0}


def _rows():
    with vl._conn() as conn:
        return {r['start_time']: dict(r) for r in conn.execute('SELECT * FROM vehicle_trips ORDER BY start_time')}


def test_resync_fixes_api_rows_but_not_uploaded(db, monkeypatch):
    trips = [_trip('07:40', '08:20', 25.0), _trip('09:00', '10:00', 30.0)]
    route = [(T('07:40'), *HOME_YM), (T('08:20'), *CO), (T('09:00'), *CO), (T('10:00'), *SITE)]
    monkeypatch.setattr(up, 'trips', lambda plate, day: trips if plate == RAY else [])
    monkeypatch.setattr(up, 'route_day', lambda plate, day: route)
    # 엑셀로 올린 줄(09:00) — 다시 받아도 그대로여야 함
    with vl._conn() as conn:
        conn.execute("INSERT INTO vehicle_trips (plate, day, start_time, end_time, km, from_addr, to_addr, commute_km, "
                     "business_km, note, upload_id) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                     (RAY, DAY, '09:00', '10:00', 29.0, '엑셀 출발', '', 0, 29.0, '엑셀', 7))
    # 처음: 카카오 실패로 주소 없음, 배정 전이라 탄 사람 몰라 출퇴근 0
    monkeypatch.setattr(vas, 'reverse', lambda lat, lon: '')
    r1 = vas.sync_day(DAY)
    assert r1['no_addr'] == 2 and r1['notes'] and r1['added'] == 1
    row = _rows()['07:40']
    assert (row['from_addr'], row['commute_km'], row['business_km']) == ('', 0, 25.0)
    # 다시 받을 때: 주소가 구해지고 배정도 생김 → 고쳐짐
    va.save_day(DAY, [{'plate': RAY, 'people': ['YM']}], ASSIGNER, confirm=True)
    monkeypatch.setattr(vas, 'reverse', lambda lat, lon: {CO: '경기 김포시 걸포동 1', HOME_YM: '경기 고양시 장항동 2',
                                                           SITE: '서울 강남구 역삼동 3'}[(lat, lon)])
    r2 = vas.sync_day(DAY)
    assert r2['added'] == 0 and r2['fixed'] == 1 and r2['no_addr'] == 0
    rows = _rows()
    assert (rows['07:40']['from_addr'], rows['07:40']['to_addr'], rows['07:40']['commute_km'],
            rows['07:40']['business_km']) == ('경기 고양시 장항동 2', '경기 김포시 걸포동 1', 25.0, 0)
    assert (rows['09:00']['from_addr'], rows['09:00']['km'], rows['09:00']['note']) == ('엑셀 출발', 29.0, '엑셀')
    # 세 번째: 경로를 못 받았어도(좌표 없음) 맞는 출퇴근·주소를 0·빈 칸으로 덮지 않음
    monkeypatch.setattr(up, 'route_day', lambda plate, day: [])
    vas.sync_day(DAY)
    row = _rows()['07:40']
    assert (row['from_addr'], row['commute_km']) == ('경기 고양시 장항동 2', 25.0)
