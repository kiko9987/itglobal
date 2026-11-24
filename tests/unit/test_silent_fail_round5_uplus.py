"""11-24 조용한 실패 점검 5차 — U+커넥트 차량 연동(스위치를 켜는 순간 조용히 틀어질 곳).

① 운전자 확인 DM 을 한 통도 못 보내면 ERROR + 'DM 못 보냄'(-2)으로 따로 적고 다음 점검 때 다시, 미확인 목록에 표시
② 차량 아침 알림 — DM 이 간 뒤에만 만기 '보냄' 표시, 한 통도 못 보내면 error(스케줄러 ERROR)·다음 날 다시
④ 안전운전 짝이 없는 운행은 급가속 등을 None(0 회 아님), 운전 습관·운행일지가 '-'
⑤ UPLUS_APP_* 를 잘못 적으면 ERROR(한 번) + '설정 오류'(승인 전과 구별), 메뉴는 보이게
⑥ GPS 위치가 비어 온 차는 위치 받은 시각(pos_at) 그대로 — '방금'·운행 중으로 안 보이게
U+·슬랙 API 는 모두 가짜. 이름·좌표도 가짜."""
import logging
import sqlite3
from datetime import date, datetime, timedelta

import pytest

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

REAL_ENABLED = up.enabled
ASSIGNER = 'sj@example.com'
STAFF = [{'code': c, 'name': n, 'slack_id': sid, 'dept': ''} for c, n, sid in [
    ('YM', '김가람', 'UYM'), ('LSH', '이나래', 'ULSH'), ('TH', '서준', 'UTH'), ('NO', '없음', ''), ('NO2', '없음둘', '')]]
DAY = '2026-11-25'
RAY, SPARK = '231허6854', '53우1783'
BASE = (37.5800, 126.6600)
SITE = (37.4800, 127.0300)


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
    monkeypatch.setattr(va, 'CAR_DM_ON', True)
    monkeypatch.setattr(tw, 'END_ASK_ON', True)
    monkeypatch.setattr(vd, 'DRIVER_ASK_ON', True)
    monkeypatch.setattr(vas, 'reverse', lambda lat, lon: '어딘가')
    state = {'ok': True, 'calls': []}

    def fake_slack(method, payload):
        state['calls'].append((method, payload))
        if not state['ok']:
            return {'ok': False, 'error': 'channel_not_found'}
        return {'ok': True, 'channel': 'D' + payload.get('channel', 'X'), 'ts': f"1.{len(state['calls'])}"}
    monkeypatch.setattr(vd, '_slack', fake_slack)
    gw._invalidate()
    yield state
    gw._invalidate()


def _st(plate, pos, at=None):
    return {'plate': plate, 'lat': pos[0], 'lon': pos[1], 'at': at, 'speed': 0.0, 'key_state': '', 'odometer': None}


def _t(hm, day=DAY):
    return datetime.strptime(f'{day} {hm}', '%Y-%m-%d %H:%M')


def _posts(state):
    return [p for m, p in state['calls'] if m == 'chat.postMessage']


def _errors(caplog, needle=''):
    return [r for r in caplog.records if r.levelno >= logging.ERROR and needle in r.getMessage()]


def _end_ask(plate, kind='park', day=DAY):
    with tw._conn() as conn:
        r = conn.execute('SELECT sent FROM vehicle_end_ask WHERE day = ? AND plate = ? AND kind = ?', (day, plate, kind)).fetchone()
    return r['sent'] if r else None


def _park_ray(people=('YM', 'LSH')):
    va.save_day(DAY, [{'plate': RAY, 'people': list(people)}], ASSIGNER, confirm=True)
    tw.poll(_t('14:00'), [_st(RAY, BASE)])
    tw.poll(_t('15:10'), [_st(RAY, SITE)])           # 15:10 현장에 섬


# ── ① 운전자 확인 DM ─────────────────────────────────────────────────────
def test_park_dm_fail_is_recorded_retried_and_shown(db, monkeypatch, caplog):
    _park_ray()
    db['ok'] = False
    r = tw.poll(_t('16:00'), [_st(RAY, SITE)])
    assert [(e['plate'], e['sent']) for e in r['ended']] == [(RAY, vd.DM_FAILED)]
    assert _end_ask(RAY) == tw.DM_FAILED == vd.DM_FAILED
    errs = _errors(caplog, '한 통도 못 보냄')
    assert len(errs) == 1 and getattr(errs[0], 'force_alert', False) is True and 'channel_not_found' in errs[0].getMessage()
    # 미확인 목록 — '물음'이 아니라 DM 못 보냄
    monkeypatch.setattr(tw, 'date', type('D', (date,), {'today': classmethod(lambda c: date(2026, 11, 26))}))
    m = tw.missing('2026-11-01', '2026-11-30')
    assert [(x['plate'], x['asked'], x['dm_failed']) for x in m] == [(RAY, 'park', True)]
    # 다음 점검(10분 뒤) 때 다시 묻고, 되면 '물음'으로 바뀌고 그 뒤로는 안 물음
    db['ok'] = True
    db['calls'].clear()
    r = tw.poll(_t('16:10'), [_st(RAY, SITE)])
    assert [(e['plate'], e['sent']) for e in r['ended']] == [(RAY, 2)] and len(_posts(db)) == 2
    assert _end_ask(RAY) == 2
    assert tw.poll(_t('16:20'), [_st(RAY, SITE)])['ended'] == []
    assert [(x['asked'], x['dm_failed']) for x in tw.missing('2026-11-01', '2026-11-30')] == [('park', False)]


def test_switch_off_is_not_dm_failure(db, monkeypatch, caplog):
    _park_ray()
    monkeypatch.setattr(tw, 'END_ASK_ON', False)
    r = tw.poll(_t('16:00'), [_st(RAY, SITE)])
    assert [(e['plate'], e['sent']) for e in r['ended']] == [(RAY, 0)] and not _posts(db)
    assert not _errors(caplog)
    assert tw.poll(_t('16:10'), [_st(RAY, SITE)])['ended'] == []          # 기록만 — 다시 안 물음(아침에 물음)
    monkeypatch.setattr(tw, 'date', type('D', (date,), {'today': classmethod(lambda c: date(2026, 11, 26))}))
    assert [x['dm_failed'] for x in tw.missing('2026-11-01', '2026-11-30')] == [False]


def test_no_slack_id_is_failure_partial_is_warning(db, caplog):
    _park_ray(people=('NO', 'NO2'))                   # 둘 다 슬랙 ID 없음
    r = tw.poll(_t('16:00'), [_st(RAY, SITE)])
    assert r['ended'][0]['sent'] == vd.DM_FAILED and not _posts(db)
    assert '슬랙 ID 없음' in _errors(caplog, '한 통도 못 보냄')[0].getMessage()
    caplog.clear()
    # 한 명은 보내지고 한 명은 슬랙 ID 없음 — 보낸 수, ERROR 아님(WARNING)
    with caplog.at_level(logging.WARNING):
        n = vd.ask_car('2026-11-24', SPARK, ['YM', 'NO'])
    assert n == 1 and not _errors(caplog)
    assert any('일부 못 보냄' in x.getMessage() for x in caplog.records if x.levelno == logging.WARNING)
    # 이미 보낸 사람만 남았으면 실패 아님(0)
    assert vd.ask_car('2026-11-24', SPARK, ['YM', 'NO']) == 0 and not _errors(caplog)


def test_morning_dm_fail_counts_and_retries(db, monkeypatch, caplog):
    va.save_day(DAY, [{'plate': RAY, 'people': ['YM', 'LSH']}, {'plate': SPARK, 'people': ['TH']}], ASSIGNER, confirm=True)
    monkeypatch.setattr(tw, 'active', lambda: True)
    db['ok'] = False
    with caplog.at_level(logging.INFO):
        r = tw.remind_morning(date(2026, 11, 26))
    assert r == {'day': DAY, 'sent': 0, 'auto': 1, 'failed': 1}
    assert _end_ask(RAY, 'morning') == vd.DM_FAILED and _errors(caplog, '한 통도 못 보냄')
    assert any('DM 못 보낸 차 1대' in x.getMessage() for x in caplog.records)
    db['ok'] = True
    assert tw.remind_morning(date(2026, 11, 26)) == {'day': DAY, 'sent': 2, 'auto': 0}   # 못 보낸 차만 다시
    caplog.clear()
    with caplog.at_level(logging.INFO):
        assert tw.remind_morning(date(2026, 11, 26)) == {'day': DAY, 'sent': 0, 'auto': 0}
    assert any('아침 다시 묻기 DM 0건' in x.getMessage() for x in caplog.records)      # 보낸 게 없어도 한 줄


def test_send_asks_logs_zero_and_failed(db, monkeypatch, caplog):
    va.save_day(DAY, [{'plate': RAY, 'people': ['YM', 'LSH']}], ASSIGNER, confirm=True)
    monkeypatch.setattr(tw, 'active', lambda: False)
    db['ok'] = False
    with caplog.at_level(logging.INFO):
        assert vd.send_asks(DAY) == {'day': DAY, 'sent': 0, 'auto': 0, 'failed': 1}
    assert _errors(caplog, '한 통도 못 보냄')
    db['ok'] = True
    caplog.clear()
    with caplog.at_level(logging.INFO):
        assert vd.send_asks('2026-11-27') == {'day': '2026-11-27', 'sent': 0, 'auto': 0}   # 확정한 배정 없음
    assert any('운전자 묻기 DM 0건' in x.getMessage() for x in caplog.records)


# ── ② 차량 아침 알림 ───────────────────────────────────────────────────────
def test_alert_marks_only_after_dm(db, monkeypatch, caplog):
    vm.update_info(RAY, {'insurance_end': '2026-12-24'}, 'x')
    monkeypatch.setattr(vas, 'ALERT_ON', True)
    monkeypatch.setattr(vas, 'alert_recipients', lambda: ['UADMIN'])
    monkeypatch.setattr(va, '_dm', lambda sid, text: False)
    real = vas.send_alerts
    r = real(date(2026, 11, 24))
    assert r['sent'] == 0 and '한 통도 못 보냈어요' in r['error'] and '내일 다시' in r['error']
    assert not vas._sent(f'exp:{RAY}:보험:2026-12-24:30')                      # 못 보냈으니 '보냄' 표시 없음
    # 스케줄러는 error 를 ERROR 로
    monkeypatch.setattr(sched, '_vehicle_api_fails', {})
    monkeypatch.setattr(vas, 'send_alerts', lambda: real(date(2026, 11, 24)))
    sched._safe_vehicle_alerts()
    assert _errors(caplog, '차량 알림 실패')
    # 되면 표시 → 같은 줄은 다시 안 감
    texts = []
    monkeypatch.setattr(va, '_dm', lambda sid, text: texts.append(text) or True)
    assert real(date(2026, 11, 24))['sent'] == 1 and '보험 만기' in texts[0]
    assert vas._sent(f'exp:{RAY}:보험:2026-12-24:30')
    texts.clear()
    real(date(2026, 11, 24))
    assert not any('보험 만기' in t for t in texts)


def test_alert_no_recipients_is_error(db, monkeypatch):
    vm.update_info(RAY, {'insurance_end': '2026-12-24'}, 'x')
    monkeypatch.setattr(vas, 'ALERT_ON', True)
    monkeypatch.setattr(vas, 'alert_recipients', lambda: [])
    r = vas.send_alerts(date(2026, 11, 24))
    assert r['sent'] == 0 and '받는 사람이 없어요' in r['error']
    assert not vas._sent(f'exp:{RAY}:보험:2026-12-24:30')


def test_alert_missing_read_failure_warns(db, monkeypatch, caplog):
    monkeypatch.setattr(tw, 'missing', lambda s, e: (_ for _ in ()).throw(sqlite3.OperationalError('database is locked')))
    with caplog.at_level(logging.WARNING):
        vas.alert_items(date(2026, 11, 24))
    assert any(x.levelno >= logging.WARNING and '운전자 미확인 읽기 실패' in x.getMessage() for x in caplog.records)


# ── ④ 안전운전 짝 없음 ─────────────────────────────────────────────────────
def test_trips_without_safe_pair_are_none(monkeypatch):
    monkeypatch.setenv('UPLUS_SECURE_KEY', 'test-key')
    monkeypatch.setenv('UPLUS_APP_ECO', '7')
    monkeypatch.delenv('UPLUS_APP_SAFE', raising=False)          # 안전운전 승인 전
    monkeypatch.setattr(up, '_tokens', {})

    def fake_post(path, body):
        if path == '/v1/session':
            return {'extensions': {'code': '20000000', 'data': [{'token': 't'}]}}
        assert path == up.APIS['eco']
        return {'extensions': {'code': '20000000', 'totalRecord': 1, 'data': [
            {'use_start_dt': '2026-11-23 09:00:00', 'use_end_dt': '2026-11-23 10:00:00', 'mileage': 30.0,
             'echo_driving_index': 95.0}]}}
    monkeypatch.setattr(up, '_post', fake_post)
    t = up.trips(RAY, '2026-11-23')[0]
    assert all(t[k] is None for k, _ in up.SAFE_EVENTS) and t['safe'] is None and t['eco'] == 95.0


def _metric(plate, hm, km, driver, **ev):
    with vl._conn() as conn:
        conn.execute("INSERT INTO vehicle_trips (plate, day, start_time, end_time, km, note) VALUES (?,?,?,?,?,?)",
                     (plate, '2026-11-23', hm, hm, km, 'U+ API'))
    keys = ('rapid_acc', 'rapid_dec', 'sudden_stop', 'rapid_turn', 'rapid_u_turn', 'rapid_lane', 'rapid_pass')
    with vas._conn() as conn:
        conn.execute('INSERT INTO vehicle_trip_metrics (plate, day, start_time, max_speed, driver, commute, '
                     + ', '.join(keys) + ') VALUES (?,?,?,?,?,0,' + ','.join('?' * len(keys)) + ')',
                     (plate, '2026-11-23', hm, 80, driver, *[ev.get(k) for k in keys]))


def test_habits_unknown_events_are_none(db):
    _metric(RAY, '09:00', 50.0, 'YM')                                  # 짝 없음
    _metric(SPARK, '09:00', 40.0, 'LSH', rapid_acc=2, rapid_dec=0)     # 짝 있음
    _metric(SPARK, '13:00', 60.0, 'LSH')                               # 짝 없음 — 100km당은 아는 40km 로만
    h = vas.habits('2026-11-01', '2026-11-30')
    ym = next(p for p in h['people'] if p['key'] == 'YM')
    lsh = next(p for p in h['people'] if p['key'] == 'LSH')
    assert (ym['acc'], ym['dec'], ym['stop'], ym['rapid'], ym['rapid_per100']) == (None, None, None, None, None)
    assert (lsh['acc'], lsh['dec'], lsh['rapid'], lsh['rapid_per100'], lsh['ev_trips']) == (2, 0, 2, 5.0, 1)
    days = vl.vehicle_days('2026-11', RAY, [])['days']
    assert days[0]['trips'][0]['events'] is None                      # 운행일지 펼침창도 0 회로 안 보임
    days = vl.vehicle_days('2026-11', SPARK, [])['days']
    assert {t['start']: t['events'] for t in days[0]['trips']}['09:00']['rapid_acc'] == 2


# ── ⑤ U+ 설정 오류 ─────────────────────────────────────────────────────────
def test_bad_app_id_is_config_error(db, monkeypatch, caplog):
    monkeypatch.setattr(up, 'enabled', REAL_ENABLED)
    monkeypatch.setattr(up, '_bad_logged', set())
    monkeypatch.setenv('UPLUS_SECURE_KEY', 'test-key')
    for n in up.APIS:
        monkeypatch.delenv(f'UPLUS_APP_{n.upper()}', raising=False)
    assert up.state('status') == 'off' and up.config_errors() == [] and gw._uplus_on() is False   # 승인 전
    monkeypatch.setenv('UPLUS_APP_STATUS', '12a')
    assert up.enabled('status') is False and up.state('status') == 'config_error'
    up.enabled('status')
    up.app('status')
    errs = _errors(caplog, 'UPLUS_APP_STATUS')
    assert len(errs) == 1 and getattr(errs[0], 'force_alert', False) is True          # 같은 값은 한 번만
    monkeypatch.setattr(vas, 'refresh_state', lambda max_age_sec=0: {'skipped': 'no_api'})
    v = vas.now_view()
    assert v['enabled'] is False and 'UPLUS_APP_STATUS' in v['config_error']
    assert gw._uplus_on() is True                                                      # 메뉴는 보여 '설정 오류'를 볼 수 있게
    monkeypatch.setenv('UPLUS_APP_STATUS', '12')
    assert up.state('status') == 'on' and vas.now_view()['config_error'] == ''


# ── ⑥ GPS 위치가 비어 온 차 ────────────────────────────────────────────────
def _state_row(plate):
    with vas._conn() as conn:
        return dict(conn.execute('SELECT * FROM vehicle_state WHERE plate = ?', (plate,)).fetchone())


def test_empty_gps_keeps_position_time(db, monkeypatch):
    monkeypatch.setattr(vas, 'refresh_state', lambda max_age_sec=0: {'cached': True})
    monkeypatch.setattr(up, 'enabled', lambda name='status': True)
    old = (datetime.now() - timedelta(hours=3)).replace(microsecond=0)
    vas.save_state([{'plate': RAY, 'lat': SITE[0], 'lon': SITE[1], 'at': old, 'speed': 0.0}], now=old)
    assert _state_row(RAY)['pos_at'] == old.strftime('%Y-%m-%d %H:%M:%S')
    now = datetime.now().replace(microsecond=0)
    vas.save_state([{'plate': RAY, 'lat': None, 'lon': None, 'at': now, 'speed': 40.0}], now=now)   # 신호만, 위치 빔
    row = _state_row(RAY)
    assert (row['lat'], row['pos_at'], row['at']) == (SITE[0], old.strftime('%Y-%m-%d %H:%M:%S'), now.strftime('%Y-%m-%d %H:%M:%S'))
    c = next(x for x in vas.now_view()['cars'] if x['plate'] == RAY)
    assert c['mins'] >= 179 and c['moving'] is False and c['pos_stale'] is True       # '방금'·운행 중 아님
    # 위치가 다시 오면 그 시각으로
    vas.save_state([{'plate': RAY, 'lat': BASE[0], 'lon': BASE[1], 'at': now, 'speed': 40.0}], now=now)
    c = next(x for x in vas.now_view()['cars'] if x['plate'] == RAY)
    assert c['mins'] <= 1 and c['moving'] is True and c['pos_stale'] is False


def test_old_table_gets_pos_at_column(tmp_path, monkeypatch):
    """칸을 더하기 전의 vehicle_state — 열면 pos_at 이 붙고, 위치가 비어 오면 옛 at 을 위치 시각으로."""
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(vas, '_db', lambda: d)
    monkeypatch.setattr(vas, '_table_ready', False)
    with d._get_connection() as conn:
        conn.execute('CREATE TABLE vehicle_state (plate TEXT PRIMARY KEY, at TEXT, seen_at TEXT, lat REAL, lon REAL, '
                     'speed REAL, key_state TEXT, odometer REAL, fuel REAL, voltage REAL, tpms TEXT, soc REAL)')
        conn.execute("INSERT INTO vehicle_state (plate, at, seen_at, lat, lon) VALUES (?,?,?,?,?)",
                     (RAY, '2026-11-24 08:00:00', '2026-11-24 08:00:00', SITE[0], SITE[1]))
        conn.commit()
    vas.save_state([{'plate': RAY, 'lat': None, 'lon': None, 'at': datetime(2026, 11, 24, 12, 0)}],
                   now=datetime(2026, 11, 24, 12, 0))
    row = _state_row(RAY)
    assert (row['lat'], row['pos_at'], row['at']) == (SITE[0], '2026-11-24 08:00:00', '2026-11-24 12:00:00')
