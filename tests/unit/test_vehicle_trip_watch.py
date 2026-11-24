"""운행 종료 감지 → 운전자 확인(2026-11-24) — U+ 차량 상태로 15시 뒤 45분 넘게 선 차를 찾아 탄 사람에게 묻기,
짐작한 운전자 [맞아요], 다음 날 아침 다시 묻기, 운전자 미확인 목록, U+ API 응답 읽기. 이름은 모두 가짜."""
from datetime import date, datetime

import pytest

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import uplus_api as up
from dashboard.services import vehicle_assign as va
from dashboard.services import vehicle_driver as vd
from dashboard.services import vehicle_log as vl
from dashboard.services import vehicle_trip_watch as tw

ASSIGNER = 'sj@example.com'
STAFF = [{'code': c, 'name': n, 'slack_id': 'U' + c, 'dept': ''} for c, n in [
    ('YM', '김가람'), ('LSH', '이나래'), ('TH', '서준'), ('LJW', '한별'), ('SJ', '오늘')]]
DAY = '2026-11-25'
RAY, SPARK = '231허6854', '53우1783'
BASE = (37.5800, 126.6600)         # 가짜 좌표
SITE = (37.4800, 127.0300)


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, vl, va, hr, vd, tw):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setattr(va, 'staff', lambda: [dict(p) for p in STAFF])
    monkeypatch.setattr(va, 'assigners', lambda: [ASSIGNER])
    monkeypatch.setattr(va, '_dm', lambda sid, text: True)
    monkeypatch.setattr(va, 'CAR_DM_ON', True)
    monkeypatch.setattr(tw, 'END_ASK_ON', True)
    calls = []

    def fake_slack(method, payload):
        calls.append((method, payload))
        return {'ok': True, 'channel': 'D' + payload.get('channel', 'X'), 'ts': f'1.{len(calls)}'}
    monkeypatch.setattr(vd, '_slack', fake_slack)
    gw._invalidate()
    yield calls
    gw._invalidate()


def _st(plate, pos, at=None):
    return {'plate': plate, 'lat': pos[0], 'lon': pos[1], 'at': at, 'speed': 0.0, 'key_state': '', 'odometer': None}


def _t(hm, day=DAY):
    return datetime.strptime(f'{day} {hm}', '%Y-%m-%d %H:%M')


def _assign(confirm=True):
    va.save_day(DAY, [{'plate': RAY, 'people': ['YM', 'LSH']}, {'plate': SPARK, 'people': ['TH']}], ASSIGNER, confirm=confirm)


def _posts(calls):
    return [p for m, p in calls if m == 'chat.postMessage']


def test_meters():
    assert tw.meters(*BASE, *BASE) == 0
    assert 100 < tw.meters(37.5800, 126.6600, 37.5810, 126.6600) < 120     # 위도 0.001° ≈ 111m


def test_park_after_15_then_ask_with_guess(db):
    calls = db
    _assign()
    tw.poll(_t('08:00'), [_st(RAY, BASE), _st(SPARK, BASE)])                 # 아침 — 회사에 서 있음(처음 봄)
    tw.poll(_t('10:00'), [_st(RAY, SITE), _st(SPARK, SITE)])                 # 현장으로
    r = tw.poll(_t('12:00'), [_st(RAY, SITE), _st(SPARK, SITE)])             # 낮에 오래 서 있어도 15시 전이라 아님
    assert r['ended'] == [] and not _posts(calls)
    tw.poll(_t('16:10'), [_st(RAY, BASE, _t('16:02')), _st(SPARK, SITE)])    # 레이 16:02 회사에 섬 / 스파크는 낮부터 그대로
    assert tw.poll(_t('16:40'), [_st(RAY, BASE), _st(SPARK, SITE)])['ended'] == []   # 38분 — 아직
    r = tw.poll(_t('16:50'), [_st(RAY, BASE), _st(SPARK, SITE)])
    assert [e['plate'] for e in r['ended']] == [RAY] and r['ended'][0]['since'] == f'{DAY} 16:02:00'
    posts = _posts(calls)
    assert {p['channel'] for p in posts} == {'UYM', 'ULSH'}
    head = posts[0]['blocks'][0]['text']['text']
    assert '운행을 마친 것 같아요' in head and '16:02부터 서 있음' in head and '*김가람* 님이 했나요' in head   # 기록 없으면 배정 맨 앞
    btn = posts[0]['blocks'][1]['elements'][0]
    assert btn['text']['text'] == '맞아요 (김가람)' and btn['style'] == 'primary'
    # [맞아요] → 운전자 적힘, 두 DM 모두 '답했어요'
    calls.clear()
    vd.handle_action(btn['action_id'], btn['value'], 'ULSH', 'DULSH', '1.2')
    assert (vd.answer(DAY, RAY)['am'], vd.answer(DAY, RAY)['pm']) == ('YM', 'YM')
    assert len([1 for m, _ in calls if m == 'chat.update']) == 2
    # 다시 움직였다 서도 다시 묻지 않음
    calls.clear()
    tw.poll(_t('17:00'), [_st(RAY, SITE), _st(SPARK, SITE)])
    assert tw.poll(_t('18:00'), [_st(RAY, SITE), _st(SPARK, SITE)])['ended'] == [] and not _posts(calls)


def test_single_rider_auto_and_unmoved_car_skipped(db):
    calls = db
    _assign()
    tw.poll(_t('14:00'), [_st(RAY, BASE), _st(SPARK, BASE)])
    tw.poll(_t('15:30'), [_st(RAY, BASE), _st(SPARK, SITE, _t('15:20'))])   # 스파크(혼자)만 움직여 15:20 섬
    r = tw.poll(_t('16:10'), [_st(RAY, BASE), _st(SPARK, SITE)])
    assert [(e['plate'], e['sent']) for e in r['ended']] == [(SPARK, -1)]   # 혼자 — 묻지 않고 적음
    assert vd.answer(DAY, SPARK)['am'] == 'TH' and not _posts(calls)        # 레이는 오늘 안 움직여 안 물음


def test_not_confirmed_or_answered_or_switch_off(db, monkeypatch):
    calls = db
    _assign(confirm=False)
    tw.poll(_t('14:00'), [_st(RAY, BASE)])
    tw.poll(_t('15:10'), [_st(RAY, SITE)])
    assert tw.poll(_t('16:30'), [_st(RAY, SITE)])['ended'] == []            # 확정 안 한 날
    _assign()
    vd.end_trip(DAY, RAY, 'YM', 'LSH', 'LSH', today=date.fromisoformat(DAY))   # 그룹웨어 [운행 종료]로 먼저 적음
    assert tw.poll(_t('16:40'), [_st(RAY, SITE)])['ended'] == []
    # 스위치 꺼짐 — 감지는 기록, DM 은 안 감
    monkeypatch.setattr(tw, 'END_ASK_ON', False)
    with tw._conn() as conn:
        conn.execute('DELETE FROM vehicle_driver')
    calls.clear()
    r = tw.poll(_t('16:50'), [_st(RAY, SITE)])
    assert [(e['plate'], e['sent']) for e in r['ended']] == [(RAY, 0)] and not _posts(calls)


def test_guess_uses_recent_drives(db):
    _assign()
    for d in ('2026-11-20', '2026-11-21', '2026-11-23'):
        vd._save(d, RAY, 'LSH', 'LSH', 'LSH')
    vd._save('2026-11-24', RAY, 'YM', 'YM', 'YM')
    assert vd.guess_driver(DAY, RAY, ['YM', 'LSH']) == 'LSH'
    assert vd.guess_driver(DAY, RAY, ['YM', 'TH']) == 'YM'
    assert vd.guess_driver(DAY, SPARK, ['TH', 'YM']) == 'TH'                # 기록 없으면 맨 앞


def test_morning_reask_once_and_missing(db, monkeypatch):
    calls = db
    _assign()
    tw.poll(_t('14:00'), [_st(RAY, BASE)])
    tw.poll(_t('15:10'), [_st(RAY, SITE)])
    tw.poll(_t('16:00'), [_st(RAY, SITE)])                                   # 레이 물음(답 없음), 스파크는 감지 못 함
    assert len(_posts(calls)) == 2
    monkeypatch.setattr(tw, 'date', type('D', (date,), {'today': classmethod(lambda c: date(2026, 11, 26))}))
    m = tw.missing('2026-11-01', '2026-11-30')
    assert [(x['plate'], x['asked']) for x in m] == [(RAY, 'park'), (SPARK, '')]   # 스파크(혼자)는 감지 못 해 아직 미확인
    calls.clear()
    assert tw.remind_morning(date(2026, 11, 26)) == {'skipped': 'off'}       # U+ 키가 없으면 쉼(저녁 일괄 묻기 쪽)
    monkeypatch.setattr(tw, 'active', lambda: True)
    r = tw.remind_morning(date(2026, 11, 26))
    assert r == {'day': DAY, 'sent': 2, 'auto': 1}                           # 레이 두 명 다시, 스파크는 자동
    assert '운전자를 아직 못 받았어요' in _posts(calls)[0]['blocks'][0]['text']['text']
    assert tw.remind_morning(date(2026, 11, 26))['sent'] == 0                # 한 번만
    assert [(x['plate'], x['asked']) for x in tw.missing('2026-11-01', '2026-11-30')] == [(RAY, 'morning')]
    assert tw._prev_day(date(2026, 11, 26)) == DAY and tw._prev_day(date(2026, 11, 27)) == DAY   # 이틀 전까지(월요일→토요일)


def test_evening_ask_rests_when_watch_active(db, monkeypatch):
    _assign()
    monkeypatch.setattr(vd, 'DRIVER_ASK_ON', True)
    monkeypatch.setattr(tw, 'active', lambda: True)
    assert vd.send_asks(DAY) == {'skipped': 'trip_watch'}
    monkeypatch.setattr(tw, 'active', lambda: False)
    assert vd.send_asks(DAY)['sent'] == 2


def test_no_keys_does_nothing(monkeypatch):
    for k in ('UPLUS_APP_STATUS', 'UPLUS_SECURE_KEY', 'UPLUS_KEY_STATUS'):
        monkeypatch.delenv(k, raising=False)
    assert not up.enabled('status') and tw.poll() == {'skipped': 'no_api'}
    assert tw.active() is False


def test_uplus_parse_and_token_retry(monkeypatch):
    monkeypatch.setenv('UPLUS_APP_STATUS', '12')
    monkeypatch.setenv('UPLUS_SECURE_KEY', 'test-key')
    monkeypatch.setattr(up, '_tokens', {})
    sent = []
    state = {'expired': True}

    def fake_post(path, body):
        sent.append((path, dict(body)))
        if path == '/v1/session':
            return {'message': '성공', 'extensions': {'code': '20000000', 'data': [{'result': 'success', 'token': f'tok{len(sent)}'}]}}
        if state['expired']:
            state['expired'] = False
            return {'message': '인증 토큰이 유효하지 않습니다', 'extensions': {'code': '70000009'}, '_http': 500}
        return {'message': '성공', 'extensions': {'code': '20000000', 'totalRecord': 2, 'data': [
            {'vehicle_number': RAY, 'latitude': '37.5598480', 'longitude': '126.8334403', 'created_at': '2026-11-25 16:02:00.0',
             'velocity': '0.0', 'key_state': '3', 'cumulative_driving_distance': '148672.0'},
            {'vehicle_number': SPARK, 'latitude': '0', 'longitude': '0', 'created_at': '2026-11-25T15:20:00.000+00:00'}]}}
    monkeypatch.setattr(up, '_post', fake_post)
    assert up.enabled('status')
    out = up.car_status()
    assert [p for p, _ in sent] == ['/v1/session', up.APIS['status'], '/v1/session', up.APIS['status']]   # 토큰 다시 받고 한 번 더
    assert sent[1][1]['application_id'] == 12 and sent[3][1]['token'] == 'tok3'
    assert out[0]['plate'] == RAY and out[0]['lat'] == pytest.approx(37.559848) and out[0]['at'] == datetime(2026, 11, 25, 16, 2)
    assert out[1]['lat'] is None and out[1]['at'] == datetime(2026, 11, 25, 15, 20)       # 위치 0 → 없음, +00:00 은 그대로
    assert up.parse_time('') is None and up.parse_time('2026-11-25 16:02') == datetime(2026, 11, 25, 16, 2)
