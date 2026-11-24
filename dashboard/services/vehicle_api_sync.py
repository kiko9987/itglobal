"""U+ API 로 차량 기능 채우기 (2026-11-24 부장님 — U+커넥트 기능을 그룹웨어 '차량'으로 모두 옮김. 승인 전에 미리 만듦).

① 운행 기록 자동 수집 — 매일 새벽 어제(+놓친 날) 운행을 경제운전(시작·끝·계기판 km)·안전운전(급가속 등)으로 받아
   하루 경로 좌표로 출발·도착 주소(카카오)를 붙여 vehicle_trips 에 넣는다(zip 올리기와 같은 표 — 같은 운행은 한 번).
   출퇴근 구분(부장님 11-24): 한쪽 끝이 회사, 다른 끝이 그날 그 차에 탄 사람의 집(그룹웨어 인사 프로필 주소)이면 출퇴근,
   나머지는 일반 업무. 집 주소는 카카오로 좌표를 한 번 찾아 둔다(vehicle_home_geo).
⑤ 차량 상태 — 30분마다(운행 종료 감지 시간엔 10분마다) 위치·누적 km·연료·배터리 전압·타이어를 vehicle_state 에.
⑨ 알림 — 아침 08:40 경영지원에게 한 번: 어제 과속, 배터리 낮음, 단말 신호 끊김, 보험·검사 만기(30·7·0일 전, 지남),
   운전자 미확인. 없으면 안 보냄. ALERT_ON=False 동안은 안 보냄.

키가 없으면(uplus_api.enabled) 모두 쉰다. 만기 알림만은 키 없이도 된다.
"""
from __future__ import annotations

import json
import math
import re
import time
import urllib.parse
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from typing import Dict, Iterable, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

COMPANY_ADDR = '경기도 김포시 걸포로 134-84'
NEAR_M = 300                 # 회사·집 '근처' 반경
SPEED_LIMIT = 120            # 과속으로 볼 최고 속도(km/h) — 설정 vehicle_speed_limit 로 바꿀 수 있음
LOW_VOLT = 11.8              # 12V 차 배터리 낮음
SILENT_DAYS = 3              # 단말이 이만큼 아무것도 안 보내면 '신호 끊김'
EXPIRY_NOTICE = (30, 7, 0)   # 만기 며칠 전에 알릴지
SYNC_BACK_DAYS = 3           # 새벽에 거슬러 다시 받아 볼 날 수(늦게 올라온 운행)

ALERT_ON: bool = False       # 경영지원 아침 알림 DM 스위치 — 부장님 확인 뒤 켬

_DDL = [
    '''CREATE TABLE IF NOT EXISTS vehicle_trip_metrics (
        plate TEXT NOT NULL, day TEXT NOT NULL, start_time TEXT NOT NULL,
        end_at TEXT, minutes REAL,
        from_lat REAL, from_lon REAL, to_lat REAL, to_lon REAL,
        eco REAL, safe REAL, avg_speed REAL, max_speed REAL, fuel_ml REAL,
        rapid_acc INTEGER, rapid_dec INTEGER, sudden_stop INTEGER, rapid_turn INTEGER,
        rapid_u_turn INTEGER, rapid_lane INTEGER, rapid_pass INTEGER,
        driver TEXT, commute INTEGER NOT NULL DEFAULT 0,
        PRIMARY KEY (plate, day, start_time)
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicle_api_days (
        day TEXT PRIMARY KEY, synced_at TEXT, trips INTEGER, error TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicle_state (
        plate TEXT PRIMARY KEY, at TEXT, seen_at TEXT, lat REAL, lon REAL, speed REAL, key_state TEXT,
        odometer REAL, fuel REAL, voltage REAL, tpms TEXT, soc REAL
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicle_device (
        plate TEXT PRIMARY KEY, model TEXT, status TEXT, fault_at TEXT, memo TEXT, checked_at TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicle_geo_cache (
        k TEXT PRIMARY KEY, v TEXT, at TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicle_home_geo (
        code TEXT PRIMARY KEY, addr TEXT, lat REAL, lon REAL, at TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicle_alert_sent (
        k TEXT PRIMARY KEY, at TEXT
    )''',
]
_table_ready = False


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


@contextmanager
def _conn():
    global _table_ready
    with _db()._get_connection() as conn:
        if not _table_ready:
            for ddl in _DDL:
                conn.execute(ddl)
            conn.commit()
            _table_ready = True
        yield conn
        conn.commit()


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def meters(lat1, lon1, lat2, lon2) -> float:
    r = 6371000.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    a = (math.sin((p2 - p1) / 2) ** 2
         + math.cos(p1) * math.cos(p2) * math.sin(math.radians(lon2 - lon1) / 2) ** 2)
    return 2 * r * math.asin(math.sqrt(a))


# ── 카카오(주소 ↔ 좌표) ─────────────────────────────────────────────────────
def _cache_get(k: str) -> Optional[str]:
    with _conn() as conn:
        r = conn.execute('SELECT v FROM vehicle_geo_cache WHERE k = ?', (k,)).fetchone()
    return r['v'] if r else None


def _cache_put(k: str, v: str) -> None:
    with _conn() as conn:
        conn.execute('INSERT OR REPLACE INTO vehicle_geo_cache (k, v, at) VALUES (?,?,?)', (k, v, _now()))


def geocode(addr: str) -> Optional[Tuple[float, float]]:
    """주소 → (위도, 경도). 상세(동·호수)가 붙어 못 찾으면 끝 낱말을 하나씩 떼고 다시. 결과는 저장해 둔다."""
    addr = ' '.join(str(addr or '').split())
    if not addr:
        return None
    k = 'g:' + addr
    hit = _cache_get(k)
    if hit is not None:
        return tuple(json.loads(hit)) if hit else None
    from dashboard.services.address_resolver import _kakao_search
    words = addr.replace(',', ' ').split()
    found = None
    for n in range(len(words), 1, -1):
        doc = _kakao_search(' '.join(words[:n]))
        if doc and doc.get('x') and doc.get('y'):
            found = (float(doc['y']), float(doc['x']))
            break
    _cache_put(k, json.dumps(found) if found else '')
    return found


def reverse(lat: float, lon: float) -> str:
    """좌표 → 지번 주소('경기 김포시 걸포동 123-4'). 4째 자리(약 10m)로 묶어 저장."""
    if lat is None or lon is None:
        return ''
    k = f'r:{lat:.4f},{lon:.4f}'
    hit = _cache_get(k)
    if hit is not None:
        return hit
    from dashboard.services.address_resolver import _kakao_get_json
    url = 'https://dapi.kakao.com/v2/local/geo/coord2address.json?' + urllib.parse.urlencode({'x': lon, 'y': lat})
    try:
        data = _kakao_get_json(url) or {}
    except Exception:
        return ''
    docs = data.get('documents') or []
    out = ''
    if docs:
        a = docs[0].get('address') or {}
        ra = docs[0].get('road_address') or {}
        out = a.get('address_name') or ra.get('address_name') or ''
        b = (ra.get('building_name') or '').strip()
        if b and b not in out:
            out = f'{out} ({b})'
    _cache_put(k, out)
    return out


def company() -> Optional[Tuple[float, float]]:
    return geocode(COMPANY_ADDR)


def home_of(code: str) -> Optional[Tuple[float, float]]:
    """직원 집(그룹웨어 인사 프로필 주소) 좌표 — 주소가 바뀌면 다시 찾는다."""
    from dashboard.services import hr
    from dashboard.services import vehicle_assign as va
    p = va._who_map().get(code)
    if not p:
        return None
    rows = hr._rows("SELECT address FROM hr_people WHERE name = ? AND status = '재직'", (p['name'],))
    addr = ' '.join(str((rows[0] if rows else {}).get('address') or '').split())
    if not addr:
        return None
    with _conn() as conn:
        r = conn.execute('SELECT addr, lat, lon FROM vehicle_home_geo WHERE code = ?', (code,)).fetchone()
    if r and r['addr'] == addr:
        return (r['lat'], r['lon']) if r['lat'] is not None else None
    pt = geocode(addr)
    with _conn() as conn:
        conn.execute('INSERT OR REPLACE INTO vehicle_home_geo (code, addr, lat, lon, at) VALUES (?,?,?,?,?)',
                     (code, addr, pt[0] if pt else None, pt[1] if pt else None, _now()))
    return pt


# ── ① 운행 기록 ─────────────────────────────────────────────────────────
def _nearest(pts: List[Tuple[datetime, float, float]], t: Optional[datetime], within_min: int = 10):
    if not pts or not t:
        return None
    best = min(pts, key=lambda p: abs((p[0] - t).total_seconds()))
    return best if abs((best[0] - t).total_seconds()) <= within_min * 60 else None


def _riders_and_driver(day: str, plate: str, start: datetime) -> Tuple[List[str], str]:
    """그날 그 차에 탄 사람(이니셜)과 이 운행 운전자 — 운전자 답 > 배정 맨 앞(13시 뒤는 오후 맨 앞) > 고정 사용자."""
    from dashboard.services import vehicle_assign as va
    from dashboard.services import vehicle_driver as vd
    rows = va.car_rows(day)
    r = rows.get(plate)
    riders = va.riders(r) if r else []
    pm = start.strftime('%H:%M') >= va.HALF_AT
    a = vd.answer(day, plate) or {}
    drv = (a.get('pm') if pm else a.get('am')) or ''
    if not drv and r:
        side = va.pm_of(r) if pm else r['people']
        drv = side[0] if side else (riders[0] if riders else '')
    if not riders:
        fu = va.fixed_users().get(plate)
        if fu:
            riders, drv = [fu], drv or fu
    return riders, drv


def is_commute(a: Optional[Tuple[float, float]], b: Optional[Tuple[float, float]], riders: List[str]) -> bool:
    """한쪽 끝이 회사, 다른 끝이 탄 사람 중 누구의 집이면 출퇴근(부장님 11-24)."""
    co = company()
    if not a or not b or not co or not riders:
        return False
    homes = [h for h in (home_of(c) for c in riders) if h]
    if not homes:
        return False

    def near(p, q):
        return meters(p[0], p[1], q[0], q[1]) <= NEAR_M
    return any((near(a, co) and near(b, h)) or (near(b, co) and near(a, h)) for h in homes)


def sync_day(day: str, plates: Optional[Iterable[str]] = None) -> Dict:
    """하루 운행을 U+ 에서 받아 넣는다. 같은 운행(차·날짜·출발 시각)은 다시 안 넣고 지표만 고친다."""
    from dashboard.services import uplus_api as up
    from dashboard.services import vehicle_assign as va
    from dashboard.services import vehicle_log as vl
    if not up.enabled('eco'):
        return {'skipped': 'no_api'}
    cars = [c for c in va.fleet(True) if c.get('uplus') and c['plate'] != va.OWN]
    if plates:
        want = set(plates)
        cars = [c for c in cars if c['plate'] in want]
    added = seen = 0
    errors = []
    for c in cars:
        plate = c['plate']
        try:
            trips = up.trips(plate, day)
            if not trips:
                continue
            pts = up.route_day(plate, day) if up.enabled('route_day') else []
        except Exception as e:
            errors.append(f'{plate}: {e}')
            continue
        for t in trips:
            p1, p2 = _nearest(pts, t['start']), _nearest(pts, t['end'])
            a = (p1[1], p1[2]) if p1 else None
            b = (p2[1], p2[2]) if p2 else None
            riders, drv = _riders_and_driver(day, plate, t['start'])
            commute = is_commute(a, b, riders)
            st = t['start'].strftime('%H:%M')
            minutes = round((t['end'] - t['start']).total_seconds() / 60, 1) if t['end'] else None
            with vl._conn() as conn:
                cur = conn.execute(
                    'INSERT OR IGNORE INTO vehicle_trips (plate, day, start_time, end_time, odo_start, odo_end, km, '
                    'from_addr, to_addr, commute_km, business_km, note, upload_id) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,NULL)',
                    (plate, day, st, t['end'].strftime('%H:%M') if t['end'] else '', t['odo_start'], t['odo_end'], t['km'],
                     reverse(*a)[:80] if a else '', reverse(*b)[:80] if b else '',
                     t['km'] if commute else 0, 0 if commute else t['km'], 'U+ API'))
                added += cur.rowcount
                seen += 1
            with _conn() as conn:
                conn.execute(
                    'INSERT OR REPLACE INTO vehicle_trip_metrics (plate, day, start_time, end_at, minutes, from_lat, from_lon, '
                    'to_lat, to_lon, eco, safe, avg_speed, max_speed, fuel_ml, rapid_acc, rapid_dec, sudden_stop, rapid_turn, '
                    'rapid_u_turn, rapid_lane, rapid_pass, driver, commute) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                    (plate, day, st, t['end'].strftime('%Y-%m-%d %H:%M:%S') if t['end'] else None, minutes,
                     a[0] if a else None, a[1] if a else None, b[0] if b else None, b[1] if b else None,
                     t['eco'], t['safe'], t['avg_speed'], t['max_speed'], t['fuel_ml'],
                     t['rapid_acc'], t['rapid_dec'], t['sudden_stop'], t['rapid_turn'], t['rapid_u_turn'],
                     t['rapid_lane'], t['rapid_pass'], drv, int(commute)))
    with _conn() as conn:
        conn.execute('INSERT OR REPLACE INTO vehicle_api_days (day, synced_at, trips, error) VALUES (?,?,?,?)',
                     (day, _now(), seen, '; '.join(errors)[:500]))
    if errors:
        logger.warning(f'[U+운행] {day} 일부 실패: {errors}')
    return {'day': day, 'trips': seen, 'added': added, 'errors': errors}


def sync_recent(today: Optional[date] = None) -> Dict:
    """새벽 — 어제부터 거슬러 SYNC_BACK_DAYS 일(늦게 올라온 운행까지)."""
    from dashboard.services import uplus_api as up
    if not up.enabled('eco'):
        return {'skipped': 'no_api'}
    today = today or date.today()
    out = []
    for back in range(1, SYNC_BACK_DAYS + 1):
        out.append(sync_day((today - timedelta(days=back)).isoformat()))
    return {'days': out}


def metrics(start: str, end: str, plate: str = '') -> Dict[Tuple[str, str, str], Dict]:
    """(차, 날짜, 출발 'HH:MM') → 지표."""
    q = 'SELECT * FROM vehicle_trip_metrics WHERE day BETWEEN ? AND ?' + (' AND plate = ?' if plate else '')
    with _conn() as conn:
        rows = conn.execute(q, (start, end, plate) if plate else (start, end)).fetchall()
    return {(r['plate'], r['day'], r['start_time']): dict(r) for r in rows}


# ── ⑤ 차량 상태 ─────────────────────────────────────────────────────────
def save_state(statuses: List[Dict], now: Optional[datetime] = None) -> int:
    now = now or datetime.now()
    n = 0
    with _conn() as conn:
        for s in statuses:
            if not s.get('plate'):
                continue
            at = s.get('at')
            prev = conn.execute('SELECT lat, lon FROM vehicle_state WHERE plate = ?', (s['plate'],)).fetchone()
            lat = s.get('lat') if s.get('lat') is not None else (prev['lat'] if prev else None)
            lon = s.get('lon') if s.get('lon') is not None else (prev['lon'] if prev else None)
            conn.execute('INSERT OR REPLACE INTO vehicle_state (plate, at, seen_at, lat, lon, speed, key_state, odometer, '
                         'fuel, voltage, tpms, soc) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)',
                         (s['plate'], at.strftime('%Y-%m-%d %H:%M:%S') if isinstance(at, datetime) else None,
                          now.strftime('%Y-%m-%d %H:%M:%S'), lat, lon, s.get('speed'), s.get('key_state') or '',
                          s.get('odometer'), s.get('fuel'), s.get('voltage'), json.dumps(s.get('tpms') or []),
                          s.get('soc')))
            n += 1
    return n


def refresh_state(max_age_sec: int = 0) -> Dict:
    """U+ 차량 상태를 받아 저장. max_age_sec 안에 받은 적 있으면 그대로(화면이 자주 불러도 API 를 아끼게)."""
    from dashboard.services import uplus_api as up
    if not up.enabled('status'):
        return {'skipped': 'no_api'}
    if max_age_sec:
        with _conn() as conn:
            r = conn.execute('SELECT MAX(seen_at) AS s FROM vehicle_state').fetchone()
        if r and r['s'] and (datetime.now() - datetime.strptime(r['s'], '%Y-%m-%d %H:%M:%S')).total_seconds() < max_age_sec:
            return {'cached': True}
    try:
        st = up.car_status()
    except Exception as e:
        logger.warning(f'[U+상태] 조회 실패: {e}')
        return {'error': str(e)}
    return {'saved': save_state(st)}


def states() -> Dict[str, Dict]:
    with _conn() as conn:
        rows = conn.execute('SELECT * FROM vehicle_state').fetchall()
    out = {}
    for r in rows:
        d = dict(r)
        try:
            d['tpms'] = json.loads(d.get('tpms') or '[]')
        except ValueError:
            d['tpms'] = []
        out[d['plate']] = d
    return out


def check_devices(day: Optional[str] = None) -> Dict:
    from dashboard.services import uplus_api as up
    if not up.enabled('device_status'):
        return {'skipped': 'no_api'}
    day = day or (date.today() - timedelta(days=1)).isoformat()
    try:
        rows = up.device_status(day)
    except Exception as e:
        return {'error': str(e)}
    with _conn() as conn:
        for d in rows:
            conn.execute('INSERT OR REPLACE INTO vehicle_device (plate, model, status, fault_at, memo, checked_at) '
                         'VALUES (?,?,?,?,?,?)',
                         (d['plate'], d['model'], d['status'],
                          d['fault_at'].strftime('%Y-%m-%d %H:%M:%S') if d['fault_at'] else None, d['memo'], _now()))
    return {'devices': len(rows)}


# ── ⑨ 아침 알림 ─────────────────────────────────────────────────────────
def _speed_limit() -> int:
    try:
        from dashboard.services import groupware as gw
        v = gw.get_setting('vehicle_speed_limit')
        return int(v) if v else SPEED_LIMIT
    except Exception:
        return SPEED_LIMIT


def alert_items(today: Optional[date] = None) -> List[str]:
    """아침 알림 줄들(없으면 []). 만기 알림은 같은 줄을 한 번만(30·7·0일 전, 지남은 매주 월요일)."""
    from dashboard.services import vehicle_assign as va
    from dashboard.services import vehicle_manage as vm
    today = today or date.today()
    y = (today - timedelta(days=1)).isoformat()
    cars = {c['plate']: c for c in va.fleet(True)}
    who = va._who_map()

    def label(p):
        return (cars.get(p) or {}).get('label') or p
    lines: List[str] = []
    lim = _speed_limit()
    with _conn() as conn:
        fast = conn.execute('SELECT plate, start_time, max_speed, driver FROM vehicle_trip_metrics WHERE day = ? AND max_speed >= ? '
                            'ORDER BY max_speed DESC', (y, lim)).fetchall()
    for r in fast:
        lines.append(f":rotating_light: 과속 {label(r['plate'])} {r['start_time']} 출발 운행 최고 {int(r['max_speed'])}km/h"
                     + (f" ({va.display(r['driver'], who)})" if r['driver'] else ''))
    st = states()
    for p, s in st.items():
        c = cars.get(p) or {}
        if c.get('active') == 0:
            continue
        if s.get('voltage') and 0 < s['voltage'] < LOW_VOLT and '전기' not in (c.get('fuel') or ''):
            lines.append(f":battery: 배터리 낮음 {label(p)} {s['voltage']:.1f}V — 방전 위험, 시동 확인")
        last = s.get('at')
        if last and (datetime.now() - datetime.strptime(last, '%Y-%m-%d %H:%M:%S')).days >= SILENT_DAYS:
            lines.append(f":satellite: 단말 신호 끊김 {label(p)} — {last[5:16].replace('-', '/')} 뒤로 안 들어옴")
    with _conn() as conn:
        for r in conn.execute('SELECT plate, fault_at, memo FROM vehicle_device WHERE fault_at >= ?', (y,)).fetchall():
            lines.append(f":satellite: 단말 장애 {label(r['plate'])} {r['fault_at'][5:16].replace('-', '/')} {r['memo'] or ''}".rstrip())
    for c in vm.overview(today)['cars']:
        for what, end, dd in (('보험', c['insurance_end'], c['insurance_d']), ('검사', c['inspection_end'], c['inspection_d'])):
            if dd is None:
                continue
            if dd in EXPIRY_NOTICE or (dd < 0 and today.weekday() == 0):
                k = f'exp:{c["plate"]}:{what}:{end}:{dd if dd >= 0 else "over" + today.isoformat()}'
                if _sent(k):
                    continue
                _mark(k)
                when = '오늘 만기' if dd == 0 else (f'{dd}일 남음' if dd > 0 else f'{-dd}일 지남')
                lines.append(f":page_facing_up: {what} 만기 {c['label']} {end} ({when})")
    try:
        from dashboard.services import vehicle_trip_watch as tw
        miss = tw.missing((today - timedelta(days=7)).isoformat(), y)
        if miss:
            lines.append(f":grey_question: 운전자 미확인 {len(miss)}건(최근 7일) — 운행일지 화면 위에서 볼 수 있어요")
    except Exception:
        pass
    return lines


def _sent(k: str) -> bool:
    with _conn() as conn:
        return bool(conn.execute('SELECT 1 FROM vehicle_alert_sent WHERE k = ?', (k,)).fetchone())


def _mark(k: str) -> None:
    with _conn() as conn:
        conn.execute('INSERT OR REPLACE INTO vehicle_alert_sent (k, at) VALUES (?,?)', (k, _now()))


def alert_recipients() -> List[str]:
    """받는 사람 슬랙 ID — 설정 vehicle_alert_to(메일 목록), 없으면 인사 관리자."""
    from dashboard.services import hr
    emails = hr._setting_list('vehicle_alert_to', hr.hr_admins())
    out = []
    for e in emails:
        rows = hr._rows('SELECT slack_id FROM hr_people WHERE (login_email = ? OR work_email = ?) AND slack_id != ""', (e, e))
        if rows and rows[0].get('slack_id'):
            out.append(rows[0]['slack_id'])
    return list(dict.fromkeys(out))


def send_alerts(today: Optional[date] = None) -> Dict:
    if not ALERT_ON:
        return {'skipped': 'off'}
    lines = alert_items(today)
    if not lines:
        return {'sent': 0}
    from dashboard.services import vehicle_assign as va
    text = ':car: *차량 알림*\n' + '\n'.join(f'• {x}' for x in lines[:30])
    sent = sum(1 for sid in alert_recipients() if va._dm(sid, text))
    return {'sent': sent, 'lines': len(lines)}


# ── ⑥ 운전 습관 · ⑦ 가동률 ──────────────────────────────────────────────────
def habits(start: str, end: str) -> Dict:
    """사람별·차량별 운전 습관 — 운행 수·km·안전/경제 점수(거리 가중 평균)·급가속 등 합계·최고 속도·과속 수."""
    from dashboard.services import vehicle_assign as va
    who = va._who_map()
    cars = {c['plate']: c for c in va.fleet(True)}
    lim = _speed_limit()
    with _conn() as conn:
        rows = conn.execute('SELECT m.*, t.km FROM vehicle_trip_metrics m LEFT JOIN vehicle_trips t '
                            'ON t.plate = m.plate AND t.day = m.day AND t.start_time = m.start_time '
                            'WHERE m.day BETWEEN ? AND ?', (start, end)).fetchall()

    def bucket():
        return {'trips': 0, 'km': 0.0, 'w': 0.0, 'safe_w': 0.0, 'eco_w': 0.0, 'rapid': 0, 'acc': 0, 'dec': 0,
                'stop': 0, 'max_speed': 0, 'speeding': 0}
    people: Dict[str, Dict] = {}
    by_car: Dict[str, Dict] = {}
    for r in rows:
        km = r['km'] or 0
        w = max(km, 0.1)
        for key, store in ((r['driver'] or '', people), (r['plate'], by_car)):
            if not key:
                continue
            b = store.setdefault(key, bucket())
            b['trips'] += 1
            b['km'] += km
            if r['safe'] is not None:
                b['safe_w'] += r['safe'] * w
            if r['eco'] is not None:
                b['eco_w'] += r['eco'] * w
            b['w'] += w
            b['acc'] += r['rapid_acc'] or 0
            b['dec'] += r['rapid_dec'] or 0
            b['stop'] += r['sudden_stop'] or 0
            b['rapid'] += sum((r[k] or 0) for k in ('rapid_acc', 'rapid_dec', 'sudden_stop', 'rapid_turn',
                                                    'rapid_u_turn', 'rapid_lane', 'rapid_pass'))
            b['max_speed'] = max(b['max_speed'], r['max_speed'] or 0)
            b['speeding'] += int((r['max_speed'] or 0) >= lim)

    def done(key, b, name):
        per100 = round(b['rapid'] / b['km'] * 100, 1) if b['km'] >= 1 else None
        return {'key': key, 'name': name, 'trips': b['trips'], 'km': round(b['km'], 1),
                'safe': round(b['safe_w'] / b['w'], 1) if b['w'] and b['safe_w'] else None,
                'eco': round(b['eco_w'] / b['w'], 1) if b['w'] and b['eco_w'] else None,
                'acc': b['acc'], 'dec': b['dec'], 'stop': b['stop'], 'rapid': b['rapid'], 'rapid_per100': per100,
                'max_speed': int(b['max_speed']), 'speeding': b['speeding']}
    ppl = sorted((done(k, b, va.display(k, who)) for k, b in people.items()), key=lambda x: -x['km'])
    cs = sorted((done(k, b, (cars.get(k) or {}).get('label') or k) for k, b in by_car.items()), key=lambda x: -x['km'])
    return {'people': ppl, 'cars': cs, 'speed_limit': lim, 'has_data': bool(rows)}


WORK_MIN = 9 * 60    # 하루 업무 시간(분) — 가동률 분모


def utilization(start: str, end: str) -> Dict:
    """차량별 가동률 — 운행한 날, 운행 시간(분), 가동률(운행 시간 / 근무일 × 9시간). 근무일 = 월~토 중 공휴일 아닌 날."""
    from dashboard.services import vehicle_assign as va
    d, last = date.fromisoformat(start), min(date.fromisoformat(end), date.today() - timedelta(days=1))
    workdays = 0
    while d <= last:
        if d.weekday() < 6 and not _holiday(d):
            workdays += 1
        d += timedelta(days=1)
    with _conn() as conn:
        rows = conn.execute('SELECT plate, COUNT(DISTINCT day) AS days, SUM(minutes) AS mins, COUNT(*) AS n '
                            'FROM vehicle_trip_metrics WHERE day BETWEEN ? AND ? GROUP BY plate', (start, end)).fetchall()
    got = {r['plate']: dict(r) for r in rows}
    out = []
    for c in va.fleet(True):
        if not c.get('uplus') or c['plate'] == va.OWN:
            continue
        g = got.get(c['plate']) or {}
        mins = g.get('mins') or 0
        out.append({'plate': c['plate'], 'label': c['label'], 'days': g.get('days') or 0, 'trips': g.get('n') or 0,
                    'hours': round(mins / 60, 1),
                    'rate': round(mins / (workdays * WORK_MIN) * 100, 1) if workdays else None,
                    'day_rate': round((g.get('days') or 0) / workdays * 100) if workdays else None})
    out.sort(key=lambda x: -(x['rate'] or 0))
    return {'workdays': workdays, 'cars': out, 'has_data': bool(rows)}


def _holiday(d: date) -> bool:
    try:
        from dashboard.services.leave import _is_holiday
        return bool(_is_holiday(d))
    except Exception:
        return False


# ── ③ 경로 · ④ 지금 위치 ───────────────────────────────────────────────────
_route_cache: Dict[Tuple[str, str], Tuple[float, list]] = {}


def route(plate: str, day: str, start: str = '', end: str = '') -> List[List[float]]:
    """그날(또는 start~end 'HH:MM') 경로 좌표 [[위도, 경도], ...] — 지도 선. 10분 동안 기억."""
    from dashboard.services import uplus_api as up
    if not up.enabled('route_day'):
        raise ValueError('U+ 경로 API 승인 전이에요.')
    k = (plate, day)
    hit = _route_cache.get(k)
    if hit and time.time() - hit[0] < 600:
        pts = hit[1]
    else:
        pts = up.route_day(plate, day)
        _route_cache[k] = (time.time(), pts)
    if start or end:
        s = datetime.strptime(f'{day} {start or "00:00"}', '%Y-%m-%d %H:%M')
        e = datetime.strptime(f'{day} {end or "23:59"}', '%Y-%m-%d %H:%M') + timedelta(minutes=1)
        pts = [p for p in pts if s <= p[0] <= e]
    step = max(1, len(pts) // 1500)       # 점이 너무 많으면 솎아 냄(5초마다 → 하루 수천 점)
    return [[round(p[1], 6), round(p[2], 6)] for p in pts[::step]]


def now_view() -> Dict:
    """지금 위치 — 차마다 위치·주소·운행 중/정지·마지막 시각·오늘 탄 사람. 상태는 1분 안에 받은 것을 다시 씀."""
    from dashboard.services import uplus_api as up
    from dashboard.services import vehicle_assign as va
    from dashboard.services import vehicle_driver as vd
    refresh_state(max_age_sec=60)
    st = states()
    today = date.today().isoformat()
    rows = va.car_rows(today)
    who = va._who_map()
    out = []
    for c in va.fleet(True):
        if not c.get('uplus') or c['plate'] == va.OWN or not c.get('active', 1):
            continue
        s = st.get(c['plate']) or {}
        at = s.get('at') or ''
        mins = None
        if at:
            mins = int((datetime.now() - datetime.strptime(at, '%Y-%m-%d %H:%M:%S')).total_seconds() // 60)
        moving = bool(s.get('speed') and s['speed'] > 3 and mins is not None and mins <= 5)
        r = rows.get(c['plate'])
        a = vd.answer(today, c['plate']) or {}
        out.append({'plate': c['plate'], 'label': c['label'], 'model': c.get('model') or '',
                    'crew': bool(re.search('포터|트럭', f"{c.get('model') or ''} {c['label']}")),   # 시공팀 차(배정 화면과 같은 기준)
                    'lat': s.get('lat'), 'lon': s.get('lon'),
                    'addr': reverse(s['lat'], s['lon']) if s.get('lat') is not None else '',
                    'at': at[5:16].replace('-', '/') if at else '', 'mins': mins, 'moving': moving,
                    'speed': s.get('speed'), 'fuel': s.get('fuel'),
                    'riders': [va.display(x, who) for x in (va.riders(r) if r else [])],
                    'ended': bool(a.get('ended_at'))})
    return {'enabled': up.enabled('status'), 'cars': out, 'company': company(),
            'checked_at': datetime.now().strftime('%H:%M')}


def nearest(addr: str) -> Dict:
    """주소 → 가까운 차 순서(직선 거리 km)."""
    pt = geocode(addr)
    if not pt:
        raise ValueError('주소를 찾지 못했어요. 시·구부터 다시 적어 주세요.')
    v = now_view()
    cars = [dict(c, km=round(meters(pt[0], pt[1], c['lat'], c['lon']) / 1000, 1)) for c in v['cars'] if c['lat'] is not None]
    cars.sort(key=lambda c: c['km'])
    return {'point': list(pt), 'cars': cars}
