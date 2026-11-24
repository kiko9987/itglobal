"""운행 종료 감지 → 운전자 확인 (2026-11-24 부장님, U+ API 승인 전에 미리 만듦).

부장님 생각: '회사 주소에서 일정 시간 이동이 없으면 복귀로 보고 운전자를 묻자'.
회사로 안 돌아오고 바로 퇴근하는 차가 많아(11월 운행일지: 카니발·스파크) 회사 대신 '그날 마지막으로 선 자리'로 본다.

- 10분마다(14~22시) U+ 차량 상태 API 로 차마다 위치를 받아 둔다(vehicle_park).
  150m 넘게 움직였으면 '오늘 움직임' + 그 자리에 선 시각을 새로 적는다.
- 15시 뒤에 선 차가 45분 넘게 그대로면 운행 종료로 본다 → 그날 확정한 배정에서 그 차에 탄 사람에게 DM.
  혼자 탄 차는 묻지 않고 그 사람으로 적는다. 여럿이면 짐작한 운전자로 '○○ 님이 운전했나요? [맞아요]'.
  다시 움직여도 다시 묻지 않는다(마지막 답이 맞음 — DM 버튼으로 고칠 수 있음).
- 다음 날 아침(월~토 09:10) 어제 답이 없는 차 — 안 물은 차는 묻고, 물었는데 답 없는 차는 한 번 더.
- 그래도 남으면 운전자 미확인(missing)으로 경영지원 화면에.
- 그룹웨어 [운행 종료]로 먼저 적으면 묻지 않는다(답이 있으니).
- 감지가 켜지면 저녁 18:30 일괄 묻기(vehicle_driver.send_asks)는 쉰다.

스위치: U+ 키가 없으면 아무것도 안 함. 키가 있어도 END_ASK_ON=False 동안은 감지만 기록하고 DM 은 안 보낸다
(승인 뒤 며칠 감지 기록을 실제 운행과 맞춰 보고 켠다).
"""
from __future__ import annotations

import math
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

PARK_RADIUS_M = 150         # 이만큼 안에서 움직이면 그대로 서 있는 것(GPS 흔들림·주차장 안 이동)
PARK_MIN = 45               # 이만큼 서 있으면 운행 종료
END_FROM = '15:00'          # 이 시각 뒤에 선 것만 '그날 마지막' 으로 본다(현장에 오래 있는 낮 시간 빼기)
POLL_HOURS = (14, 22)       # 이 시간대만 조회(호출 수 줄이기)

END_ASK_ON: bool = False    # DM 스위치 — 승인 뒤 감지 기록을 실제와 맞춰 본 다음 True + 재시작

_DDL = [
    '''CREATE TABLE IF NOT EXISTS vehicle_park (
        plate TEXT PRIMARY KEY,
        lat REAL, lon REAL,
        since TEXT,          -- 이 자리에 선 시각
        seen_at TEXT,        -- 마지막으로 본 시각
        moved_day TEXT       -- 마지막으로 움직인 날(오늘 움직인 차만 운행 종료로 본다)
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicle_end_ask (
        day TEXT NOT NULL,
        plate TEXT NOT NULL,
        kind TEXT NOT NULL,  -- park(선 것을 보고) / morning(다음 날 아침)
        at TEXT,
        park_since TEXT,
        lat REAL, lon REAL,
        sent INTEGER,        -- 보낸 DM 수(-1 = 혼자라 자동, 0 = 스위치 꺼져 기록만)
        PRIMARY KEY (day, plate, kind)
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


def _fmt(t: datetime) -> str:
    return t.strftime('%Y-%m-%d %H:%M:%S')


def _parse(s: str) -> Optional[datetime]:
    try:
        return datetime.strptime(s or '', '%Y-%m-%d %H:%M:%S')
    except ValueError:
        return None


def meters(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    r = 6371000.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = p2 - p1, math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def active() -> bool:
    """감지가 실제로 DM 까지 보내는 중인지 — 저녁 일괄 묻기가 쉴지 정할 때."""
    from dashboard.services import uplus_api as up
    from dashboard.services import vehicle_assign as va
    return bool(END_ASK_ON and va.CAR_DM_ON and up.enabled('status'))


# ── 위치 받아 두기 ─────────────────────────────────────────────────────────
def track(statuses: List[Dict], now: datetime) -> int:
    """차 상태 목록으로 vehicle_park 를 고친다. 돌려주는 값 = 움직인 차 수."""
    from dashboard.services import vehicle_assign as va
    plates = {c['plate'] for c in va.fleet()}
    today = now.date().isoformat()
    moved = 0
    with _conn() as conn:
        for s in statuses:
            plate, lat, lon = s.get('plate') or '', s.get('lat'), s.get('lon')
            if plate not in plates or lat is None or lon is None:
                continue
            at = s.get('at')
            # 단말이 마지막으로 보낸 시각 — 시동을 끄면 보내기를 멈춰 '선 시각'에 가깝다. 미래·이상값이면 지금.
            stop_at = at if isinstance(at, datetime) and at <= now + timedelta(minutes=5) else now
            row = conn.execute('SELECT * FROM vehicle_park WHERE plate = ?', (plate,)).fetchone()
            if not row:
                conn.execute('INSERT INTO vehicle_park (plate, lat, lon, since, seen_at, moved_day) VALUES (?,?,?,?,?,?)',
                             (plate, lat, lon, _fmt(stop_at), _fmt(now), ''))
                continue
            if meters(row['lat'], row['lon'], lat, lon) > PARK_RADIUS_M:
                last_seen = _parse(row['seen_at']) or now
                since = stop_at if stop_at > last_seen else now     # 지난번 본 뒤에 섰을 것
                conn.execute('UPDATE vehicle_park SET lat = ?, lon = ?, since = ?, seen_at = ?, moved_day = ? WHERE plate = ?',
                             (lat, lon, _fmt(since), _fmt(now), today, plate))
                moved += 1
            else:
                conn.execute('UPDATE vehicle_park SET seen_at = ? WHERE plate = ?', (_fmt(now), plate))
    return moved


# ── 운행 종료 보기 ────────────────────────────────────────────────────────
def _confirmed(day: str) -> bool:
    from dashboard.services import vehicle_assign as va
    return bool((va._day_info(day) or {}).get('confirmed_at'))


def ended_cars(now: datetime) -> List[Dict]:
    """지금 운행 종료로 볼 차 — 오늘 움직였고, 15시 뒤에 선 자리에서 45분 넘게 그대로이고, 아직 답·감지 기록이 없는 차.
    오늘 확정한 배정에 탄 사람이 있는 차만."""
    from dashboard.services import vehicle_assign as va
    from dashboard.services import vehicle_driver as vd
    day = now.date().isoformat()
    if not _confirmed(day):
        return []
    rows = va.car_rows(day)
    out = []
    with _conn() as conn:
        parks = {r['plate']: dict(r) for r in conn.execute('SELECT * FROM vehicle_park').fetchall()}
        done = {r['plate'] for r in conn.execute("SELECT plate FROM vehicle_end_ask WHERE day = ? AND kind = 'park'",
                                                (day,)).fetchall()}
    for plate, r in rows.items():
        riders = va.riders(r)
        p = parks.get(plate)
        if not riders or not p or plate in done or p.get('moved_day') != day:
            continue
        since = _parse(p.get('since') or '')
        if not since or since.date().isoformat() != day or since.strftime('%H:%M') < END_FROM:
            continue
        if now - since < timedelta(minutes=PARK_MIN):
            continue
        if vd.answer(day, plate):
            continue
        out.append({'day': day, 'plate': plate, 'riders': riders, 'since': p['since'], 'lat': p['lat'], 'lon': p['lon']})
    return out


def _record(day: str, plate: str, kind: str, sent: int, park: Optional[Dict] = None) -> None:
    park = park or {}
    with _conn() as conn:
        conn.execute('INSERT OR REPLACE INTO vehicle_end_ask (day, plate, kind, at, park_since, lat, lon, sent) '
                     'VALUES (?,?,?,?,?,?,?,?)',
                     (day, plate, kind, _fmt(datetime.now()), park.get('since'), park.get('lat'), park.get('lon'), sent))


def poll(now: Optional[datetime] = None, statuses: Optional[List[Dict]] = None) -> Dict:
    """스케줄러(10분마다) — 위치 받아 두고, 운행 종료로 보이는 차의 운전자를 묻는다."""
    from dashboard.services import uplus_api as up
    from dashboard.services import vehicle_assign as va
    from dashboard.services import vehicle_driver as vd
    now = now or datetime.now()
    if statuses is None:
        if not up.enabled('status'):
            return {'skipped': 'no_api'}
        if not (POLL_HOURS[0] <= now.hour < POLL_HOURS[1]):
            return {'skipped': 'hours'}
        try:
            statuses = up.car_status()
        except Exception as e:
            logger.warning(f'[운행종료] U+ 차량 상태 조회 실패: {e}')
            return {'error': str(e)}
        try:        # 받은 김에 차량 상태도 저장(지금 위치·차량 관리 — API 아끼기)
            from dashboard.services import vehicle_api_sync as vas
            vas.save_state(statuses, now)
        except Exception as e:
            logger.warning(f'[운행종료] 차량 상태 저장 실패: {e}')
    moved = track(statuses, now)
    ends = ended_cars(now)
    asked = []
    send = END_ASK_ON and va.CAR_DM_ON
    for e in ends:
        cars = {c['plate']: c for c in va.fleet(True)}
        if send:
            label = (cars.get(e['plate']) or {}).get('label') or e['plate']
            n = vd.ask_car(e['day'], e['plate'], e['riders'],
                           guess=vd.guess_driver(e['day'], e['plate'], e['riders']),
                           lead=f"{va.day_label(e['day'])} *{label}* ({e['plate']}) 운행을 마친 것 같아요 "
                                f"({e['since'][11:16]}부터 서 있음).")
        else:
            n = 0
        _record(e['day'], e['plate'], 'park', n, e)
        asked.append({'plate': e['plate'], 'since': e['since'], 'sent': n})
    if asked:
        logger.info(f"[운행종료] {now:%m-%d %H:%M} 운행 종료 {len(asked)}대{'' if send else '(기록만 — DM 꺼짐)'}: {asked}")
    return {'cars': len(statuses), 'moved': moved, 'ended': asked}


# ── 다음 날 아침 ─────────────────────────────────────────────────────────
def _prev_day(today: date) -> Optional[str]:
    """어제 — 일요일은 배정이 없어 월요일 아침엔 토요일까지 거슬러 본다(확정한 날만)."""
    for back in (1, 2):
        d = (today - timedelta(days=back)).isoformat()
        if _confirmed(d):
            return d
    return None


def remind_morning(today: Optional[date] = None) -> Dict:
    """월~토 09:10 — 어제 운전자 답이 없는 차. 안 물은 차는 묻고, 물었는데 답 없는 차는 한 번 더(한 번만)."""
    from dashboard.services import vehicle_assign as va
    from dashboard.services import vehicle_driver as vd
    if not active():
        return {'skipped': 'off'}
    today = today or date.today()
    day = _prev_day(today)
    if not day:
        return {'skipped': 'no_day'}
    cars = {c['plate']: c for c in va.fleet(True)}
    with _conn() as conn:
        done = {r['plate'] for r in conn.execute("SELECT plate FROM vehicle_end_ask WHERE day = ? AND kind = 'morning'",
                                                (day,)).fetchall()}
    sent = auto = 0
    for plate, r in va.car_rows(day).items():
        riders = va.riders(r)
        if not riders or vd.answer(day, plate) or plate in done:
            continue
        label = (cars.get(plate) or {}).get('label') or plate
        n = vd.ask_car(day, plate, riders, guess=vd.guess_driver(day, plate, riders),
                       lead=f"{va.day_label(day)} *{label}* ({plate}) 운전자를 아직 못 받았어요.", again=True)
        _record(day, plate, 'morning', n)
        if n < 0:
            auto += 1
        else:
            sent += n
    if sent or auto:
        logger.info(f'[운행종료] {day} 아침 다시 묻기 DM {sent}건, 혼자 탄 차 {auto}대 자동')
    return {'day': day, 'sent': sent, 'auto': auto}


# ── 경영지원 — 운전자 미확인 ──────────────────────────────────────────────
def missing(start: str, end: str) -> List[Dict]:
    """기간 안 확정한 날 중 운전자 답이 없는 차 — [{'day', 'plate', 'car', 'riders', 'asked'}] (오래된 날 먼저).
    오늘은 아직 운행 중일 수 있어 뺀다."""
    from dashboard.services import vehicle_assign as va
    from dashboard.services import vehicle_driver as vd
    who = va._who_map()
    cars = {c['plate']: c for c in va.fleet(True)}
    today = date.today().isoformat()
    out = []
    d, last = date.fromisoformat(start), date.fromisoformat(end)
    while d <= last:
        day = d.isoformat()
        d += timedelta(days=1)
        if day >= today or not _confirmed(day):
            continue
        with _conn() as conn:
            asked = {r['plate']: r['kind'] for r in conn.execute(
                'SELECT plate, kind FROM vehicle_end_ask WHERE day = ? ORDER BY kind DESC', (day,)).fetchall()}
        for plate, r in va.car_rows(day).items():
            riders = va.riders(r)
            if not riders or vd.answer(day, plate):
                continue
            out.append({'day': day, 'plate': plate, 'car': (cars.get(plate) or {}).get('label') or plate,
                        'riders': [va.display(c, who) for c in riders], 'asked': asked.get(plate, '')})
    return out
