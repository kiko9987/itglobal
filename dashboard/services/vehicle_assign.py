"""차량 배정 (2026-11-20 부장님) — 다음 날 어느 차를 누가 가지고 나가는지.

지금까지는 빈승정 대리(SJ)가 오후 4~5시 카톡 단체방에 'YM+LSH 레이 / JW+LJW+KMW 8021 / SJ 전기트럭'처럼
올렸다. 레이가 3대라 어느 레이인지 남지 않고, 운행일지(U+커넥트)의 운전자 칸도 비어 있어 누가 탔는지 알 수 없었다.
→ 카톡 배정을 그만두고 이 화면에서 배정한다(부장님 11-20: 카톡 글 붙여넣기는 카톡을 계속 쓰게 하므로 넣지 않는다).

- 차 기준으로 사람을 붙인다(한 차에 여러 명, 맨 앞 사람 = 운행기록부 '사용자').
- 공용 차(레이·쏘렌토·포터·스파크·모닝·카니발)는 매일 배정(카니발은 평소 YG 차지만 목록엔 넣는다 — 부장님 11-20).
- 파나메라(JW)·벤츠(김영숙)·랜드로버는 현장에 쓰지 않아 배정 화면에서 아예 뺀다(usage='off', 부장님 11-20).
  목록(vehicle_fleet)엔 남겨 두고 나중에 km 입력·차량 정보에서 쓴다. FIXED(평소 사용자 + 빌려 쓰기)는
  지금 쓰는 차가 없지만 화면은 그대로 둔다.
- 확정하면 탄 사람마다 슬랙 DM. 확정 뒤 고치면 바뀐 사람에게만 다시 DM.
- /일정확정이 끝나면 배정하는 사람(설정 vehicle_assigners, 기본 SJ)에게 '내일 차량 배정해 주세요' DM.
- 그날 방문 일정이 있는데 아직 차가 없는 사람을 화면 위에 보여 준다(리드 시트 방문 예정일·영업 담당자).
"""
from __future__ import annotations

import json
import os
import re
import urllib.request
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from typing import Dict, Iterable, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

# 자차(11-23 부장님: 차가 모자라면 자기 차로 이동 — 11/24 JH) — 배정 표에 번호판 대신 '자차' 줄로 날짜마다 남긴다.
#   차 목록(vehicle_fleet)에는 없고 배정 화면 맨 끝에 칸 하나로만 보인다. 운행일지·U+ 대조·저녁 운전자 묻기는 이 줄을 뺀다.
OWN = '자차'
OWN_CAR = {'plate': OWN, 'label': '자차', 'model': '자차', 'fuel': '', 'corp': '개인', 'usage': 'own', 'fixed_user': '',
           'needs_log': 0, 'uplus': 0}
SHARED, FIXED, OFF = 'shared', 'fixed', 'off'      # OFF = 배정 화면에 안 나옴(현장에 안 쓰는 차)

# 법인 차량(11-20 U+커넥트 차량 등록 화면 + 공유 드라이브 ★법인 차량관리폴더 기준).
# (번호판, 부르는 이름, 차종, 연료, 법인, 공용/고정, 고정 사용자, 운행일지 필요, U+ 단말)
_SEED = [
    ('237허5358', '레이 5358', '레이', '가솔린', '글로벌', SHARED, '', 1, 1),
    ('231허6854', '레이 6854', '레이', '가솔린', '글로벌', SHARED, '', 1, 1),
    ('231허6873', '레이 6873', '레이', '가솔린', '글로벌', SHARED, '', 1, 1),
    ('03모8021', '쏘렌토 8021', '쏘렌토', '디젤', '글로벌', SHARED, '', 1, 1),
    ('16소4041', '쏘렌토 4041', '쏘렌토', '디젤', '글로벌', SHARED, '', 1, 1),
    ('817루4904', '포터 4904', '포터', '디젤', '글로벌', SHARED, '', 1, 1),
    ('821저4919', '포터 4919 (LPG)', '포터', 'LPG', '글로벌', SHARED, '', 1, 1),
    ('90머9131', '포터 9131 (전기)', '포터 일렉트릭', '전기', '글로벌', SHARED, '', 1, 1),
    ('53우1783', '스파크 1783', '스파크', '가솔린', '개인(MW)', FIXED, 'MW', 0, 1),    # MW 개인 차 — 회사가 유류비 지원, U+ 단말 연결(11-22)
    ('60거5484', '모닝 5484', '모닝', '가솔린', '글로벌', SHARED, '', 1, 1),
    ('143고9687', '카니발 9687', '카니발', '디젤', '글로벌', FIXED, 'YG', 0, 1),   # YG 차 — 차가 모자랄 때만 빌려 씀(11-22)
    ('204두8988', '파나메라 8988', '파나메라', '', '글로벌', OFF, 'JW', 1, 0),      # 현장 안 씀 — 배정에서 뺌
    ('260나7350', '벤츠 7350', '벤츠', '', '그룹', OFF, '김영숙', 1, 0),
    ('133가6633', '랜드로버 6633', '랜드로버', '디젤', '플랜트', OFF, '', 1, 0),
]

_DDL = [
    '''CREATE TABLE IF NOT EXISTS vehicle_fleet (
        plate TEXT PRIMARY KEY,
        label TEXT NOT NULL,
        model TEXT,
        fuel TEXT,
        corp TEXT,
        usage TEXT NOT NULL DEFAULT 'shared',
        fixed_user TEXT,
        needs_log INTEGER NOT NULL DEFAULT 1,
        uplus INTEGER NOT NULL DEFAULT 1,
        active INTEGER NOT NULL DEFAULT 1,
        sort_no INTEGER NOT NULL DEFAULT 0,
        updated_by TEXT,
        updated_at TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicle_assign (
        day TEXT NOT NULL,
        plate TEXT NOT NULL,
        people TEXT NOT NULL DEFAULT '',
        note TEXT,
        updated_by TEXT,
        updated_at TEXT,
        PRIMARY KEY (day, plate)
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicle_assign_day (
        day TEXT PRIMARY KEY,
        confirmed_by TEXT,
        confirmed_at TEXT,
        requested_at TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicle_assign_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        day TEXT NOT NULL,
        plate TEXT NOT NULL,
        before TEXT,
        after TEXT,
        by TEXT,
        at TEXT
    )''',
]
_table_ready = False
# 차량 DM 스위치 — 11-23 꺼 둠("차량은 더 다듬어야 하니"), 11-24 켬(부장님: 오늘부터 방문 일정 → 차량 배정 라이브).
#   켜면: 확정한 사람들에게 가는 차량 DM·바뀐 사람 DM, 배정 담당 요청·방문 일정 변경 DM.
#   저녁 운전자 묻기는 따로 vehicle_driver.DRIVER_ASK_ON(운행일지가 아직 개발 중이라 꺼 둠).
CAR_DM_ON: bool = True

# 빈승정 대리 + JW(11-24 부장님 — 방문 일정을 확정하고 바로 차량 배정 탭에서 차를 넣음). 인사 관리자는 따로 늘 가능
_DEFAULT_ASSIGNERS = ('sj@itg-aircon.com', 'jw@itg-aircon.com')


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
            # 온라인 당번(캔버스2 '온라인' 섹션, /일정확정 때 저장 — 11-21)
            cols = {r[1] for r in conn.execute('PRAGMA table_info(vehicle_assign_day)').fetchall()}
            if 'online_duty' not in cols:
                conn.execute('ALTER TABLE vehicle_assign_day ADD COLUMN online_duty TEXT')
            # 한 차를 오전·오후 다른 사람이 운전(11-22 부장님: 4041 대표님 오전 → 승정 대리 오후)
            acols = {r[1] for r in conn.execute('PRAGMA table_info(vehicle_assign)').fetchall()}
            for col in ('shift_driver', 'shift_at', 'pm_people'):   # pm_people = 오후(없으면 오전과 같음, 11-22)
                if col not in acols:
                    conn.execute(f'ALTER TABLE vehicle_assign ADD COLUMN {col} TEXT')
            for i, s in enumerate(_SEED):
                conn.execute('INSERT OR IGNORE INTO vehicle_fleet (plate, label, model, fuel, corp, usage, fixed_user, '
                             'needs_log, uplus, active, sort_no, updated_by, updated_at) '
                             'VALUES (?,?,?,?,?,?,?,?,?,1,?,?,?)', (*s, i, 'seed', _now()))
            # 11-22 부장님: 카니발=YG 차(모자랄 때만 빌려 씀), 스파크=MW 개인 차(유류비 지원) → 고정 차로.
            #   처음 시드 그대로인 줄만 바꾼다(화면에서 고친 줄은 그대로).
            for plate, corp, user, need in (('143고9687', '글로벌', 'YG', 0), ('53우1783', '개인(MW)', 'MW', 0)):
                conn.execute("UPDATE vehicle_fleet SET usage = ?, fixed_user = ?, corp = ?, needs_log = ?, "
                             "updated_by = 'seed-1122', updated_at = ? WHERE plate = ? AND updated_by = 'seed'",
                             (FIXED, user, corp, need, _now(), plate))
            conn.commit()
            _table_ready = True
        yield conn
        conn.commit()


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


# ── 권한 ────────────────────────────────────────────────────────────────
def assigners() -> List[str]:
    from dashboard.services.hr import _setting_list
    return _setting_list('vehicle_assigners', _DEFAULT_ASSIGNERS)


def can_assign(email: str) -> bool:
    from dashboard.services import hr
    email = (email or '').strip().lower()
    # 관리자(admin·super_admin)는 모두(11-25 부장님 '어드민한테는 다 열어줘')
    return bool(email) and (hr.is_hr_admin(email) or email in assigners() or hr.is_admin_email(email))


# ── 차량·직원 ───────────────────────────────────────────────────────────
def fleet(include_inactive: bool = False) -> List[Dict]:
    """배정할 차(운용 중 + 배정 화면에 나오는 차). include_inactive=True 면 뺀 차까지 전부."""
    q = ('SELECT * FROM vehicle_fleet' + ('' if include_inactive else f" WHERE active = 1 AND usage != '{OFF}'")
         + ' ORDER BY sort_no, plate')
    with _conn() as conn:
        return [dict(r) for r in conn.execute(q).fetchall()]


def update_car(plate: str, values: Dict, by: str) -> Dict:
    """공용/고정·고정 사용자·부르는 이름만 고친다."""
    allowed = {}
    if 'usage' in values:
        if values['usage'] not in (SHARED, FIXED, OFF):
            raise ValueError('공용·고정·배정 안 함 중에서 골라 주세요.')
        allowed['usage'] = values['usage']
    if 'fixed_user' in values:
        allowed['fixed_user'] = (values.get('fixed_user') or '').strip()[:20]
    if 'label' in values:
        lab = (values.get('label') or '').strip()[:20]
        if not lab:
            raise ValueError('부르는 이름을 적어 주세요.')
        allowed['label'] = lab
    if not allowed:
        raise ValueError('고칠 내용이 없어요.')
    with _conn() as conn:
        if not conn.execute('SELECT 1 FROM vehicle_fleet WHERE plate = ?', (plate,)).fetchone():
            raise ValueError('없는 차량이에요.')
        sets = ', '.join(f'{k} = ?' for k in allowed)
        conn.execute(f'UPDATE vehicle_fleet SET {sets}, updated_by = ?, updated_at = ? WHERE plate = ?',
                     (*allowed.values(), by, _now(), plate))
    return next(c for c in fleet(True) if c['plate'] == plate)


OFFICE_DEPTS = ('경영지원',)


def staff() -> List[Dict]:
    """배정할 수 있는 사람 — 재직 인사카드 중 이니셜(공사 코드 접미사)이 있는 현장 인원(사무직 제외).
    [{code, name, slack_id, dept}] (code = 이니셜)"""
    from dashboard.services import hr
    from dashboard.services.staff_directory import suffix_map
    from dashboard.services.staff_directory import non_sales_owners
    sm = suffix_map()
    office = set(non_sales_owners())
    rows = hr._rows("SELECT name, slack_id, department FROM hr_people WHERE status = '재직' ORDER BY name")
    out, seen = [], set()
    for r in rows:
        code = (sm.get(r['name']) or '').upper()
        # 사무직(경영지원·비영업 명단)은 현장 차를 안 탄다 — 샛별 등(부장님 11-20)
        if not code or code in seen or r['name'] in office or (r.get('department') or '') in OFFICE_DEPTS:
            continue
        seen.add(code)
        out.append({'code': code, 'name': r['name'], 'slack_id': r.get('slack_id') or '', 'dept': r.get('department') or ''})
    out.sort(key=lambda p: p['code'])
    return out


def _who_map() -> Dict[str, Dict]:
    return {p['code']: p for p in staff()}


def display(code: str, who: Optional[Dict[str, Dict]] = None) -> str:
    """이니셜 → 이름(모르면 그대로 — '김영숙'처럼 이름으로 적은 사람)."""
    who = who if who is not None else _who_map()
    p = who.get((code or '').upper())
    return p['name'] if p else code


# ── 날짜 ────────────────────────────────────────────────────────────────
def next_workday(today: Optional[date] = None) -> str:
    """기본 배정 날짜 = 내일(일요일이면 월요일)."""
    d = (today or date.today()) + timedelta(days=1)
    if d.weekday() == 6:
        d += timedelta(days=1)
    return d.isoformat()


_DAY_RE = re.compile(r'^\d{4}-\d{2}-\d{2}$')


def valid_day(s: str) -> bool:
    if not _DAY_RE.match(s or ''):
        return False
    try:
        date.fromisoformat(s)
        return True
    except ValueError:
        return False


# ── 배정 읽기·쓰기 ───────────────────────────────────────────────────────
def _split(people: str) -> List[str]:
    return [p for p in (people or '').split(',') if p]


HALF_AT = '13:00'   # 오전·오후 나눈 차 — 이 시각 이후 운행은 오후 사람이 운전(운행일지)


def _empty() -> Dict:
    return {'people': [], 'pm': None, 'note': ''}


def _row(people: List[str], note: str = '', pm: Optional[List[str]] = None) -> Dict:
    """people = 오전(나누지 않았으면 종일), pm = 오후(None 이면 오전과 같음)."""
    return {'people': list(people), 'pm': (list(pm) if pm is not None else None), 'note': note or ''}


def pm_of(row: Dict) -> List[str]:
    return row['pm'] if row.get('pm') is not None else row['people']


def riders(row: Dict) -> List[str]:
    """그 차에 하루 동안 탄 사람(오전 먼저, 오후에만 탄 사람 뒤)."""
    return list(dict.fromkeys(row['people'] + pm_of(row)))


def _day_rows(day: str) -> Dict[str, Dict]:
    with _conn() as conn:
        return {r['plate']: _row(_split(r['people']), r['note'],
                                 _split(r['pm_people']) if r['pm_people'] is not None else None)
                for r in conn.execute('SELECT plate, people, pm_people, note FROM vehicle_assign WHERE day = ?',
                                      (day,)).fetchall()}


def _day_info(day: str) -> Optional[Dict]:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM vehicle_assign_day WHERE day = ?', (day,)).fetchone()
    return dict(r) if r else None


BIG_TEAM = 3   # 이 인원 이상인 일정확정 팀은 방문 뒤 현장별로 갈라질 수 있다(11-22 부장님: JW+YM+TH+LJW → 8021·레이 2대)


def day_view(day: str, teams: Optional[List[Dict]] = None) -> Dict:
    """배정 화면 한 날짜 — 차량마다 사람·메모. 저장한 것만 보인다.
    초안(최근 탄 차·짝 기록·전날 배정 복사)은 없앴다 — 타는 차·인원이 매일 달라 맞지 않음(11-23 부장님).
    teams 는 예전 호출과 맞추려고 받기만 한다."""
    cars = fleet()
    rows = _day_rows(day)
    info = _day_info(day)
    out = []
    for c in cars + [OWN_CAR]:          # 자차 칸은 맨 끝(11-23)
        r = rows.get(c['plate'], _empty())
        out.append({**{k: c[k] for k in ('plate', 'label', 'model', 'fuel', 'corp', 'usage', 'fixed_user',
                                         'needs_log', 'uplus')},
                    'people': r['people'], 'pm': r.get('pm'), 'note': r['note']})
    return {'day': day, 'cars': out, 'draft_from': None, 'draft_kind': None, 'base': day_version(day),
            'confirmed_at': (info or {}).get('confirmed_at'), 'confirmed_by': (info or {}).get('confirmed_by')}


class StaleSave(Exception):
    """화면을 연 뒤 다른 사람이 그날 배정을 먼저 저장·확정함 — 그대로 저장하면 앞 사람 변경을 말없이 덮어쓴다."""


STALE_MSG = '다른 분이 먼저 고쳤어요. 새로 불러온 뒤 다시 해 주세요.'


def _version(conn, day: str) -> str:
    """그날 배정의 판(版) — 바뀔 때마다 늘어나는 로그 마지막 id + 확정 시각.
    updated_at(초 단위·지운 줄은 안 남음)은 같은 초에 두 번 저장하거나 차를 비우면 그대로라 쓰지 않는다."""
    r = conn.execute('SELECT MAX(id) FROM vehicle_assign_log WHERE day = ?', (day,)).fetchone()
    d = conn.execute('SELECT confirmed_at FROM vehicle_assign_day WHERE day = ?', (day,)).fetchone()
    return f"{(r[0] if r and r[0] else 0)}|{(d[0] if d and d[0] else '')}"


def day_version(day: str) -> str:
    """불러오기 응답의 base — 저장할 때 그대로 돌려받아 그 사이 남이 고쳤는지 본다(11-24 조용한 실패 점검 17번)."""
    with _conn() as conn:
        return _version(conn, day)


def _clean_rows(rows: Iterable[Dict], plates: set, codes: set) -> Dict[str, Dict]:
    """화면에서 온 배정 → {plate: {people, pm, note}}. 같은 사람이 같은 반나절(오전·오후)에 두 차면 ValueError.
    pm 이 없거나 오전과 똑같으면 '나누지 않음'(None)."""
    def clean(lst):
        out = []
        for p in lst or []:
            p = (p or '').strip()
            if not p:
                continue
            p = p.upper() if p.upper() in codes else p[:20]
            if p not in out:
                out.append(p)
        return out

    out: Dict[str, Dict] = {}
    am_at: Dict[str, str] = {}
    pm_at: Dict[str, str] = {}
    for r in rows or []:
        plate = (r.get('plate') or '').strip()
        if plate not in plates:
            raise ValueError(f'없는 차량이에요: {plate}')
        people = clean(r.get('people'))
        pm = clean(r.get('pm')) if r.get('pm') is not None else None
        if pm == people:
            pm = None
        for p in people:
            if p in am_at and am_at[p] != plate:
                raise ValueError(f'{p} 님이 오전에 두 차에 들어 있어요. 한 차만 골라 주세요.')
            am_at[p] = plate
        for p in (pm if pm is not None else people):
            if p in pm_at and pm_at[p] != plate:
                raise ValueError(f'{p} 님이 오후에 두 차에 들어 있어요. 한 차만 골라 주세요.')
            pm_at[p] = plate
        note = (r.get('note') or '').strip()[:60]
        if people or pm or note:
            out[plate] = _row(people, note, pm)
    return out


def save_day(day: str, rows: Iterable[Dict], by: str, confirm: bool = False, notify: bool = True,
             base: Optional[str] = None) -> Dict:
    """배정 저장. confirm=True 면 확정(DM 보냄). 확정된 날을 고치면 바뀐 사람에게만 DM.
    notify=False 면 확정된 날을 고쳐도 DM 을 보내지 않는다(메모 오타 등 — 화면 [알리지 않고 저장], 11-21).
    base = 화면이 불러올 때 받은 판(day_version). 그 사이 다른 사람이 저장·확정했으면 StaleSave —
      배정 담당 두 명(SJ·JW)이 같은 날을 같이 고치면 화면이 모든 차 줄을 통째로 보내 나중 저장이 앞 사람 변경을
      말없이 덮어쓰고, 확정 뒤라면 '빠졌어요' DM 까지 나가던 것(11-24 조용한 실패 점검 17번).
      None(base 를 안 보내는 옛 화면 — 재시작 직후 열려 있던 것)은 예전처럼 그대로 저장.
    돌려주는 값: {'changed': [...], 'notified': n, 'confirmed': bool, 'base': 저장 뒤 새 판}"""
    if not valid_day(day):
        raise ValueError('날짜가 올바르지 않아요.')
    cars = {c['plate']: c for c in fleet()}
    cars[OWN] = OWN_CAR
    who = _who_map()
    new = _clean_rows(rows, set(cars), set(who))
    old = _day_rows(day)
    info = _day_info(day)
    was_confirmed = bool(info and info.get('confirmed_at'))
    if confirm or (was_confirmed and notify):
        bad = off_riders(day, new)
        if bad:
            raise ValueError(f"휴무인 {', '.join(bad)} 님이 차에 있어요. 빼야 확정·알림이 돼요.")
    now = _now()
    changed = []
    with _conn() as conn:     # 연결이 프로세스 안에서 잠금(RLock)을 쥐므로 판 비교 → 쓰기가 한 덩어리
        if base is not None and _version(conn, day) != str(base):
            logger.info(f'[차량배정] {day} 저장 거절 — 화면을 연 뒤 다른 분이 먼저 고침(by {by})')
            raise StaleSave(STALE_MSG)
        for plate in set(old) | set(new):
            o = old.get(plate, _empty())
            n = new.get(plate, _empty())
            if o == n:
                continue
            changed.append(plate)
            conn.execute('INSERT INTO vehicle_assign_log (day, plate, before, after, by, at) VALUES (?,?,?,?,?,?)',
                         (day, plate, json.dumps(o, ensure_ascii=False), json.dumps(n, ensure_ascii=False), by, now))
            if n['people'] or n['pm'] or n['note']:
                conn.execute('INSERT OR REPLACE INTO vehicle_assign (day, plate, people, pm_people, note, updated_by, updated_at) '
                             'VALUES (?,?,?,?,?,?,?)',
                             (day, plate, ','.join(n['people']), ','.join(n['pm']) if n['pm'] is not None else None,
                              n['note'], by, now))
            else:
                conn.execute('DELETE FROM vehicle_assign WHERE day = ? AND plate = ?', (day, plate))
        if confirm and not was_confirmed:
            conn.execute('INSERT INTO vehicle_assign_day (day, confirmed_by, confirmed_at) VALUES (?,?,?) '
                         'ON CONFLICT(day) DO UPDATE SET confirmed_by = excluded.confirmed_by, '
                         'confirmed_at = excluded.confirmed_at', (day, by, now))
        new_base = _version(conn, day)     # 같은 사람이 연달아 저장해도 막히지 않게 새 판을 돌려준다
    notified, failed, retry = 0, [], []
    if confirm and not was_confirmed:
        notified = notify_people(day, cars, new, who, everyone=True, failed=failed, retry=retry)
    elif was_confirmed and changed and notify:
        notified = notify_changes(day, cars, old, new, who, failed=failed, retry=retry)
    if failed:
        # 못 받은 사람은 화면에 이름으로 — 이미 '확정'이라 다시 눌러도 안 간다(11-24 조용한 실패 점검 3번)
        log = logger.error if not notified else logger.warning
        log(f"[차량배정] {day} DM 못 받은 사람 {len(failed)}명: {', '.join(failed)} (받은 사람 {notified}명)",
            **({'extra': {'force_alert': True}} if not notified else {}))
    return {'changed': changed, 'notified': notified, 'failed': failed, 'retry': retry,
            'confirmed': confirm or was_confirmed, 'dm_off': not CAR_DM_ON, 'base': new_base}


def resend(day: str, codes: Iterable[str]) -> Dict:
    """[다시 보내기] — 못 받은 사람에게 그날 지금 배정을 다시 보낸다(확정한 날만, 11-24 조용한 실패 점검 추가분)."""
    if not valid_day(day):
        raise ValueError('날짜가 올바르지 않아요.')
    info = _day_info(day)
    if not (info and info.get('confirmed_at')):
        raise ValueError('아직 확정하지 않은 날이에요.')
    cars = {c['plate']: c for c in fleet()}
    cars[OWN] = OWN_CAR
    failed: List[str] = []
    retry: List[str] = []
    n = notify_people(day, cars, _day_rows(day), _who_map(), only={str(c).upper() for c in codes or []},
                      failed=failed, retry=retry)
    if failed:
        logger.error(f"[차량배정] {day} 다시 보내기에도 못 받은 사람: {', '.join(failed)}", extra={'force_alert': True})
    return {'notified': n, 'failed': failed, 'retry': retry}


def person_car(day: str, code: str) -> Optional[Dict]:
    """그날 이 사람이 탈 차(오전 차 먼저, 오전에 없으면 오후 차)."""
    rows = _day_rows(day)
    for half in ('am', 'pm'):
        for plate, r in rows.items():
            if code in (r['people'] if half == 'am' else pm_of(r)):
                return {'plate': plate, 'half': half, **r}
    return None


def assigned_people(start: str, end: str) -> Dict[Tuple[str, str], List[str]]:
    """운행일지용 — (날짜, 번호판) → 탄 사람 이니셜(맨 앞 = 사용자). 배정 없는 고정 차는 고정 사용자."""
    return {k: riders(v) for k, v in assigned_halves(start, end).items() if riders(v)}


def assigned_halves(start: str, end: str) -> Dict[Tuple[str, str], Dict]:
    """운행일지용 — (날짜, 번호판) → {'people': 오전, 'pm': 오후 또는 None}. 13:00(HALF_AT) 이후 운행은 오후 사람."""
    with _conn() as conn:
        rows = conn.execute('SELECT day, plate, people, pm_people FROM vehicle_assign WHERE day BETWEEN ? AND ? AND plate != ?',
                            (start, end, OWN)).fetchall()
    return {(r['day'], r['plate']): _row(_split(r['people']), '', _split(r['pm_people']) if r['pm_people'] is not None else None)
            for r in rows}


def fixed_users() -> Dict[str, str]:
    return {c['plate']: c['fixed_user'] for c in fleet(True) if c['usage'] == FIXED and c['fixed_user']}


def visitors(day: str, leads: Iterable[dict]) -> List[Dict]:
    """그날 방문 일정이 있는 사람(리드 시트 방문 예정일·영업 담당자) — 배정 화면에서 '차 없는 사람'을 짚는 데 쓴다.
    [{code, name, visits, places:[{label, addr}]}] — 이니셜이 없는 이름(외부 인원 등)은 뺀다.
    places = 그 사람이 갈 곳(이름·방문 주소) — 배정 화면 코스 카드에 보인다.
    이름은 방문 일정과 같은 도로명 기준(시·구·읍·면, '용인 수지구'). 행안부로 찾은 지번 동('용인 풍덕천동')은
    구주소라 쓰지 않는다(11-24 부장님: 방문 일정은 도로명) — 동은 운행일지 운전자 추정에만.
    고객명은 넣지 않는다(차 고르는 데는 위치만 있으면 됨)."""
    from dashboard.services import vehicle_log as vl
    from dashboard.services.visit_plan import _where
    by_name = {p['name']: p for p in staff()}
    count: Dict[str, int] = {}
    places: Dict[str, List[Dict]] = {}
    for v in vl.visits_by_day(leads, day, day).get(day, []):
        for n in v['who']:
            if n in by_name:
                count[n] = count.get(n, 0) + 1
                mine = places.setdefault(n, [])
                if all(p['addr'] != v['addr'] for p in mine):
                    mine.append({'label': _where(v['addr'], vl)['place'] or v['addr'][:20], 'addr': v['addr']})
    return [{'code': by_name[n]['code'], 'name': n, 'visits': c, 'places': places.get(n, [])}
            for n, c in sorted(count.items())]


def day_off(day: str) -> Dict[str, List[Dict]]:
    """그날 쉬는 사람 — 그룹웨어 휴가(승인된 것)와 토요일 휴무표(11-21 부장님: 캔버스 말고 그룹웨어에서).
    {'off': [{code, name, why}], 'half': [{code, name, why}]} — 종일은 명단에서 '휴무'로 빼고, 반차는 이름 옆 꼬리표만.
    배정 명단(staff)에 없는 사람은 뺀다."""
    from dashboard.services import leave_view, sat_off
    by_name = {p['name']: p['code'] for p in staff()}
    d = date.fromisoformat(day)
    off: Dict[str, Dict] = {}
    half: Dict[str, Dict] = {}
    for e in leave_view.off_on(d):
        code = by_name.get(e['name'])
        if not code:
            continue
        if e.get('half'):
            half.setdefault(code, {'code': code, 'name': e['name'], 'why': f"{e['half']} 반차"})
        else:
            off[code] = {'code': code, 'name': e['name'], 'why': e.get('type') or '휴가'}
    if d.weekday() == 5:
        month = day[:7]
        sure = sat_off.month_row(month).get('status') == sat_off.CONFIRMED
        for name in sat_off.picks(month).get(day, []):
            code = by_name.get(name)
            if code and code not in off:
                off[code] = {'code': code, 'name': name, 'why': '스케줄 근무 쉬는 날' + ('' if sure else '(확정 전)'),
                             'soft': not sure}     # 확정 전 휴무표는 바뀔 수 있어 방문 일정에서 막지 않고 경고만
    half = {c: v for c, v in half.items() if c not in off}
    return {'off': sorted(off.values(), key=lambda x: x['name']), 'half': sorted(half.values(), key=lambda x: x['name'])}


def record_online_duty(day: Optional[str], codes: Iterable[str], shifts: Optional[Dict[str, str]] = None) -> bool:
    """/일정확정 때 그날 온라인 당번(캔버스2 '온라인' 섹션 — 사무실에 남는 문의 당번)을 저장(11-21 부장님).
    다시 확정하면 덮어쓴다. 배정 화면은 이 사람들을 '차 없음'에서 빼고, 방문도 있는 겸업이면 꼬리표만 붙인다."""
    day = day if day and valid_day(day) else next_workday()     # 방문 DM 이 없던 날도(request_assignment 와 같은 기준)
    shifts = shifts or {}
    rows = [{'code': c, 'shift': shifts.get(c, '')} for c in dict.fromkeys(codes or []) if c]
    with _conn() as conn:
        conn.execute('INSERT INTO vehicle_assign_day (day, online_duty) VALUES (?,?) '
                     'ON CONFLICT(day) DO UPDATE SET online_duty = excluded.online_duty',
                     (day, json.dumps(rows, ensure_ascii=False)))
    return True


def online_duty(day: str) -> List[Dict]:
    """그날 온라인 당번 [{code, name, shift}] — /일정확정 전이면 []. 배정 명단에 없는 사람은 뺀다."""
    raw = (_day_info(day) or {}).get('online_duty')
    try:
        rows = json.loads(raw) if raw else []
    except ValueError:
        rows = []
    names = {p['code']: p['name'] for p in staff()}
    return [{'code': r['code'], 'name': names[r['code']], 'shift': r.get('shift') or ''}
            for r in rows if r.get('code') in names]


_NOTE_SEP = re.compile(r'\s*─{2,}\s*')
_NOTE_HEAD = re.compile(r'^\[[^\]]{0,60}\]\s*')


def _latest_note(text: str) -> str:
    """상담 내용(K열)의 마지막 내용 — '내용 ─── [11.17 09:18 MJ · 방문 완료] 다음 내용' 처럼 쌓인다.
    머리말([…])만 있고 내용이 빈 회차(방문 완료 표시 등)는 건너뛴다."""
    parts = [_NOTE_HEAD.sub('', p).strip() for p in _NOTE_SEP.split(str(text or '').strip())]
    parts = [p for p in parts if p]
    return parts[-1] if parts else ''


_SCHED_CATS = ('기타', '거래처', '온라인 방문')
_SCHED_TITLES = {'기타': '기타 (AS·계약서 작성 등)', '거래처': '거래처', '온라인 방문': '온라인 방문'}


def day_schedule(day: str, leads: Iterable[dict]) -> List[Dict]:
    """그날 방문 일정 — 캔버스1(봇 자동, visit_canvas_sync)과 같은 묶음·표시, 시트에서 바로 그린다(11-21 부장님).
    캔버스1은 오늘 이후·방문 완료 전만 싣지만, 여기는 고른 날짜에 걸친 방문을 다 보여 준다(지난 날 포함).
    가지 않은 리드(유선·부재·취소·보류)는 뺀다. 줄마다 영업 담당자(/일정확정이 넣은 값, 없으면 '')를 붙인다.
    [{'cat', 'title', 'items': [{lead_no, date, phone, addr, text, who, marks:[...]}]}]"""
    from dashboard.services import vehicle_log as vl
    from dashboard.services import visit_assignment_sync as vas
    from dashboard.services import visit_canvas_sync as vcs
    d = date.fromisoformat(day)
    buckets: Dict[str, List[Tuple[tuple, Dict]]] = {c: [] for c in _SCHED_CATS}
    for lead in leads:
        vd = lead.get('방문 예정일')
        st, en = vas._parse_visit_date_start(vd), vas._parse_visit_date_end(vd)
        if not st or not (st <= d <= (en if en and en >= st else st)):
            continue
        status = str(lead.get('상태') or '').strip()
        if any(w in status for w in vl._NOT_VISITED):
            continue
        cat = vcs._categorize(lead) or '온라인 방문'
        raw_phone = str(lead.get('고객 연락처') or '').strip()
        text = re.sub(r'\s*\n\s*', ' ', _latest_note(str(lead.get('상담 내용') or lead.get('문의 내용') or '')))
        who = [w.strip() for w in re.split(r'[,/·\s]+', str(lead.get('영업 담당자') or '')) if vl._HAS_NAME.search(w)]
        marks = []      # '공사확정' 꼬리표는 뺌(11-24 부장님: 차량 배정·방문 일정 모두 필요 없음)
        if '본인 방문 필수' in str(lead.get('본인 방문 여부') or ''):
            marks.append('본인 방문 필수')
        # 플랫폼은 줄마다 꼬리표로(당근만 붙이던 것 → 모두, 11-21 부장님). 기타 방문(ETC-)은 '기타'
        platform = str(lead.get('플랫폼') or '').strip() or ('기타' if cat == '기타' else '')
        item = {'lead_no': str(lead.get('리드 No') or '').strip(), 'date': vcs._fmt_visit_date(vd), 'platform': platform,
                'phone': raw_phone if re.search(r'\d', raw_phone) else '',
                'addr': re.sub(r'\s*\n\s*', ' ', str(lead.get('방문 주소') or '').strip()),
                'text': text[:200] + ('…' if len(text) > 200 else ''), 'who': who, 'marks': marks}
        buckets[cat].append((vcs._visit_date_sort_key(vd), item))
    out = []
    for c in _SCHED_CATS:
        if buckets[c]:
            items = [it for _, it in sorted(buckets[c], key=lambda x: (x[0], x[1]['lead_no']))]
            out.append({'cat': c, 'title': _SCHED_TITLES[c], 'items': items})
    return out


# ── 슬랙 ────────────────────────────────────────────────────────────────
_WEEK = '월화수목금토일'


def day_label(day: str) -> str:
    d = date.fromisoformat(day)
    return f'{d.month}/{d.day}({_WEEK[d.weekday()]})'


def _token() -> str:
    return (os.getenv('SLACK_VISIT_BOT_TOKEN', '').strip() or os.getenv('SLACK_BOT_TOKEN', '').strip())


def _page_url(day: str) -> str:
    base = ((os.getenv('GW_PUBLIC_URL') or '').strip() or 'https://pm.itg-aircon.com').rstrip('/')
    return f'{base}/vehicles/assign?day={day}'


def _dm(slack_id: str, text: str) -> bool:
    if not slack_id or not _token():
        return False
    try:
        body = json.dumps({'channel': slack_id, 'text': text, 'unfurl_links': False,
                           'unfurl_media': False}).encode('utf-8')
        req = urllib.request.Request('https://slack.com/api/chat.postMessage', data=body,
                                     headers={'Authorization': f'Bearer {_token()}',
                                              'Content-Type': 'application/json; charset=utf-8'})
        r = json.loads(urllib.request.urlopen(req, timeout=10).read().decode('utf-8'))
        if not r.get('ok'):
            logger.warning(f"[차량배정] DM 실패 {slack_id}: {r.get('error')}")
        return bool(r.get('ok'))
    except Exception as e:
        logger.warning(f'[차량배정] DM 예외 {slack_id}: {e}')
        return False


def car_rows(day: str) -> Dict[str, Dict]:
    """그날 회사 차 배정만(자차 줄 뺌) — 운전자 묻기·운행 종료처럼 실제 차가 있어야 하는 곳."""
    return {p: r for p, r in _day_rows(day).items() if p != OWN}


def _car_text(plate: str, cars: Dict[str, Dict]) -> str:
    if plate == OWN:
        return '*본인 차량*으로 이동'      # 화면 꼬리표와 같은 말(11-24)
    c = cars.get(plate) or {}
    return f"*{c.get('label') or plate}* ({plate})"


def _where(rows: Dict[str, Dict]) -> Dict[str, Dict]:
    """사람 → {'am': 번호판, 'pm': 번호판} (그 반나절에 차가 없으면 '')."""
    out: Dict[str, Dict] = {}
    for plate, r in rows.items():
        for c in r['people']:
            out.setdefault(c, {'am': '', 'pm': ''})['am'] = plate
        for c in pm_of(r):
            out.setdefault(c, {'am': '', 'pm': ''})['pm'] = plate
    return out


def _car_line(plate: str, people: List[str], code: str, cars: Dict[str, Dict], who: Dict[str, Dict]) -> str:
    others = [display(p, who) for p in people if p != code] if plate != OWN else []   # 자차는 같이 타는 게 아님
    return _car_text(plate, cars) + (f" · 함께 {', '.join(others)}" if others else '')


def _person_text(day: str, code: str, rows: Dict[str, Dict], cars: Dict[str, Dict], who: Dict[str, Dict]) -> str:
    w = _where(rows).get(code) or {'am': '', 'pm': ''}
    lines = []
    if w['am'] and w['am'] == w['pm'] and rows[w['am']].get('pm') is None:
        r = rows[w['am']]
        lines.append(f":car: {day_label(day)} 차량: {_car_text(w['am'], cars)}")
        others = [display(p, who) for p in r['people'] if p != code] if w['am'] != OWN else []
        if others:
            lines.append(f"함께 타는 사람: {', '.join(others)}")
        notes = [r['note']] if r.get('note') else []
    elif bool(w['am']) != bool(w['pm']) and (w['am'] or w['pm']) != OWN:
        # 차 한 대를 두 코스가 오전·오후로 나눠 씀(11-24) — '오전만' + 반대 반나절에 그 차를 쓰는 사람(차를 넘겨줄 상대)
        half = 'am' if w['am'] else 'pm'
        plate = w[half]
        r = rows[plate]
        mine = r['people'] if half == 'am' else pm_of(r)
        theirs = [p for p in (pm_of(r) if half == 'am' else r['people']) if p not in mine]
        label = '오전' if half == 'am' else '오후'
        lines.append(f":car: {day_label(day)} 차량: {_car_text(plate, cars)} {label}만")
        others = [display(p, who) for p in mine if p != code]
        if others:
            lines.append(f"함께 타는 사람: {', '.join(others)}")
        if theirs:
            lines.append(f"{'오후' if half == 'am' else '오전'}에는 {', '.join(display(p, who) for p in theirs)} 님이 이 차를 써요.")
        notes = [r['note']] if r.get('note') else []
    else:
        lines.append(f":car: {day_label(day)} 차량 — 오전·오후가 달라요")
        notes = []
        for half, label in (('am', '오전'), ('pm', '오후')):
            plate = w[half]
            if not plate:
                lines.append(f"{label}: 배정 없음")
                continue
            r = rows[plate]
            lines.append(f"{label}: " + _car_line(plate, r['people'] if half == 'am' else pm_of(r), code, cars, who))
            if r.get('note') and r['note'] not in notes:
                notes.append(r['note'])
    for n in notes:
        lines.append(f"메모: {n}")
    # 같이 타는 차면 [운행 종료] 안내(11-22) — 화면을 모두에게 연 뒤부터(개발 중엔 못 여는 화면이라 안 붙임)
    from dashboard.services import groupware as gw
    shared = any(len(riders(rows[p])) > 1 for p in {w['am'], w['pm']} if p and p in rows and p != OWN)
    if shared and not gw.DEV_PREVIEW_USERS:
        lines.append("운행을 마치면 그룹웨어 차량 배정에서 [운행 종료]를 눌러 운전한 사람을 적어 주세요.")
    return '\n'.join(lines)


def _person_view(code: str, rows: Dict[str, Dict]) -> tuple:
    """알림 비교용 — 그 사람이 받을 내용이 바뀌었는지(차·같이 탄 사람·운전·메모)."""
    w = _where(rows).get(code)
    if not w:
        return ()
    out = []
    for half in ('am', 'pm'):
        plate = w[half]
        r = rows.get(plate) if plate else None
        out.append((plate, tuple((r['people'] if half == 'am' else pm_of(r)) if r else ()), r['note'] if r else ''))
    return tuple(out)


def _dm_or_miss(p: Optional[Dict], code: str, text: str, failed: Optional[list],
                retry: Optional[list] = None) -> bool:
    """DM 하나 — 못 보내면(사람 정보·슬랙 ID 없음 포함) failed 에 이름을 넣는다.
    슬랙 ID 는 있는데 보내다 실패한 사람은 retry 에 이니셜 — 화면 [다시 보내기] 대상(11-24 추가분)."""
    if p and _dm(p.get('slack_id') or '', text):
        return True
    if failed is not None:
        failed.append((p or {}).get('name') or code)
    if retry is not None and p and (p.get('slack_id') or '').strip():
        retry.append(code)
    return False


def notify_people(day: str, cars: Dict[str, Dict], rows: Dict[str, Dict], who: Dict[str, Dict],
                  everyone: bool = True, only: Optional[set] = None, failed: Optional[list] = None,
                  retry: Optional[list] = None) -> int:
    if not CAR_DM_ON:
        return 0
    n = 0
    for code in _where(rows):
        if only is not None and code not in only:
            continue
        p = who.get(code)
        if _dm_or_miss(p, code, _person_text(day, code, rows, cars, who) if p else '', failed, retry):
            n += 1
    return n


def notify_changes(day: str, cars: Dict[str, Dict], old: Dict[str, Dict], new: Dict[str, Dict],
                   who: Dict[str, Dict], failed: Optional[list] = None, retry: Optional[list] = None) -> int:
    """확정 뒤 고친 경우 — 차(오전·오후)·같이 타는 사람·운전·메모가 바뀐 사람에게만."""
    if not CAR_DM_ON:
        return 0
    ow, nw = _where(old), _where(new)
    n = 0
    for code in set(ow) | set(nw):
        if _person_view(code, old) == _person_view(code, new):
            continue
        p = who.get(code)
        if not p:
            if failed is not None:
                failed.append(code)
            continue
        if code in nw:
            text = ':arrows_counterclockwise: 차량 배정이 바뀌었어요\n' + _person_text(day, code, new, cars, who)
        else:
            was = [x for x in dict.fromkeys((ow[code]['am'], ow[code]['pm'])) if x]
            text = f":arrows_counterclockwise: {day_label(day)} 차량 배정에서 빠졌어요 ({', '.join(_car_text(x, cars) for x in was)})"
        if _dm_or_miss(p, code, text, failed, retry):
            n += 1
    return n


class SentCount(int):
    """보낸 수(int 그대로 비교·JSON) + 못 보낸 사람 .failed(이메일 목록) — '대상 있는데 다 실패'와
    '대상 없음(안 보내도 되는 날)'을 구별하려고(11-24 조용한 실패 점검). 호출처·테스트 가짜는 int 를 돌려줘도 되게
    읽을 땐 failed_of() 로."""
    failed: List[str] = []

    def __new__(cls, n: int, failed: Optional[List[str]] = None):
        obj = super().__new__(cls, n)
        obj.failed = list(failed or [])
        return obj


def failed_of(res) -> List[str]:
    """SentCount 의 못 보낸 사람 — 보통 int(가짜 등)면 []."""
    return list(getattr(res, 'failed', None) or [])


def _dm_assigners(text: str, skip: str = '') -> SentCount:
    """배정 담당(설정 명단)에게 DM — 보낸 수 + 못 보낸 사람. 못 보낸 사람이 있거나 명단이 비면 ERROR(11-24)."""
    from dashboard.services import hr
    n, failed = 0, []
    every = assigners()
    targets = [e for e in every if e != skip]
    for email in targets:
        sid = ((hr.find_by_email(email) or {}).get('slack_id')) or ''
        if _dm(sid, text):
            n += 1
        else:
            failed.append(email)
    if failed:
        logger.error(f"[차량배정] 배정 담당 DM 못 보냄 {len(failed)}/{len(targets)}명({', '.join(failed)}) — "
                     f"슬랙 ID·토큰을 확인하고 직접 알려 주세요")
    elif not every:
        logger.error('[차량배정] 배정 담당 명단이 비어 차량 배정 DM 을 아무에게도 못 보냈어요')
    return SentCount(n, failed)


def request_assignment(day: Optional[str] = None, by: str = '') -> int:
    """/일정확정 뒤 — 배정하는 사람에게 '내일 차량 배정해 주세요' DM(그날 이미 확정했거나 이미 보냈으면 안 보냄).
    by = 방문 일정을 확정한 사람. 그 사람이 배정 담당(JW 등)이면 보내지 않는다 — 확정하고 바로 차량 배정 탭으로 넘어가
    직접 넣으므로(11-24 부장님). 대표님처럼 배정 담당이 아닌 사람이 확정하면 그대로 보낸다.
    돌려주는 값 = SentCount(보낸 수, .failed = 못 보낸 배정 담당) — 0 이어도 failed 가 비면 '보낼 필요 없음'(11-24)."""
    day = day if day and valid_day(day) else next_workday()
    if not CAR_DM_ON:
        return 0
    if by and by.strip().lower() in assigners():
        return 0
    info = _day_info(day) or {}
    if info.get('confirmed_at') or info.get('requested_at'):
        return 0
    # 차량 배정 화면은 배정 담당에게 열렸다(11-24) — 개발 중 명단 거르기 없앰
    n = _dm_assigners(f":car: 일정확정이 끝났어요. {day_label(day)} 차량 배정을 해 주세요.\n<{_page_url(day)}|차량 배정 열기>")
    if n:
        with _conn() as conn:
            conn.execute('INSERT INTO vehicle_assign_day (day, requested_at) VALUES (?,?) '
                         'ON CONFLICT(day) DO UPDATE SET requested_at = excluded.requested_at', (day, _now()))
    return n


def last_saved_at(day: str) -> str:
    """그날 차량 배정을 마지막으로 저장한 때('' = 저장한 적 없음) — 방문 일정이 그 뒤에 다시 확정됐는지 보려고."""
    with _conn() as conn:
        r = conn.execute('SELECT MAX(updated_at) FROM vehicle_assign WHERE day = ?', (day,)).fetchone()
    return (r[0] if r and r[0] else '') or ''


def notify_plan_changed(day: str, change: str, by: str = '') -> int:
    """방문 일정을 다시 확정해 코스 사람이 바뀜(11-23 부장님) — 차량 배정을 이미 저장한 날만 배정 담당에게 DM.
    아직 배정이 없으면 보내지 않는다(화면을 열면 바뀐 코스대로 초안이 잡힘).
    by = 다시 확정한 사람 — 그 사람은 확정 뒤 바로 차량 배정 탭으로 넘어오므로 빼고 보낸다(11-24).
    돌려주는 값 = SentCount(보낸 수, .failed = 못 보낸 배정 담당) — 배정 안 한 날의 0 과 다 실패한 0 을 구별(11-24)."""
    if not CAR_DM_ON:
        return 0
    if not any(r['people'] or r.get('pm') for r in _day_rows(day).values()):
        return 0
    # 배정 담당은 화면을 열 수 있다(11-24, request_assignment 와 같음)
    return _dm_assigners(f":arrows_counterclockwise: {day_label(day)} 방문 일정이 다시 확정됐어요 — {change}.\n"
                         f"차량 배정을 확인해 주세요. <{_page_url(day)}|차량 배정 열기>",
                         skip=(by or '').strip().lower())


def off_riders(day: str, rows: Dict[str, Dict]) -> List[str]:
    """차에 넣은 사람 중 그날 종일 쉬는 사람(승인 휴가·확정된 토요 휴무) 이름 — 확정·알리기를 막는다(방문 일정과 같게, 11-23).
    확정 전 토요 휴무(soft)는 막지 않는다. 휴가를 못 읽으면 막지 않는다."""
    try:
        off = {o['code']: o['name'] for o in day_off(day)['off'] if not o.get('soft')}
    except Exception as e:
        logger.warning(f'[차량배정] 휴가 확인 실패(막지 않음): {e}')
        return []
    seen = list(dict.fromkeys(p for r in rows.values() for p in riders(r)))
    return [off[p] for p in seen if p in off]


def teams(day: str, leads: Iterable[dict]) -> List[Dict]:
    """일정확정에서 함께 가는 사람 묶음 — 리드 시트 그날 방문의 영업 담당자 조합(예: YM+LSH).
    한 사람이 두 조합에 걸치면 방문 많은 조합 → 작은 조합 순으로 먼저 잡는다(11-21: TH+LJW·MW+TH+LJW → TH+LJW, MW).
    [{'members': [이니셜(적힌 순서)], 'visits': n}] — 직원 명단에 없는 이름은 뺀다."""
    from collections import Counter
    from dashboard.services import vehicle_log as vl
    by_name = {p['name']: p['code'] for p in staff()}
    combos: Counter = Counter()
    for v in vl.visits_by_day(leads, day, day).get(day, []):
        codes = tuple(dict.fromkeys(by_name[n] for n in v['who'] if n in by_name))
        if codes:
            combos[codes] += 1
    return teams_from_combos(combos)


def teams_from_combos(combos) -> List[Dict]:
    """조합별 방문 수(Counter{(이니셜,...): n}) → 팀. 한 사람이 두 조합에 걸치면 방문 많은 조합 → 작은 조합 순으로 먼저."""
    taken: set = set()
    out: Dict[tuple, int] = {}
    for combo, n in sorted(combos.items(), key=lambda kv: (-kv[1], len(kv[0]), kv[0])):
        rest = tuple(p for p in combo if p not in taken)
        if not rest:
            continue
        taken.update(rest)
        out[rest] = out.get(rest, 0) + n
    # 남은 사람이 이미 있는 팀과 같으면 합친다(위 out 키가 같으면 이미 합쳐짐)
    return [{'members': list(k), 'visits': v, 'big': len(k) >= BIG_TEAM} for k, v in out.items()]
