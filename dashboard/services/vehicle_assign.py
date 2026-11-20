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
    ('821저4919', '포터 4919(LPG)', '포터', 'LPG', '글로벌', SHARED, '', 1, 1),
    ('90머9131', '전기트럭 9131', '포터 일렉트릭', '전기', '글로벌', SHARED, '', 1, 1),
    ('53우1783', '스파크 1783', '스파크', '가솔린', '글로벌', SHARED, '', 1, 1),
    ('60거5484', '모닝 5484', '모닝', '가솔린', '글로벌', SHARED, '', 1, 1),
    ('143고9687', '카니발 9687', '카니발', '디젤', '글로벌', SHARED, 'YG', 0, 1),   # 평소 YG, 매일 배정 목록엔 넣음(11-20)
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
_DEFAULT_ASSIGNERS = ('sj@itg-aircon.com',)      # 빈승정 대리 — 인사 관리자는 따로 늘 가능


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
            for i, s in enumerate(_SEED):
                conn.execute('INSERT OR IGNORE INTO vehicle_fleet (plate, label, model, fuel, corp, usage, fixed_user, '
                             'needs_log, uplus, active, sort_no, updated_by, updated_at) '
                             'VALUES (?,?,?,?,?,?,?,?,?,1,?,?,?)', (*s, i, 'seed', _now()))
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
    return bool(email) and (hr.is_hr_admin(email) or email in assigners())


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


def _day_rows(day: str) -> Dict[str, Dict]:
    with _conn() as conn:
        return {r['plate']: {'people': _split(r['people']), 'note': r['note'] or ''}
                for r in conn.execute('SELECT plate, people, note FROM vehicle_assign WHERE day = ?', (day,)).fetchall()}


def _day_info(day: str) -> Optional[Dict]:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM vehicle_assign_day WHERE day = ?', (day,)).fetchone()
    return dict(r) if r else None


def _previous_day_with_rows(day: str) -> Optional[str]:
    with _conn() as conn:
        r = conn.execute("SELECT MAX(day) FROM vehicle_assign WHERE day < ? AND people != ''", (day,)).fetchone()
    return r[0] if r and r[0] else None


def _recent_cars(day: str, back_days: int = 30) -> Dict[str, str]:
    """사람 → 그날 이전에 마지막으로 탄 차."""
    start = (date.fromisoformat(day) - timedelta(days=back_days)).isoformat()
    with _conn() as conn:
        rows = conn.execute('SELECT day, plate, people FROM vehicle_assign WHERE day >= ? AND day < ? ORDER BY day',
                            (start, day)).fetchall()
    out: Dict[str, str] = {}
    for r in rows:
        for p in _split(r['people']):
            out[p] = r['plate']
    return out


def _team_draft(day: str, teams: List[Dict], cars: List[Dict]) -> Dict[str, Dict]:
    """일정확정 팀마다 그 팀 사람이 최근에 탄 차(비어 있으면)에 넣은 초안. 탄 적 없으면 넣지 않는다."""
    recent = _recent_cars(day)
    shared = {c['plate'] for c in cars if c['usage'] == SHARED}
    rows: Dict[str, Dict] = {}
    for t in teams:
        plate = next((recent[p] for p in t['members'] if recent.get(p) in shared and recent[p] not in rows), None)
        if plate:
            rows[plate] = {'people': list(t['members']), 'note': ''}
    return rows


def day_view(day: str, teams: Optional[List[Dict]] = None) -> Dict:
    """배정 화면 한 날짜 — 차량마다 사람·메모. 아직 아무것도 없으면 초안:
    일정확정 팀이 있으면 팀대로(각 팀이 최근 탄 차), 없으면 직전 배정(공용 차만)."""
    cars = fleet()
    rows = _day_rows(day)
    info = _day_info(day)
    draft_from = draft_kind = None
    if not any(r['people'] for r in rows.values()) and teams:
        rows, draft_kind = _team_draft(day, teams, cars), 'teams'
    elif not any(r['people'] for r in rows.values()):
        prev = _previous_day_with_rows(day)
        if prev:
            draft_kind = 'prev'
            prev_rows = _day_rows(prev)
            shared = {c['plate'] for c in cars if c['usage'] == SHARED}
            rows = {p: {'people': r['people'], 'note': ''} for p, r in prev_rows.items() if p in shared and r['people']}
            draft_from = prev
    out = []
    for c in cars:
        r = rows.get(c['plate'], {'people': [], 'note': ''})
        out.append({**{k: c[k] for k in ('plate', 'label', 'model', 'fuel', 'corp', 'usage', 'fixed_user',
                                         'needs_log', 'uplus')},
                    'people': r['people'], 'note': r['note']})
    return {'day': day, 'cars': out, 'draft_from': draft_from, 'draft_kind': draft_kind,
            'confirmed_at': (info or {}).get('confirmed_at'), 'confirmed_by': (info or {}).get('confirmed_by')}


def _clean_rows(rows: Iterable[Dict], plates: set, codes: set) -> Dict[str, Dict]:
    """화면에서 온 배정 → {plate: {people, note}}. 한 사람이 두 차에 있으면 ValueError."""
    out: Dict[str, Dict] = {}
    where: Dict[str, str] = {}
    for r in rows or []:
        plate = (r.get('plate') or '').strip()
        if plate not in plates:
            raise ValueError(f'없는 차량이에요: {plate}')
        people = []
        for p in r.get('people') or []:
            p = (p or '').strip()
            if not p:
                continue
            p = p.upper() if p.upper() in codes else p[:20]
            if p in people:
                continue
            if p in where and where[p] != plate:
                raise ValueError(f'{p} 님이 두 차에 들어 있어요. 한 차만 골라 주세요.')
            where[p] = plate
            people.append(p)
        note = (r.get('note') or '').strip()[:60]
        if people or note:
            out[plate] = {'people': people, 'note': note}
    return out


def save_day(day: str, rows: Iterable[Dict], by: str, confirm: bool = False) -> Dict:
    """배정 저장. confirm=True 면 확정(DM 보냄). 확정된 날을 고치면 바뀐 사람에게만 DM.
    돌려주는 값: {'changed': [...], 'notified': n, 'confirmed': bool}"""
    if not valid_day(day):
        raise ValueError('날짜가 올바르지 않아요.')
    cars = {c['plate']: c for c in fleet()}
    who = _who_map()
    new = _clean_rows(rows, set(cars), set(who))
    old = _day_rows(day)
    info = _day_info(day)
    was_confirmed = bool(info and info.get('confirmed_at'))
    now = _now()
    changed = []
    with _conn() as conn:
        for plate in set(old) | set(new):
            o = old.get(plate, {'people': [], 'note': ''})
            n = new.get(plate, {'people': [], 'note': ''})
            if o == n:
                continue
            changed.append(plate)
            conn.execute('INSERT INTO vehicle_assign_log (day, plate, before, after, by, at) VALUES (?,?,?,?,?,?)',
                         (day, plate, json.dumps(o, ensure_ascii=False), json.dumps(n, ensure_ascii=False), by, now))
            if n['people'] or n['note']:
                conn.execute('INSERT OR REPLACE INTO vehicle_assign (day, plate, people, note, updated_by, updated_at) '
                             'VALUES (?,?,?,?,?,?)', (day, plate, ','.join(n['people']), n['note'], by, now))
            else:
                conn.execute('DELETE FROM vehicle_assign WHERE day = ? AND plate = ?', (day, plate))
        if confirm and not was_confirmed:
            conn.execute('INSERT INTO vehicle_assign_day (day, confirmed_by, confirmed_at) VALUES (?,?,?) '
                         'ON CONFLICT(day) DO UPDATE SET confirmed_by = excluded.confirmed_by, '
                         'confirmed_at = excluded.confirmed_at', (day, by, now))
    notified = 0
    if confirm and not was_confirmed:
        notified = notify_people(day, cars, new, who, everyone=True)
    elif was_confirmed and changed:
        notified = notify_changes(day, cars, old, new, who)
    return {'changed': changed, 'notified': notified, 'confirmed': confirm or was_confirmed}


def person_car(day: str, code: str) -> Optional[Dict]:
    """그날 이 사람이 탈 차(배정, 없으면 고정 차)."""
    rows = _day_rows(day)
    for plate, r in rows.items():
        if code in r['people']:
            return {'plate': plate, **r}
    return None


def assigned_people(start: str, end: str) -> Dict[Tuple[str, str], List[str]]:
    """운행일지용 — (날짜, 번호판) → 탄 사람 이니셜(맨 앞 = 사용자). 배정 없는 고정 차는 고정 사용자."""
    with _conn() as conn:
        rows = conn.execute('SELECT day, plate, people FROM vehicle_assign WHERE day BETWEEN ? AND ?',
                            (start, end)).fetchall()
    return {(r['day'], r['plate']): _split(r['people']) for r in rows if r['people']}


def fixed_users() -> Dict[str, str]:
    return {c['plate']: c['fixed_user'] for c in fleet(True) if c['usage'] == FIXED and c['fixed_user']}


def visitors(day: str, leads: Iterable[dict]) -> List[Dict]:
    """그날 방문 일정이 있는 사람(리드 시트 방문 예정일·영업 담당자) — 배정 화면에서 '차 없는 사람'을 짚는 데 쓴다.
    [{code, name, visits}] — 이니셜이 없는 이름(외부 인원 등)은 뺀다."""
    from dashboard.services import vehicle_log as vl
    by_name = {p['name']: p for p in staff()}
    count: Dict[str, int] = {}
    for v in vl.visits_by_day(leads, day, day).get(day, []):
        for n in v['who']:
            if n in by_name:
                count[n] = count.get(n, 0) + 1
    return [{'code': by_name[n]['code'], 'name': n, 'visits': c} for n, c in sorted(count.items())]


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


def _car_text(plate: str, cars: Dict[str, Dict]) -> str:
    c = cars.get(plate) or {}
    return f"*{c.get('label') or plate}* ({plate})"


def _person_text(day: str, code: str, plate: str, row: Dict, cars: Dict[str, Dict], who: Dict[str, Dict]) -> str:
    others = [display(p, who) for p in row['people'] if p != code]
    lines = [f":car: {day_label(day)} 차량: {_car_text(plate, cars)}"]
    if others:
        lines.append(f"함께 타는 사람: {', '.join(others)}")
    if row.get('note'):
        lines.append(f"메모: {row['note']}")
    return '\n'.join(lines)


def notify_people(day: str, cars: Dict[str, Dict], rows: Dict[str, Dict], who: Dict[str, Dict],
                  everyone: bool = True, only: Optional[set] = None) -> int:
    n = 0
    for plate, row in rows.items():
        for code in row['people']:
            if only is not None and code not in only:
                continue
            p = who.get(code)
            if p and _dm(p['slack_id'], _person_text(day, code, plate, row, cars, who)):
                n += 1
    return n


def notify_changes(day: str, cars: Dict[str, Dict], old: Dict[str, Dict], new: Dict[str, Dict],
                   who: Dict[str, Dict]) -> int:
    """확정 뒤 고친 경우 — 차가 바뀌었거나, 같이 타는 사람·메모가 바뀐 사람에게만."""
    def where(rows):
        return {c: (plate, r) for plate, r in rows.items() for c in r['people']}
    o, nw = where(old), where(new)
    n = 0
    for code in set(o) | set(nw):
        before, after = o.get(code), nw.get(code)
        if before == after:
            continue
        p = who.get(code)
        if not p:
            continue
        if after:
            text = _person_text(day, code, after[0], after[1], cars, who)
            if before and before[0] != after[0]:
                text = f":arrows_counterclockwise: 차량이 바뀌었어요 ({_car_text(before[0], cars)} → {_car_text(after[0], cars)})\n" + text
            else:
                text = ':arrows_counterclockwise: 차량 배정이 바뀌었어요\n' + text
        else:
            text = f":arrows_counterclockwise: {day_label(day)} 차량 배정에서 빠졌어요 ({_car_text(before[0], cars)})"
        if _dm(p['slack_id'], text):
            n += 1
    return n


def request_assignment(day: Optional[str] = None) -> int:
    """/일정확정 뒤 — 배정하는 사람에게 '내일 차량 배정해 주세요' DM(그날 이미 확정했거나 이미 보냈으면 안 보냄)."""
    day = day if day and valid_day(day) else next_workday()
    info = _day_info(day) or {}
    if info.get('confirmed_at') or info.get('requested_at'):
        return 0
    from dashboard.services import hr
    n = 0
    for email in assigners():
        p = hr.find_by_email(email)
        sid = (p or {}).get('slack_id') or ''
        if _dm(sid, f":car: 일정확정이 끝났어요. {day_label(day)} 차량 배정을 해 주세요.\n<{_page_url(day)}|차량 배정 열기>"):
            n += 1
    if n:
        with _conn() as conn:
            conn.execute('INSERT INTO vehicle_assign_day (day, requested_at) VALUES (?,?) '
                         'ON CONFLICT(day) DO UPDATE SET requested_at = excluded.requested_at', (day, _now()))
    return n


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
    taken: set = set()
    out: Dict[tuple, int] = {}
    for combo, n in sorted(combos.items(), key=lambda kv: (-kv[1], len(kv[0]), kv[0])):
        rest = tuple(p for p in combo if p not in taken)
        if not rest:
            continue
        taken.update(rest)
        out[rest] = out.get(rest, 0) + n
    # 남은 사람이 이미 있는 팀과 같으면 합친다(위 out 키가 같으면 이미 합쳐짐)
    return [{'members': list(k), 'visits': v} for k, v in out.items()]
