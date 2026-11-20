"""차량 배정 (2026-11-20 부장님) — 다음 날 어느 차를 누가 가지고 나가는지.

지금까지는 빈승정 대리(SJ)가 오후 4~5시 카톡 단체방에 'YM+LSH 레이 / JW+LJW+KMW 8021 / SJ 전기트럭'처럼
올렸다. 레이가 3대라 어느 레이인지 남지 않고, 운행일지(U+커넥트)의 운전자 칸도 비어 있어 누가 탔는지 알 수 없었다.

- 차 기준으로 사람을 붙인다(한 차에 여러 명, 맨 앞 사람 = 운행기록부 '사용자').
- 공용 차(레이·쏘렌토·포터·스파크·모닝·카니발)는 매일 배정(카니발은 평소 YG 차지만 목록엔 넣는다 — 부장님 11-20).
  고정 차(파나메라=JW, 벤츠=김영숙, 랜드로버)는 평소 그 사람 차, 차가 모자란 날만 '빌려 쓰기'로 그날 배정을 넣는다.
- 확정하면 탄 사람마다 슬랙 DM. 확정 뒤 고치면 바뀐 사람에게만 다시 DM.
- /일정확정이 끝나면 배정하는 사람(설정 vehicle_assigners, 기본 SJ)에게 '내일 차량 배정해 주세요' DM.
- 카톡에 쓰던 글을 그대로 붙여 넣으면 채워 준다. 어느 차인지 모호하면(그냥 '레이') 그 사람이 최근에 탄 차를
  먼저 골라 두고 '확인 필요'로 표시한다.
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

SHARED, FIXED = 'shared', 'fixed'

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
    ('204두8988', '파나메라 8988', '파나메라', '', '글로벌', FIXED, 'JW', 1, 0),
    ('260나7350', '벤츠 7350', '벤츠', '', '그룹', FIXED, '김영숙', 1, 0),
    ('133가6633', '랜드로버 6633', '랜드로버', '디젤', '플랜트', FIXED, '', 1, 0),
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
    q = 'SELECT * FROM vehicle_fleet' + ('' if include_inactive else ' WHERE active = 1') + ' ORDER BY sort_no, plate'
    with _conn() as conn:
        return [dict(r) for r in conn.execute(q).fetchall()]


def update_car(plate: str, values: Dict, by: str) -> Dict:
    """공용/고정·고정 사용자·부르는 이름만 고친다."""
    allowed = {}
    if 'usage' in values:
        if values['usage'] not in (SHARED, FIXED):
            raise ValueError('공용 또는 고정만 고를 수 있어요.')
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


def staff() -> List[Dict]:
    """배정할 수 있는 사람 — 재직 인사카드 중 이니셜(공사 코드 접미사)이 있는 사람.
    [{code, name, slack_id, dept}] (code = 이니셜)"""
    from dashboard.services import hr
    from dashboard.services.staff_directory import suffix_map
    sm = suffix_map()
    rows = hr._rows("SELECT name, slack_id, department FROM hr_people WHERE status = '재직' ORDER BY name")
    out, seen = [], set()
    for r in rows:
        code = (sm.get(r['name']) or '').upper()
        if not code or code in seen:
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


def day_view(day: str) -> Dict:
    """배정 화면 한 날짜 — 차량마다 사람·메모. 아직 아무것도 없으면 직전 배정(공용 차만)을 초안으로."""
    cars = fleet()
    rows = _day_rows(day)
    info = _day_info(day)
    draft_from = None
    if not any(r['people'] for r in rows.values()):
        prev = _previous_day_with_rows(day)
        if prev:
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
    return {'day': day, 'cars': out, 'draft_from': draft_from,
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


# ── 카톡 글 해석 ─────────────────────────────────────────────────────────
# 'YM+LSH 레이' / 'JW+LJW+KMW 8021' / 'TH+JSH 레이 5358' / 'SJ 전기트럭' / 'TH 5358(유리용접)'
_LINE_RE = re.compile(r'^\s*([A-Za-z]{1,4}(?:\s*[+,/]\s*[A-Za-z]{1,4})*)\s*(.*?)\s*$')
_NOTE_RE = re.compile(r'[(\[]([^)\]]*)[)\]]')
_KEYWORDS = [   # (글자, 차종 이름에 들어가는 말) — 앞에 있는 것이 먼저
    ('전기', '포터 일렉트릭'), ('일렉', '포터 일렉트릭'), ('가스', 'LPG'), ('lpg', 'LPG'),
    ('레이', '레이'), ('쏘렌토', '쏘렌토'), ('소렌토', '쏘렌토'), ('포터', '포터'), ('트럭', '포터'),
    ('스파크', '스파크'), ('모닝', '모닝'), ('카니발', '카니발'), ('파나메라', '파나메라'),
    ('벤츠', '벤츠'), ('랜드로버', '랜드로버'), ('디펜더', '랜드로버'),
]


def _candidates(spec: str, cars: List[Dict]) -> List[str]:
    s = spec.lower().replace(' ', '')
    digits = re.findall(r'\d{4}', s)
    if digits:
        hit = [c['plate'] for c in cars if c['plate'].endswith(digits[0])]
        if hit:
            return hit
    for word, kind in _KEYWORDS:
        if word in s:
            if kind == 'LPG':
                return [c['plate'] for c in cars if c['fuel'] == 'LPG']
            if kind == '포터':     # 그냥 '트럭'·'포터'는 전기트럭까지 셋 다
                return [c['plate'] for c in cars if c['model'].startswith('포터')]
            return [c['plate'] for c in cars if c['model'] == kind]
    return []


def parse_text(text: str, day: Optional[str] = None) -> Dict:
    """카톡 글 → 화면에 채울 배정. {'rows': [{plate, people, note, choices, unsure, line}], 'skipped': [줄]}
    같은 차가 두 번 나오면 뒤 줄이 이긴다(카톡에서 고쳐 쓴 줄). 모호하면 그 사람의 최근 차를 먼저 고른다."""
    cars = fleet()
    codes = {p['code'] for p in staff()}
    recent = _recent_cars(day or date.today().isoformat())
    rows: List[Dict] = []
    skipped: List[str] = []
    for raw in (text or '').splitlines():
        line = raw.strip()
        if not line:
            continue
        m = _LINE_RE.match(line)
        if not m:
            skipped.append(line)
            continue
        people = [p.upper() for p in re.split(r'\s*[+,/]\s*', m.group(1)) if p]
        if not all(p in codes for p in people):
            skipped.append(line)
            continue
        rest = m.group(2)
        note = '; '.join(x.strip() for x in _NOTE_RE.findall(rest) if x.strip())
        spec = _NOTE_RE.sub(' ', rest).strip()
        choices = _candidates(spec, cars)
        if not choices:
            skipped.append(line)
            continue
        rows.append({'people': people, 'note': note, 'choices': choices, 'line': line})
    # 차 고르기 — 확실한 줄 먼저, 모호한 줄은 남은 차 중 그 사람이 최근에 탄 차
    taken = set()
    for r in rows:
        if len(r['choices']) == 1:
            r['plate'], r['unsure'] = r['choices'][0], False
            taken.add(r['plate'])
    for r in rows:
        if 'plate' in r:
            continue
        free = [p for p in r['choices'] if p not in taken] or r['choices']
        pick = next((recent[p] for p in r['people'] if recent.get(p) in free), free[0])
        r['plate'], r['unsure'] = pick, True
        taken.add(pick)
    # 줄 합치기 — 같은 차는 한 줄로(사람·메모 합침), 한 사람이 다른 차로 또 나오면 뒤 줄(고쳐 쓴 줄)이 이긴다
    merged: Dict[str, Dict] = {}
    for r in rows:
        m = merged.setdefault(r['plate'], {'plate': r['plate'], 'people': [], 'notes': [], 'unsure': False,
                                           'choices': r['choices'], 'lines': []})
        for p in r['people']:
            for other in merged.values():
                if other is not m and p in other['people']:
                    other['people'].remove(p)
            if p not in m['people']:
                m['people'].append(p)
        if r['note'] and r['note'] not in m['notes']:
            m['notes'].append(r['note'])
        m['unsure'] = m['unsure'] or r['unsure']
        m['lines'].append(r['line'])
    out = [{'plate': m['plate'], 'people': m['people'], 'note': '; '.join(m['notes']), 'unsure': m['unsure'],
            'choices': m['choices'], 'line': ' / '.join(m['lines'])} for m in merged.values() if m['people']]
    return {'rows': out, 'skipped': skipped}


def _recent_cars(day: str, back_days: int = 30) -> Dict[str, str]:
    """사람 → 최근에 배정된 차(그날 이전)."""
    start = (date.fromisoformat(day) - timedelta(days=back_days)).isoformat()
    with _conn() as conn:
        rows = conn.execute('SELECT day, plate, people FROM vehicle_assign WHERE day >= ? AND day < ? ORDER BY day',
                            (start, day)).fetchall()
    out: Dict[str, str] = {}
    for r in rows:
        for p in _split(r['people']):
            out[p] = r['plate']
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
