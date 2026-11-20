"""토요일 휴무(격주휴무) — 기술영업이 쉬는 토요일을 그룹웨어 달력에서 고르고 결재받는다(2026-11-20 부장님).

예전 방식(2024-12 격주휴무 회의, 드라이브 '기술영업 휴무자료'): 매월 첫째 주 각자 빈승정 대리에게 알림 → 빈 대리가
엑셀 「격주휴무 신청서」로 취합 → 담당실장·대표 결재. 한 사람 한 달 2번. 공사가 겹치면 일이 먼저(다른 주로 미룸).

그룹웨어:
- 대상 = 인사카드 부서 '기술영업'(재직). 달력의 토요일을 눌러 자기 이름을 넣고 뺀다.
- 달 상태: 작성 중 → (취합 담당이 [결재 올리기]) 결재 대기 → (박정우·박용구 둘 중 한 분 승인) 확정.
  돌려보내면 다시 작성 중. 확정 뒤 바꾸는 건 변경 신청 → 둘 중 한 분 승인(부장님 11-20).
- 한 달 2번을 넘게 골라도 막지 않고 경고만(전체휴무 같은 예외가 있음).
- 확정된 달은 휴가 일수에 쓰인다: 기술영업이 그달 토요일에 휴가를 내면, 휴무표에서 쉬는 토요일은 빼고 근무 토요일은 센다
  (leave.count_days). 확정 전 달은 예전 규칙(휴가 첫날·마지막 날로 고른 토요일만 셈).
- 승인권자(박정우·박용구)는 어느 상태에서든 직접 고칠 수 있다(기록이 남음).
표는 휴가 장부와 같은 DB(leave._conn).
"""
from __future__ import annotations

import json
import os
import threading
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Tuple

from dashboard.services import leave as lv
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)
_LOG = '[토요휴무]'

MEMBER_DEPT = '기술영업'
MONTHLY_LIMIT = 2
AHEAD_MONTHS = 6                       # 몇 달 앞까지 고를 수 있나(결혼식·해외여행은 3달 전에 알리기로 함)
_DEFAULT_COORDINATORS = ('U0BHDBCEUHL',)   # 빈승정 대리 — 취합 담당(2024-12 격주휴무 회의)

OPEN, WAIT, CONFIRMED = 'open', 'wait', 'confirmed'
STATUS_LABELS = {OPEN: '작성 중', WAIT: '결재 대기', CONFIRMED: '확정'}
CH_WAIT, CH_OK, CH_NO, CH_CANCEL = 'wait', 'approved', 'rejected', 'cancelled'
CHANGE_LABELS = {CH_WAIT: '결재 대기', CH_OK: '승인', CH_NO: '반려', CH_CANCEL: '취소'}

DDL = [
    '''CREATE TABLE IF NOT EXISTS sat_off_days (
        day TEXT NOT NULL,
        name TEXT NOT NULL,
        slack_id TEXT,
        added_by TEXT,
        added_at TEXT,
        PRIMARY KEY (day, name)
    )''',
    '''CREATE TABLE IF NOT EXISTS sat_off_months (
        month TEXT PRIMARY KEY,
        status TEXT NOT NULL DEFAULT 'open',
        submitted_by TEXT,
        submitted_at TEXT,
        decided_by TEXT,
        decided_at TEXT,
        note TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS sat_off_changes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        month TEXT NOT NULL,
        name TEXT NOT NULL,
        slack_id TEXT,
        add_days TEXT NOT NULL DEFAULT '[]',
        remove_days TEXT NOT NULL DEFAULT '[]',
        reason TEXT,
        status TEXT NOT NULL DEFAULT 'wait',
        requested_by TEXT,
        requested_at TEXT,
        decided_by TEXT,
        decided_at TEXT,
        decide_note TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS sat_off_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        at TEXT NOT NULL,
        actor TEXT,
        action TEXT NOT NULL,
        month TEXT,
        detail TEXT
    )''',
]


class SatOffError(ValueError):
    """화면에 그대로 보여 줄 수 있는 거절 사유."""


# ── 사람·역할 ──────────────────────────────────────────────────────────
def _ids(env: str, default: tuple) -> tuple:
    raw = os.getenv(env, '').strip()
    return tuple(x.strip() for x in raw.split(',') if x.strip()) if raw else default


def coordinator_ids() -> tuple:
    return _ids('SAT_OFF_COORDINATOR_IDS', _DEFAULT_COORDINATORS)


def approver_ids() -> tuple:
    return lv.approver_ids()


def manager_ids() -> tuple:
    """작성 중인 달에서 남의 이름도 넣고 빼고, 결재를 올릴 수 있는 사람 — 취합 담당·승인권자·경영지원."""
    return tuple(dict.fromkeys(coordinator_ids() + approver_ids() + lv.manager_ids()))


def members() -> List[dict]:
    """대상자 — 인사카드 부서가 기술영업인 재직자 [{name, slack_id}] (이름순)."""
    with lv._conn() as conn:
        if not conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='hr_people'").fetchone():
            return []
        rows = conn.execute('SELECT name, slack_id FROM hr_people WHERE department=? AND status != ? ORDER BY name',
                            (MEMBER_DEPT, lv._HR_RESIGNED)).fetchall()
    return [{'name': r[0], 'slack_id': r[1] or ''} for r in rows]


def is_member(name: str) -> bool:
    return bool(name) and lv._department(name) == MEMBER_DEPT


def _member(name: str) -> Optional[dict]:
    return next((m for m in members() if m['name'] == name), None)


def role(sid: str, name: str = '') -> Dict[str, bool]:
    approver = bool(sid) and sid in approver_ids()
    coordinator = bool(sid) and sid in coordinator_ids()
    manager = bool(sid) and sid in manager_ids()
    member = bool(name) and _member(name) is not None
    return {'member': member, 'coordinator': coordinator, 'approver': approver, 'manager': manager}


# ── 날짜 ──────────────────────────────────────────────────────────────
def _month_of(d: date) -> str:
    return d.strftime('%Y-%m')


def parse_month(s: str) -> Optional[str]:
    try:
        y, m = (int(x) for x in (s or '').split('-'))
        return date(y, m, 1).strftime('%Y-%m')
    except (ValueError, TypeError):
        return None


def _first(month: str) -> date:
    y, m = (int(x) for x in month.split('-'))
    return date(y, m, 1)


def saturdays(month: str) -> List[date]:
    d = _first(month)
    d += timedelta(days=(5 - d.weekday()) % 7)
    out = []
    while d.strftime('%Y-%m') == month:
        out.append(d)
        d += timedelta(days=7)
    return out


def holiday_name(d: date) -> str:
    if (d.month, d.day) == (5, 1):
        return '노동절'
    try:
        import holidays
        return holidays.KR(years=d.year).get(d) or ''
    except Exception:
        return ''


def kday(d: date) -> str:
    return f'{d.month}/{d.day}(토)' if d.weekday() == 5 else lv.kdate(d, weekday=True)


def _add_months(month: str, n: int) -> str:
    return _month_of(lv.add_months(_first(month), n))


def editable_until(today: date) -> str:
    return _add_months(_month_of(today), AHEAD_MONTHS)


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


# ── 읽기 ──────────────────────────────────────────────────────────────
def month_row(month: str) -> dict:
    with lv._conn() as conn:
        r = conn.execute('SELECT * FROM sat_off_months WHERE month=?', (month,)).fetchone()
    return dict(r) if r else {'month': month, 'status': OPEN}


def picks(month: str) -> Dict[str, List[str]]:
    """{'YYYY-MM-DD': [이름…]} — 그달 토요일마다 쉬는 사람."""
    out = {d.isoformat(): [] for d in saturdays(month)}
    with lv._conn() as conn:
        rows = conn.execute('SELECT day, name FROM sat_off_days WHERE day LIKE ? ORDER BY day, name',
                            (month + '-%',)).fetchall()
    for day, name in rows:
        out.setdefault(day, []).append(name)
    return out


def counts(month: str) -> Dict[str, int]:
    c: Dict[str, int] = {}
    for names in picks(month).values():
        for n in names:
            c[n] = c.get(n, 0) + 1
    return c


def work_saturdays(name: str, start: date, end: date) -> Dict[date, bool]:
    """휴가 일수용 — 기간 안 토요일 중 휴무표가 확정된 달의 것만 {토요일: 근무하는 날인가}.
    대상자가 아니거나 확정 전 달이면 넣지 않는다(→ 예전 규칙)."""
    sats = []
    d = start + timedelta(days=(5 - start.weekday()) % 7)
    while d <= end:
        sats.append(d)
        d += timedelta(days=7)
    if not sats or not is_member(name):
        return {}
    months = sorted({_month_of(s) for s in sats})
    with lv._conn() as conn:
        q = ','.join('?' * len(months))
        done = {r[0] for r in conn.execute(f'SELECT month FROM sat_off_months WHERE status=? AND month IN ({q})',
                                           (CONFIRMED, *months))}
        if not done:
            return {}
        off = {r[0] for r in conn.execute('SELECT day FROM sat_off_days WHERE name=? AND day BETWEEN ? AND ?',
                                          (name, sats[0].isoformat(), sats[-1].isoformat()))}
    return {s: s.isoformat() not in off for s in sats if _month_of(s) in done}


def _changes(month: str, where: str = '', args: tuple = ()) -> List[dict]:
    with lv._conn() as conn:
        rows = conn.execute(f'SELECT * FROM sat_off_changes WHERE month=? {where} ORDER BY id DESC',
                            (month, *args)).fetchall()
    out = []
    for r in rows:
        c = dict(r)
        c['add_days'] = json.loads(c['add_days'] or '[]')
        c['remove_days'] = json.loads(c['remove_days'] or '[]')
        c['status_label'] = CHANGE_LABELS.get(c['status'], c['status'])
        c['what'] = change_text(c)
        out.append(c)
    return out


def get_change(cid: int) -> Optional[dict]:
    with lv._conn() as conn:
        r = conn.execute('SELECT month FROM sat_off_changes WHERE id=?', (cid,)).fetchone()
    if not r:
        return None
    return next((c for c in _changes(r[0], 'AND id=?', (cid,))), None)


def change_text(c: dict) -> str:
    rm = [kday(lv._d(x)) for x in c.get('remove_days') or []]
    ad = [kday(lv._d(x)) for x in c.get('add_days') or []]
    if len(rm) == 1 and len(ad) == 1:
        return f'{rm[0]} → {ad[0]}'
    parts = []
    if rm:
        parts.append(', '.join(rm) + ' 빼기')
    if ad:
        parts.append(', '.join(ad) + ' 넣기')
    return ' · '.join(parts)


def recent_log(month: str, limit: int = 20) -> List[dict]:
    with lv._conn() as conn:
        rows = conn.execute('SELECT at, actor, action, detail FROM sat_off_log WHERE month=? ORDER BY id DESC LIMIT ?',
                            (month, limit)).fetchall()
    return [dict(r) for r in rows]


def view(month: str, sid: str, name: str, today: Optional[date] = None) -> dict:
    """달력 화면 한 달치 — 토요일마다 쉬는 사람, 사람별 횟수, 내 역할로 할 수 있는 것."""
    today = today or date.today()
    row = month_row(month)
    st = row['status']
    r = role(sid, name)
    p = picks(month)
    c = counts(month)
    sats = []
    for d in saturdays(month):
        hol = holiday_name(d)
        sats.append({'day': d.isoformat(), 'label': kday(d), 'holiday': hol, 'past': d < today,
                     'names': p.get(d.isoformat(), []), 'mine': name in p.get(d.isoformat(), [])})
    in_range = _month_of(today) <= month <= editable_until(today)
    can_edit_open = st == OPEN and in_range and (r['member'] or r['manager'])
    changes = _changes(month)
    show_all = r['approver'] or r['manager']
    my_wait = any(c['status'] == CH_WAIT and c['name'] == name for c in changes)   # 대기 중이면 새 신청은 그 뒤에
    return {
        'month': month, 'status': st, 'status_label': STATUS_LABELS[st],
        'submitted_by': row.get('submitted_by') or '', 'submitted_at': row.get('submitted_at') or '',
        'decided_by': row.get('decided_by') or '', 'decided_at': row.get('decided_at') or '',
        'note': row.get('note') or '',
        'saturdays': sats, 'limit': MONTHLY_LIMIT,
        'members': [{'name': m['name'], 'count': c.get(m['name'], 0)} for m in members()],
        'me': {'name': name, **r},
        'can_toggle': can_edit_open or r['approver'],
        'pick_others': r['manager'],                       # 남의 이름도 넣고 뺄 수 있음(작성 중일 때)
        'can_submit': st == OPEN and r['manager'] and in_range,
        'can_decide': st == WAIT and r['approver'],
        'can_request_change': st == CONFIRMED and (r['member'] or r['manager']) and not r['approver'] and not my_wait,
        'changes': [x for x in changes if show_all or x['name'] == name],
        'log': recent_log(month) if show_all else [],
    }


def calendar_month(year: int, month: int) -> List[dict]:
    """전사 휴가 달력용 — 그달 토요일 휴무 [{day, names, status}]. 비어 있는 토요일은 뺀다."""
    m = f'{year:04d}-{month:02d}'
    st = month_row(m)['status']
    return [{'day': d, 'names': n, 'status': st, 'status_label': STATUS_LABELS[st]}
            for d, n in picks(m).items() if n]


# ── 고치기 ────────────────────────────────────────────────────────────
def _log(conn, actor: str, action: str, month: str, detail: str) -> None:
    conn.execute('INSERT INTO sat_off_log (at, actor, action, month, detail) VALUES (?, ?, ?, ?, ?)',
                 (_now(), actor, action, month, detail))


def _check_day(d: Optional[date], today: date, approver: bool) -> str:
    if not d:
        raise SatOffError('날짜를 다시 골라 주세요.')
    if d.weekday() != 5:
        raise SatOffError('토요일만 고를 수 있어요.')
    if holiday_name(d):
        raise SatOffError(f'{kday(d)}은 공휴일({holiday_name(d)})이라 모두 쉬는 날이에요.')
    month = _month_of(d)
    if not approver:
        if d < today:
            raise SatOffError('지난 토요일은 바꿀 수 없어요.')
        if month > editable_until(today):
            raise SatOffError(f'{AHEAD_MONTHS}달 뒤까지만 고를 수 있어요.')
    return month


def toggle(day: Optional[date], target: str, on: bool, sid: str, actor_name: str,
           today: Optional[date] = None) -> dict:
    """토요일 칸에 이름 넣기(on)·빼기. 반환 {names, count, warn}."""
    today = today or date.today()
    r = role(sid, actor_name)
    month = _check_day(day, today, r['approver'])
    m = _member(target)
    if not m:
        raise SatOffError(f'{target} 님은 토요일 휴무 대상(기술영업)이 아니에요.')
    if target != actor_name and not (r['manager'] or r['approver']):
        raise PermissionError('다른 사람 이름은 넣거나 뺄 수 없어요.')
    if target == actor_name and not (r['member'] or r['manager']):
        raise PermissionError('토요일 휴무 대상이 아니에요.')
    st = month_row(month)['status']
    if st != OPEN and not r['approver']:
        raise SatOffError('결재 대기 중이거나 확정된 달이에요. 바꾸려면 [변경 신청]을 해 주세요.'
                          if st == CONFIRMED else '결재 대기 중인 달이라 지금은 바꿀 수 없어요.')
    key = day.isoformat()
    with lv._conn() as conn:
        if on:
            conn.execute('INSERT OR IGNORE INTO sat_off_days (day, name, slack_id, added_by, added_at) '
                         'VALUES (?, ?, ?, ?, ?)', (key, target, m['slack_id'], actor_name, _now()))
        else:
            conn.execute('DELETE FROM sat_off_days WHERE day=? AND name=?', (key, target))
        how = '넣음' if on else '뺌'
        _log(conn, actor_name, 'toggle', month,
             f'{kday(day)} {target} {how}' + (f' (확정 뒤 직접 수정)' if st != OPEN else ''))
    n = counts(month).get(target, 0)
    warn = f'{target} 님은 이번 달 {n}번째예요. 한 달 {MONTHLY_LIMIT}번이 기본이에요.' if on and n > MONTHLY_LIMIT else ''
    return {'names': picks(month).get(key, []), 'count': n, 'warn': warn}


def submit(month: str, sid: str, actor_name: str, today: Optional[date] = None) -> dict:
    """작성 중 → 결재 대기."""
    today = today or date.today()
    if not role(sid, actor_name)['manager']:
        raise PermissionError('결재는 취합 담당(빈승정 대리)이 올려요.')
    row = month_row(month)
    if row['status'] != OPEN:
        raise SatOffError(f"이미 {STATUS_LABELS[row['status']]}인 달이에요.")
    if month < _month_of(today):
        raise SatOffError('지난 달은 결재를 올릴 수 없어요.')
    if not any(picks(month).values()):
        raise SatOffError('아직 아무도 고르지 않았어요.')
    now = _now()
    with lv._conn() as conn:
        conn.execute('INSERT INTO sat_off_months (month, status, submitted_by, submitted_at) VALUES (?, ?, ?, ?) '
                     'ON CONFLICT(month) DO UPDATE SET status=excluded.status, submitted_by=excluded.submitted_by, '
                     'submitted_at=excluded.submitted_at, decided_by=NULL, decided_at=NULL, note=NULL',
                     (month, WAIT, actor_name, now))
        _log(conn, actor_name, 'submit', month, '결재 올림')
    return month_row(month)


def decide(month: str, approve: bool, sid: str, actor_name: str, note: str = '') -> dict:
    """결재 대기 → 확정(승인) 또는 작성 중(돌려보냄). 박정우·박용구 둘 중 한 분."""
    if sid not in approver_ids():
        raise PermissionError('승인은 박정우 실장님이나 대표님만 할 수 있어요.')
    note = (note or '').strip()[:300]
    row = month_row(month)
    if row['status'] != WAIT:
        raise SatOffError('결재 대기 중인 달이 아니에요. 이미 다른 분이 처리했을 수 있어요. 새로 고쳐 주세요.')
    if not approve and not note:
        raise SatOffError('돌려보내는 이유를 적어 주세요.')
    with lv._conn() as conn:
        cur = conn.execute('UPDATE sat_off_months SET status=?, decided_by=?, decided_at=?, note=? '
                           'WHERE month=? AND status=?',
                           (CONFIRMED if approve else OPEN, actor_name, _now(), note or None, month, WAIT))
        if cur.rowcount != 1:
            raise SatOffError('이미 다른 분이 처리했어요. 새로 고쳐 주세요.')
        _log(conn, actor_name, 'approve' if approve else 'return', month, note or ('승인' if approve else ''))
    return month_row(month)


def _norm_days(days, month: str) -> List[str]:
    out = []
    for x in days or []:
        d = lv._d(x)
        if not d or _month_of(d) != month:
            raise SatOffError('그달 토요일만 고를 수 있어요.')
        if d.isoformat() not in out:
            out.append(d.isoformat())
    return sorted(out)


def _check_change(month: str, name: str, add: List[str], remove: List[str], today: date, approver: bool) -> None:
    if not add and not remove:
        raise SatOffError('바꿀 토요일을 골라 주세요.')
    if set(add) & set(remove):
        raise SatOffError('같은 날을 넣고 빼려고 했어요. 다시 골라 주세요.')
    p = picks(month)
    for x in add:
        _check_day(lv._d(x), today, approver)
        if name in p.get(x, []):
            raise SatOffError(f'{kday(lv._d(x))}은 이미 쉬는 날이에요.')
    for x in remove:
        _check_day(lv._d(x), today, approver)
        if name not in p.get(x, []):
            raise SatOffError(f'{kday(lv._d(x))}은 쉬는 날이 아니에요.')


def _apply(conn, month: str, name: str, slack_id: str, add: List[str], remove: List[str], actor: str) -> None:
    now = _now()
    for x in remove:
        conn.execute('DELETE FROM sat_off_days WHERE day=? AND name=?', (x, name))
    for x in add:
        conn.execute('INSERT OR IGNORE INTO sat_off_days (day, name, slack_id, added_by, added_at) VALUES (?, ?, ?, ?, ?)',
                     (x, name, slack_id, actor, now))


def request_change(month: str, target: str, add, remove, reason: str, sid: str, actor_name: str,
                   today: Optional[date] = None) -> dict:
    """확정된 달 바꾸기 — 승인권자가 아니면 결재 대기로 남고, 승인권자가 하면 바로 반영."""
    today = today or date.today()
    r = role(sid, actor_name)
    if month_row(month)['status'] != CONFIRMED:
        raise SatOffError('확정된 달만 변경 신청을 해요. 작성 중인 달은 달력에서 바로 바꾸면 돼요.')
    m = _member(target)
    if not m:
        raise SatOffError(f'{target} 님은 토요일 휴무 대상(기술영업)이 아니에요.')
    if target != actor_name and not (r['manager'] or r['approver']):
        raise PermissionError('다른 사람의 변경 신청은 할 수 없어요.')
    reason = (reason or '').strip()[:300]
    if not reason:
        raise SatOffError('바꾸는 이유를 적어 주세요.')
    add, remove = _norm_days(add, month), _norm_days(remove, month)
    _check_change(month, target, add, remove, today, r['approver'])
    if _changes(month, 'AND name=? AND status=?', (target, CH_WAIT)):
        raise SatOffError('이미 결재 대기 중인 변경 신청이 있어요. 그 신청을 취소하고 다시 해 주세요.')
    now = _now()
    direct = r['approver']
    with lv._conn() as conn:
        cur = conn.execute('INSERT INTO sat_off_changes (month, name, slack_id, add_days, remove_days, reason, status, '
                           'requested_by, requested_at, decided_by, decided_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                           (month, target, m['slack_id'], json.dumps(add), json.dumps(remove), reason,
                            CH_OK if direct else CH_WAIT, actor_name, now,
                            actor_name if direct else None, now if direct else None))
        cid = cur.lastrowid
        if direct:
            _apply(conn, month, target, m['slack_id'], add, remove, actor_name)
        _log(conn, actor_name, 'change_direct' if direct else 'change_request', month,
             f"{target} {change_text({'add_days': add, 'remove_days': remove})}")
    return get_change(cid)


def decide_change(cid: int, approve: bool, sid: str, actor_name: str, note: str = '',
                  today: Optional[date] = None) -> dict:
    today = today or date.today()
    if sid not in approver_ids():
        raise PermissionError('승인은 박정우 실장님이나 대표님만 할 수 있어요.')
    c = get_change(cid)
    if not c:
        raise LookupError('없는 변경 신청이에요.')
    if c['status'] != CH_WAIT:
        raise SatOffError(f"이미 {c['status_label']} 처리된 신청이에요. 새로 고쳐 주세요.")
    note = (note or '').strip()[:300]
    if not approve and not note:
        raise SatOffError('반려하는 이유를 적어 주세요.')
    if approve:
        _check_change(c['month'], c['name'], c['add_days'], c['remove_days'], today, True)   # 그사이 바뀌었는지
    with lv._conn() as conn:
        cur = conn.execute('UPDATE sat_off_changes SET status=?, decided_by=?, decided_at=?, decide_note=? '
                           'WHERE id=? AND status=?',
                           (CH_OK if approve else CH_NO, actor_name, _now(), note or None, cid, CH_WAIT))
        if cur.rowcount != 1:
            raise SatOffError('이미 다른 분이 처리했어요. 새로 고쳐 주세요.')
        if approve:
            _apply(conn, c['month'], c['name'], c['slack_id'], c['add_days'], c['remove_days'], actor_name)
        _log(conn, actor_name, 'change_approve' if approve else 'change_reject', c['month'],
             f"{c['name']} {c['what']}" + (f' — {note}' if note else ''))
    return get_change(cid)


def cancel_change(cid: int, sid: str, actor_name: str) -> dict:
    c = get_change(cid)
    if not c:
        raise LookupError('없는 변경 신청이에요.')
    if c['name'] != actor_name and c['requested_by'] != actor_name and not role(sid, actor_name)['manager']:
        raise PermissionError('본인 신청만 취소할 수 있어요.')
    with lv._conn() as conn:
        cur = conn.execute('UPDATE sat_off_changes SET status=?, decided_by=?, decided_at=? WHERE id=? AND status=?',
                           (CH_CANCEL, actor_name, _now(), cid, CH_WAIT))
        if cur.rowcount != 1:
            raise SatOffError('이미 처리된 신청이에요. 새로 고쳐 주세요.')
        _log(conn, actor_name, 'change_cancel', c['month'], f"{c['name']} {c['what']}")
    return get_change(cid)


# ── 휴가 일수 영향 ────────────────────────────────────────────────────
def affected_leaves(month: str) -> List[dict]:
    """확정·변경으로 일수가 달라지는 휴가 — 그달 토요일이 낀 대상자 휴가 중 저장된 일수와 다시 센 일수가 다른 것.
    자동으로 고치지 않는다(연차 차감 배분이 얽힘) — 경영지원에 알려 휴가 관리에서 고치게."""
    sats = saturdays(month)
    names = {m['name'] for m in members()}
    if not sats or not names:
        return []
    a, b = sats[0].isoformat(), sats[-1].isoformat()
    with lv._conn() as conn:
        q = ','.join('?' * len(lv.HOLDING))
        rows = conn.execute(f'SELECT * FROM leave_requests WHERE status IN ({q}) AND start <= ? AND "end" >= ?',
                            (*lv.HOLDING, b, a)).fetchall()
    out = []
    for r in rows:
        r = dict(r)
        if r['name'] not in names or r['ltype'] in lv.HALF_TYPES:
            continue
        s, e = lv._d(r['start']), lv._d(r['end'])
        if not s or not e:
            continue
        n = lv.count_days(r['ltype'], s, e, r['name'])
        if abs(n - float(r['days'] or 0)) > 1e-9:
            out.append({'id': r['id'], 'name': r['name'], 'what': lv.period_text(r), 'saved': r['days'], 'now': n})
    return out


# ── 슬랙 알림 ─────────────────────────────────────────────────────────
def _run_bg(fn) -> None:
    """화면이 슬랙 응답을 기다리지 않게 뒤에서. 시험에서는 바로 돌도록 바꿔 끼운다."""
    threading.Thread(target=fn, daemon=True).start()


def _page_url(month: str) -> str:
    from dashboard.services import gw_login
    return f'{gw_login.public_url()}/leave/saturday?month={month}'


def _month_title(month: str) -> str:
    y, m = month.split('-')
    return f'{int(y)}년 {int(m)}월'


def _blocks(text: str, month: str) -> list:
    return [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': text}},
            {'type': 'actions', 'elements': [{'type': 'button', 'action_id': 'satoff_open',
                                              'text': {'type': 'plain_text', 'text': '그룹웨어에서 보기'},
                                              'url': _page_url(month)}]}]


def _dm(client, uid: str, text: str, month: str) -> None:
    if not client or not uid:
        return
    try:
        ch = client.conversations_open(users=uid)['channel']['id']
        client.chat_postMessage(channel=ch, text=text, blocks=_blocks(text, month))
    except Exception as e:
        logger.warning(f'{_LOG} DM 실패 ({uid}): {e}')


def _post(client, text: str) -> None:
    if not client:
        return
    try:
        client.chat_postMessage(channel=lv.announce_channel(), text=text)
    except Exception as e:
        logger.warning(f'{_LOG} #휴가_일정 공지 실패: {e}')


def _client():
    from dashboard.blueprints.slack_leave import leave_client
    return leave_client()


def schedule_text(month: str) -> str:
    lines = []
    for d in saturdays(month):
        names = picks(month).get(d.isoformat(), [])
        hol = holiday_name(d)
        if hol:
            lines.append(f'• {kday(d)} 공휴일({hol})')
        elif names:
            lines.append(f"• {kday(d)} {', '.join(names)}")
    return '\n'.join(lines) or '• 쉬는 사람 없음'


def _who_sid(name: str) -> str:
    m = _member(name)
    if m and m['slack_id']:
        return m['slack_id']
    st = lv.staff_by_name(name)
    return (st or {}).get('slack_id') or ''


def notify(kind: str, month: str, actor: str, change: Optional[dict] = None, note: str = '') -> None:
    """kind: submit·approve·return·change_request·change_approve·change_reject. 실패해도 화면 처리는 그대로."""
    def run():
        client = _client()
        if not client:
            logger.info(f'{_LOG} 휴가 봇 토큰 없음 — 알림 건너뜀 ({kind} {month})')
            return
        t = _month_title(month)
        row = month_row(month)
        if kind == 'submit':
            text = (f'📅 *{t} 토요일 휴무표* 결재 요청이 왔어요. ({actor} 님이 올림)\n'
                    f'{schedule_text(month)}\n두 분 중 한 분이 승인하면 확정돼요.')
            for uid in approver_ids():
                _dm(client, uid, text, month)
        elif kind in ('approve', 'return'):
            sub = _who_sid(row.get('submitted_by') or '')
            to = [u for u in dict.fromkeys(((sub,) if sub else ()) + coordinator_ids())]
            if kind == 'approve':
                for uid in to:
                    _dm(client, uid, f'✅ {t} 토요일 휴무표를 {actor} 님이 승인했어요. 이제 확정이에요.', month)
                _post(client, f'📅 {t} 토요일 휴무가 확정됐어요.\n{schedule_text(month)}')
                bad = affected_leaves(month)
                if bad:
                    lines = '\n'.join(f"• {b['name']} {b['what']}: {lv.fmt_days(b['saved'])} → {lv.fmt_days(b['now'])}"
                                      for b in bad)
                    for uid in lv.manager_ids():
                        _dm(client, uid, f'ℹ️ {t} 토요일 휴무표가 확정돼 휴가 일수가 달라지는 신청이 있어요. '
                                         f'휴가 관리에서 일수를 고쳐 주세요.\n{lines}', month)
            else:
                for uid in to:
                    _dm(client, uid, f'↩️ {t} 토요일 휴무표를 {actor} 님이 돌려보냈어요.\n사유: {note}\n'
                                     '고쳐서 다시 올려 주세요.', month)
        elif kind == 'change_request' and change:
            text = (f"🔁 *{change['name']} 님 토요일 휴무 변경 신청* ({t})\n{change['what']}\n"
                    f"사유: {change.get('reason') or ''}\n두 분 중 한 분이 승인하면 바뀌어요.")
            for uid in approver_ids():
                _dm(client, uid, text, month)
        elif kind in ('change_approve', 'change_reject') and change:
            who = change.get('slack_id') or _who_sid(change['name'])
            if kind == 'change_approve':
                _dm(client, who, f"✅ 토요일 휴무 변경을 {actor} 님이 승인했어요. {change['what']}", month)
                _post(client, f"🔁 토요일 휴무 변경: {change['name']} {change['what']}")
                if affected_leaves(month):
                    for uid in lv.manager_ids():
                        _dm(client, uid, f"ℹ️ {change['name']} 님 토요일 휴무가 바뀌어 휴가 일수가 달라지는 신청이 있어요. "
                                         '휴가 관리에서 확인해 주세요.', month)
            else:
                _dm(client, who, f"↩️ 토요일 휴무 변경 신청을 {actor} 님이 반려했어요. {change['what']}\n사유: {note}",
                    month)
    _run_bg(lambda: _safe(run, kind, month))


def _safe(fn, kind: str, month: str) -> None:
    try:
        fn()
    except Exception as e:
        logger.error(f'{_LOG} 알림 실패 ({kind} {month}): {e}', exc_info=True)
