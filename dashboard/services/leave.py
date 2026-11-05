# -*- coding: utf-8 -*-
"""휴가 — 플렉스 휴가 이전 (2026-11-05).

결정(KiKO 11-05):
  - 종류는 플렉스 그대로. 신청 창 = 연차·반차(오전/오후)·병가·군소집훈련·경조사(세부를 고르면 일수 한도).
    관리자가 부여하는 휴가(배우자 출산·보상 등)는 플렉스 직원 화면처럼 아무에게도 안 보인다(경영지원 포함).
  - 단위는 종일·반차만(반반차·시간 없음).
  - 승인: 요청자 → JW → YG. JW 가 신청하면 YG 만, YG 가 신청하면 바로 확정.
  - 전자결재와 따로 '휴가' 봇(slack_leave). 입구 = #휴가_일정 워크플로 버튼·봇 홈 탭.
  - 신청 카드는 채널에 올리지 않고 DM 으로만. 확정되면 #휴가_일정에 '누가 언제 휴가/반차'만(종류·사유 없음).
  - 슬랙이 없는 직원(미화·시공)은 경영지원이 대신 신청한다.
  - 연차는 지금 입사일 기준. 2027-01-01 회계연도 기준으로 바꾼다(정산표는 따로 — FISCAL_FROM).

장부: 연차는 '부여(grant)' 줄마다 쓸 수 있는 기간이 있고, 신청은 그 휴가 첫날에 살아 있는 부여에서
소멸이 빠른 것부터 떼어 쓴다(leave_uses). 승인 대기 중에도 떼어 두어 두 번 신청해 넘치는 일이 없게 하고,
반려·취소면 돌려놓는다.
"""
import json
import logging
import os
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# ── 사람 ──────────────────────────────────────────────────────────────
_DEFAULT_APPROVERS = ('U051ENK8AAY', 'U050XNG5VBR')   # 박정우(JW) → 박용구(YG)
ANNOUNCE_CHANNEL_DEFAULT = 'C0BDR06RALR'              # #휴가_일정


def _env_ids(name: str) -> tuple:
    return tuple(x.strip() for x in os.getenv(name, '').split(',') if x.strip())


def approver_ids() -> tuple:
    """승인 순서대로(첫 줄 → 다음 줄)."""
    return _env_ids('LEAVE_APPROVER_IDS') or _DEFAULT_APPROVERS


def manager_ids() -> tuple:
    """경영지원 — 다른 직원 대신 신청·관리자 부여 휴가·확정 뒤 취소. 정산 승인 명단과 같다."""
    from dashboard.services.settlement_checkers import checker_ids
    return checker_ids()


def announce_channel() -> str:
    return os.getenv('LEAVE_ANNOUNCE_CHANNEL', '').strip() or ANNOUNCE_CHANNEL_DEFAULT


# ── 종류 ──────────────────────────────────────────────────────────────
ANNUAL, HALF_AM, HALF_PM = 'annual', 'half_am', 'half_pm'
SICK, MILITARY, EVENT = 'sick', 'military', 'event'
SICK_DAYS_PER_YEAR = 30

# key: (이름, 연차에서 빠지나, 증빙 필요, 관리자 부여)
TYPES: Dict[str, Tuple[str, bool, bool, bool]] = {
    ANNUAL: ('연차', True, False, False),
    HALF_AM: ('반차 (오전)', True, False, False),
    HALF_PM: ('반차 (오후)', True, False, False),
    SICK: ('병가', False, False, False),
    MILITARY: ('군소집훈련', False, True, False),
    EVENT: ('경조사', False, True, False),
    'spouse_birth': ('배우자 출산', False, True, True),
    'reward': ('보상', False, False, True),
    'health': ('보건', False, False, True),
    'maternity': ('산전후', False, True, True),
    'infertility': ('난임 치료', False, True, True),
    'miscarriage': ('유산·사산', False, True, True),
}
HALF_TYPES = (HALF_AM, HALF_PM)

# 경조사 세부: (이름, 한도 일수) — 플렉스 설정 그대로
EVENT_SUBS: Dict[str, Tuple[str, int]] = {
    'mourn1': ('조의 - 부모/배우자/자녀', 5),
    'mourn2': ('조의 - 조부모/형제/자매', 3),
    'wed_self': ('결혼 - 본인', 5),
    'wed_child': ('결혼 - 자녀', 1),
    'spouse_loss': ('배우자 유산·사산', 5),
}


def type_label(ltype: str, sub: str = '') -> str:
    base = TYPES.get(ltype, (ltype,))[0]
    if ltype == EVENT and sub in EVENT_SUBS:
        return f'{base} ({EVENT_SUBS[sub][0]})'
    return base


def deducts(ltype: str) -> bool:
    return TYPES.get(ltype, ('', False))[1]


def needs_proof(ltype: str) -> bool:
    return TYPES.get(ltype, ('', False, False))[2]


def type_options(for_manager: bool = False) -> List[Tuple[str, str]]:
    """신청 창 종류 — 관리자 부여 휴가(배우자 출산·보상·보건·산전후·난임·유산사산)는 경영지원에게도 안 보인다
    (11-05 KiKO: 플렉스 직원 화면과 같게). 표(TYPES)엔 남겨 두어 나중에 다시 켤 수 있게."""
    return [(k, v[0]) for k, v in TYPES.items() if not v[3]]


# ── 날짜 ──────────────────────────────────────────────────────────────
def is_workday(d: date) -> bool:
    if d.weekday() >= 5:
        return False
    try:
        import holidays
        return d not in holidays.KR(years=d.year)
    except Exception:
        return True


def count_days(ltype: str, start: date, end: Optional[date]) -> float:
    """쓰는 일수 — 주말·공휴일은 빼고, 반차는 0.5."""
    if ltype in HALF_TYPES:
        return 0.5 if is_workday(start) else 0.0
    end = end or start
    n, d = 0, start
    while d <= end:
        if is_workday(d):
            n += 1
        d += timedelta(days=1)
    return float(n)


def add_months(d: date, n: int) -> date:
    y, m = divmod(d.month - 1 + n, 12)
    y, m = d.year + y, m + 1
    for day in (d.day, 30, 29, 28):
        try:
            return date(y, m, day)
        except ValueError:
            continue
    raise ValueError(d)


def add_years(d: date, n: int) -> date:
    try:
        return d.replace(year=d.year + n)
    except ValueError:            # 2/29 입사
        return d.replace(year=d.year + n, day=28)


def fmt_days(x: float) -> str:
    return f'{x:g}일'


def _d(s) -> Optional[date]:
    if isinstance(s, date):
        return s
    try:
        return datetime.strptime(str(s)[:10], '%Y-%m-%d').date()
    except (TypeError, ValueError):
        return None


_WD = '월화수목금토일'


def kdate(d: date, weekday: bool = False) -> str:
    s = f'{d.month}월 {d.day}일'
    return f'{s} ({_WD[d.weekday()]})' if weekday else s


def period_text(row: dict, weekday: bool = True) -> str:
    s, e = _d(row['start']), _d(row.get('end') or row['start'])
    if row['ltype'] in HALF_TYPES:
        return f"{kdate(s, weekday)} {'오전' if row['ltype'] == HALF_AM else '오후'}"
    if e == s:
        return kdate(s, weekday)
    return f'{kdate(s, weekday)} ~ {kdate(e, weekday)}'


# ── 장부 ──────────────────────────────────────────────────────────────
_DDL = [
    '''CREATE TABLE IF NOT EXISTS leave_staff (
        name TEXT PRIMARY KEY,
        slack_id TEXT,
        hire_date TEXT,
        active INTEGER NOT NULL DEFAULT 1,
        accrue_from TEXT,
        note TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS leave_grants (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        days REAL NOT NULL,
        valid_from TEXT NOT NULL,
        valid_to TEXT NOT NULL,
        source TEXT UNIQUE,
        note TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    )''',
    'CREATE INDEX IF NOT EXISTS idx_leave_grants_name ON leave_grants(name)',
    '''CREATE TABLE IF NOT EXISTS leave_requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        slack_id TEXT,
        submitter_id TEXT,
        submitter_name TEXT,
        ltype TEXT NOT NULL,
        sub TEXT,
        start TEXT NOT NULL,
        "end" TEXT NOT NULL,
        days REAL NOT NULL DEFAULT 0,
        reason TEXT,
        status TEXT NOT NULL,
        dm_refs TEXT NOT NULL DEFAULT '[]',
        announce_ch TEXT,
        announce_ts TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    )''',
    'CREATE INDEX IF NOT EXISTS idx_leave_requests_name ON leave_requests(name)',
    '''CREATE TABLE IF NOT EXISTS leave_uses (
        request_id INTEGER NOT NULL,
        grant_id INTEGER NOT NULL,
        days REAL NOT NULL
    )''',
    'CREATE INDEX IF NOT EXISTS idx_leave_uses_req ON leave_uses(request_id)',
    '''CREATE TABLE IF NOT EXISTS leave_steps (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        request_id INTEGER NOT NULL,
        step TEXT NOT NULL,
        actor_id TEXT,
        actor_name TEXT,
        result TEXT NOT NULL,
        reason TEXT,
        at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    )''',
]
_table_ready = False

WAIT_1, WAIT_2, APPROVED, REJECTED, CANCELLED = 'wait_1', 'wait_2', 'approved', 'rejected', 'cancelled'
OPEN = (WAIT_1, WAIT_2)
HOLDING = (WAIT_1, WAIT_2, APPROVED)   # 연차를 잡고 있는 상태
STATUS_LABELS = {
    WAIT_1: ('🟡', '승인 대기'),
    WAIT_2: ('🟡', '승인 대기'),
    APPROVED: ('🟢', '승인 완료'),
    REJECTED: ('🔴', '반려'),
    CANCELLED: ('⚪', '취소'),
}


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


@contextmanager
def _conn():
    """사용자 DB 연결 — 처음 한 번 표를 만들고, 오류 없이 끝나면 저장한다."""
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


# ── 직원 ──────────────────────────────────────────────────────────────
def upsert_staff(name: str, hire_date: str, slack_id: str = '', accrue_from: str = '', note: str = '') -> None:
    with _conn() as conn:
        conn.execute(
            '''INSERT INTO leave_staff (name, slack_id, hire_date, active, accrue_from, note) VALUES (?,?,?,1,?,?)
               ON CONFLICT(name) DO UPDATE SET slack_id=excluded.slack_id, hire_date=excluded.hire_date,
               accrue_from=COALESCE(NULLIF(excluded.accrue_from, ''), leave_staff.accrue_from),
               note=COALESCE(NULLIF(excluded.note, ''), leave_staff.note)''',
            (name, slack_id or None, hire_date, accrue_from or None, note or None))


def staff(active_only: bool = True) -> List[dict]:
    with _conn() as conn:
        q = 'SELECT * FROM leave_staff' + (' WHERE active=1' if active_only else '') + ' ORDER BY name'
        return [dict(r) for r in conn.execute(q).fetchall()]


def staff_by_slack(uid: str) -> Optional[dict]:
    if not uid:
        return None
    with _conn() as conn:
        r = conn.execute('SELECT * FROM leave_staff WHERE slack_id=? AND active=1', (uid,)).fetchone()
        return dict(r) if r else None


def staff_by_name(name: str) -> Optional[dict]:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM leave_staff WHERE name=?', (name,)).fetchone()
        return dict(r) if r else None


# ── 부여 ──────────────────────────────────────────────────────────────
def add_grant(name: str, days: float, valid_from: date, valid_to: date, source: str = '', note: str = '') -> bool:
    """부여 한 줄. 같은 source 가 있으면 넣지 않는다(자동 발생을 매일 돌려도 한 번만). 넣었으면 True."""
    with _conn() as conn:
        cur = conn.execute(
            'INSERT OR IGNORE INTO leave_grants (name, days, valid_from, valid_to, source, note, created_at) '
            'VALUES (?,?,?,?,?,?,?)',
            (name, float(days), valid_from.isoformat(), valid_to.isoformat(), source or None, note or None, _now()))
        return cur.rowcount == 1


def grants(name: str) -> List[dict]:
    with _conn() as conn:
        return [dict(r) for r in conn.execute(
            'SELECT g.*, COALESCE((SELECT SUM(u.days) FROM leave_uses u WHERE u.grant_id=g.id), 0) AS used '
            'FROM leave_grants g WHERE g.name=? ORDER BY g.valid_to, g.id', (name,)).fetchall()]


def balance(name: str, on: Optional[date] = None) -> float:
    """그날 쓸 수 있는 연차 = 그날 살아 있는 부여 − 이미 잡힌 것(승인 대기 포함)."""
    on = on or date.today()
    total = 0.0
    for g in grants(name):
        if g['valid_from'] <= on.isoformat() <= g['valid_to']:
            total += g['days'] - g['used']
    return round(total, 2)


def _allocate(conn, rid: int, name: str, on: date, need: float) -> bool:
    """신청 rid 에 연차를 떼어 준다 — 그날 살아 있는 부여 중 소멸이 빠른 것부터. 모자라면 False(아무것도 안 뗌)."""
    rows = conn.execute(
        'SELECT g.id, g.days - COALESCE((SELECT SUM(u.days) FROM leave_uses u WHERE u.grant_id=g.id), 0) AS left '
        'FROM leave_grants g WHERE g.name=? AND g.valid_from<=? AND g.valid_to>=? ORDER BY g.valid_to, g.id',
        (name, on.isoformat(), on.isoformat())).fetchall()
    plan, rest = [], need
    for r in rows:
        if rest <= 0:
            break
        take = min(rest, r['left'])
        if take > 0:
            plan.append((r['id'], take))
            rest -= take
    if rest > 1e-9:
        return False
    for gid, take in plan:
        conn.execute('INSERT INTO leave_uses (request_id, grant_id, days) VALUES (?,?,?)', (rid, gid, take))
    return True


# ── 자동 발생 (입사일 기준) ──────────────────────────────────────────
FISCAL_FROM = date(2027, 1, 1)   # 이날부터 1년 이상 근속자의 연차는 회계연도 기준(별도 정산표로 부여)


def annual_days(years: int) -> int:
    """근속 years 년 차 연차 — 15일, 3년째부터 2년마다 1일, 최대 25일 (근로기준법 60조)."""
    return min(25, 15 + max(0, years - 1) // 2)


def due_grants(person: dict, today: date) -> List[dict]:
    """오늘까지 생겼어야 할 입사일 기준 부여(accrue_from 이후 것만 — 그 전 몫은 플렉스 잔여로 옮겨 넣었다).

    1년 미만: 한 달 개근마다 1일(최대 11), 입사 1년이 되는 날까지 쓴다.
    1년 이상: 입사 기념일에 annual_days, 다음 기념일 전날까지 쓴다. FISCAL_FROM 이후 기념일은 만들지 않는다."""
    hire = _d(person.get('hire_date'))
    if not hire:
        return []
    start = _d(person.get('accrue_from')) or hire
    name, out = person['name'], []
    first_year_end = add_years(hire, 1) - timedelta(days=1)
    for m in range(1, 12):
        d = add_months(hire, m)
        if d > first_year_end or d > today:
            break
        if d >= start:
            out.append({'name': name, 'days': 1, 'valid_from': d, 'valid_to': first_year_end,
                        'source': f'monthly:{name}:{d.isoformat()}', 'note': f'입사 {m}개월 개근'})
    n = 1
    while True:
        d = add_years(hire, n)
        if d > today or d >= FISCAL_FROM:
            break
        if d >= start:
            out.append({'name': name, 'days': annual_days(n), 'valid_from': d,
                        'valid_to': add_years(hire, n + 1) - timedelta(days=1),
                        'source': f'annual:{name}:{d.isoformat()}', 'note': f'입사 {n}년'})
        n += 1
    return out


def accrue(today: Optional[date] = None) -> List[dict]:
    """매일 돌린다 — 새로 생긴 부여만 넣고 그 목록을 돌려준다."""
    today = today or date.today()
    made = []
    for p in staff():
        for g in due_grants(p, today):
            if add_grant(g['name'], g['days'], g['valid_from'], g['valid_to'], g['source'], g['note']):
                made.append(g)
    return made


# ── 신청 ──────────────────────────────────────────────────────────────
def _row(conn, rid: int) -> Optional[dict]:
    r = conn.execute('SELECT * FROM leave_requests WHERE id=?', (rid,)).fetchone()
    return dict(r) if r else None


def get(rid: int) -> Optional[dict]:
    with _conn() as conn:
        return _row(conn, rid)


def steps(rid: int) -> List[dict]:
    with _conn() as conn:
        return [dict(r) for r in conn.execute(
            'SELECT * FROM leave_steps WHERE request_id=? ORDER BY id', (rid,)).fetchall()]


def sick_used(name: str, year: int, exclude_id: int = 0) -> float:
    with _conn() as conn:
        r = conn.execute(
            f'SELECT COALESCE(SUM(days), 0) FROM leave_requests WHERE name=? AND ltype=? AND substr(start,1,4)=? '
            f'AND status IN ({",".join("?" * len(HOLDING))}) AND id<>?',
            (name, SICK, str(year), *HOLDING, exclude_id)).fetchone()
        return float(r[0] or 0)


def overlaps(name: str, start: date, end: date) -> List[dict]:
    with _conn() as conn:
        return [dict(r) for r in conn.execute(
            f'SELECT * FROM leave_requests WHERE name=? AND status IN ({",".join("?" * len(HOLDING))}) '
            f'AND start<=? AND "end">=?', (name, *HOLDING, end.isoformat(), start.isoformat())).fetchall()]


def validate(name: str, ltype: str, sub: str, start: Optional[date], end: Optional[date],
             for_manager: bool = False) -> Dict[str, str]:
    """창 칸 이름 → 오류 문구."""
    errs: Dict[str, str] = {}
    if ltype not in dict(type_options(for_manager)):
        errs['type'] = '휴가 종류를 골라 주세요.'
        return errs
    if ltype == EVENT and sub not in EVENT_SUBS:
        errs['sub'] = '경조사 종류를 골라 주세요.'
    if not start:
        errs['start'] = '날짜를 골라 주세요.'
        return errs
    if ltype not in HALF_TYPES:
        if not end:
            errs['end'] = '마지막 날을 골라 주세요.'
            return errs
        if end < start:
            errs['end'] = '마지막 날이 시작일보다 빨라요.'
            return errs
    end = start if ltype in HALF_TYPES else end
    days = count_days(ltype, start, end)
    if days <= 0:
        errs['start'] = '주말·공휴일이라 쓸 휴가가 없어요.'
        return errs
    clash = [r for r in overlaps(name, start, end)
             if not (ltype in HALF_TYPES and r['ltype'] in HALF_TYPES and r['ltype'] != ltype)]
    if clash:
        c = clash[0]
        errs['start'] = f"이미 신청한 휴가와 겹쳐요: {type_label(c['ltype'], c.get('sub') or '')} {period_text(c, False)}"
        return errs
    if deducts(ltype):
        left = balance(name, start)
        if days > left + 1e-9:
            errs['start'] = (f'{kdate(start)} 기준 남은 연차가 {fmt_days(left)}이라 {fmt_days(days)}를 쓸 수 없어요. '
                             '연차가 더 있어야 한다면 경영지원에 문의해 주세요.')
    elif ltype == EVENT and sub in EVENT_SUBS and days > EVENT_SUBS[sub][1]:
        errs['end'] = f'{EVENT_SUBS[sub][0]}은 {EVENT_SUBS[sub][1]}일까지예요 (주말·공휴일 빼고 {fmt_days(days)}).'
    elif ltype == SICK:
        used = sick_used(name, start.year)
        if used + days > SICK_DAYS_PER_YEAR:
            errs['end'] = f'병가는 한 해 {SICK_DAYS_PER_YEAR}일까지예요 (올해 쓴 병가 {fmt_days(used)}).'
    return errs


def create(name: str, subject_slack_id: str, submitter_id: str, submitter_name: str, ltype: str, sub: str,
           start: date, end: Optional[date], reason: str = '') -> Optional[dict]:
    """신청 만들기. 첫 승인자가 신청하면 다음 승인자부터, 마지막 승인자가 신청하면 바로 확정.
    연차가 모자라면(그 사이 다른 신청이 잡아감) None."""
    end = start if ltype in HALF_TYPES else (end or start)
    days = count_days(ltype, start, end)
    appr = approver_ids()
    who = subject_slack_id or ''
    status = WAIT_1
    auto: List[Tuple[str, str]] = []
    if who and appr and who == appr[-1]:
        status = APPROVED
        auto = [('approve_1', '요청자가 승인자'), ('approve_2', '요청자가 승인자')] if len(appr) > 1 \
            else [('approve_1', '요청자가 승인자')]
    elif who and len(appr) > 1 and who == appr[0]:
        status = WAIT_2
        auto = [('approve_1', '요청자가 승인자')]
    now = _now()
    with _conn() as conn:
        cur = conn.execute(
            'INSERT INTO leave_requests (name, slack_id, submitter_id, submitter_name, ltype, sub, start, "end", '
            'days, reason, status, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)',
            (name, subject_slack_id or None, submitter_id, submitter_name, ltype, sub or None, start.isoformat(),
             end.isoformat(), days, reason or None, status, now, now))
        rid = cur.lastrowid
        if deducts(ltype) and not _allocate(conn, rid, name, start, days):
            conn.rollback()
            return None
        conn.execute('INSERT INTO leave_steps (request_id, step, actor_id, actor_name, result, at) '
                     'VALUES (?,?,?,?,?,?)', (rid, 'request', submitter_id, submitter_name, 'requested', now))
        for step, why in auto:
            conn.execute('INSERT INTO leave_steps (request_id, step, actor_id, actor_name, result, reason, at) '
                         'VALUES (?,?,?,?,?,?,?)', (rid, step, who, name, 'auto', why, now))
        conn.commit()
        return _row(conn, rid)


def current_approver(row: dict) -> str:
    appr = approver_ids()
    if row['status'] == WAIT_1 and appr:
        return appr[0]
    if row['status'] == WAIT_2 and len(appr) > 1:
        return appr[1]
    return ''


def can_act(action: str, row: dict, uid: str, today: Optional[date] = None) -> bool:
    today = today or date.today()
    if action in ('approve', 'reject'):
        return bool(uid) and uid == current_approver(row)
    if action == 'cancel':
        if uid in manager_ids():
            return row['status'] in HOLDING
        mine = uid in (row.get('slack_id'), row.get('submitter_id'))
        if row['status'] in OPEN:
            return mine
        # 확정된 휴가는 시작 전까지만 본인이 취소(지난 휴가는 경영지원)
        return mine and row['status'] == APPROVED and _d(row['start']) > today
    return False


def transition(rid: int, action: str, uid: str, actor_name: str, reason: str = '') -> Optional[dict]:
    """상태 바꾸기(두 번 눌러도 한 번만). 반려·취소면 잡아 둔 연차를 돌려놓는다."""
    now = _now()
    with _conn() as conn:
        row = _row(conn, rid)
        if not row:
            return None
        appr = approver_ids()
        if action == 'approve':
            frm = row['status']
            if frm == WAIT_1:
                nxt, step = (WAIT_2 if len(appr) > 1 else APPROVED), 'approve_1'
            elif frm == WAIT_2:
                nxt, step = APPROVED, 'approve_2'
            else:
                return None
            result = 'approved'
        elif action == 'reject':
            frm, nxt, result = row['status'], REJECTED, 'rejected'
            if frm not in OPEN:
                return None
            step = 'approve_1' if frm == WAIT_1 else 'approve_2'
        elif action == 'cancel':
            frm, nxt, step, result = row['status'], CANCELLED, 'cancel', 'cancelled'
            if frm not in HOLDING:
                return None
        else:
            return None
        cur = conn.execute('UPDATE leave_requests SET status=?, updated_at=? WHERE id=? AND status=?',
                           (nxt, now, rid, frm))
        if cur.rowcount != 1:
            conn.rollback()
            return None
        if nxt in (REJECTED, CANCELLED):
            conn.execute('DELETE FROM leave_uses WHERE request_id=?', (rid,))
        conn.execute('INSERT INTO leave_steps (request_id, step, actor_id, actor_name, result, reason, at) '
                     'VALUES (?,?,?,?,?,?,?)', (rid, step, uid, actor_name, result, reason or None, now))
        conn.commit()
        return _row(conn, rid)


def add_dm_ref(rid: int, channel: str, ts: str, role: str, uid: str) -> None:
    with _conn() as conn:
        r = conn.execute('SELECT dm_refs FROM leave_requests WHERE id=?', (rid,)).fetchone()
        refs = json.loads((r['dm_refs'] if r else None) or '[]')
        refs.append({'ch': channel, 'ts': ts, 'role': role, 'user': uid})
        conn.execute('UPDATE leave_requests SET dm_refs=? WHERE id=?', (json.dumps(refs), rid))


def dm_refs(row: dict) -> List[dict]:
    return json.loads(row.get('dm_refs') or '[]')


def set_announce(rid: int, channel: str, ts: str) -> None:
    with _conn() as conn:
        conn.execute('UPDATE leave_requests SET announce_ch=?, announce_ts=? WHERE id=?', (channel, ts, rid))


def requests_of(name: str, limit: int = 15) -> List[dict]:
    with _conn() as conn:
        return [dict(r) for r in conn.execute(
            'SELECT * FROM leave_requests WHERE name=? ORDER BY start DESC, id DESC LIMIT ?', (name, limit)).fetchall()]


# ── 문구 ──────────────────────────────────────────────────────────────
def announce_text(row: dict) -> str:
    """#휴가_일정 — 누가 언제 휴가/반차인지만(종류·사유 없음). 플렉스 알림과 같은 모양."""
    s, e = _d(row['start']), _d(row['end'])
    if row['ltype'] in HALF_TYPES:
        what = f"{kdate(s)} {'오전' if row['ltype'] == HALF_AM else '오후'} 반차입니다."
    elif s == e:
        what = f'{kdate(s)} 하루종일 휴가입니다.'
    else:
        what = f'{kdate(s)} ~ {kdate(e)} 휴가입니다.'
    return f"🌴 [{row['name']}] - {what}"


def cancelled_announce_text(row: dict) -> str:
    return f'~{announce_text(row)}~ (취소)'


def step_line(row: dict, step_rows: List[dict], names: Tuple[str, str] = ('박정우', '박용구')) -> str:
    done = {s['step']: s for s in step_rows}
    parts = []
    for i, key in enumerate(('approve_1', 'approve_2')):
        if i >= len(approver_ids()):
            break
        s = done.get(key)
        who = names[i] if i < len(names) else ''
        if s and s['result'] in ('approved', 'auto'):
            parts.append(f'✅ {who}')
        elif s and s['result'] == 'rejected':
            parts.append(f'❌ {who}')
        elif (row['status'] == WAIT_1 and i == 0) or (row['status'] == WAIT_2 and i == 1):
            parts.append(f'⏳ {who}')
        else:
            parts.append(f'· {who}')
    return '  →  '.join(parts)


def card_text(row: dict, step_rows: List[dict], role: str, left: Optional[float] = None) -> str:
    emoji, label = STATUS_LABELS.get(row['status'], ('', row['status']))
    lines = [f"🌴 *휴가 신청 — {row['name']}*  {emoji} {label}",
             f"• 종류: {type_label(row['ltype'], row.get('sub') or '')}",
             f"• 날짜: {period_text(row)}  ({fmt_days(row['days'])})"]
    if row.get('reason'):
        lines.append(f"• 사유: {row['reason']}")
    if row.get('submitter_name') and row.get('submitter_name') != row['name']:
        lines.append(f"• 대신 신청: {row['submitter_name']}")
    if deducts(row['ltype']) and left is not None:
        lines.append(f'• 남은 연차: {fmt_days(left)}' + (' (이 신청 반영)' if row['status'] in HOLDING else ''))
    lines.append(f'• 승인: {step_line(row, step_rows)}')
    if needs_proof(row['ltype']) and row['status'] in HOLDING:
        lines.append('📎 증빙(청첩장·부고·훈련 통지서 등)은 경영지원에 보내 주세요.')
    for s in step_rows:
        if s['result'] == 'rejected' and s.get('reason'):
            lines.append(f"❌ 반려 사유 ({s.get('actor_name') or ''}): {s['reason']}")
    return '\n'.join(lines)


A_APPROVE, A_REJECT, A_CANCEL = 'leave_approve', 'leave_reject', 'leave_cancel'
A_NEW, A_MINE = 'leave_new', 'leave_mine'
FN_OPEN_LEAVE = 'leave_open'   # 휴가 봇 워크플로 단계 「휴가 신청 열기」 — docs/slack_leave_manifest.yml 키와 같게
_DEFAULT_PILOTS = ('U04UL2ZLJAX',)   # 고광일


def is_open(uid: str) -> bool:
    """신청 받기 — 플렉스에서 옮기는 날(LEAVE_ENABLED=true) 전까지는 시험하는 사람(LEAVE_PILOT_IDS, 기본 KiKO)만."""
    if os.getenv('LEAVE_ENABLED', '').strip().lower() in ('1', 'true', 'yes'):
        return True
    return bool(uid) and uid in (_env_ids('LEAVE_PILOT_IDS') or _DEFAULT_PILOTS)
CB_SUBMIT, CB_REJECT = 'leave_submit', 'leave_reject_submit'


def _btn(text: str, action: str, rid: int, style: str = '', confirm: Optional[dict] = None) -> dict:
    b = {'type': 'button', 'action_id': action, 'value': str(rid), 'text': {'type': 'plain_text', 'text': text}}
    if style:
        b['style'] = style
    if confirm:
        b['confirm'] = confirm
    return b


def card_blocks(row: dict, step_rows: List[dict], role: str, left: Optional[float] = None) -> List[dict]:
    """role: approver(지금 승인할 사람) / requester / info."""
    blocks = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': card_text(row, step_rows, role, left)}}]
    els = []
    if role == 'approver' and row['status'] in OPEN:
        els = [_btn('✅ 승인', A_APPROVE, row['id'], 'primary'), _btn('반려', A_REJECT, row['id'], 'danger')]
    elif role == 'requester' and row['status'] in HOLDING and \
            (row['status'] in OPEN or _d(row['start']) > date.today()):
        els = [_btn('신청 취소', A_CANCEL, row['id'], confirm={
            'title': {'type': 'plain_text', 'text': '휴가 신청 취소'},
            'text': {'type': 'plain_text', 'text': '이 휴가 신청을 취소할까요? 쓴 연차는 돌아와요.'},
            'confirm': {'type': 'plain_text', 'text': '취소하기'},
            'deny': {'type': 'plain_text', 'text': '그대로 두기'}})]
    if els:
        blocks.append({'type': 'actions', 'block_id': f'leave_{row["id"]}', 'elements': els})
    return blocks


def mine_text(name: str, today: Optional[date] = None) -> str:
    """[내 휴가] — 남은 연차·소멸 예정·최근 신청."""
    today = today or date.today()
    left = balance(name, today)
    lines = [f'*{name} 님 남은 연차: {fmt_days(left)}*']
    live = [g for g in grants(name) if g['valid_from'] <= today.isoformat() <= g['valid_to']
            and g['days'] - g['used'] > 0]
    for g in live:
        lines.append(f"• {fmt_days(g['days'] - g['used'])}: {kdate(_d(g['valid_to']))}까지 사용")
    rows = requests_of(name)
    if rows:
        lines.append('\n*최근 신청*')
        for r in rows:
            emoji, label = STATUS_LABELS.get(r['status'], ('', r['status']))
            lines.append(f"{emoji} {period_text(r, False)} · {type_label(r['ltype'], r.get('sub') or '')} "
                         f"· {fmt_days(r['days'])} · {label}")
    else:
        lines.append('\n아직 신청한 휴가가 없어요.')
    return '\n'.join(lines)
