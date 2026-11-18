# -*- coding: utf-8 -*-
"""휴가 — 플렉스 휴가 이전 (2026-11-05).

결정(KiKO 11-05):
  - 종류는 플렉스 그대로. 신청 창 = 연차·반차(오전/오후)·병가·군소집훈련·경조사(세부를 고르면 일수 한도).
    관리자가 부여하는 휴가(배우자 출산·보상 등)는 플렉스 직원 화면처럼 아무에게도 안 보인다(경영지원 포함).
  - 단위는 종일·반차만(반반차·시간 없음).
  - 승인: 요청자 → JW → YG. JW 가 신청하면 YG 만, YG 가 신청하면 바로 확정.
  - 전자결재와 따로 '휴가' 봇(slack_leave). 입구 = #휴가_일정 워크플로 버튼·봇 홈 탭.
  - 신청 카드는 채널에 올리지 않고 DM 으로만. 확정되면 #휴가_일정에 '누가 언제 휴가/반차'만(종류·사유 없음).
  - 직원 신청 창은 본인 휴가만. 경영지원은 휴가 관리 기능으로 등록·고치기·취소·연차 조정(11-18).
  - 연차는 2026년까지 입사일 기준, **2027-01-01 부터 회계연도(1/1~12/31) 기준**(11-18 부장님 확정 — 11-17 의 2028 기본값을 당김).
    2027-01-01 = 마지막 입사 기념일~2026-12-31 근무 비례분(0.5일 올림), 2028 부터 매년 1/1 만 근속 기준(fiscal_grants).
    퇴직 때 입사일 기준보다 모자라면 그 일수만 수당 정산(resign_settlement — 고용노동부 근기 68207-620·근로개선정책과-5352,
    더 준 몫을 깎는 정산은 취업규칙 단서 없인 안 됨 — 임금근로시간정책팀-489). 취업규칙 개정·신고는 따로(경영지원).

장부: 연차는 '부여(grant)' 줄마다 쓸 수 있는 기간이 있고, 신청은 그 휴가 첫날에 살아 있는 부여에서
소멸이 빠른 것부터 떼어 쓴다(leave_uses). 승인 대기 중에도 떼어 두어 두 번 신청해 넘치는 일이 없게 하고,
반려·취소면 돌려놓는다.
"""
import json
import logging
import math
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
    """경영지원 — 이미 시작한 휴가도 취소할 수 있는 사람. 정산 승인 명단과 같다."""
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


def type_options() -> List[Tuple[str, str]]:
    """신청 창 종류 — 관리자 부여 휴가(배우자 출산·보상·보건·산전후·난임·유산사산)는 경영지원에게도 안 보인다
    (11-05 KiKO: 플렉스 직원 화면과 같게). 표(TYPES)엔 남겨 두어 나중에 다시 켤 수 있게."""
    return [(k, v[0]) for k, v in TYPES.items() if not v[3]]


# ── 날짜 ──────────────────────────────────────────────────────────────
def is_workday(d: date) -> bool:
    if d.weekday() >= 5:
        return False
    if (d.month, d.day) == (5, 1):   # 노동절 — 회사 휴무(11-05 KiKO). holidays.KR 엔 없음
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
    # 맞춤 휴가 부여(11-18, 플렉스 '맞춤 휴가 부여') — 관리자 부여 휴가(보건·난임 등)를 일수로 먼저 주면 직원이 신청.
    #   쓴 날은 따로 떼어 두지 않고 그 종류 신청(승인 대기 포함)을 기간 안에서 더해 본다(취소·반려가 바로 돌아오게)
    '''CREATE TABLE IF NOT EXISTS leave_custom_grants (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        ltype TEXT NOT NULL,
        days REAL NOT NULL,
        valid_from TEXT NOT NULL,
        valid_to TEXT NOT NULL,
        note TEXT,
        created_by TEXT,
        created_by_name TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        cancelled_at TEXT,
        cancelled_by_name TEXT
    )''',
    'CREATE INDEX IF NOT EXISTS idx_leave_custom_name ON leave_custom_grants(name)',
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
            # 변경 신청(11-18): 승인된 휴가를 바꾸는 새 신청이 어느 휴가를 대신하는지
            cols = {r[1] for r in conn.execute('PRAGMA table_info(leave_requests)')}
            if 'replaces' not in cols:
                conn.execute('ALTER TABLE leave_requests ADD COLUMN replaces INTEGER')
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


_HR_RESIGNED = '퇴사'   # dashboard.services.hr.STATUS_RESIGNED


def _hr_people_by_slack() -> Dict[str, dict]:
    """같은 DB 의 인사카드(hr_people) — 슬랙 ID → {status, hire_date}. 같은 슬랙 ID 카드가 여러 장이면 퇴사가 아닌
    카드(재입사). 인사카드 표가 없으면(테스트·첫 설치) 빈 dict."""
    with _conn() as conn:
        if not conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='hr_people'").fetchone():
            return {}
        rows = conn.execute("SELECT status, hire_date, slack_id FROM hr_people "
                            "WHERE slack_id IS NOT NULL AND slack_id != '' ORDER BY id").fetchall()
    out: Dict[str, dict] = {}
    for r in rows:
        r = dict(r)
        cur = out.get(r['slack_id'])
        if cur is None or (cur['status'] == _HR_RESIGNED and r['status'] != _HR_RESIGNED):
            out[r['slack_id']] = r
    return out


def _hr_overlay(rows: List[dict]) -> List[dict]:
    """사람 정보는 인사카드가 정본 — 인사카드가 퇴사면 빼고, 입사일은 인사카드 값(2026-11-16 점검 묶음 G).
    휴가 명단은 플렉스 이관 때 복사해 둔 사본이라, 인사카드를 고쳐도(퇴사·입사일 정정) 연차 발생·슬랙 신청은
    옛 값으로 돌았다. 인사카드를 못 읽으면 명단 그대로(휴가 기능을 멈추지 않게)."""
    try:
        hr = _hr_people_by_slack()
    except Exception as exc:
        logger.warning(f'[LEAVE] 인사카드 읽기 실패 — 휴가 명단 그대로 사용: {exc}')
        return rows
    out = []
    for r in rows:
        p = hr.get(r.get('slack_id') or '')
        if p:
            if p.get('status') == _HR_RESIGNED:
                continue
            if p.get('hire_date') and p['hire_date'] != r.get('hire_date'):
                r = dict(r, hire_date=p['hire_date'])
        out.append(r)
    return out


def staff(active_only: bool = True) -> List[dict]:
    with _conn() as conn:
        q = 'SELECT * FROM leave_staff' + (' WHERE active=1' if active_only else '') + ' ORDER BY name'
        rows = [dict(r) for r in conn.execute(q).fetchall()]
    return _hr_overlay(rows) if active_only else rows


def staff_by_slack(uid: str) -> Optional[dict]:
    if not uid:
        return None
    with _conn() as conn:
        r = conn.execute('SELECT * FROM leave_staff WHERE slack_id=? AND active=1', (uid,)).fetchone()
    rows = _hr_overlay([dict(r)]) if r else []
    return rows[0] if rows else None


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


def balance(name: str, on: Optional[date] = None, until: Optional[date] = None, exclude_id: int = 0) -> float:
    """그날 쓸 수 있는 연차 = 그날 살아 있는 부여 − 이미 잡힌 것(승인 대기 포함).
    until 이면 그날까지 소멸하지 않는 부여만(휴가 도중 소멸하는 연차로 뒷날을 채우지 않게, 11-05 점검).
    exclude_id: 그 신청이 잡은 몫은 없는 셈 치고(고치기·변경 신청 때 자기 몫을 다시 쓸 수 있게, 11-18)."""
    on = on or date.today()
    last = (until or on).isoformat()
    total = 0.0
    alive = set()
    for g in grants(name):
        if g['valid_from'] <= on.isoformat() and last <= g['valid_to']:
            total += g['days'] - g['used']
            alive.add(g['id'])
    if exclude_id and alive:
        with _conn() as conn:
            for r in conn.execute('SELECT grant_id, days FROM leave_uses WHERE request_id=?', (exclude_id,)):
                if r['grant_id'] in alive:
                    total += r['days']
    return round(total, 2)


def _alloc_until(end: date, today: Optional[date] = None) -> date:
    """새 신청이 쓸 연차가 살아 있어야 하는 마지막 날 — 휴가 마지막 날과 오늘 중 늦은 날(지난 날짜 신청 대비)."""
    return max(end, today or date.today())


def _allocate(conn, rid: int, name: str, on: date, need: float, until: Optional[date] = None) -> bool:
    """신청 rid 에 연차를 떼어 준다 — 첫날 살아 있고 마지막 날까지 소멸하지 않는 부여 중 소멸이 빠른 것부터.
    모자라면 False(아무것도 안 뗌)."""
    rows = conn.execute(
        'SELECT g.id, g.days - COALESCE((SELECT SUM(u.days) FROM leave_uses u WHERE u.grant_id=g.id), 0) AS left '
        'FROM leave_grants g WHERE g.name=? AND g.valid_from<=? AND g.valid_to>=? ORDER BY g.valid_to, g.id',
        (name, on.isoformat(), (until or on).isoformat())).fetchall()
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
FISCAL_FROM = date(2027, 1, 1)   # 이날부터 회계연도 기준(11-18 부장님 확정). 이날·그 뒤 입사 기념일 부여는 만들지 않고 fiscal_grants 로
FISCAL_PREGRANT_DAYS = 17        # 다음 해 1/1 부여를 12-15 부터 미리 넣는다(12월에 1월 휴가 신청 가능, 설계 D10)


def annual_days(years: int) -> int:
    """근속 years 년 차 연차 — 15일, 3년째부터 2년마다 1일, 최대 25일 (근로기준법 60조)."""
    return min(25, 15 + max(0, years - 1) // 2)


def _prorate(base: float, days: int) -> float:
    """비례분 — base × 일수/365, 0.5일 단위로 올림(설계 D2)."""
    return math.ceil(base * days / 365 * 2 - 1e-9) / 2


def completed_years(hire: date, on: date) -> int:
    """on 날까지 채운 근속 연수(기념일이 on 이하인 횟수)."""
    n = 0
    while add_years(hire, n + 1) <= on:
        n += 1
    return n


def fiscal_grants(person: dict, today: date) -> List[dict]:
    """회계연도 부여(11-18 확정) — 해 Y 부여는 Y-1 년 12-15 부터 미리 만든다.
      2027(전환 해): 1년 이상 = annual_days(다음 기념일 근속연수) × (마지막 기념일~2026-12-31 일수)/365,
                    1년 미만 = 15 × (입사일~2026-12-31 일수)/365. 모두 0.5일 올림.
      2028~: 1/1 기준 만 근속 1년 이상 = annual_days(만 근속), 전년도 입사자 = 15 × 전년도 근무일/365.
      그해 1/1 이후 입사자는 그해 몫 없음(1년이 될 때까지 월차는 due_grants 가 그대로)."""
    hire = _d(person.get('hire_date'))
    if not hire:
        return []
    name, out = person['name'], []
    y = FISCAL_FROM.year
    while today >= date(y, 1, 1) - timedelta(days=FISCAL_PREGRANT_DAYS):
        first, prev_end = date(y, 1, 1), date(y - 1, 12, 31)
        if hire < first:
            if y == FISCAL_FROM.year:
                k = completed_years(hire, prev_end)
                frm = add_years(hire, k) if k >= 1 else hire
                base = annual_days(k + 1) if k >= 1 else 15
                span = (prev_end - frm).days + 1
                days = _prorate(base, span)
                note = f'회계연도 전환 비례분 ({frm.isoformat()}~{prev_end.isoformat()} {span}일, 기준 {base}일)'
            else:
                k = completed_years(hire, first)
                if k >= 1:
                    days, note = float(annual_days(k)), f'{y}년 회계연도 연차 (1/1 기준 만 {k}년)'
                else:
                    span = (prev_end - hire).days + 1
                    days, note = _prorate(15, span), f'{y}년 회계연도 연차 (전년도 입사 {span}일 비례)'
            if days > 0:
                out.append({'name': name, 'days': days, 'valid_from': first, 'valid_to': date(y, 12, 31),
                            'source': f'fiscal:{name}:{y}', 'note': note})
        y += 1
    return out


def hire_basis_grants(person: dict, until: date) -> List[dict]:
    """입사일 기준이었다면 until 날까지 생겼을 부여 — 월차(1개월 개근 다음 날)·기념일 연차(1년 근로 다음 날,
    대법원 2021다227100: 그 전에 그만두면 그 해 몫은 없다). 퇴직 정산용이라 FISCAL_FROM·accrue_from 을 보지 않는다."""
    return _hire_grants(person, until, cut=None, start=_d(person.get('hire_date')))


def due_grants(person: dict, today: date) -> List[dict]:
    """오늘까지 생겼어야 할 부여 — 입사일 기준 월차·기념일(accrue_from 이후, FISCAL_FROM 전 기념일만) + 회계연도 부여.

    1년 미만: 한 달 개근마다 1일(최대 11), 입사 1년이 되는 날까지 쓴다.
    1년 이상: 입사 기념일에 annual_days, 다음 기념일 전날까지 쓴다. FISCAL_FROM 이후 기념일은 만들지 않는다(fiscal_grants)."""
    start = _d(person.get('accrue_from')) or _d(person.get('hire_date'))
    return _hire_grants(person, today, cut=FISCAL_FROM, start=start) + fiscal_grants(person, today)


def _hire_grants(person: dict, today: date, cut: Optional[date], start: Optional[date]) -> List[dict]:
    hire = _d(person.get('hire_date'))
    if not hire:
        return []
    start = start or hire
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
        if d > today or (cut and d >= cut):
            break
        if d >= start:
            out.append({'name': name, 'days': annual_days(n), 'valid_from': d,
                        'valid_to': add_years(hire, n + 1) - timedelta(days=1),
                        'source': f'annual:{name}:{d.isoformat()}', 'note': f'입사 {n}년'})
        n += 1
    return out


ACCRUE_HOLD_OVER = 3   # 한 사람에게 한 번에 이보다 많이 생기면 넣지 않고 보류(소급 몰림 방지, 11-16)


def _missing_grants(person: dict, today: date) -> List[dict]:
    with _conn() as conn:
        have = {r[0] for r in conn.execute('SELECT source FROM leave_grants WHERE name=?', (person['name'],))}
    return [g for g in due_grants(person, today) if g['source'] not in have]


def accrue(today: Optional[date] = None, held: Optional[List[dict]] = None,
           allow_backfill: bool = False) -> List[dict]:
    """매일 돌린다 — 새로 생긴 부여만 넣고 그 목록을 돌려준다.

    한 사람에게 ACCRUE_HOLD_OVER 줄보다 많이 생기면(휴가 기능을 켜는 날 옮겨 넣은 시점·accrue_from 이 비어
    입사일부터 몰려 생기는 경우 등) 넣지 않고 held 에 담는다 — 관리자가 확인 뒤 allow_backfill=True 로 넣는다
    (2026-11-16 점검: 켜는 날 소급 부여 + 사람마다 DM 이 한꺼번에 나갈 수 있었다)."""
    today = today or date.today()
    made = []
    for p in staff():
        due = _missing_grants(p, today)
        if len(due) > ACCRUE_HOLD_OVER and not allow_backfill:
            if held is not None:
                held.append({'name': p['name'], 'count': len(due),
                             'days': sum(g['days'] for g in due), 'accrue_from': p.get('accrue_from') or ''})
            continue
        for g in due:
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


def days_in_year(ltype: str, start: date, end: date, year: int) -> float:
    """기간 중 그 해에 걸친 날만 센 일수 — 12/29~1/3 병가는 두 해에 나눠 잡는다."""
    s, e = max(start, date(year, 1, 1)), min(end, date(year, 12, 31))
    return count_days(ltype, s, e) if s <= e else 0.0


def used_in_year(name: str, ltype: str, year: int, sub: str = '', exclude_id: int = 0) -> float:
    """그 해 누적 사용(대기 중 포함) — 병가 30일·경조사 종류별 한도를 신청 여러 건 합으로 본다."""
    q = (f'SELECT ltype, start, "end" FROM leave_requests WHERE name=? AND ltype=? AND COALESCE(sub,"")=? '
         f'AND status IN ({",".join("?" * len(HOLDING))}) AND id<>? AND start<=? AND "end">=?')
    with _conn() as conn:
        rows = conn.execute(q, (name, ltype, sub or '', *HOLDING, exclude_id,
                                f'{year}-12-31', f'{year}-01-01')).fetchall()
    return sum(days_in_year(r['ltype'], _d(r['start']), _d(r['end']), year) for r in rows)


def sick_used(name: str, year: int, exclude_id: int = 0) -> float:
    return used_in_year(name, SICK, year, exclude_id=exclude_id)


def overlaps(name: str, start: date, end: date, exclude_id: int = 0) -> List[dict]:
    with _conn() as conn:
        return [dict(r) for r in conn.execute(
            f'SELECT * FROM leave_requests WHERE name=? AND status IN ({",".join("?" * len(HOLDING))}) '
            f'AND id<>? AND start<=? AND "end">=?',
            (name, *HOLDING, exclude_id, end.isoformat(), start.isoformat())).fetchall()]


def admin_type_options() -> List[Tuple[str, str]]:
    """경영지원이 등록하는 휴가(배우자 출산·보상·보건·산전후·난임·유산사산, 11-18) — 직원 신청 창엔 안 보임."""
    return [(k, v[0]) for k, v in TYPES.items() if v[3]]


def is_admin_type(ltype: str) -> bool:
    return bool(TYPES.get(ltype, ('', False, False, False))[3])


def manager_registered(row: dict) -> bool:
    """경영지원이 대신 넣은 줄인가 — 신청한 사람이 휴가 주인이 아님."""
    return bool(row.get('submitter_id')) and row.get('submitter_id') != row.get('slack_id')


PROMO_DESIGNATE_REASON = '연차 사용 시기 지정(근로기준법 제61조)'   # leave_promo.designate 가 넣는 사유


def manager_only(row: dict) -> bool:
    """본인이 취소·변경할 수 없는 줄 — 경영지원이 등록한 관리자 부여 휴가, 회사가 지정한 촉진 휴가(법적 지정 통보)."""
    if row.get('reason') == PROMO_DESIGNATE_REASON:
        return True
    return is_admin_type(row['ltype']) and manager_registered(row)


# ── 맞춤 휴가 부여 ─────────────────────────────────────────────────────
ANYTIME = '9999-12-31'   # 쓸 수 있는 기간 '언제든'
CUSTOM_MAX_DAYS = 120    # 산전후 90일(다태아 120일)


def _custom_alive(name: str, ltype: str, on: date) -> List[dict]:
    with _conn() as conn:
        return [dict(r) for r in conn.execute(
            'SELECT * FROM leave_custom_grants WHERE name=? AND ltype=? AND cancelled_at IS NULL '
            'AND valid_from<=? AND valid_to>=? ORDER BY valid_to, id', (name, ltype, on.isoformat(), on.isoformat()))]


def _custom_alloc(name: str, ltype: str, exclude_rid: int = 0, skip_gid: int = 0) -> Tuple[List[dict], Dict[int, float], float]:
    """맞춤 휴가 부여(취소 안 된 것)와 신청(승인 대기 포함)을 맞춰 본다 — 신청마다 그 첫날 살아 있는 부여 중
    소멸이 빠른 것부터 뗀다(연차 _allocate 와 같은 규칙). → (부여들, {부여 id: 뗀 일수}, 어느 부여로도 못 덮은 일수).
    skip_gid: 그 부여가 없다고 치고(부여 취소 검사)."""
    with _conn() as conn:
        every = [dict(r) for r in conn.execute(
            'SELECT * FROM leave_custom_grants WHERE name=? AND ltype=? AND cancelled_at IS NULL '
            'ORDER BY valid_to, id', (name, ltype))]
        gs = [g for g in every if g['id'] != skip_gid]
        reqs = [dict(r) for r in conn.execute(
            f'SELECT id, start, days FROM leave_requests WHERE name=? AND ltype=? AND id<>? '
            f'AND status IN ({",".join("?" * len(HOLDING))}) ORDER BY start, id',
            (name, ltype, exclude_rid, *HOLDING))]
    taken: Dict[int, float] = {g['id']: 0.0 for g in gs}
    short = 0.0
    for r in reqs:
        need = float(r['days'] or 0)
        for g in gs:
            if need <= 1e-9:
                break
            if g['valid_from'] <= r['start'] <= g['valid_to']:
                t = min(need, g['days'] - taken[g['id']])
                if t > 0:
                    taken[g['id']] += t
                    need -= t
        # 부여 기간 밖 신청(경영지원이 부여 없이 등록한 것 등)은 셈에 넣지 않는다
        if any(g['valid_from'] <= r['start'] <= g['valid_to'] for g in every):
            short += max(0.0, need)
    return gs, taken, round(short, 2)


def custom_left(name: str, ltype: str, on: Optional[date] = None, exclude_id: int = 0) -> float:
    """그날 쓸 수 있는 맞춤 휴가 일수 = 그날 살아 있는 부여마다 (일수 − 거기서 뗀 신청)."""
    if not name or not is_admin_type(ltype):
        return 0.0
    d = (on or date.today()).isoformat()
    gs, taken, _ = _custom_alloc(name, ltype, exclude_rid=exclude_id)
    return round(sum(g['days'] - taken[g['id']] for g in gs if g['valid_from'] <= d <= g['valid_to']), 2)


def _custom_covers(name: str, ltype: str, start: date, end: date) -> bool:
    """휴가 마지막 날까지 쓸 수 있는 기간 안인가."""
    return any(g['valid_to'] >= end.isoformat() for g in _custom_alive(name, ltype, start))


def custom_grants(name: str, include_cancelled: bool = True) -> List[dict]:
    """한 사람의 맞춤 휴가 부여 — 줄마다 쓴 일수(used)·남은 일수(left)."""
    with _conn() as conn:
        rows = [dict(r) for r in conn.execute(
            'SELECT * FROM leave_custom_grants WHERE name=?' + ('' if include_cancelled else ' AND cancelled_at IS NULL')
            + ' ORDER BY created_at DESC, id DESC', (name,))]
    allocs = {}
    for g in rows:
        if not g['cancelled_at'] and g['ltype'] not in allocs:
            allocs[g['ltype']] = _custom_alloc(name, g['ltype'])[1]
        g['label'] = type_label(g['ltype'])
        g['used'] = 0.0 if g['cancelled_at'] else round(allocs[g['ltype']].get(g['id'], 0.0), 2)
        g['left'] = 0.0 if g['cancelled_at'] else round(max(0.0, g['days'] - g['used']), 2)
        g['anytime'] = g['valid_to'] == ANYTIME
    return rows


def custom_balances(name: str, today: Optional[date] = None) -> List[dict]:
    """지금 쓸 수 있는 맞춤 휴가(종류마다 한 줄) — [{ltype, label, left, valid_to}]."""
    today = today or date.today()
    out = []
    for k, label in admin_type_options():
        alive = _custom_alive(name, k, today)
        if alive:
            out.append({'ltype': k, 'label': label, 'left': custom_left(name, k, today),
                        'valid_to': max(g['valid_to'] for g in alive)})
    return out


def type_options_for(name: str, today: Optional[date] = None) -> List[Tuple[str, str]]:
    """그 사람 신청 창 종류 — 보통 종류 + 남은 맞춤 휴가(이름에 남은 일수)."""
    extra = [(c['ltype'], f"{c['label']} (남은 {fmt_days(c['left'])})")
             for c in custom_balances(name, today) if c['left'] > 0] if name else []
    return type_options() + extra


def custom_grant(name: str, ltype: str, days, valid_from: Optional[date], valid_to: Optional[date], note: str,
                 uid: str, actor_name: str, today: Optional[date] = None) -> Tuple[Optional[dict], str]:
    """맞춤 휴가 부여 → (줄, '') 또는 (None, 오류). valid_to 가 없으면 언제든."""
    today = today or date.today()
    if uid not in manager_ids():
        return None, '경영지원만 부여할 수 있어요.'
    if not staff_by_name(name):
        return None, '휴가 명단에 없는 직원이에요.'
    if not is_admin_type(ltype):
        return None, '부여할 휴가 종류를 골라 주세요.'
    try:
        days = float(days)
    except (TypeError, ValueError):
        return None, '일수를 숫자로 적어 주세요.'
    if days <= 0 or abs(days * 2 - round(days * 2)) > 1e-9 or days > CUSTOM_MAX_DAYS:
        return None, f'일수는 0.5일 단위로, {CUSTOM_MAX_DAYS}일까지 적어 주세요.'
    valid_from = valid_from or today
    if valid_to and valid_to < valid_from:
        return None, '쓸 수 있는 마지막 날이 시작일보다 빨라요.'
    if valid_to and valid_to < today:
        return None, '이미 지난 기간이에요.'
    note = (note or '').strip()
    if len(note) > 200:
        return None, '메모는 200자까지 적을 수 있어요.'
    with _conn() as conn:
        cur = conn.execute(
            'INSERT INTO leave_custom_grants (name, ltype, days, valid_from, valid_to, note, created_by, created_by_name, '
            'created_at) VALUES (?,?,?,?,?,?,?,?,?)',
            (name, ltype, days, valid_from.isoformat(), (valid_to.isoformat() if valid_to else ANYTIME),
             note or None, uid, actor_name, _now()))
        g = dict(conn.execute('SELECT * FROM leave_custom_grants WHERE id=?', (cur.lastrowid,)).fetchone())
    g['label'] = type_label(ltype)
    return g, ''


def cancel_custom_grant(gid: int, uid: str, actor_name: str) -> Tuple[Optional[dict], str]:
    """맞춤 휴가 부여 취소 — 그 기간에 이미 신청한 것이 다른 부여로 다 덮이지 않으면 안 된다(먼저 그 휴가를 취소)."""
    if uid not in manager_ids():
        return None, '경영지원만 취소할 수 있어요.'
    with _conn() as conn:
        g = conn.execute('SELECT * FROM leave_custom_grants WHERE id=?', (int(gid),)).fetchone()
    if not g or g['cancelled_at']:
        return None, '취소할 수 있는 부여가 아니에요.'
    g = dict(g)
    # 이 부여가 없을 때 덮이지 않는 신청이 늘면 안 된다(다른 부여의 남은 몫까지 따져서)
    _, taken, short_now = _custom_alloc(g['name'], g['ltype'])
    _, _, short_after = _custom_alloc(g['name'], g['ltype'], skip_gid=g['id'])
    if short_after > short_now + 1e-9:
        used = round(short_after - short_now, 2)
        return None, f"이 부여로 {fmt_days(used)}을 이미 신청했어요. 그 휴가를 먼저 취소해 주세요."
    with _conn() as conn:
        cur = conn.execute('UPDATE leave_custom_grants SET cancelled_at=?, cancelled_by_name=? WHERE id=? '
                           'AND cancelled_at IS NULL', (_now(), actor_name, g['id']))
        if cur.rowcount != 1:
            return None, '그 사이 바뀌었어요. 다시 열어 주세요.'
    g['label'] = type_label(g['ltype'])
    return g, ''


def validate(name: str, ltype: str, sub: str, start: Optional[date], end: Optional[date],
             exclude_id: int = 0, allow_admin: bool = False) -> Dict[str, str]:
    """창 칸 이름 → 오류 문구.
    exclude_id: 고치는 신청 자신(또는 변경 신청이 대신할 휴가)은 겹침·남은 연차·한도에서 빼고 본다(11-18).
    allow_admin: 경영지원이 관리자 부여 휴가를 등록·수정할 때."""
    errs: Dict[str, str] = {}
    allowed = dict(type_options())
    if allow_admin:
        allowed.update(dict(admin_type_options()))
    elif is_admin_type(ltype) and custom_left(name, ltype, start or date.today(), exclude_id) > 0:
        allowed[ltype] = type_label(ltype)     # 경영지원이 일수를 먼저 준 맞춤 휴가(11-18)
    if ltype not in allowed:
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
    clash = [r for r in overlaps(name, start, end, exclude_id)
             if not (ltype in HALF_TYPES and r['ltype'] in HALF_TYPES and r['ltype'] != ltype)]
    if clash:
        c = clash[0]
        errs['start'] = f"이미 신청한 휴가와 겹쳐요: {type_label(c['ltype'], c.get('sub') or '')} {period_text(c, False)}"
        return errs
    if deducts(ltype):
        # 지난 날짜 신청은 오늘까지 살아 있는 연차에서만 — 휴가 첫날 기준으로만 보면 이미 소멸한 연차가
        #   되살아나 쓰였다(2026-11-16 상태 머신 점검 위험 13)
        last = _alloc_until(end)
        left = balance(name, start, last, exclude_id)
        if days > left + 1e-9:
            whole = balance(name, start, exclude_id=exclude_id)
            why = (f'{kdate(start)} 기준 남은 연차가 {fmt_days(whole)}이지만 그중 {fmt_days(round(whole - left, 2))}은 '
                   f'{kdate(last)} 전에 소멸해서 이 기간엔 {fmt_days(left)}만 쓸 수 있어요. '
                   if whole > left + 1e-9 else f'{kdate(start)} 기준 남은 연차가 {fmt_days(left)}이라 ')
            errs['start'] = why + f'{fmt_days(days)}를 쓸 수 없어요. 연차가 더 있어야 한다면 경영지원에 문의해 주세요.'
    elif is_admin_type(ltype) and not allow_admin:
        left = custom_left(name, ltype, start, exclude_id)
        if days > left + 1e-9:
            errs['start'] = (f'{type_label(ltype)}는 {kdate(start)} 기준 {fmt_days(left)}만 쓸 수 있어요. '
                             '더 필요하면 경영지원에 문의해 주세요.')
        elif not _custom_covers(name, ltype, start, end):
            errs['end'] = f'{type_label(ltype)}를 쓸 수 있는 기간을 넘어요. 기간 안 날짜로 골라 주세요.'
    elif ltype in (EVENT, SICK):
        # 한도는 신청 건마다가 아니라 그 해 누적(11-05 KiKO) — 나눠 신청해도, 해를 넘겨도 해마다 합으로
        if ltype == EVENT and sub not in EVENT_SUBS:
            return errs
        label, cap = (EVENT_SUBS[sub] if ltype == EVENT else ('병가', SICK_DAYS_PER_YEAR))
        for y in range(start.year, end.year + 1):
            here = days_in_year(ltype, start, end, y)
            used = used_in_year(name, ltype, y, sub if ltype == EVENT else '', exclude_id=exclude_id)
            if here and used + here > cap:
                errs['end'] = (f'{label}은 한 해 {cap}일까지예요 ({y}년에 이미 신청한 {fmt_days(used)}'
                               f' + 이번 {fmt_days(here)}). 더 필요하면 경영지원에 문의해 주세요.')
                break
    return errs


def _initial_status(who: str) -> Tuple[str, List[Tuple[str, str]]]:
    """처음(또는 고친 뒤 다시) 승인받을 상태 — 첫 승인자 본인이면 2차부터, 마지막 승인자 본인이면 바로 확정."""
    appr = approver_ids()
    if who and appr and who == appr[-1]:
        return APPROVED, ([('approve_1', '요청자가 승인자'), ('approve_2', '요청자가 승인자')] if len(appr) > 1
                          else [('approve_1', '요청자가 승인자')])
    if who and len(appr) > 1 and who == appr[0]:
        return WAIT_2, [('approve_1', '요청자가 승인자')]
    return WAIT_1, []


def create(name: str, subject_slack_id: str, submitter_id: str, submitter_name: str, ltype: str, sub: str,
           start: date, end: Optional[date], reason: str = '', replaces: int = 0) -> Optional[dict]:
    """신청 만들기. 첫 승인자가 신청하면 다음 승인자부터, 마지막 승인자가 신청하면 바로 확정.
    연차가 모자라면(그 사이 다른 신청이 잡아감) None.
    replaces: 승인된 휴가를 바꾸는 변경 신청(11-18) — 원래 휴가는 이 신청이 승인될 때 취소된다. 그때까지 둘 다
    연차를 잡고 있되, 이 신청의 몫은 원래 휴가가 잡은 몫을 쓸 수 있는 셈으로 뗀다(잠깐 남은 연차가 덜 보일 수 있음)."""
    end = start if ltype in HALF_TYPES else (end or start)
    days = count_days(ltype, start, end)
    who = subject_slack_id or ''
    status, auto = _initial_status(who)
    now = _now()
    with _conn() as conn:
        cur = conn.execute(
            'INSERT INTO leave_requests (name, slack_id, submitter_id, submitter_name, ltype, sub, start, "end", '
            'days, reason, status, created_at, updated_at, replaces) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
            (name, subject_slack_id or None, submitter_id, submitter_name, ltype, sub or None, start.isoformat(),
             end.isoformat(), days, reason or None, status, now, now, replaces or None))
        rid = cur.lastrowid
        if deducts(ltype):
            held = []
            if replaces:   # 원래 휴가 몫을 잠깐 빼고 뗀 뒤 그대로 되돌린다
                held = [tuple(r) for r in conn.execute(
                    'SELECT request_id, grant_id, days FROM leave_uses WHERE request_id=?', (replaces,))]
                conn.execute('DELETE FROM leave_uses WHERE request_id=?', (replaces,))
            ok = _allocate(conn, rid, name, start, days, _alloc_until(end))
            for h in held:
                conn.execute('INSERT INTO leave_uses (request_id, grant_id, days) VALUES (?,?,?)', h)
            if not ok:
                conn.rollback()
                return None
        conn.execute('INSERT INTO leave_steps (request_id, step, actor_id, actor_name, result, at) '
                     'VALUES (?,?,?,?,?,?)', (rid, 'request', submitter_id, submitter_name, 'requested', now))
        for step, why in auto:
            conn.execute('INSERT INTO leave_steps (request_id, step, actor_id, actor_name, result, reason, at) '
                         'VALUES (?,?,?,?,?,?,?)', (rid, step, who, name, 'auto', why, now))
        if status == APPROVED and replaces:
            _replace_original(conn, replaces, rid, who, name, now)   # 마지막 승인자 본인의 변경 신청 — 바로 바뀜
        conn.commit()
        return _row(conn, rid)


def _replace_original(conn, orig: int, rid: int, uid: str, actor_name: str, now: str) -> None:
    """변경 신청 rid 가 확정됐을 때 원래 휴가 orig 취소(같은 연결·커밋은 부르는 쪽)."""
    if conn.execute(f'UPDATE leave_requests SET status=?, updated_at=? WHERE id=? AND status IN ({",".join("?" * len(HOLDING))})',
                    (CANCELLED, now, orig, *HOLDING)).rowcount == 1:
        conn.execute('DELETE FROM leave_uses WHERE request_id=?', (orig,))
        conn.execute('INSERT INTO leave_steps (request_id, step, actor_id, actor_name, result, reason, at) '
                     'VALUES (?,?,?,?,?,?,?)',
                     (orig, 'cancel', uid, actor_name, 'replaced', f'변경 신청 #{rid} 승인으로 바뀜', now))


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
        if manager_only(row):
            # 경영지원이 등록한 휴가(배우자 출산 등)는 경영지원만 취소(11-18). 맞춤 휴가를 본인이 신청한 것은 보통 규칙
            return uid in manager_ids() and row['status'] in HOLDING
        mine = uid in (row.get('slack_id'), row.get('submitter_id'))
        if mine and row['status'] == APPROVED and _d(row['end']) >= today:
            # 본인: 끝나기 전까지(쓰는 중이면 오늘부터 남은 날만 취소 — cancel_mode). 끝난 휴가는 경영지원
            return True
        if uid in manager_ids():
            return row['status'] in HOLDING
        return mine and row['status'] in OPEN
    return False


def cancel_mode(row: dict, uid: str, today: Optional[date] = None) -> str:
    """'rest' = 쓰는 중인 휴가를 본인이 취소 → 어제까지만 남기고 오늘부터 취소. 나머지는 'all'."""
    today = today or date.today()
    mine = uid in (row.get('slack_id'), row.get('submitter_id'))
    if mine and row['status'] == APPROVED and _d(row['start']) < today <= _d(row['end']) \
            and count_days(row['ltype'], _d(row['start']), today - timedelta(days=1)) > 0:
        return 'rest'
    return 'all'


def cut_rest(rid: int, uid: str, actor_name: str, today: Optional[date] = None) -> Optional[dict]:
    """쓰는 중인 휴가의 오늘부터 남은 날 취소 — 기간을 어제까지로 줄이고 연차를 다시 뗀다(남은 날 몫은 돌아옴)."""
    today = today or date.today()
    now = _now()
    with _conn() as conn:
        row = _row(conn, rid)
        if not row or cancel_mode(row, uid, today) != 'rest':
            return None
        new_end = today - timedelta(days=1)
        start = _d(row['start'])
        days = count_days(row['ltype'], start, new_end)
        cur = conn.execute('UPDATE leave_requests SET "end"=?, days=?, updated_at=? WHERE id=? AND status=? '
                           'AND "end"=?', (new_end.isoformat(), days, now, rid, APPROVED, row['end']))
        if cur.rowcount != 1:
            conn.rollback()
            return None
        conn.execute('DELETE FROM leave_uses WHERE request_id=?', (rid,))
        if deducts(row['ltype']) and not _allocate(conn, rid, row['name'], start, days, new_end):
            conn.rollback()
            return None
        conn.execute('INSERT INTO leave_steps (request_id, step, actor_id, actor_name, result, reason, at) '
                     'VALUES (?,?,?,?,?,?,?)',
                     (rid, 'cancel', uid, actor_name, 'shortened',
                      f"{kdate(today)}부터 취소 (원래 {kdate(_d(row['end']))}까지)", now))
        conn.commit()
        return _row(conn, rid)


def transition(rid: int, action: str, uid: str, actor_name: str, reason: str = '',
               today: Optional[date] = None) -> Optional[dict]:
    """상태 바꾸기(두 번 눌러도 한 번만). 반려·취소면 잡아 둔 연차를 돌려놓는다.

    권한은 여기서 '지금 상태' 기준으로 다시 본다(2026-11-13 보안 점검 2차 #11): 1차 승인자가 승인을 빠르게
    두 번 누르면 두 번째 클릭이 화면을 띄울 때의 상태(1차 대기)로 권한을 통과한 뒤, 이미 2차 대기로 바뀐 행에서
    2차 승인까지 처리할 수 있었다. 같은 연결에서 다시 확인 + UPDATE 의 상태 조건으로 한 번만 바뀐다."""
    now = _now()
    with _conn() as conn:
        row = _row(conn, rid)
        if not row:
            return None
        if not can_act(action, row, uid, today):
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
        if nxt == APPROVED and row.get('replaces'):
            # 변경 신청이 확정되면 원래 휴가는 취소(같은 연결 — 둘 다 남거나 둘 다 잡히는 틈 없게, 11-18)
            _replace_original(conn, int(row['replaces']), rid, uid, actor_name, now)
        conn.commit()
        return _row(conn, rid)


# ── 고치기·변경 신청·연차 조정·관리자 등록 (2026-11-18, 플렉스 관리자 기능 맞추기) ─────────────
def _what(row: dict) -> str:
    return f"{type_label(row['ltype'], row.get('sub') or '')} {period_text(row, False)} ({fmt_days(row['days'])})"


def pending_change(rid: int) -> Optional[dict]:
    """승인된 휴가 rid 를 바꾸려는, 아직 승인 대기 중인 변경 신청."""
    with _conn() as conn:
        r = conn.execute(f'SELECT * FROM leave_requests WHERE replaces=? AND status IN (?,?) ORDER BY id DESC',
                         (rid, *OPEN)).fetchone()
        return dict(r) if r else None


def own_edit_mode(row: dict, uid: str = '', today: Optional[date] = None) -> str:
    """본인 고치기 단추 — 'edit'(승인 대기: 바로 고치고 승인 다시) · 'change'(승인 뒤 끝나기 전: 변경 신청) · ''."""
    today = today or date.today()
    who = uid or row.get('slack_id') or row.get('submitter_id') or ''
    if not who or who not in (row.get('slack_id'), row.get('submitter_id')):
        return ''
    if row['status'] in OPEN:
        return 'edit'
    if row['status'] == APPROVED and _d(row['start']) > today and not pending_change(row['id']) \
            and not manager_only(row):
        return 'change'      # 시작 전 휴가만(쓰는 중이면 [남은 날 취소] 뒤 다시 신청). 관리자 등록 휴가는 경영지원이
    return ''


def edit_request(rid: int, ltype: str, sub: str, start: Optional[date], end: Optional[date], reason: str,
                 uid: str, actor_name: str, as_admin: bool = False,
                 today: Optional[date] = None) -> Tuple[Optional[dict], Dict[str, str]]:
    """신청 고치기 → (고친 줄, {}) 또는 (None, {칸: 오류}).
    - 본인: 승인 대기 중인 것만. 고치면 승인을 처음부터 다시 받는다(바뀐 내용으로 다시 보게).
    - 경영지원(as_admin): 대기·승인 어느 상태든 그대로 두고 바로 고친다(기록 남김 — 알림은 슬랙 쪽)."""
    row = get(rid)
    if not row:
        return None, {'_': '없는 신청이에요.'}
    if as_admin:
        if uid not in manager_ids():
            return None, {'_': '경영지원만 고칠 수 있어요.'}
        if row['status'] not in HOLDING:
            return None, {'_': '취소·반려된 휴가는 고칠 수 없어요.'}
    else:
        if uid not in (row.get('slack_id'), row.get('submitter_id')):
            return None, {'_': '본인 휴가만 고칠 수 있어요.'}
        if row['status'] not in OPEN:
            return None, {'_': '승인된 휴가는 [변경 신청]으로 바꿔 주세요.'}
    end = start if ltype in HALF_TYPES else (end or start)
    errs = validate(row['name'], ltype, sub, start, end, exclude_id=rid, allow_admin=as_admin)
    if errs:
        return None, (_admin_errs(errs) if as_admin else errs)
    sub = sub if ltype == EVENT else ''
    days = count_days(ltype, start, end)
    before = _what(row)
    now = _now()
    if as_admin:
        status, auto = row['status'], []
    else:
        status, auto = _initial_status(row.get('slack_id') or '')
    with _conn() as conn:
        cur = conn.execute('UPDATE leave_requests SET ltype=?, sub=?, start=?, "end"=?, days=?, reason=?, status=?, '
                           'updated_at=? WHERE id=? AND status=? AND updated_at=?',
                           (ltype, sub or None, start.isoformat(), end.isoformat(), days, reason or row.get('reason'),
                            status, now, rid, row['status'], row['updated_at']))
        if cur.rowcount != 1:
            conn.rollback()
            return None, {'_': '그 사이 신청이 바뀌었어요. 다시 열어 주세요.'}
        conn.execute('DELETE FROM leave_uses WHERE request_id=?', (rid,))
        if deducts(ltype) and not _allocate(conn, rid, row['name'], start, days, _alloc_until(end, today)):
            conn.rollback()
            return None, {'start': '남은 연차가 모자라요. 다시 확인해 주세요.'}
        conn.execute('INSERT INTO leave_steps (request_id, step, actor_id, actor_name, result, reason, at) '
                     'VALUES (?,?,?,?,?,?,?)',
                     (rid, 'admin_edit' if as_admin else 'edit', uid, actor_name, 'edited', f'고치기 전: {before}', now))
        for step, why in auto:
            conn.execute('INSERT INTO leave_steps (request_id, step, actor_id, actor_name, result, reason, at) '
                         'VALUES (?,?,?,?,?,?,?)', (rid, step, uid, actor_name, 'auto', why, now))
        conn.commit()
        return _row(conn, rid), {}


def change_request(rid: int, ltype: str, sub: str, start: Optional[date], end: Optional[date], reason: str,
                   uid: str, actor_name: str, today: Optional[date] = None) -> Tuple[Optional[dict], Dict[str, str]]:
    """승인된(시작 전) 휴가의 변경 신청 → (새 신청, {}). 승인되면 원래 휴가가 취소되고, 반려·취소되면 원래 그대로."""
    row = get(rid)
    if not row:
        return None, {'_': '없는 신청이에요.'}
    if own_edit_mode(row, uid, today) != 'change':
        if pending_change(rid):
            return None, {'_': '이미 변경 신청이 승인을 기다리고 있어요.'}
        return None, {'_': '시작 전인 본인 휴가만 변경 신청할 수 있어요.'}
    end = start if ltype in HALF_TYPES else (end or start)
    errs = validate(row['name'], ltype, sub, start, end, exclude_id=rid)
    if errs:
        return None, errs
    new = create(row['name'], row.get('slack_id') or uid, uid, actor_name, ltype, sub if ltype == EVENT else '',
                 start, end, reason or row.get('reason') or '', replaces=rid)
    if not new:
        return None, {'start': '그 사이 남은 연차가 바뀌었어요. 다시 확인해 주세요.'}
    return new, {}


def _period_end(name: str, today: date) -> date:
    """조정 연차를 쓸 수 있는 마지막 날 — 지금 살아 있는 부여 중 가장 늦게 끝나는 날, 없으면 다음 입사 기념일 전날."""
    live = [g['valid_to'] for g in grants(name) if g['valid_from'] <= today.isoformat() <= g['valid_to']]
    if live:
        return _d(max(live))
    s = staff_by_name(name) or {}
    hire = _d(s.get('hire_date')) or today
    n = 1
    while add_years(hire, n) <= today:
        n += 1
    return add_years(hire, n) - timedelta(days=1)


ADJUST_WINDOW_DAYS = 366   # 적용일은 오늘 앞뒤 1년 안(잘못 친 연도 막기)


def _check_adjust(name: str, days, reason: str, applies_on: date, today: date) -> Tuple[float, str]:
    """연차 조정 한 줄 검사 → (일수, '') 또는 (0, 오류). 저장은 안 한다(일괄 조정이 모두 맞을 때만 넣게)."""
    if not staff_by_name(name):
        return 0.0, '휴가 명단에 없는 직원이에요.'
    try:
        days = float(days)
    except (TypeError, ValueError):
        return 0.0, '일수를 숫자로 적어 주세요.'
    if not days or abs(days * 2 - round(days * 2)) > 1e-9 or abs(days) > 30:
        return 0.0, '일수는 0.5일 단위로, 30일까지 적어 주세요(빼려면 −).'
    reason = (reason or '').strip()
    if not reason:
        return 0.0, '사유를 적어 주세요.'
    if len(reason) > 200:
        return 0.0, '사유는 200자까지 적을 수 있어요.'
    if abs((applies_on - today).days) > ADJUST_WINDOW_DAYS:
        return 0.0, '적용일은 오늘 앞뒤 1년 안으로 골라 주세요.'
    left = balance(name, applies_on)
    if days < 0 and left + days < -1e-9:
        return 0.0, f'{kdate(applies_on)} 기준 남은 연차({fmt_days(left)})보다 많이 뺄 수 없어요.'
    return days, ''


def _add_adjust(name: str, days: float, reason: str, actor_name: str, applies_on: date) -> Optional[dict]:
    until = _period_end(name, applies_on)
    src = f"adjust:{name}:{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
    sign = '+' if days > 0 else '−'
    note = f'{sign}{fmt_days(abs(days))} 조정: {reason.strip()} ({actor_name})'
    if not add_grant(name, days, applies_on, until, src, note):
        return None
    with _conn() as conn:
        return dict(conn.execute('SELECT * FROM leave_grants WHERE source=?', (src,)).fetchone())


def adjust(name: str, days: float, reason: str, uid: str, actor_name: str,
           today: Optional[date] = None, applies_on: Optional[date] = None) -> Tuple[Optional[dict], str]:
    """경영지원 연차 조정(+/−) → (부여 줄, '') 또는 (None, 오류).
    적용일(applies_on, 플렉스 '적용일')부터 그날 살아 있는 연차와 같은 날 소멸. 비우면 오늘.
    날짜가 있는 휴가(병원·반차 등)는 조정이 아니라 [휴가 등록]으로 — 그래야 휴가 기록·일정에 남는다."""
    today = today or date.today()
    applies_on = applies_on or today
    if uid not in manager_ids():
        return None, '경영지원만 조정할 수 있어요.'
    days, err = _check_adjust(name, days, reason, applies_on, today)
    if err:
        return None, err
    g = _add_adjust(name, days, reason, actor_name, applies_on)
    return (g, '') if g else (None, '조정을 저장하지 못했어요. 다시 해 주세요.')


def adjust_bulk(items: List[dict], reason: str, uid: str, actor_name: str, today: Optional[date] = None,
                applies_on: Optional[date] = None) -> Tuple[List[dict], Dict[str, str]]:
    """전 직원 연차 부여·차감 표(플렉스 같은 화면) — items=[{name, days, reason?}] 를 모두 검사해 하나라도 틀리면
    아무것도 넣지 않고 {이름: 오류}. 모두 맞으면 넣은 부여 줄 목록. 사람별 사유가 비면 공통 사유."""
    today = today or date.today()
    applies_on = applies_on or today
    if uid not in manager_ids():
        return [], {'_': '경영지원만 조정할 수 있어요.'}
    todo, errs, seen = [], {}, set()
    for it in items or []:
        name = str(it.get('name') or '').strip()
        if not name or str(it.get('days') or '').strip() in ('', '0'):
            continue
        if name in seen:
            errs[name] = '같은 직원이 두 번 들어 있어요.'
            continue
        seen.add(name)
        why = str(it.get('reason') or '').strip() or (reason or '')
        days, err = _check_adjust(name, it.get('days'), why, applies_on, today)
        if err:
            errs[name] = err
        else:
            todo.append((name, days, why))
    if errs:
        return [], errs
    if not todo:
        return [], {'_': '조정할 일수를 한 명 이상 적어 주세요.'}
    made = []
    stamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
    with _conn() as conn:
        try:
            for i, (name, days, why) in enumerate(todo):
                until = _period_end(name, applies_on)
                src = f'adjust:{name}:{stamp}{i:03d}'
                sign = '+' if days > 0 else '−'
                conn.execute('INSERT INTO leave_grants (name, days, valid_from, valid_to, source, note, created_at) '
                             'VALUES (?,?,?,?,?,?,?)',
                             (name, float(days), applies_on.isoformat(), until.isoformat(), src,
                              f'{sign}{fmt_days(abs(days))} 조정: {why.strip()} ({actor_name})', _now()))
                made.append(dict(conn.execute('SELECT * FROM leave_grants WHERE source=?', (src,)).fetchone()))
        except Exception:
            conn.rollback()
            raise
        conn.commit()
    return made, {}


def cancel_adjust(gid: int, uid: str, actor_name: str) -> Tuple[Optional[dict], str]:
    """연차 조정 취소(플렉스 [조정 취소하기]) — 줄은 지우지 않고 0일로 바꾸고 '취소' 기록을 남긴다.
    늘린 조정에서 이미 떼어 쓴 휴가가 있으면 취소하지 않는다(그 휴가가 쓸 연차가 사라지므로 빼기 조정으로 맞춤)."""
    if uid not in manager_ids():
        return None, '경영지원만 취소할 수 있어요.'
    with _conn() as conn:
        g = conn.execute('SELECT g.*, COALESCE((SELECT SUM(u.days) FROM leave_uses u WHERE u.grant_id=g.id), 0) AS used '
                         'FROM leave_grants g WHERE g.id=?', (int(gid),)).fetchone()
        if not g or not str(g['source'] or '').startswith('adjust:'):
            return None, '취소할 수 있는 연차 조정이 아니에요.'
        if g['used'] > 1e-9:
            return None, (f"이 조정으로 늘린 연차 중 {fmt_days(g['used'])}을 이미 휴가에 썼어요. "
                          '취소하지 말고 빼기 조정으로 맞춰 주세요.')
        g = dict(g)
    if g['days'] > 0:
        on = max(date.today(), _d(g['valid_from']))
        if on <= _d(g['valid_to']) and balance(g['name'], on) - g['days'] < -1e-9:
            return None, '취소하면 남은 연차가 0보다 적어져요. 빼기 조정을 먼저 취소해 주세요.'
    with _conn() as conn:
        note = f"{g.get('note') or ''} → 취소({actor_name} {datetime.now().strftime('%Y-%m-%d %H:%M')})"
        cur = conn.execute("UPDATE leave_grants SET days=0, source='cancelled:'||source, note=? "
                           "WHERE id=? AND source LIKE 'adjust:%'", (note, g['id']))
        if cur.rowcount != 1:
            conn.rollback()
            return None, '그 사이 바뀌었어요. 다시 열어 주세요.'
        conn.commit()
        return dict(conn.execute('SELECT * FROM leave_grants WHERE id=?', (g['id'],)).fetchone(), orig_days=g['days']), ''


def adjustments(name: str) -> List[dict]:
    """연차 조정 기록 — 취소한 것도(cancelled=True, days=0, 원래 일수는 메모에)."""
    with _conn() as conn:
        rows = [dict(r) for r in conn.execute(
            "SELECT * FROM leave_grants WHERE name=? AND (source LIKE 'adjust:%' OR source LIKE 'cancelled:adjust:%') "
            'ORDER BY created_at DESC, id DESC', (name,)).fetchall()]
    for r in rows:
        r['cancelled'] = str(r.get('source') or '').startswith('cancelled:')
    return rows


def monthly_detail(name: str, year: int, today: Optional[date] = None) -> dict:
    """직원별 월별 연차 상세(플렉스 '연차 상세') — 달마다 부여·조정·사용·소멸·잔여.

    잔여 = 다음 달 1일에 살아 있는 연차 − 그 전에 시작한 휴가로 뗀 것. 그래서 늘 '앞 달 잔여 + 부여 + 조정 − 사용 − 소멸'.
    사용은 휴가 첫날이 든 달(승인 대기 포함 — 신청하면 바로 빠지는 플렉스와 같음). 소멸은 쓸 수 있는 마지막 날이 든 달.
    오늘 뒤 달은 지금 잡힌 휴가만 반영한 예상(projected)."""
    today = today or date.today()
    with _conn() as conn:
        gs = [dict(r) for r in conn.execute('SELECT * FROM leave_grants WHERE name=? AND days<>0', (name,)).fetchall()]
        uses = [dict(r) for r in conn.execute(
            f'SELECT u.grant_id, u.days, r.start FROM leave_uses u JOIN leave_requests r ON r.id=u.request_id '
            f'WHERE r.name=? AND r.status IN ({",".join("?" * len(HOLDING))})', (name, *HOLDING)).fetchall()]
    if not gs:
        return {'year': year, 'months': [], 'opening': 0.0}
    by_g: Dict[int, list] = {}
    for u in uses:
        by_g.setdefault(u['grant_id'], []).append(u)

    def used(gid, before: str = '9999-12-31') -> float:
        return sum(u['days'] for u in by_g.get(gid, []) if u['start'] < before)

    def left_on(d: str) -> float:
        """d 날 0시 — d 전에 생겨 d 에도 살아 있는 부여(그날 생기는 것은 그달 '부여'로 센다)."""
        return round(sum(g['days'] - used(g['id'], d) for g in gs if g['valid_from'] < d <= g['valid_to']), 2)

    first_month = min(_d(g['valid_from']) for g in gs)
    months = []
    for m in range(1, 13):
        a = date(year, m, 1)
        b = date(year + (m == 12), m % 12 + 1, 1)
        if b <= date(first_month.year, first_month.month, 1):
            continue
        A, B = a.isoformat(), b.isoformat()
        is_adj = lambda g: str(g.get('source') or '').startswith('adjust:')   # noqa: E731
        months.append({
            'month': m,
            'grant': round(sum(g['days'] for g in gs if A <= g['valid_from'] < B and not is_adj(g)), 2),
            'adjust': round(sum(g['days'] for g in gs if A <= g['valid_from'] < B and is_adj(g)), 2),
            'used': round(sum(u['days'] for u in uses if A <= u['start'] < B), 2),
            'expired': round(sum(g['days'] - used(g['id']) for g in gs if A <= g['valid_to'] < B), 2),
            'left': left_on(B),
            'projected': a > today,
        })
    opening = left_on(date(year, months[0]['month'], 1).isoformat()) if months else 0.0
    return {'year': year, 'months': months, 'opening': opening}


_SETTLE_SOURCES = ('fiscal:', 'monthly:', 'annual:')


def resign_settlement(name: str, last_day: date) -> dict:
    """퇴직 정산(11-18) — 회계연도 기준으로 실제 준 연차와 입사일 기준이었다면 받았을 연차를 2027-01-01 뒤 몫끼리 비교.
    2026년까지는 두 기준이 같다(플렉스·장부 모두 입사일 기준). 모자라면 그 일수를 수당으로(더 준 몫은 깎지 않음).
    미사용 연차 = 마지막 근무일에 살아 있는 연차 − 그날까지 시작한 휴가로 쓴 몫(수당 대상)."""
    st = staff_by_name(name)
    if not st:
        return {'ok': False, 'why': '휴가 명단에 없는 직원이에요.'}
    hire = _d(st.get('hire_date'))
    if not hire or last_day < hire:
        return {'ok': False, 'why': '마지막 근무일이 입사일보다 빨라요.'}
    lo, hi = FISCAL_FROM.isoformat(), last_day.isoformat()
    basis = [g for g in hire_basis_grants(st, last_day) if g['valid_from'] >= FISCAL_FROM]
    actual = [g for g in grants(name) if str(g['source'] or '').startswith(_SETTLE_SOURCES)
              and lo <= g['valid_from'] <= hi]
    b_total = round(sum(g['days'] for g in basis), 2)
    a_total = round(sum(g['days'] for g in actual), 2)
    nxt = (last_day + timedelta(days=1)).isoformat()
    with _conn() as conn:
        used_before = {r[0]: r[1] for r in conn.execute(
            f'SELECT u.grant_id, SUM(u.days) FROM leave_uses u JOIN leave_requests r ON r.id=u.request_id '
            f'WHERE r.name=? AND r.status IN ({",".join("?" * len(HOLDING))}) AND r.start<? GROUP BY u.grant_id',
            (name, *HOLDING, nxt))}
        later = [dict(r) for r in conn.execute(
            f'SELECT * FROM leave_requests WHERE name=? AND status IN ({",".join("?" * len(HOLDING))}) AND start>? '
            'ORDER BY start', (name, *HOLDING, hi))]
    alive = [g for g in grants(name) if g['valid_from'] <= hi <= g['valid_to']]
    unused = round(sum(g['days'] - used_before.get(g['id'], 0.0) for g in alive), 2)

    def rows(gs):
        return [{'from': (g['valid_from'].isoformat() if isinstance(g['valid_from'], date) else g['valid_from']),
                 'days': g['days'], 'note': g.get('note') or ''} for g in gs]
    return {
        'ok': True, 'name': name, 'hire_date': hire.isoformat(), 'last_day': hi,
        'basis': rows(basis), 'basis_total': b_total, 'actual': rows(actual), 'actual_total': a_total,
        'short': round(max(0.0, b_total - a_total), 2), 'unused': max(0.0, unused),
        'later': [{'id': r['id'], 'what': f"{type_label(r['ltype'], r.get('sub') or '')} {period_text(r, False)}",
                   'days': r['days']} for r in later],
        'before_fiscal': last_day < FISCAL_FROM,
    }


_ASK_MANAGER = ' 연차가 더 있어야 한다면 경영지원에 문의해 주세요.'


def _admin_errs(errs: Dict[str, str]) -> Dict[str, str]:
    """경영지원이 보는 오류 — '경영지원에 문의' 대신 할 일([연차 조정])로."""
    return {k: (v.replace(_ASK_MANAGER, ' 먼저 [연차 조정]으로 늘린 뒤 다시 해 주세요.')
                .replace(' 더 필요하면 경영지원에 문의해 주세요.', '')) for k, v in errs.items()}


def admin_register(name: str, ltype: str, sub: str, start: Optional[date], end: Optional[date], reason: str,
                   uid: str, actor_name: str) -> Tuple[Optional[dict], Dict[str, str]]:
    """경영지원 휴가 관리 — 휴가 등록 → 승인 없이 바로 확정(줄, {}).
    - 관리자 부여 휴가(배우자 출산·보상·보건 등, 11-18 첫 판)
    - 연차·반차·병가·경조사·군소집도(11-18 부장님: 본인이 급히 자리를 비운 사이 경영지원이 반영) — 연차·반차는
      그 자리에서 연차를 뗀다. 지난 날짜도 된다(이미 쉰 날을 나중에 반영). 모자라면 [연차 조정]으로 먼저 늘린다."""
    if uid not in manager_ids():
        return None, {'_': '경영지원만 등록할 수 있어요.'}
    st = staff_by_name(name)
    if not st:
        return None, {'_': '휴가 명단에 없는 직원이에요.'}
    if ltype not in dict(type_options() + admin_type_options()):
        return None, {'type': '등록할 휴가 종류를 골라 주세요.'}
    end = start if ltype in HALF_TYPES else (end or start)
    sub = sub if ltype == EVENT else ''
    errs = validate(name, ltype, sub, start, end, allow_admin=True)
    if errs:
        return None, _admin_errs(errs)
    days = count_days(ltype, start, end)
    now = _now()
    with _conn() as conn:
        cur = conn.execute(
            'INSERT INTO leave_requests (name, slack_id, submitter_id, submitter_name, ltype, sub, start, "end", '
            'days, reason, status, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)',
            (name, st.get('slack_id') or None, uid, actor_name, ltype, sub or None, start.isoformat(), end.isoformat(),
             days, (reason or '').strip() or None, APPROVED, now, now))
        rid = cur.lastrowid
        if deducts(ltype) and not _allocate(conn, rid, name, start, days, _alloc_until(end)):
            conn.rollback()
            return None, {'start': '그 사이 남은 연차가 바뀌었어요. 다시 확인해 주세요.'}
        for step, result, why in (('request', 'registered', '경영지원 등록'), ('approve_1', 'auto', '경영지원 등록'),
                                  ('approve_2', 'auto', '경영지원 등록')):
            conn.execute('INSERT INTO leave_steps (request_id, step, actor_id, actor_name, result, reason, at) '
                         'VALUES (?,?,?,?,?,?,?)', (rid, step, uid, actor_name, result, why, now))
        conn.commit()
        return _row(conn, rid), {}


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
    """#휴가_일정 — 누가 언제 휴가/반차인지만(종류·사유 없음).
    11-18 부장님: 플렉스 문구 '하루종일 휴가입니다'가 어색 → '이름 · 날짜(요일) 휴가/오전 반차'로 짧게."""
    return f"🌴 {row['name']} · {_announce_what(row)}"


def _announce_what(row: dict) -> str:
    s, e = _d(row['start']), _d(row['end'])
    if row['ltype'] in HALF_TYPES:
        return f"{kdate(s, True)} {'오전' if row['ltype'] == HALF_AM else '오후'} 반차"
    if s == e:
        return f'{kdate(s, True)} 휴가'
    return f'{kdate(s, True)} ~ {kdate(e, True)} 휴가'


def cancelled_announce_text(row: dict) -> str:
    """처음 올린 글을 고칠 때 — 취소선 + (취소). 달력처럼 채널을 훑는 사람이 지난 글만 봐도 알게."""
    return f'~{announce_text(row)}~ (취소)'


def cancel_notice_text(row: dict) -> str:
    """취소 새 글(11-18 부장님 — 플렉스처럼 채널에 새로 알림. 글을 고치기만 하면 알림이 안 가서 못 봄)."""
    return f"❗ {row['name']} · {_announce_what(row)} 취소"


def cut_notice_text(row: dict, today: Optional[date] = None) -> str:
    """쓰는 중인 휴가를 오늘부터 취소했을 때 새 글."""
    today = today or date.today()
    return f"❗ {row['name']} · {kdate(today, True)}부터 휴가 취소 (어제까지만 휴가)"


def step_line(row: dict, step_rows: List[dict], names: Tuple[str, str] = ('박정우', '박용구')) -> str:
    done: Dict[str, dict] = {}
    for s in step_rows:
        if s['step'] == 'edit':      # 본인이 고치면 승인을 처음부터 다시(11-18) — 그 전 승인은 지운 셈
            done = {}
            continue
        done[s['step']] = s
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
    title = '휴가 변경 신청' if row.get('replaces') else '휴가 신청'
    lines = [f"🌴 *{title} — {row['name']}*  {emoji} {label}",
             f"• 종류: {type_label(row['ltype'], row.get('sub') or '')}",
             f"• 날짜: {period_text(row)}  ({fmt_days(row['days'])})"]
    if row.get('replaces'):
        orig = get(int(row['replaces']))
        if orig:
            lines.append(f"• 바꾸기 전: {_what(orig)} — 이 신청이 승인되면 원래 휴가는 취소돼요")
    if row.get('reason'):
        lines.append(f"• 사유: {row['reason']}")
    if deducts(row['ltype']) and left is not None:
        lines.append(f'• 남은 연차: {fmt_days(left)}' + (' (이 신청 반영)' if row['status'] in HOLDING else ''))
    lines.append(f'• 승인: {step_line(row, step_rows)}')
    if needs_proof(row['ltype']) and row['status'] in HOLDING:
        lines.append('📎 증빙(청첩장·부고·훈련 통지서 등)은 경영지원에 보내 주세요.')
    for s in step_rows:
        if s['result'] == 'rejected' and s.get('reason'):
            lines.append(f"❌ 반려 사유 ({s.get('actor_name') or ''}): {s['reason']}")
        elif s['result'] == 'shortened':
            lines.append(f"⚪ 남은 날 취소 ({s.get('actor_name') or ''}): {s.get('reason') or ''}")
        elif s['result'] == 'edited':
            who = '경영지원 ' if s['step'] == 'admin_edit' else ''
            lines.append(f"✏️ {who}고침 ({s.get('actor_name') or ''}): {s.get('reason') or ''}")
        elif s['result'] == 'replaced':
            lines.append(f"⚪ {s.get('reason') or '변경 신청으로 바뀜'}")
        elif s['result'] == 'registered':
            lines.append(f"🗂 경영지원 등록 ({s.get('actor_name') or ''})")
    return '\n'.join(lines)


A_APPROVE, A_REJECT, A_CANCEL = 'leave_approve', 'leave_reject', 'leave_cancel'
A_EDIT = 'leave_edit'   # 요청자 카드 [✏️ 고치기]·[✏️ 변경 신청] (11-18)
A_NEW, A_MINE = 'leave_new', 'leave_mine'
FN_OPEN_LEAVE = 'leave_open'   # 휴가 봇 워크플로 단계 「휴가 신청 열기」 — docs/slack_leave_manifest.yml 키와 같게
_DEFAULT_PILOTS = ('U04UL2ZLJAX',)   # 고광일


def enabled() -> bool:
    """플렉스에서 옮긴 뒤(LEAVE_ENABLED=true) — 전원 신청·연차 자동 발생."""
    return os.getenv('LEAVE_ENABLED', '').strip().lower() in ('1', 'true', 'yes')


def is_open(uid: str) -> bool:
    """신청 받기 — 플렉스에서 옮기는 날(LEAVE_ENABLED=true) 전까지는 시험하는 사람(LEAVE_PILOT_IDS, 기본 KiKO)만."""
    if enabled():
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


# 본인 취소 단추 — 슬랙 요청자 카드·슬랙 [내 휴가]·그룹웨어 내 휴가가 같은 규칙과 문구를 쓴다(11-15)
CANCEL_TITLE, CANCEL_YES, CANCEL_NO = '휴가 취소', '취소하기', '그대로 두기'


def edit_label(mode: str) -> str:
    return '✏️ 변경 신청' if mode == 'change' else '✏️ 고치기'


def cancel_label(mode: str) -> str:
    return '남은 날 취소' if mode == 'rest' else '신청 취소'


def cancel_confirm_text(mode: str) -> str:
    return ('오늘부터 남은 휴가를 취소할까요? 어제까지 쓴 날은 그대로 남고, 나머지 연차는 돌아와요.' if mode == 'rest'
            else '이 휴가를 취소할까요? 쓴 연차는 돌아와요.')


def own_cancel_mode(row: dict, uid: str = '', today: Optional[date] = None) -> str:
    """본인 취소 단추를 보일지 — 'all'·'rest', 안 보이면 ''. 승인 대기이거나 아직 안 끝난 휴가만.
    uid 를 주면 그 사람이 본인일 때만(그룹웨어·[내 휴가]), 안 주면 신청자 기준(요청자 DM 카드)."""
    today = today or date.today()
    who = uid or row.get('slack_id') or row.get('submitter_id') or ''
    if not who or who not in (row.get('slack_id'), row.get('submitter_id')):
        return ''
    if row['status'] not in HOLDING or not (row['status'] in OPEN or _d(row['end']) >= today):
        return ''
    if not can_act('cancel', row, who, today):
        return ''
    return cancel_mode(row, who, today)


def cancel_button(row: dict, uid: str = '', today: Optional[date] = None) -> Optional[dict]:
    mode = own_cancel_mode(row, uid, today)
    if not mode:
        return None
    return _btn(cancel_label(mode), A_CANCEL, row['id'], confirm={
        'title': {'type': 'plain_text', 'text': CANCEL_TITLE},
        'text': {'type': 'plain_text', 'text': cancel_confirm_text(mode)},
        'confirm': {'type': 'plain_text', 'text': CANCEL_YES},
        'deny': {'type': 'plain_text', 'text': CANCEL_NO}})


def card_blocks(row: dict, step_rows: List[dict], role: str, left: Optional[float] = None) -> List[dict]:
    """role: approver(지금 승인할 사람) / requester / info."""
    blocks = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': card_text(row, step_rows, role, left)[:2900]}}]
    els = []
    if role == 'approver' and row['status'] in OPEN:
        els = [_btn('✅ 승인', A_APPROVE, row['id'], 'primary'), _btn('반려', A_REJECT, row['id'], 'danger')]
    elif role == 'requester':
        mode = own_edit_mode(row)
        if mode:
            els.append(_btn(edit_label(mode), A_EDIT, row['id']))
        b = cancel_button(row)
        if b:
            els.append(b)
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
