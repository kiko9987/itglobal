"""토요일 휴무(격주휴무) — 기술영업이 쉬는 토요일을 그룹웨어 달력에서 고르고 결재받는다(2026-11-20 부장님).

예전 방식(2024-12 격주휴무 회의, 드라이브 '기술영업 휴무자료'): 매월 첫째 주 각자 빈승정 대리에게 알림 → 빈 대리가
엑셀 「격주휴무 신청서」로 취합 → 담당실장·대표 결재. 한 사람 한 달 2번. 공사가 겹치면 일이 먼저(다른 주로 미룸).

그룹웨어:
- 대상 = 인사카드 부서 '기술영업'(재직). 달력의 토요일을 눌러 자기 이름을 넣고 뺀다.
- 달 상태: 작성 중 → (취합 담당이 [결재 올리기]) 결재 대기 → (박정우·박용구 둘 중 한 분 승인) 확정.
  돌려보내면 다시 작성 중. 확정 뒤 바꾸는 건 변경 신청 → 둘 중 한 분 승인(부장님 11-20).
- 한 달 몇 번인지는 정하지 않는다(11-24 부장님 — 예전엔 2번 기준 경고). 사람별 횟수만 보여 준다.
- 휴무 출근(11-24 부장님): 쉬기로 한 토요일에 급한 주말 공사로 나오면 박정우·박용구 중 한 분께 구두로 승인받고
  본인이 [휴무 출근 기록]으로 남긴다(결재 단계 없음). 그날은 쉬는 사람에서 빠지고 휴무 출근으로 남는다.
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
# 한 달 몇 번 쉬는지는 정해 두지 않는다 — 급한 주말 공사로 쉬는 날에도 나오는 일이 많아(11-24 부장님) 횟수는 보여 주기만
WORKED_BACK_DAYS = 31                  # 휴무 출근은 지난 날도 기록할 수 있게(다음 주에 적는 경우) — 한 달까지
AHEAD_MONTHS = 6                      # 몇 달 앞까지 고를 수 있나(결혼식·해외여행은 3달 전에 알리기로 함)
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
    # 휴무 출근(11-24 부장님) — 쉬기로 한 토요일에 급한 주말 공사로 나온 날. 박정우·박용구께 구두로 승인받고 본인이 기록.
    # 쉬는 사람 표(sat_off_days)에서는 빠져(차량·방문·휴가 일수는 근무로 봄) 여기 남는다
    '''CREATE TABLE IF NOT EXISTS sat_off_worked (
        day TEXT NOT NULL,
        name TEXT NOT NULL,
        slack_id TEXT,
        verbal_by TEXT,
        reason TEXT,
        recorded_by TEXT,
        recorded_at TEXT,
        PRIMARY KEY (day, name)
    )''',
    # 전체 휴무(11-25 부장님) — 관리자가 토요일 하루를 모두 쉬는 날로. 대상자를 그날 쉬는 사람(sat_off_days, added_by=FULL_TAG)으로
    # 넣으므로 근무표·횟수(격주 1일 차감)·차량·방문·휴가 일수가 따로 손대지 않아도 맞는다. 이 표는 표시·풀기용
    '''CREATE TABLE IF NOT EXISTS sat_off_full (
        day TEXT PRIMARY KEY,
        note TEXT,
        set_by TEXT,
        set_at TEXT
    )''',
    # 휴무 메모(11-25 부장님) — 엑셀 「격주휴무 신청서」 메모란 대신. 각자 '꼭 쉬어야 하는 토요일'·'못 쉬는 토요일'과 사유를
    # 먼저 적어 두면 다른 매니저가 보고 양보하고 피해서 잡는다. day='' + name='' 은 그달 전체 메모(취합 담당·승인권자)
    '''CREATE TABLE IF NOT EXISTS sat_off_memos (
        month TEXT NOT NULL,
        day TEXT NOT NULL DEFAULT '',
        name TEXT NOT NULL DEFAULT '',
        kind TEXT NOT NULL,
        reason TEXT NOT NULL,
        written_by TEXT,
        written_at TEXT,
        PRIMARY KEY (month, day, name)
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


def _tech_sales(on: Optional[date] = None) -> Tuple[List[dict], List[dict]]:
    """(대상자, 인센티브) — 인사카드 부서가 기술영업인 재직자를 인센티브 시작일(on 기준, 기본 오늘)로 가른다(이름순).
    인센티브은 격주휴무가 없어 스케줄 근무 대상에서 빠진다(11-25 부장님). 인센티브 칸이 아직 없는 DB(재시작 전)면 모두 대상."""
    on_iso = (on or date.today()).isoformat()
    with lv._conn() as conn:
        if not conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='hr_people'").fetchone():
            return [], []
        has_inc = 'incentive_from' in {r[1] for r in conn.execute('PRAGMA table_info(hr_people)')}
        rows = conn.execute(f"SELECT name, slack_id, {'incentive_from' if has_inc else 'NULL'} FROM hr_people "
                            'WHERE department=? AND status != ? ORDER BY name', (MEMBER_DEPT, lv._HR_RESIGNED)).fetchall()
    mem, inc = [], []
    for name, sid, frm in rows:
        p = {'name': name, 'slack_id': sid or ''}
        (inc if frm and str(frm)[:10] <= on_iso else mem).append(p)
    return mem, inc


def members(on: Optional[date] = None) -> List[dict]:
    """대상자 — 인사카드 부서가 기술영업인 재직자 중 인센티브이 아닌 사람 [{name, slack_id}] (이름순)."""
    return _tech_sales(on)[0]


def incentive_members(on: Optional[date] = None) -> List[dict]:
    """기술영업 중 인센티브(on 기준) — 스케줄 근무 대상이 아님. 화면에 '인센티브' 으로 따로 보여 준다."""
    return _tech_sales(on)[1]


def is_member(name: str, on: Optional[date] = None) -> bool:
    return bool(name) and _member(name, on) is not None


def _member(name: str, on: Optional[date] = None) -> Optional[dict]:
    return next((m for m in members(on) if m['name'] == name), None)


def role(sid: str, name: str = '', on: Optional[date] = None) -> Dict[str, bool]:
    approver = bool(sid) and sid in approver_ids()
    coordinator = bool(sid) and sid in coordinator_ids()
    manager = bool(sid) and sid in manager_ids()
    member = bool(name) and _member(name, on) is not None
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
    """회사 쉬는 날 이름(법정 공휴일 + 회사 지정, 11-25 하나로) — 아니면 ''."""
    return lv.holiday_name(d)


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


FULL_TAG = '전체 휴무'


def full_days(month: str) -> Dict[str, str]:
    """그달 전체 휴무 토요일 {날짜: 메모}."""
    with lv._conn() as conn:
        rows = conn.execute('SELECT day, note FROM sat_off_full WHERE day LIKE ?', (month + '-%',)).fetchall()
    return {r['day']: r['note'] or '' for r in rows}


def set_full_off(day: Optional[date], on: bool, actor_name: str, note: str = '') -> dict:
    """전체 휴무 지정(on)·풀기 — 관리자만(권한은 화면 API 가 본다). 지정하면 그달 대상자(인센티브 제외) 중 그날 휴무 출근을
    기록하지 않은 사람을 모두 쉬는 사람으로 넣는다(절반 제한 없음). 풀면 전체 휴무로 넣은 이름만 뺀다(각자 고른 건 남김)."""
    if not day or day.weekday() != 5:
        raise SatOffError('토요일만 고를 수 있어요.')
    if holiday_name(day):
        raise SatOffError(f'{kday(day)}는 공휴일({holiday_name(day)})이라 이미 모두 쉬는 날이에요.')
    key, month = day.isoformat(), _month_of(day)
    note = (note or '').strip()[:100]
    now = _now()
    with lv._conn() as conn:
        if on:
            conn.execute('INSERT OR REPLACE INTO sat_off_full (day, note, set_by, set_at) VALUES (?, ?, ?, ?)',
                         (key, note or None, actor_name, now))
            worked_names = {r[0] for r in conn.execute('SELECT name FROM sat_off_worked WHERE day=?', (key,))}
            n = 0
            for m in members(_first(month)):
                if m['name'] in worked_names:
                    continue
                cur = conn.execute('INSERT OR IGNORE INTO sat_off_days (day, name, slack_id, added_by, added_at) '
                                   'VALUES (?, ?, ?, ?, ?)', (key, m['name'], m['slack_id'], FULL_TAG, now))
                n += cur.rowcount
            _log(conn, actor_name, 'full_off', month, f'{kday(day)} 전체 휴무로 지정 — {n}명 넣음' + (f' ({note})' if note else ''))
        else:
            if not conn.execute('SELECT 1 FROM sat_off_full WHERE day=?', (key,)).fetchone():
                raise SatOffError(f'{kday(day)}는 전체 휴무가 아니에요. 새로 고쳐 주세요.')
            conn.execute('DELETE FROM sat_off_full WHERE day=?', (key,))
            n = conn.execute('DELETE FROM sat_off_days WHERE day=? AND added_by=?', (key, FULL_TAG)).rowcount
            _log(conn, actor_name, 'full_off_undo', month, f'{kday(day)} 전체 휴무 풂 — {n}명 뺌')
    return {'day': key, 'on': on, 'count': n}


def worked(month: str) -> List[dict]:
    """그달 휴무 출근 기록 [{day, label, name, verbal_by, reason, recorded_by, recorded_at}] (날짜·이름순)."""
    with lv._conn() as conn:
        rows = conn.execute('SELECT day, name, verbal_by, reason, recorded_by, recorded_at FROM sat_off_worked '
                            'WHERE day LIKE ? ORDER BY day, name', (month + '-%',)).fetchall()
    return [dict(r, label=kday(lv._d(r['day']))) for r in rows]


def approvers() -> List[dict]:
    """구두 승인할 수 있는 분 [{name, slack_id}] — 휴가 승인권자(박정우·박용구) 순서대로, 인사카드 이름."""
    ids = approver_ids()
    if not ids:
        return []
    with lv._conn() as conn:
        if not conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='hr_people'").fetchone():
            return []
        q = ','.join('?' * len(ids))
        rows = conn.execute(f'SELECT name, slack_id FROM hr_people WHERE slack_id IN ({q}) AND status != ?',
                            (*ids, lv._HR_RESIGNED)).fetchall()
    by_sid = {r[1]: r[0] for r in rows}
    return [{'name': by_sid[s], 'slack_id': s} for s in ids if s in by_sid]


def work_saturdays(name: str, start: date, end: date) -> Dict[date, bool]:
    """휴가 일수용 — 기간 안 토요일 중 휴무표가 확정된 달의 것만 {토요일: 근무하는 날인가}.
    대상자가 아니거나 확정 전 달이면 넣지 않는다(→ 예전 규칙)."""
    sats = []
    d = start + timedelta(days=(5 - start.weekday()) % 7)
    while d <= end:
        sats.append(d)
        d += timedelta(days=7)
    if not sats or not is_member(name, sats[0].replace(day=1)):
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
        parts.append(', '.join(rm) + ' 휴무 취소')
    if ad:
        parts.append(', '.join(ad) + ' 휴무 신청')
    return ' · '.join(parts)


def recent_log(month: str, limit: int = 20) -> List[dict]:
    with lv._conn() as conn:
        rows = conn.execute('SELECT at, actor, action, detail FROM sat_off_log WHERE month=? ORDER BY id DESC LIMIT ?',
                            (month, limit)).fetchall()
    return [dict(r) for r in rows]


# ── 휴무 메모(11-25) ────────────────────────────────────────────────────
MEMO_MUST, MEMO_CANT = 'must', 'cant'
MEMO_MONTH = 'month'   # 예전 공통 메모 — 11-25 부장님 결정으로 없앰(본인 메모로 충분). 읽을 때 건너뜀
MEMO_LABELS = {MEMO_MUST: '꼭 쉬어야 해요', MEMO_CANT: '못 쉬어요'}
MEMO_MAX = 60


def memos(month: str) -> dict:
    """{'items': [{day, label, name, kind, kind_label, reason, written_by, written_at}]} (토요일·이름순)."""
    with lv._conn() as conn:
        rows = [dict(r) for r in conn.execute(
            'SELECT * FROM sat_off_memos WHERE month=? ORDER BY day, kind, name', (month,)).fetchall()]
    items = []
    for r in rows:
        if r['kind'] not in MEMO_LABELS:     # 예전 공통 메모 등
            continue
        d = lv._d(r['day'])
        items.append({'day': r['day'], 'label': kday(d) if d else r['day'], 'name': r['name'], 'kind': r['kind'],
                      'kind_label': MEMO_LABELS.get(r['kind'], r['kind']), 'reason': r['reason'],
                      'written_by': r.get('written_by') or '', 'written_at': r.get('written_at') or ''})
    return {'items': items}


def _memo_reason(reason: str, limit: int) -> str:
    text = ' '.join(str(reason or '').split())
    if not text:
        raise SatOffError('사유를 적어 주세요. 예: 가족여행, 결혼식, 현장 일정')
    if len(text) > limit:
        raise SatOffError(f'사유는 {limit}자까지 적을 수 있어요.')
    return text


def write_memo(day: Optional[date], target: str, kind: str, reason: str, sid: str, actor_name: str,
               today: Optional[date] = None, admin: bool = False) -> dict:
    """본인 토요일 메모 — '꼭 쉬어야 해요'·'못 쉬어요' + 사유. 로그인한 본인만(11-25 부장님 — 대신 적기 없음). 같은 토요일은 고쳐 씀.
    '꼭 쉬어야 해요'는 작성 중인 달이면 그 토요일에 이름도 넣는다(절반이 차 있으면 메모만 남기고 알려 줌 — 다른 분 양보 대기)."""
    today = today or date.today()
    if kind not in (MEMO_MUST, MEMO_CANT):
        raise SatOffError('메모 종류를 골라 주세요.')
    on = day.replace(day=1) if day else None
    r = role(sid, actor_name, on)
    edit = r['approver'] or bool(admin)
    month = _check_day(day, today, edit)
    if not _member(target, on):
        raise SatOffError(f'{target} 님은 스케줄 근무 대상(기술영업)이 아니에요.')
    if target != actor_name:
        raise PermissionError('메모는 본인 것만 적을 수 있어요.')
    if not r['member']:
        raise PermissionError('스케줄 근무 대상(기술영업)만 메모를 적을 수 있어요.')
    text = _memo_reason(reason, MEMO_MAX)
    key = day.isoformat()
    with lv._conn() as conn:
        conn.execute('INSERT OR REPLACE INTO sat_off_memos (month, day, name, kind, reason, written_by, written_at) '
                     'VALUES (?, ?, ?, ?, ?, ?, ?)', (month, key, target, kind, text, actor_name, _now()))
        _log(conn, actor_name, 'memo', month, f'{kday(day)} {target} 메모 — {MEMO_LABELS[kind]}: {text}')
    picked, warn = target in picks(month).get(key, []), ''
    st = month_row(month)['status']
    if kind == MEMO_MUST and not picked:
        if st == OPEN or edit:
            try:
                toggle(day, target, True, sid, actor_name, today=today, admin=admin)
                picked = True
            except SatOffError as e:          # 절반이 찼음 — 메모는 남고, 보고 양보해 주면 넣을 수 있게
                warn = f'{e} 다른 분이 보고 양보하면 그때 이름을 넣어 주세요.'
        elif st == CONFIRMED:
            warn = '확정된 달이라 이름은 넣지 않았어요. 쉬려면 [변경 신청]을 해 주세요.'
        else:
            warn = '결재 대기 중인 달이라 이름은 넣지 않았어요.'
    return {'day': key, 'name': target, 'kind': kind, 'reason': text, 'picked': picked, 'warn': warn}


def delete_memo(day: Optional[date], target: str, sid: str, actor_name: str, admin: bool = False) -> dict:
    """토요일 메모 지우기 — 본인 것만(11-25 부장님). 휴무 이름은 그대로 둔다(빼려면 달력에서)."""
    if not day:
        raise SatOffError('날짜를 다시 골라 주세요.')
    if target != actor_name:
        raise PermissionError('메모는 본인 것만 지울 수 있어요.')
    month = _month_of(day)
    with lv._conn() as conn:
        n = conn.execute('DELETE FROM sat_off_memos WHERE month=? AND day=? AND name=? AND kind<>?',
                         (month, day.isoformat(), target, MEMO_MONTH)).rowcount
        if not n:
            raise LookupError('이미 지워진 메모예요. 새로 고쳐 주세요.')
        _log(conn, actor_name, 'memo', month, f'{kday(day)} {target} 메모 지움')
    return {'day': day.isoformat(), 'name': target}


def view(month: str, sid: str, name: str, today: Optional[date] = None, admin: bool = False) -> dict:
    """달력 화면 한 달치 — 토요일마다 쉬는 사람, 사람별 횟수, 내 역할로 할 수 있는 것."""
    today = today or date.today()
    row = month_row(month)
    st = row['status']
    on = _first(month)                       # 인센티브는 그달 1일 기준(고르기·휴가 일수와 같게)
    r = role(sid, name, on)
    p = picks(month)
    c = counts(month)
    wk = worked(month)
    wc: Dict[str, int] = {}
    for w in wk:
        wc[w['name']] = wc.get(w['name'], 0) + 1
    sats = []
    for d in saturdays(month):
        hol = holiday_name(d)
        sats.append({'day': d.isoformat(), 'label': kday(d), 'holiday': hol, 'past': d < today,
                     'names': p.get(d.isoformat(), []), 'mine': name in p.get(d.isoformat(), []),
                     'worked': [w['name'] for w in wk if w['day'] == d.isoformat()]})
    in_range = _month_of(today) <= month <= editable_until(today)
    can_edit_open = st == OPEN and in_range and (r['member'] or r['manager'])
    changes = _changes(month)
    edit = r['approver'] or bool(admin)          # 확정된 달도 바로 고치는 사람 — 승인권자·관리자(11-25)
    show_all = edit or r['manager']
    my_wait = any(c['status'] == CH_WAIT and c['name'] == name for c in changes)   # 대기 중이면 새 신청은 그 뒤에
    mm = memos(month)
    return {
        'month': month, 'status': st, 'status_label': STATUS_LABELS[st],
        'submitted_by': row.get('submitted_by') or '', 'submitted_at': row.get('submitted_at') or '',
        'decided_by': row.get('decided_by') or '', 'decided_at': row.get('decided_at') or '',
        'note': row.get('note') or '',
        'saturdays': sats,
        'max_off': len(members(on)) // 2,        # 한 토요일 최대 쉬는 인원(절반, 11-25) — max_off(month) 와 같음
        'members': [{'name': m['name'], 'count': c.get(m['name'], 0), 'worked': wc.get(m['name'], 0)} for m in members(on)],
        'incentive': [m['name'] for m in incentive_members(on)],      # 기술영업 인센티브 — 스케줄 대상 아님(11-25)
        'me': {'name': name, **r, 'editor': edit},
        # 휴무 출근(11-24) — 결재 올린 뒤(대기·확정)부터. 작성 중엔 달력에서 이름을 빼면 됨
        'can_mark_worked': st in (WAIT, CONFIRMED) and (r['member'] or r['manager'] or r['approver']),
        'approvers': [a['name'] for a in approvers()],
        'worked': [dict(w, can_undo=w['name'] == name or r['manager'] or r['approver']) for w in wk],
        'full_days': full_days(month),           # 전체 휴무 토요일 {날짜: 메모}(11-25). 지정·풀기 권한은 화면 API(can_full_off)
        'can_toggle': can_edit_open or edit,
        'pick_others': r['manager'] or edit,               # 남의 이름도 넣고 뺄 수 있음(작성 중일 때)
        'can_submit': st == OPEN and r['manager'] and in_range,
        'can_decide': st == WAIT and r['approver'],
        'can_request_change': st == CONFIRMED and (r['member'] or r['manager']) and not r['approver'] and not my_wait,
        'changes': [x for x in changes if show_all or x['name'] == name],
        'log': recent_log(month) if show_all else [],
        'notify_fail': notify_fail(month),       # 마지막 처리의 슬랙 알림을 못 보냈으면 그 문구(11-24)
        # 휴무 메모(11-25) — 모두 봄. 쓰기·지우기 = 로그인한 본인(대상자)만. 공통 메모는 없앰(부장님)
        'memos': mm['items'],
        'can_memo': in_range and r['member'],          # 본인만(11-25 부장님 — 대신 적기 없음)
        'can_memo_others': False,
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
        raise SatOffError(f'{kday(d)}는 공휴일({holiday_name(d)})이라 모두 쉬는 날이에요.')
    month = _month_of(d)
    if not approver:
        if d < today:
            raise SatOffError('지난 토요일은 바꿀 수 없어요.')
        if month > editable_until(today):
            raise SatOffError(f'{AHEAD_MONTHS}달 뒤까지만 고를 수 있어요.')
    return month


def max_off(month: str) -> int:
    """한 토요일에 쉴 수 있는 최대 인원 — 그달 스케줄 근무 대상(인센티브 제외)의 절반, 내림.
    과반수 이상이 쉬지 못하게(11-25 부장님 — 2명만 근무하고 다 쉰 날이 있었음). 10명이면 5명."""
    return len(members(_first(month))) // 2


def _over_cap(conn, month: str, day_iso: str, approver: bool) -> str:
    """저장 뒤(같은 연결 안) 그날 쉬는 사람이 절반을 넘었나. 넘었으면 승인권자는 경고 문구, 그 밖엔 SatOffError(저장 안 됨).
    승인권자는 전체 휴무 같은 예외가 있어 막지 않는다."""
    n = conn.execute('SELECT COUNT(*) FROM sat_off_days WHERE day=?', (day_iso,)).fetchone()[0]
    cap = max_off(month)
    if n <= cap:
        return ''
    msg = f'{kday(lv._d(day_iso))}는 쉬는 사람이 {n}명이 돼요. 한 토요일에 절반({cap}명)까지만 쉴 수 있어요.'
    if approver:
        return msg[:-len('쉴 수 있어요.')] + '쉬는 게 기준이에요(승인권자·관리자라 그대로 저장했어요).'
    raise SatOffError(msg)


def toggle(day: Optional[date], target: str, on: bool, sid: str, actor_name: str,
           today: Optional[date] = None, admin: bool = False) -> dict:
    """토요일 칸에 이름 넣기(on)·빼기. 반환 {names, count, warn}.
    admin — 그룹웨어 관리자(admin·super_admin). 승인권자처럼 확정된 달도 바로 고친다(11-25 부장님: 변경은 본인 신청이
    원칙이지만 관리자는 직접 수정도). 결재 승인 권한은 아님."""
    today = today or date.today()
    r = role(sid, actor_name, day.replace(day=1) if day else None)    # 인센티브는 그달 1일 기준(화면과 같게)
    edit = r['approver'] or bool(admin)
    month = _check_day(day, today, edit)
    m = _member(target, day.replace(day=1) if day else None)
    if not m:
        raise SatOffError(f'{target} 님은 스케줄 근무 대상(기술영업)이 아니에요.')
    if target != actor_name and not (r['manager'] or edit):
        raise PermissionError('다른 사람 이름은 넣거나 뺄 수 없어요.')
    if target == actor_name and not (r['member'] or r['manager']):
        raise PermissionError('스케줄 근무 대상이 아니에요.')
    st = month_row(month)['status']
    if st != OPEN and not edit:
        raise SatOffError('결재 대기 중이거나 확정된 달이에요. 바꾸려면 [변경 신청]을 해 주세요.'
                          if st == CONFIRMED else '결재 대기 중인 달이라 지금은 바꿀 수 없어요.')
    key = day.isoformat()
    warn = ''
    with lv._conn() as conn:
        if on:
            cur = conn.execute('INSERT OR IGNORE INTO sat_off_days (day, name, slack_id, added_by, added_at) '
                               'VALUES (?, ?, ?, ?, ?)', (key, target, m['slack_id'], actor_name, _now()))
            if cur.rowcount:                    # 새로 넣었을 때만 절반 제한(이미 있던 이름은 그대로)
                warn = _over_cap(conn, month, key, edit)
            conn.execute('DELETE FROM sat_off_worked WHERE day=? AND name=?', (key, target))   # 다시 쉬면 출근 기록은 지움
        else:
            conn.execute('DELETE FROM sat_off_days WHERE day=? AND name=?', (key, target))
        how = '넣음' if on else '뺌'
        _log(conn, actor_name, 'toggle', month,
             f'{kday(day)} {target} {how}' + (f' (확정 뒤 직접 수정)' if st != OPEN else '') + (' — 절반 넘음' if warn else ''))
    n = counts(month).get(target, 0)
    # 한 달 몇 번 기준 경고는 뺐다(11-24). warn = 승인권자가 한 토요일 절반을 넘겨 넣었을 때만(11-25)
    return {'names': picks(month).get(key, []), 'count': n, 'warn': warn}


def mark_worked(day: Optional[date], target: str, verbal_by: str, reason: str, sid: str, actor_name: str,
                today: Optional[date] = None) -> dict:
    """휴무 출근 기록 — 쉬기로 한 토요일에 급히 나옴. 박정우·박용구 중 한 분께 구두로 승인받은 뒤 본인이 남긴다
    (결재 단계 없음, 11-24 부장님). 그날은 쉬는 사람에서 빠지고(차량·방문 배정·휴가 일수는 근무로 봄) 기록이 남는다."""
    today = today or date.today()
    r = role(sid, actor_name, day.replace(day=1) if day else None)
    if not day or day.weekday() != 5:
        raise SatOffError('토요일만 고를 수 있어요.')
    if holiday_name(day):
        raise SatOffError(f'{kday(day)}는 공휴일({holiday_name(day)})이라 모두 쉬는 날이에요.')
    if day < today - timedelta(days=WORKED_BACK_DAYS) and not r['approver']:
        raise SatOffError('한 달보다 지난 토요일은 승인권자께 고쳐 달라고 말씀해 주세요.')
    month = _month_of(day)
    st = month_row(month)['status']
    if st == OPEN:
        raise SatOffError('아직 작성 중인 달이에요. 달력에서 이름을 빼면 돼요.')
    m = _member(target, day.replace(day=1) if day else None)
    if not m:
        raise SatOffError(f'{target} 님은 스케줄 근무 대상(기술영업)이 아니에요.')
    if target != actor_name and not (r['manager'] or r['approver']):
        raise PermissionError('다른 사람의 휴무 출근은 취합 담당·승인권자만 기록할 수 있어요.')
    if target == actor_name and not (r['member'] or r['manager'] or r['approver']):
        raise PermissionError('스케줄 근무 대상이 아니에요.')
    verbal_by = (verbal_by or '').strip()
    if verbal_by not in {a['name'] for a in approvers()}:
        raise SatOffError('구두로 승인해 주신 분을 골라 주세요.')
    reason = (reason or '').strip()[:200]
    if not reason:
        raise SatOffError('어떤 일로 나오는지 적어 주세요. (예: ○○ 현장 공사)')
    key = day.isoformat()
    with lv._conn() as conn:
        cur = conn.execute('DELETE FROM sat_off_days WHERE day=? AND name=?', (key, target))
        if cur.rowcount != 1:
            raise SatOffError(f'{kday(day)}는 {target} 님이 쉬는 날이 아니에요. 새로 고쳐 주세요.')
        conn.execute('INSERT OR REPLACE INTO sat_off_worked (day, name, slack_id, verbal_by, reason, recorded_by, '
                     'recorded_at) VALUES (?, ?, ?, ?, ?, ?, ?)',
                     (key, target, m['slack_id'], verbal_by, reason, actor_name, _now()))
        _log(conn, actor_name, 'worked', month, f'{kday(day)} {target} 휴무 출근 (구두 승인 {verbal_by}) — {reason}')
    return {'day': key, 'label': kday(day), 'name': target, 'verbal_by': verbal_by, 'reason': reason}


def undo_worked(day: Optional[date], target: str, sid: str, actor_name: str) -> dict:
    """휴무 출근 기록 되돌리기(잘못 눌렀을 때) — 그날을 다시 쉬는 날로. 본인·취합 담당·승인권자."""
    r = role(sid, actor_name)
    if target != actor_name and not (r['manager'] or r['approver']):
        raise PermissionError('본인 기록만 되돌릴 수 있어요.')
    if not day:
        raise SatOffError('날짜를 다시 골라 주세요.')
    key, month = day.isoformat(), _month_of(day)
    with lv._conn() as conn:
        w = conn.execute('SELECT slack_id FROM sat_off_worked WHERE day=? AND name=?', (key, target)).fetchone()
        if not w:
            raise SatOffError('휴무 출근 기록이 없어요. 새로 고쳐 주세요.')
        conn.execute('DELETE FROM sat_off_worked WHERE day=? AND name=?', (key, target))
        conn.execute('INSERT OR IGNORE INTO sat_off_days (day, name, slack_id, added_by, added_at) VALUES (?, ?, ?, ?, ?)',
                     (key, target, w['slack_id'], actor_name, _now()))
        warn = _over_cap(conn, month, key, r['approver'])      # 다시 쉬면 그날 절반을 넘는지(11-25)
        _log(conn, actor_name, 'worked_undo', month, f'{kday(day)} {target} 휴무 출근 되돌림(다시 쉬는 날)')
    return {'day': key, 'name': target, 'warn': warn}


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
        raise PermissionError('승인은 승인권자만 할 수 있어요.')
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
            raise SatOffError(f'{int(month[5:7])}월 토요일만 고를 수 있어요.')
        if d.isoformat() not in out:
            out.append(d.isoformat())
    return sorted(out)


def _check_change(month: str, name: str, add: List[str], remove: List[str], today: date, approver: bool) -> None:
    if not add and not remove:
        raise SatOffError('바꿀 토요일을 골라 주세요.')
    if set(add) & set(remove):
        raise SatOffError('같은 날을 넣고 빼려고 했어요. 다시 골라 주세요.')
    p = picks(month)
    cap = max_off(month)
    for x in add:
        _check_day(lv._d(x), today, approver)
        if name in p.get(x, []):
            raise SatOffError(f'{kday(lv._d(x))}는 이미 쉬는 날이에요.')
        if not approver and len(p.get(x, [])) + 1 > cap:     # 한 토요일 절반 제한(11-25) — 승인권자만 예외
            raise SatOffError(f'{kday(lv._d(x))}는 이미 {len(p.get(x, []))}명이 쉬어요. '
                              f'한 토요일에 절반({cap}명)까지만 쉴 수 있어요.')
    for x in remove:
        _check_day(lv._d(x), today, approver)
        if name not in p.get(x, []):
            raise SatOffError(f'{kday(lv._d(x))}는 쉬는 날이 아니에요.')


def _apply(conn, month: str, name: str, slack_id: str, add: List[str], remove: List[str], actor: str) -> None:
    now = _now()
    for x in remove:
        conn.execute('DELETE FROM sat_off_days WHERE day=? AND name=?', (x, name))
    for x in add:
        conn.execute('INSERT OR IGNORE INTO sat_off_days (day, name, slack_id, added_by, added_at) VALUES (?, ?, ?, ?, ?)',
                     (x, name, slack_id, actor, now))
        conn.execute('DELETE FROM sat_off_worked WHERE day=? AND name=?', (x, name))   # 다시 쉬면 출근 기록은 지움


def request_change(month: str, target: str, add, remove, reason: str, sid: str, actor_name: str,
                   today: Optional[date] = None) -> dict:
    """확정된 달 바꾸기 — 승인권자가 아니면 결재 대기로 남고, 승인권자가 하면 바로 반영."""
    today = today or date.today()
    r = role(sid, actor_name, _first(month))
    if month_row(month)['status'] != CONFIRMED:
        raise SatOffError('확정된 달만 변경 신청을 해요. 작성 중인 달은 달력에서 바로 바꾸면 돼요.')
    m = _member(target, _first(month))
    if not m:
        raise SatOffError(f'{target} 님은 스케줄 근무 대상(기술영업)이 아니에요.')
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
        raise PermissionError('승인은 승인권자만 할 수 있어요.')
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
    names = {m['name'] for m in members(_first(month))}
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


def _dm(client, uid: str, text: str, month: str) -> bool:
    if not client:
        return False
    if not uid:
        logger.warning(f'{_LOG} 슬랙 ID 가 없어 DM 을 못 보냄 ({month})')
        return False
    try:
        ch = client.conversations_open(users=uid)['channel']['id']
        client.chat_postMessage(channel=ch, text=text, blocks=_blocks(text, month))
        return True
    except Exception as e:
        logger.warning(f'{_LOG} DM 실패 ({uid}): {e}')
        return False


def _dm_some(client, uids, text: str, month: str, what: str) -> None:
    """여러 사람 DM — 한 명도 못 받으면 ERROR 알림(11-24 조용한 실패 점검 5번: 결재 요청이 안 가도 아무도 몰랐음.
    그룹웨어 알림 종에도 토요일 휴무는 없다)."""
    uids = list(dict.fromkeys(u for u in (uids or []) if u is not None))
    sent = sum(1 for u in uids if _dm(client, u, text, month))
    if not sent:
        logger.error(f'{_LOG} {what} 알림을 아무에게도 못 보냄 ({month}, 대상 {len(uids)}명) — 직접 알려 주세요',
                     extra={'force_alert': True})
        _note_fail(what)


def _post(client, text: str) -> bool:
    if not client:
        return False
    try:
        client.chat_postMessage(channel=lv.announce_channel(), text=text)
        return True
    except Exception as e:
        logger.error(f'{_LOG} #휴가_일정 공지 실패 — 채널에 직접 올려 주세요: {e}', extra={'force_alert': True})
        _note_fail('#휴가_일정 공지')
        return False


# ── 못 보낸 알림을 화면에도(11-24 조용한 실패 점검 추가분) ─────────────────────────────
# 알림은 뒤에서 나가 응답에 결과가 없다. ERROR 알림은 부장님께만 가므로, 버튼을 누른 사람도 알게
# 토요일 휴무 기록에 'notify_fail' 한 줄을 남기고 화면(view 의 notify_fail)이 '직접 알려 주세요'를 띄운다.
NOTIFY_FAIL = 'notify_fail'
_fails = threading.local()


def _note_fail(what: str) -> None:
    items = getattr(_fails, 'items', None)
    if items is not None:
        items.append(what)


def _record_fail(month: str, items: List[str]) -> None:
    detail = f"슬랙 알림을 못 보냈어요({', '.join(dict.fromkeys(items))}). 직접 알려 주세요."
    try:
        with lv._conn() as conn:
            _log(conn, '슬랙 알림', NOTIFY_FAIL, month, detail)
    except Exception as e:
        logger.error(f'{_LOG} 알림 실패 기록 실패 ({month}): {e}', exc_info=True)


def notify_fail(month: str) -> str:
    """이 달 마지막 기록이 '알림 못 보냄'이면 그 문구 — 그 뒤 다른 처리가 있으면 ''(지난 일)."""
    with lv._conn() as conn:
        r = conn.execute('SELECT action, detail FROM sat_off_log WHERE month=? ORDER BY id DESC LIMIT 1',
                         (month,)).fetchone()
    return r['detail'] if r and r['action'] == NOTIFY_FAIL else ''


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
    """kind: submit·approve·return·change_request·change_approve·change_reject·worked·worked_undo.
    실패해도 화면 처리는 그대로."""
    def run():
        client = _client()
        if not client:
            logger.error(f'{_LOG} 휴가 봇 토큰이 없어 알림을 못 보냄 ({kind} {month}) — 직접 알려 주세요',
                         extra={'force_alert': True})
            _note_fail('휴가 봇 연결 없음')
            return
        t = _month_title(month)
        row = month_row(month)
        if kind == 'submit':
            text = (f'📅 *{t} 스케줄 근무표* 결재 요청이 왔어요. ({actor} 님이 올림)\n'
                    f'토요일에 쉬는 사람\n{schedule_text(month)}\n두 분 중 한 분이 승인하면 확정돼요.')
            _dm_some(client, approver_ids(), text, month, f'{t} 스케줄 근무표 결재 요청')
        elif kind in ('approve', 'return'):
            sub = _who_sid(row.get('submitted_by') or '')
            to = [u for u in dict.fromkeys(((sub,) if sub else ()) + coordinator_ids())]
            if kind == 'approve':
                _dm_some(client, to, f'✅ {t} 스케줄 근무표를 {actor} 님이 승인했어요. 이제 확정이에요.', month,
                         f'{t} 스케줄 근무표 승인')
                _post(client, f'📅 {t} 스케줄 근무표가 확정됐어요. 토요일에 쉬는 사람이에요.\n{schedule_text(month)}')
                bad = affected_leaves(month)
                if bad:
                    lines = '\n'.join(f"• {b['name']} {b['what']}: {lv.fmt_days(b['saved'])} → {lv.fmt_days(b['now'])}"
                                      for b in bad)
                    _dm_some(client, lv.manager_ids(), f'ℹ️ {t} 스케줄 근무표가 확정돼 휴가 일수가 달라지는 신청이 있어요. '
                                                       f'휴가 관리에서 일수를 고쳐 주세요.\n{lines}', month,
                             f'{t} 휴가 일수 고칠 신청 {len(bad)}건')
            else:
                _dm_some(client, to, f'↩️ {t} 스케줄 근무표를 {actor} 님이 돌려보냈어요.\n사유: {note}\n'
                                     '고쳐서 다시 올려 주세요.', month, f'{t} 스케줄 근무표 돌려보냄')
        elif kind == 'change_request' and change:
            text = (f"🔁 *{change['name']} 님 스케줄 근무 변경 신청* ({t})\n쉬는 토요일: {change['what']}\n"
                    f"사유: {change.get('reason') or ''}\n두 분 중 한 분이 승인하면 바뀌어요.")
            _dm_some(client, approver_ids(), text, month, f"{change['name']} 님 스케줄 근무 변경 신청")
        elif kind in ('change_approve', 'change_reject') and change:
            who = change.get('slack_id') or _who_sid(change['name'])
            if kind == 'change_approve':
                _dm_some(client, [who], f"✅ 스케줄 근무 변경을 {actor} 님이 승인했어요. 쉬는 토요일: {change['what']}", month,
                         f"{change['name']} 님 스케줄 근무 변경 승인")
                _post(client, f"🔁 스케줄 근무 변경: {change['name']} 쉬는 토요일 {change['what']}")
                if affected_leaves(month):
                    _dm_some(client, lv.manager_ids(), f"ℹ️ {change['name']} 님 스케줄 근무가 바뀌어 휴가 일수가 달라지는 신청이 있어요. "
                                                       '휴가 관리에서 확인해 주세요.', month,
                             f"{change['name']} 님 변경으로 휴가 일수 고칠 신청")
            else:
                _dm_some(client, [who], f"↩️ 스케줄 근무 변경 신청을 {actor} 님이 반려했어요. 쉬는 토요일: {change['what']}\n사유: {note}",
                         month, f"{change['name']} 님 스케줄 근무 변경 반려")
        elif kind == 'full_off' and change:
            # 전체 휴무 지정·풀기(11-25) — 확정된 달이라 휴가 일수가 달라질 때만 경영지원에
            if affected_leaves(month):
                _dm_some(client, lv.manager_ids(), f"ℹ️ {change['label']} 전체 휴무 {change['how']}로 휴가 일수가 달라지는 신청이 있어요. "
                                                   '휴가 관리에서 확인해 주세요.', month, f"{change['label']} 전체 휴무로 휴가 일수 고칠 신청")
        elif kind in ('worked', 'worked_undo') and change:
            # 휴무 출근은 구두 승인이라 승인권자 DM·채널 글은 없음. 휴가 일수만 달라질 수 있어 그때만 경영지원에
            if affected_leaves(month):
                how = '휴무 출근' if kind == 'worked' else '휴무 출근 되돌림'
                _dm_some(client, lv.manager_ids(), f"ℹ️ {change['name']} 님 {change['label']} {how}으로 휴가 일수가 달라지는 "
                                                   '신청이 있어요. 휴가 관리에서 확인해 주세요.', month,
                         f"{change['name']} 님 {how}으로 휴가 일수 고칠 신청")
    _run_bg(lambda: _safe(run, kind, month))


def _safe(fn, kind: str, month: str) -> None:
    _fails.items = []
    try:
        fn()
    except Exception as e:
        logger.error(f'{_LOG} 알림 실패 ({kind} {month}): {e}', exc_info=True, extra={'force_alert': True})
        _fails.items.append('알림 처리 오류')
    finally:
        items, _fails.items = _fails.items, None
    if items:
        _record_fail(month, items)
