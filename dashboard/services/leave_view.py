"""그룹웨어 휴가 화면용 조회 (2026-11-10).

휴가 신청·승인·연차 계산은 슬랙 휴가 봇 장부(services/leave.py) 그대로 쓰고, 여기서는 '보기'만 만든다.
플렉스 해지 뒤에도 지난 휴가를 볼 수 있게 플렉스 휴가 사용 내역(488건)을 gw_leave_history 에 보관한다(읽기 전용).

열람 범위: 본인=자기 휴가 / 휴가 현황 열람자·인사카드 관리자(hr.can_view_all_leave)=전 직원.
"""
from __future__ import annotations

from contextlib import contextmanager
from datetime import date, timedelta
from typing import Dict, List, Optional

from dashboard.services import hr
from dashboard.services import leave as lv
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_DDL = [
    '''CREATE TABLE IF NOT EXISTS gw_leave_history (
        flex_event_id TEXT PRIMARY KEY,
        person_id INTEGER,
        name TEXT NOT NULL,
        type_label TEXT,
        start TEXT NOT NULL,
        "end" TEXT,
        days REAL,
        unit TEXT,
        status TEXT,
        reason TEXT,
        approval TEXT,
        comment TEXT,
        requested_at TEXT
    )''',
    'CREATE INDEX IF NOT EXISTS idx_gw_leave_hist_person ON gw_leave_history(person_id)',
    'CREATE INDEX IF NOT EXISTS idx_gw_leave_hist_start ON gw_leave_history(start)',
]
_table_ready = False

# 화면에 쓰는 상태 네 가지
ST_OK, ST_WAIT, ST_REJECT, ST_CANCEL = '승인', '대기', '반려', '취소'
_FLEX_STATUS = {'승인': ST_OK, '등록': ST_OK, '승인 대기': ST_WAIT, '반려': ST_REJECT, '취소': ST_CANCEL, '취소 승인': ST_CANCEL}
_SLACK_STATUS = {lv.APPROVED: ST_OK, lv.WAIT_1: ST_WAIT, lv.WAIT_2: ST_WAIT, lv.REJECTED: ST_REJECT, lv.CANCELLED: ST_CANCEL}
_FLEX_HALF = {'HALF_DAY_AM': '오전', 'HALF_DAY_PM': '오후'}


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


# ── 플렉스 기록 보관 ───────────────────────────────────────────────────
def upsert_flex_history(rows: List[Dict]) -> int:
    """플렉스 휴가 사용 내역을 넣거나 고친다(같은 건 번호는 한 줄). 넣거나 고친 줄 수."""
    n = 0
    with _conn() as conn:
        for r in rows:
            cur = conn.execute(
                'INSERT INTO gw_leave_history (flex_event_id, person_id, name, type_label, start, "end", days, unit, '
                'status, reason, approval, comment, requested_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?) '
                'ON CONFLICT(flex_event_id) DO UPDATE SET person_id=excluded.person_id, name=excluded.name, '
                'type_label=excluded.type_label, start=excluded.start, "end"=excluded."end", days=excluded.days, '
                'unit=excluded.unit, status=excluded.status, reason=excluded.reason, approval=excluded.approval, '
                'comment=excluded.comment, requested_at=excluded.requested_at',
                (r['flex_event_id'], r.get('person_id'), r['name'], r.get('type_label'), r['start'], r.get('end'),
                 r.get('days'), r.get('unit'), r.get('status'), r.get('reason'), r.get('approval'), r.get('comment'),
                 r.get('requested_at')))
            n += cur.rowcount
    return n


def _flex_rows(person_id: Optional[int] = None, start_from: str = '', end_to: str = '') -> List[dict]:
    sql, args = 'SELECT * FROM gw_leave_history WHERE 1=1', []
    if person_id is not None:
        sql += ' AND person_id = ?'
        args.append(person_id)
    if end_to:
        sql += ' AND start <= ?'
        args.append(end_to)
    if start_from:
        sql += ' AND COALESCE("end", start) >= ?'
        args.append(start_from)
    with _conn() as conn:
        return [dict(r) for r in conn.execute(sql + ' ORDER BY start DESC', args).fetchall()]


# ── 사람 ↔ 휴가 장부 이름 ──────────────────────────────────────────────
def leave_name(person: dict) -> Optional[str]:
    """인사카드 → 휴가 장부(leave_staff) 이름. 슬랙 ID 우선, 없으면 같은 이름. 퇴사자는 장부 연결 안 함."""
    if not person or person.get('status') == hr.STATUS_RESIGNED:
        return None
    if person.get('slack_id'):
        s = lv.staff_by_slack(person['slack_id'])
        if s:
            return s['name']
    s = lv.staff_by_name(person['name'])
    return s['name'] if s and s.get('active') else None


# ── 한 줄로 맞추기 ─────────────────────────────────────────────────────
def _entry_from_slack(r: dict) -> dict:
    half = {lv.HALF_AM: '오전', lv.HALF_PM: '오후'}.get(r['ltype'], '')
    return {
        'source': 'slack', 'id': r['id'], 'name': r['name'],
        'type': lv.type_label(r['ltype'], r.get('sub') or ''),
        'annual': lv.deducts(r['ltype']), 'sick': r['ltype'] == lv.SICK,
        'start': r['start'], 'end': r.get('end') or r['start'], 'half': half,
        'days': float(r.get('days') or 0), 'status': _SLACK_STATUS.get(r['status'], r['status']),
        'reason': r.get('reason') or '', 'requested_at': (r.get('created_at') or '')[:16],
    }


def _entry_from_flex(r: dict) -> dict:
    label = r.get('type_label') or ''
    unit = r.get('unit') or ''
    return {
        'source': 'flex', 'id': r['flex_event_id'], 'name': r['name'],
        'type': label, 'annual': label == '연차', 'sick': label == '병가',
        'start': r['start'], 'end': r.get('end') or r['start'], 'half': _FLEX_HALF.get(unit, ''),
        'days': float(r.get('days') or 0), 'status': _FLEX_STATUS.get(r.get('status') or '', r.get('status') or ''),
        'reason': r.get('reason') or '', 'requested_at': (r.get('requested_at') or '')[:16],
    }


def entries_for(person: dict) -> List[dict]:
    """한 사람의 휴가 기록 — 슬랙 신청 + 플렉스 보관, 최근 것부터."""
    out = [_entry_from_flex(r) for r in _flex_rows(person_id=person['id'])]
    name = leave_name(person)
    if name:
        out += [_entry_from_slack(r) for r in lv.requests_of(name, limit=500)]
    out.sort(key=lambda e: (e['start'], e['requested_at']), reverse=True)
    return out


def _year_sum(entries: List[dict], year: int, kind: str) -> float:
    y = str(year)
    return round(sum(e['days'] for e in entries if e['status'] == ST_OK and e[kind] and e['start'][:4] == y), 2)


def summary(person: dict, today: Optional[date] = None) -> dict:
    """내 휴가 화면 — 남은 연차·소멸 예정·올해 사용·기록."""
    today = today or date.today()
    name = leave_name(person)
    entries = entries_for(person)
    live = []
    left = None
    if name:
        left = lv.balance(name, today)
        for g in lv.grants(name):
            rest = round(g['days'] - g['used'], 2)
            if g['valid_from'] <= today.isoformat() <= g['valid_to'] and rest > 0:
                live.append({'days': rest, 'valid_to': g['valid_to']})
    upcoming = [e for e in entries if e['status'] in (ST_OK, ST_WAIT) and e['end'] >= today.isoformat()]
    upcoming.sort(key=lambda e: e['start'])
    return {
        'person_id': person['id'], 'name': person['name'], 'linked': bool(name),
        'balance': left, 'grants': live,
        'annual_used': _year_sum(entries, today.year, 'annual'),
        'sick_used': _year_sum(entries, today.year, 'sick'), 'sick_limit': lv.SICK_DAYS_PER_YEAR,
        'pending': sum(1 for e in entries if e['status'] == ST_WAIT),
        'upcoming': upcoming[:5], 'entries': entries, 'year': today.year,
    }


def team(today: Optional[date] = None) -> List[dict]:
    """전사 현황 — 재직자 한 줄씩."""
    today = today or date.today()
    rows = []
    for p in hr.list_people(include_resigned=False):
        full = hr.get_person(p['id'])
        s = summary(full, today)
        nxt = s['upcoming'][0] if s['upcoming'] else None
        rows.append({
            'person_id': p['id'], 'name': p['name'], 'department': p.get('department') or '',
            'hire_date': p.get('hire_date') or '', 'access': p.get('access'), 'job_role': p.get('job_role') or '',
            'linked': s['linked'], 'balance': s['balance'],
            'annual_used': s['annual_used'], 'sick_used': s['sick_used'], 'pending': s['pending'],
            'next': {'start': nxt['start'], 'end': nxt['end'], 'half': nxt['half'], 'type': nxt['type'],
                     'status': nxt['status']} if nxt else None,
        })
    return rows


def calendar(year: int, month: int) -> List[dict]:
    """그달에 걸친 휴가(승인·대기) — 전사 달력용. 사유는 넣지 않는다."""
    first = date(year, month, 1)
    last = (date(year + (month == 12), (month % 12) + 1, 1) - timedelta(days=1))
    a, b = first.isoformat(), last.isoformat()
    out = []
    for r in _flex_rows(start_from=a, end_to=b):
        e = _entry_from_flex(r)
        if e['status'] in (ST_OK, ST_WAIT):
            out.append(e)
    names = {s['name'] for s in lv.staff(active_only=True)}
    with lv._conn() as conn:
        rows = conn.execute('SELECT * FROM leave_requests WHERE start <= ? AND "end" >= ? ORDER BY start',
                            (b, a)).fetchall()
    for r in rows:
        e = _entry_from_slack(dict(r))
        if e['status'] in (ST_OK, ST_WAIT) and e['name'] in names:
            out.append(e)
    for e in out:
        e.pop('reason', None)
    out.sort(key=lambda e: (e['start'], e['name']))
    return out


def off_on(day: date) -> List[dict]:
    """그날 휴가자(승인만) — 홈 화면용."""
    d = day.isoformat()
    return [e for e in calendar(day.year, day.month) if e['status'] == ST_OK and e['start'] <= d <= e['end']]
