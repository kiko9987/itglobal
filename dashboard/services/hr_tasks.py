# -*- coding: utf-8 -*-
"""입사·퇴사 할 일 (2026-11-20, 플렉스 「할 일」 대체) — 인사카드에 입사일·퇴사일이 들어오면 경영지원 할 일을 만든다.

플렉스는 입사·퇴직 7일 뒤 '4대보험 취득·상실 신고', '퇴직자 서류 전달'을 할 일로 만들었지만 플렉스 안에만 떠서
55건이 처리 표시 없이 쌓였다(11-20 점검). 여기서는 기한을 붙여 슬랙 DM(경영지원)·그룹웨어 홈 「지금 할 일」·인사카드에 띄우고
[완료]로 닫는다. 기한(근거는 원문 확인):
  - 4대보험 취득·상실 신고: 건강보험이 가장 빠름 — 취득일·상실일(퇴사 다음 날)부터 14일 안(국민건강보험법 제8조·제10조).
    연금·고용·산재는 각 공단 기한. 4대사회보험 정보연계센터에서 한 번에 신고할 수 있다.
  - 퇴직 정산: 퇴직 뒤 14일 안 금품 청산(근로기준법 제36조).
HR_TASKS_FROM(기본 2026-11-20) 이전 입사·퇴사는 만들지 않는다 — 플렉스 때 일은 경영지원이 이미 처리했다.
"""
import json
import os
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional

OPEN, DONE, CANCELLED = 'open', 'done', 'cancelled'
A_DONE = 'hr_task_done'

# kind → (언제 기준, 제목, 기한 일수, 설명)
KINDS = {
    'hire_docs': ('hire', '입사 서류 받기', 7,
                  '근로계약서(서면으로 써서 한 부 주기)·주민등록등본·통장 사본을 받아 인사카드 서류에 올려 주세요.'),
    'ins_acquire': ('hire', '4대보험 취득 신고', 14,
                    '건강보험은 입사일부터 14일 안에 신고해야 해요. 4대사회보험 정보연계센터(4insure.or.kr)에서 한 번에 할 수 있어요.'),
    'resign_settle': ('resign', '퇴직 정산 (급여·연차·퇴직금)', 14,
                      '퇴직 뒤 14일 안에 임금·퇴직금 등을 정산해야 해요(근로기준법 제36조). 남은 연차는 휴가 관리 [퇴직 정산]에서 확인하세요.'),
    'ins_loss': ('resign', '4대보험 상실 신고', 15,
                 '건강보험은 상실일(퇴사 다음 날)부터 14일 안에 신고해야 해요. 4대사회보험 정보연계센터(4insure.or.kr)에서 한 번에 할 수 있어요.'),
    'resign_docs': ('resign', '퇴직자 서류 전달', 14,
                    '경력증명서(인사카드 [증명서 발급])·원천징수영수증을 퇴직자 개인 메일로 보내 주세요.'),
}
ORDER = list(KINDS)

_DDL = ['''CREATE TABLE IF NOT EXISTS hr_tasks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    person_id INTEGER NOT NULL,
    name TEXT,
    kind TEXT NOT NULL,
    event_date TEXT NOT NULL,
    due_date TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'open',
    created_at TEXT NOT NULL,
    done_by TEXT,
    done_at TEXT,
    dm_refs TEXT,
    UNIQUE(person_id, kind, event_date)
)''']
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


def tasks_from() -> date:
    v = (os.getenv('HR_TASKS_FROM') or '').strip()
    try:
        return date.fromisoformat(v) if v else date(2026, 11, 20)
    except ValueError:
        return date(2026, 11, 20)


def _d(v) -> Optional[date]:
    try:
        return date.fromisoformat(str(v)[:10]) if v else None
    except ValueError:
        return None


def wanted(person: dict) -> List[tuple]:
    """이 사람에게 있어야 할 할 일 [(kind, 기준일, 기한)]."""
    out = []
    since = tasks_from()
    hire, resign = _d(person.get('hire_date')), _d(person.get('resign_date'))
    for kind, (base, _t, days, _d2) in KINDS.items():
        ev = hire if base == 'hire' else resign
        if ev and ev >= since:
            out.append((kind, ev, ev + timedelta(days=days)))
    return out


def sync(today: Optional[date] = None) -> List[dict]:
    """인사카드 전체를 보고 할 일을 맞춘다 → 새로 만든 할 일. 입사·퇴사일이 바뀌면 옛 날짜의 열린 일은 '취소'."""
    from dashboard.services import hr
    created = []
    with hr._conn() as hc:
        people = [dict(r) for r in hc.execute('SELECT id, name, hire_date, resign_date, status FROM hr_people').fetchall()]
    with _conn() as conn:
        for p in people:
            want = wanted(p)
            keys = {(k, ev.isoformat()) for k, ev, _ in want}
            for r in conn.execute('SELECT id, kind, event_date FROM hr_tasks WHERE person_id=? AND status=?',
                                  (p['id'], OPEN)).fetchall():
                if (r['kind'], r['event_date']) not in keys:
                    conn.execute('UPDATE hr_tasks SET status=?, done_at=? WHERE id=?', (CANCELLED, _now(), r['id']))
            for kind, ev, due in want:
                cur = conn.execute('INSERT OR IGNORE INTO hr_tasks (person_id, name, kind, event_date, due_date, status, '
                                   'created_at) VALUES (?,?,?,?,?,?,?)',
                                   (p['id'], p['name'], kind, ev.isoformat(), due.isoformat(), OPEN, _now()))
                if cur.rowcount:
                    created.append(dict(conn.execute('SELECT * FROM hr_tasks WHERE id=?', (cur.lastrowid,)).fetchone()))
    return created


def get(tid: int) -> Optional[dict]:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM hr_tasks WHERE id=?', (tid,)).fetchone()
        return dict(r) if r else None


def open_tasks(person_id: Optional[int] = None) -> List[dict]:
    sql, args = 'SELECT * FROM hr_tasks WHERE status=?', [OPEN]
    if person_id is not None:
        sql += ' AND person_id=?'
        args.append(person_id)
    with _conn() as conn:
        rows = [dict(r) for r in conn.execute(sql + ' ORDER BY due_date, id', args).fetchall()]
    return rows


def for_person(person_id: int) -> List[dict]:
    """인사카드용 — 열린 것 + 최근 닫은 것(취소 제외)."""
    with _conn() as conn:
        return [dict(r) for r in conn.execute(
            'SELECT * FROM hr_tasks WHERE person_id=? AND status != ? ORDER BY event_date DESC, id',
            (person_id, CANCELLED)).fetchall()]


def complete(tid: int, by: str) -> Optional[dict]:
    """열린 일만 닫는다(두 번 눌러도 한 번)."""
    with _conn() as conn:
        cur = conn.execute('UPDATE hr_tasks SET status=?, done_by=?, done_at=? WHERE id=? AND status=?',
                           (DONE, by, _now(), tid, OPEN))
        if cur.rowcount != 1:
            return None
        return dict(conn.execute('SELECT * FROM hr_tasks WHERE id=?', (tid,)).fetchone())


def add_ref(tid: int, ch: str, ts: str) -> None:
    with _conn() as conn:
        r = conn.execute('SELECT dm_refs FROM hr_tasks WHERE id=?', (tid,)).fetchone()
        refs = json.loads((r['dm_refs'] if r else None) or '[]')
        refs.append({'ch': ch, 'ts': ts})
        conn.execute('UPDATE hr_tasks SET dm_refs=? WHERE id=?', (json.dumps(refs), tid))


def refs(row: dict) -> List[dict]:
    return json.loads(row.get('dm_refs') or '[]')


def title(row: dict) -> str:
    return KINDS.get(row['kind'], ('', row['kind'], 0, ''))[1]


def due_label(row: dict, today: Optional[date] = None) -> str:
    today = today or date.today()
    due = _d(row['due_date'])
    if not due:
        return ''
    left = (due - today).days
    md = f'{due.month}/{due.day}'
    if left < 0:
        return f'기한 {md} · {-left}일 지남'
    if left == 0:
        return f'기한 {md} · 오늘까지'
    return f'기한 {md} · {left}일 남음'


def handler_ids() -> tuple:
    """DM 받는 사람 — 증명서·이체 DM 과 같은 경영지원(시험 모드도 따름)."""
    from dashboard.services import certificates as ct
    return ct.handler_ids()


# ── 슬랙 카드 ─────────────────────────────────────────────────────────
def _plain(t: str) -> dict:
    return {'type': 'plain_text', 'text': t, 'emoji': True}


def card_blocks(row: dict, today: Optional[date] = None) -> List[dict]:
    base = KINDS.get(row['kind'], ('', '', 0, ''))
    ev = _d(row['event_date'])
    when = f"{'입사' if base[0] == 'hire' else '퇴사'} {ev.month}/{ev.day}" if ev else ''
    head = f"📋 *{title(row)}* — {row.get('name') or ''} ({when})"
    if row['status'] == DONE:
        text = f"✅ ~{title(row)}~ — {row.get('name') or ''} ({when})\n완료 · {row.get('done_by') or ''}"
        return [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': text}}]
    text = f"{head}\n{due_label(row, today)}\n{base[3]}"
    return [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': text}},
            {'type': 'actions', 'elements': [
                {'type': 'button', 'action_id': A_DONE, 'value': str(row['id']), 'style': 'primary', 'text': _plain('✅ 완료')},
                {'type': 'button', 'action_id': 'hr_task_open', 'text': _plain('인사카드 열기'),
                 'url': f"{_base_url()}/hr/people#{row['person_id']}"}]}]


def _base_url() -> str:
    return ((os.getenv('GW_PUBLIC_URL') or '').strip() or 'https://pm.itg-aircon.com').rstrip('/')


def reminder_rows(today: Optional[date] = None, within: int = 3) -> List[dict]:
    """알림 대상 — 열린 일 중 기한이 within 일 안이거나 지난 것."""
    today = today or date.today()
    return [r for r in open_tasks() if (_d(r['due_date']) or today) <= today + timedelta(days=within)]


def reminder_text(rows: List[dict], today: Optional[date] = None) -> str:
    lines = [f"📋 *입사·퇴사 할 일 {len(rows)}건*이 기한이 다가왔거나 지났어요."]
    for r in rows:
        lines.append(f"• {r.get('name') or ''} — {title(r)} ({due_label(r, today)})")
    lines.append('처리했으면 할 일 카드의 [✅ 완료]나 그룹웨어 인사카드에서 닫아 주세요.')
    return '\n'.join(lines)
