"""그룹웨어 알림 🔔 (2026-11-10, 당근 비즈니스 위쪽 「알림」 빨간 점처럼).

새로 알림을 '보내는' 게 아니라 이미 쌓이는 기록에서 '나와 관련된 최근 일'을 골라 보여 준다
→ 슬랙 DM 과 어긋나지 않고, 슬랙 코드는 건드리지 않는다.
- 결재: approval_steps(요청자에게 진행 소식) + 지금 내 차례인 결재
- 휴가: leave_steps(신청자에게 승인·반려) + 지금 내 승인 차례인 신청
- 연말정산: 시작·보완 요청·확인 완료·결과(본인) / 제출(경영지원·대표)
읽음은 사람마다 '마지막으로 알림 창을 연 시각' 하나만 저장(gw_notify_seen) → 그 뒤 일이면 빨간 점.
"""
from __future__ import annotations

import time
from contextlib import contextmanager
from datetime import datetime, timedelta
from typing import Dict, List, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

DAYS = 30
LIMIT = 60
_CACHE_SECONDS = 20
_cache: Dict[str, tuple] = {}

CATS = {'approval': '결재', 'leave': '휴가', 'hr': '인사'}


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


_ready = False


@contextmanager
def _conn():
    global _ready
    with _db()._get_connection() as conn:
        if not _ready:
            conn.execute('CREATE TABLE IF NOT EXISTS gw_notify_seen (who TEXT PRIMARY KEY, seen_at TEXT)')
            conn.commit()
            _ready = True
        yield conn
        conn.commit()


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _ev(cat: str, at: str, title: str, body: str = '', href: str = '', icon: str = '', key: str = '',
        actor_id: str = '') -> dict:
    return {'cat': cat, 'cat_label': CATS[cat], 'at': (at or '')[:19], 'title': title, 'body': body, 'href': href,
            'icon': icon, 'key': key, 'actor_id': actor_id}


# ── 결재 ──────────────────────────────────────────────────────────────
_AP_TEXT = {
    ('approve', 'approved'): ('✅', '{actor}님이 승인했어요'),
    ('final', 'approved'): ('✅', '{actor}님이 최종 승인했어요'),
    ('approve', 'rejected'): ('❌', '{actor}님이 반려했어요'),
    ('final', 'rejected'): ('❌', '{actor}님이 반려했어요'),
    ('transfer', 'done'): ('💸', '{w} 완료됐어요'),
    ('revise', 'revise'): ('✏️', '{actor}님이 수정을 요청했어요'),
    ('cancel', 'cancelled'): ('⚪', '{actor}님이 요청을 취소했어요'),
}


def _approval_events(sid: str, since: str) -> List[dict]:
    if not sid:
        return []
    from dashboard.services import e_approval as ea
    from dashboard.services.gw_home import waiting_on
    out = []
    with ea._db()._get_connection() as conn:
        ea._ensure_tables(conn)
        rows = {r['id']: dict(r) for r in conn.execute(
            'SELECT * FROM approvals WHERE updated_at >= ? OR status IN (%s)' % ','.join('?' * len(ea.OPEN_STATUSES)),
            (since, *ea.OPEN_STATUSES)).fetchall()}
        steps = [dict(s) for s in conn.execute(
            'SELECT * FROM approval_steps WHERE at >= ? ORDER BY id', (since,)).fetchall()]
    # 내 결재 진행 소식(다른 사람이 한 것만)
    for s in steps:
        row = rows.get(s['approval_id'])
        if not row or row.get('requester_id') != sid or s.get('actor_id') == sid or s['result'] == 'auto':
            continue
        spec = _AP_TEXT.get((s['step'], s['result']))
        if not spec:
            continue
        icon, text = spec
        body = (s.get('reason') or '') if s['result'] in ('rejected', 'revise') else ea.what_of(row)
        out.append(_ev('approval', s['at'], f"{ea.title(row)} — {text.format(actor=s.get('actor_name') or '', w=ea.step2_word(row))}",
                       body, f"/approval?box=mine&id={row['id']}", icon, f"ap{row['id']}s{s['id']}", s.get('actor_id') or ''))
    # 지금 내 차례(차례가 된 시각 = 마지막 단계 시각)
    last_at: Dict[int, str] = {}
    for s in steps:
        last_at[s['approval_id']] = s['at']
    for row in rows.values():
        if row['status'] not in ea.OPEN_STATUSES or not waiting_on(ea, row, sid):
            continue
        what = {'pending': '승인', 'final': '최종 승인', 'approved': f'{ea.step2_word(row)} 처리', 'revise': '수정'}.get(row['status'], '처리')
        amt = f" · {int(row.get('amount') or 0):,}원" if row.get('amount') else ''
        out.append(_ev('approval', last_at.get(row['id']) or row.get('updated_at') or row.get('created_at'),
                       f"{ea.title(row)} {what} 차례예요", f"{row.get('requester_name') or ''} · {ea.what_of(row)}{amt}",
                       f"/approval?box=todo&id={row['id']}", '🔔', f"ap{row['id']}turn{row['status']}"))
    return out


# ── 휴가 ──────────────────────────────────────────────────────────────
def _leave_events(sid: str, since: str) -> List[dict]:
    if not sid:
        return []
    from dashboard.services import leave as lv
    from dashboard.services.gw_home import _leave_link
    out = []
    with lv._conn() as conn:
        rows = {r['id']: dict(r) for r in conn.execute(
            'SELECT * FROM leave_requests WHERE updated_at >= ? OR status IN (%s)' % ','.join('?' * len(lv.OPEN)),
            (since, *lv.OPEN)).fetchall()}
        steps = [dict(s) for s in conn.execute('SELECT * FROM leave_steps WHERE at >= ? ORDER BY id', (since,)).fetchall()]
    for s in steps:
        row = rows.get(s['request_id'])
        if not row or sid not in (row.get('slack_id'), row.get('submitter_id')) or s.get('actor_id') in (sid, '', None):
            continue
        txt = {'approved': ('✅', '승인했어요'), 'rejected': ('❌', '반려했어요'), 'cancelled': ('⚪', '취소했어요')}.get(s['result'])
        if not txt:
            continue
        period = lv.period_text(row)
        out.append(_ev('leave', s['at'], f"{s.get('actor_name') or ''}님이 휴가를 {txt[1]}",
                       f"{lv.type_label(row['ltype'], row.get('sub') or '')} · {period}" + (f" — {s['reason']}" if s.get('reason') else ''),
                       '/leave', txt[0], f"lv{row['id']}s{s['id']}", s.get('actor_id') or ''))
    for row in rows.values():
        if row['status'] not in lv.OPEN or lv.current_approver(row) != sid:
            continue
        out.append(_ev('leave', row.get('updated_at') or row.get('created_at'), f"{row['name']}님 휴가 승인 차례예요",
                       f"{lv.type_label(row['ltype'], row.get('sub') or '')} · {lv.period_text(row)}",
                       _leave_link(row, sid) or '/leave/team', '🔔', f"lv{row['id']}turn{row['status']}"))
    return out


# ── 연말정산 ──────────────────────────────────────────────────────────
def _yearend_events(person: Optional[dict], hr_admin: bool, since: str) -> List[dict]:
    from dashboard.services import hr
    from dashboard.services import yearend as ye
    out = []
    year = ye.current_year()
    season = ye.season(year)
    if season.get('status') == 'open' and (season.get('updated_at') or '') >= since and person:
        out.append(_ev('hr', season.get('updated_at'), f'{year}년 귀속 연말정산이 시작됐어요',
                       '홈택스 자료 제공 동의 또는 간소화 PDF 업로드', '/yearend', '🧾', f'ye{year}open'))
    if person:
        sub = ye._sub(year, person['id'])
        if sub.get('status') == ye.RETURNED and (sub.get('updated_at') or '') >= since:
            out.append(_ev('hr', sub['updated_at'], '연말정산 보완 요청이 왔어요', sub.get('returned_reason') or '',
                           '/yearend', '✏️', f'ye{year}ret{sub["updated_at"]}'))
        if sub.get('status') == ye.DONE and (sub.get('done_at') or '') >= since:
            out.append(_ev('hr', sub['done_at'], '연말정산 자료 확인이 끝났어요', '', '/yearend', '✅', f'ye{year}done'))
        for d in hr.documents(person['id'], tag=ye.tag(year)):
            if d.get('doc_type') == ye.KIND_TYPES['result'] and (d.get('created_at') or '') >= since:
                out.append(_ev('hr', d['created_at'], '연말정산 결과가 올라왔어요', d.get('file_name') or '',
                               '/yearend', '📄', f"ye{year}res{d['id']}"))
    if hr_admin and season.get('status') in ('open', 'closed'):
        for r in ye.admin_rows(year):
            if r['status'] == ye.SUBMITTED and (r.get('submitted_at') or '') >= since:
                out.append(_ev('hr', r['submitted_at'], f"{r['name']}님이 연말정산 자료를 냈어요", '확인 완료 또는 보완 요청',
                               '/yearend/admin', '📥', f"ye{year}sub{r['person_id']}{r['submitted_at']}"))
    return out


# ── 모으기·읽음 ───────────────────────────────────────────────────────
def who_key(email: str, person: Optional[dict], field: bool) -> str:
    return f"field:{(person or {}).get('id')}" if field else (email or '').lower()


def seen_at(who: str) -> str:
    with _conn() as conn:
        r = conn.execute('SELECT seen_at FROM gw_notify_seen WHERE who = ?', (who,)).fetchone()
    return r[0] if r and r[0] else ''


def mark_seen(who: str) -> str:
    now = _now()
    with _conn() as conn:
        conn.execute('INSERT INTO gw_notify_seen (who, seen_at) VALUES (?, ?) ON CONFLICT(who) DO UPDATE SET seen_at=excluded.seen_at',
                     (who, now))
    _cache.pop(who, None)
    return now


def feed(email: str, person: Optional[dict], field: bool, hr_admin: bool, days: int = DAYS) -> dict:
    who = who_key(email, person, field)
    hit = _cache.get(who)
    if hit and time.time() - hit[0] < _CACHE_SECONDS:
        return hit[1]
    since = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d %H:%M:%S')
    sid = ((person or {}).get('slack_id') or '').strip()
    items: List[dict] = []
    for name, fn in (('approval', lambda: [] if field else _approval_events(sid, since)),
                     ('leave', lambda: _leave_events(sid, since)),
                     ('yearend', lambda: _yearend_events(person, hr_admin and not field, since))):
        try:
            items += fn()
        except Exception as e:   # 한 갈래가 실패해도 나머지는 보여 줌
            logger.warning(f'[GW알림] {name} 실패: {type(e).__name__}: {e}')
    items = [i for i in items if not (sid and i.get('actor_id') == sid)]
    items.sort(key=lambda i: i['at'], reverse=True)
    items = items[:LIMIT]
    seen = seen_at(who)
    for i in items:
        i['unread'] = bool(i['at'] > seen) if seen else True
        i.pop('actor_id', None)
    out = {'items': items, 'unread': sum(1 for i in items if i['unread']), 'seen_at': seen}
    _cache[who] = (time.time(), out)
    return out
