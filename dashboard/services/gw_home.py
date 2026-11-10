"""그룹웨어 홈 (2026-11-10) — 결재·휴가·휴가자·메일을 한 화면에.

- 결재: 슬랙 전자결재 장부(services/e_approval.py)에서 '지금 내가 처리할 것'과 '내가 올린 진행 중'만. 처리는 슬랙 카드에서.
- 휴가: services/leave_view.py 그대로.
- 메일: 기존 도메인 위임 서비스 계정으로 **로그인한 본인 주소만** 읽는다(gmail.readonly, 제목·보낸 사람·미리보기만, 저장 안 함).
  회사 도메인 메일만. 60초 동안 기억해 화면을 다시 열어도 구글을 매번 부르지 않는다.
"""
from __future__ import annotations

import os
import threading
import time
from datetime import date, timedelta
from email.utils import parseaddr, parsedate_to_datetime
from typing import Dict, List, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

COMPANY_DOMAIN = '@itg-aircon.com'


# ── 결재 ──────────────────────────────────────────────────────────────
def _card_link(row: dict) -> str:
    ch, ts = row.get('card_channel'), row.get('card_ts')
    if not (ch and ts):
        return ''
    ws = os.getenv('SLACK_WORKSPACE_DOMAIN', 'w1679118856-ybv966043').strip()
    return f'https://{ws}.slack.com/archives/{ch}/p{ts.replace(".", "")}'


def _line(ea, row: dict) -> dict:
    emoji, label = ea.status_label(row)
    return {
        'id': row['id'], 'title': ea.title(row), 'what': ea.what_of(row),
        'amount': int(row.get('amount') or 0), 'status': label,
        'requester': row.get('requester_name') or '', 'created_at': (row.get('created_at') or '')[:16],
        'link': _card_link(row),
    }


def waiting_on(ea, row: dict, uid: str) -> bool:
    """이 결재가 지금 이 사람 차례인가."""
    st = row.get('status')
    if st == ea.PENDING:
        return uid in ea.approver_ids() and uid != row.get('requester_id')
    if st == ea.FINAL:
        return uid in ea.final_approver_ids()
    if st == ea.APPROVED:
        return uid in ea.executor_dm_ids(row)
    if st == ea.REVISE:
        return uid == row.get('requester_id')
    return False


def approvals(slack_id: str, limit: int = 8) -> dict:
    if not slack_id:
        return {'linked': False, 'todo': [], 'mine': [], 'todo_count': 0, 'mine_count': 0}
    from dashboard.services import e_approval as ea
    with ea._db()._get_connection() as conn:
        ea._ensure_tables(conn)
        ph = ','.join('?' * len(ea.OPEN_STATUSES))
        rows = [dict(r) for r in conn.execute(
            f'SELECT * FROM approvals WHERE status IN ({ph}) ORDER BY id DESC', ea.OPEN_STATUSES).fetchall()]
    todo = [r for r in rows if waiting_on(ea, r, slack_id)]
    mine = [r for r in rows if r.get('requester_id') == slack_id]
    return {
        'linked': True,
        'todo': [_line(ea, r) for r in todo[:limit]], 'todo_count': len(todo),
        'mine': [_line(ea, r) for r in mine[:limit]], 'mine_count': len(mine),
    }


# ── 휴가자 ────────────────────────────────────────────────────────────
def people_off(today: Optional[date] = None, days: int = 7) -> dict:
    """오늘 휴가자 + 앞으로 며칠(승인만, 종류·사유 없이 이름·반차만)."""
    from dashboard.services import leave_view
    today = today or date.today()
    end = today + timedelta(days=days)
    entries = leave_view.calendar(today.year, today.month)
    if end.month != today.month:
        entries += leave_view.calendar(end.year, end.month)
    t, e = today.isoformat(), end.isoformat()
    seen, todays, soon = set(), [], []
    for x in entries:
        if x['status'] != leave_view.ST_OK:
            continue
        key = (x['source'], x['id'])
        if key in seen:
            continue
        seen.add(key)
        item = {'name': x['name'], 'start': x['start'], 'end': x['end'], 'half': x['half']}
        if x['start'] <= t <= x['end']:
            todays.append(item)
        elif t < x['start'] <= e:
            soon.append(item)
    soon.sort(key=lambda i: (i['start'], i['name']))
    return {'today': todays, 'soon': soon}


# ── 메일 ──────────────────────────────────────────────────────────────
_MAIL_TTL = 60
_mail_cache: Dict[str, tuple] = {}
_mail_lock = threading.Lock()


def _gmail(email: str):
    """도메인 위임으로 그 사람 메일함(읽기 전용). 호출마다 새로 만든다(googleapiclient 는 스레드 공유 금지)."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds_file = (os.getenv('GOOGLE_CREDENTIALS_FILE') or os.getenv('GOOGLE_APPLICATION_CREDENTIALS')
                  or 'credentials.json')
    creds = service_account.Credentials.from_service_account_file(
        creds_file, scopes=['https://www.googleapis.com/auth/gmail.readonly'])
    try:   # 구글이 늦을 때 홈이 오래 멈추지 않게 8초 제한
        import google_auth_httplib2
        import httplib2
        http = google_auth_httplib2.AuthorizedHttp(creds.with_subject(email), http=httplib2.Http(timeout=8))
        return build('gmail', 'v1', http=http, cache_discovery=False)
    except ImportError:
        return build('gmail', 'v1', credentials=creds.with_subject(email), cache_discovery=False)


def _when(raw: str) -> str:
    try:
        dt = parsedate_to_datetime(raw)
        dt = dt.astimezone()
        return dt.strftime('%m-%d %H:%M') if dt.date() != date.today() else dt.strftime('%H:%M')
    except Exception:
        return ''


def mail_summary(email: str, limit: int = 5, service=None) -> dict:
    """안 읽은 수 + 받은편지함 최근 몇 통. 회사 메일만, 실패하면 {'ok': False}."""
    email = (email or '').strip().lower()
    if not email.endswith(COMPANY_DOMAIN):
        return {'ok': False, 'reason': 'no_company_mail'}
    now = time.time()
    with _mail_lock:
        hit = _mail_cache.get(email)
        if hit and now - hit[0] < _MAIL_TTL:
            return hit[1]
    try:
        svc = service or _gmail(email)
        label = svc.users().labels().get(userId='me', id='INBOX').execute()
        lst = svc.users().messages().list(userId='me', q='in:inbox', maxResults=limit).execute()
        items = []
        for m in lst.get('messages') or []:
            msg = svc.users().messages().get(userId='me', id=m['id'], format='metadata',
                                             metadataHeaders=['From', 'Subject', 'Date']).execute()
            h = {x['name'].lower(): x['value'] for x in (msg.get('payload') or {}).get('headers') or []}
            name, addr = parseaddr(h.get('from', ''))
            items.append({
                'from': name or addr, 'subject': h.get('subject') or '(제목 없음)',
                'snippet': (msg.get('snippet') or '')[:120], 'when': _when(h.get('date', '')),
                'unread': 'UNREAD' in (msg.get('labelIds') or []),
                'link': f'https://mail.google.com/mail/u/?authuser={email}#inbox/{msg.get("threadId") or m["id"]}',
            })
        out = {'ok': True, 'unread': int(label.get('messagesUnread') or 0), 'items': items,
               'inbox': f'https://mail.google.com/mail/u/?authuser={email}#inbox'}
    except Exception as e:
        logger.warning(f'[GW-HOME] 메일 불러오기 실패({email}): {type(e).__name__}: {e}')
        out = {'ok': False, 'reason': 'error', 'inbox': f'https://mail.google.com/mail/u/?authuser={email}#inbox'}
    with _mail_lock:
        _mail_cache[email] = (now, out)
    return out
