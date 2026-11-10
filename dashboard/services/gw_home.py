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


# ── 할 일(당근 비즈니스 홈의 「지금 할 일」 카드처럼, 11-10) ─────────────
# 지금 이 사람이 해야 할 것만, 있을 때만. 각 줄: id(닫기 기억용)·sig(내용이 바뀌면 다시 보이게)·문구·단추.
def _leave_waiting(uid: str) -> List[dict]:
    """내 승인을 기다리는 휴가 신청(슬랙 휴가 앱) — 링크는 나에게 온 DM 카드."""
    if not uid:
        return []
    from dashboard.services import leave
    with leave._conn() as conn:
        ph = ','.join('?' * len(leave.OPEN))
        rows = [dict(r) for r in conn.execute(
            f'SELECT * FROM leave_requests WHERE status IN ({ph}) ORDER BY id', leave.OPEN).fetchall()]
    return [r for r in rows if leave.current_approver(r) == uid]


def _leave_link(row: dict, uid: str) -> str:
    from dashboard.services import leave
    for ref in reversed(leave.dm_refs(row)):
        if ref.get('user') == uid and ref.get('ch') and ref.get('ts'):
            return _card_link({'card_channel': ref['ch'], 'card_ts': ref['ts']})
    return ''


def todos(person: Optional[dict], slack_id: str, field: bool, hr_admin: bool,
          approvals_data: Optional[dict] = None) -> List[dict]:
    out: List[dict] = []
    # 1) 전자결재 — 내 차례
    if not field and approvals_data and approvals_data.get('todo_count'):
        n = approvals_data['todo_count']
        first = (approvals_data.get('todo') or [{}])[0]
        out.append({'id': 'approval', 'sig': str(n), 'tone': 'act', 'icon': 'fas fa-file-signature',
                    'title': f'결재 {n}건이 기다리고 있어요',
                    'desc': '슬랙 #전자_결재 카드에서 승인하거나 반려해 주세요.',
                    'action': '결재 카드 열기', 'href': first.get('link') or '', 'external': True})
    # 2) 휴가 승인 — 내 차례
    if not field and slack_id:
        try:
            rows = _leave_waiting(slack_id)
        except Exception as e:
            logger.warning(f'[GW-HOME] 휴가 승인 대기 실패: {e}')
            rows = []
        if rows:
            names = ', '.join(dict.fromkeys(r.get('name') or '' for r in rows if r.get('name')))
            out.append({'id': 'leave-approve', 'sig': ','.join(str(r['id']) for r in rows), 'tone': 'act',
                        'icon': 'fas fa-umbrella-beach', 'title': f'휴가 승인 {len(rows)}건이 기다리고 있어요',
                        'desc': f'{names} — 슬랙 휴가 앱 DM에서 승인해 주세요.',
                        'action': '승인 카드 열기', 'href': _leave_link(rows[0], slack_id), 'external': True})
    # 3) 연말정산 — 내 것
    try:
        from dashboard.services import yearend as ye
        year = ye.current_year()
        season = ye.season(year)
        if person and season.get('status') == 'open':
            sub = ye._sub(year, person['id'])
            st = sub.get('status') or ye.DRAFT
            due_d = _ymd(season.get('due_date'))
            due = f'마감 {due_d.month}월 {due_d.day}일' if due_d else ''
            if st == ye.RETURNED:
                out.append({'id': 'ye-mine', 'sig': f'{year}:{st}', 'tone': 'warn', 'icon': 'fas fa-receipt',
                            'title': '연말정산 자료 보완 요청이 왔어요',
                            'desc': (sub.get('returned_reason') or '경영지원 메모를 확인해 주세요.') + (f' · {due}' if due else ''),
                            'action': '보완하기', 'href': '/yearend'})
            elif st == ye.DRAFT:
                out.append({'id': 'ye-mine', 'sig': f'{year}:{st}', 'tone': 'act', 'icon': 'fas fa-receipt',
                            'title': f'{year}년 귀속 연말정산 자료를 내 주세요',
                            'desc': '홈택스 자료 제공 동의 또는 간소화 PDF 업로드' + (f' · {due}' if due else ''),
                            'action': '연말정산 하기', 'href': '/yearend'})
        # 4) 연말정산 — 경영지원·대표: 제출 확인 기다리는 사람
        if hr_admin and not field and season.get('status') in ('open', 'closed'):
            waiting = [r for r in ye.admin_rows(year) if r.get('status') == ye.SUBMITTED]
            if waiting:
                out.append({'id': 'ye-admin', 'sig': str(len(waiting)), 'tone': 'info', 'icon': 'fas fa-clipboard-check',
                            'title': f'연말정산 제출 {len(waiting)}명이 확인을 기다려요',
                            'desc': '자료를 보고 「확인 완료」 또는 「보완 요청」을 눌러 주세요.',
                            'action': '현황 보기', 'href': '/yearend/admin'})
    except Exception as e:
        logger.warning(f'[GW-HOME] 연말정산 할 일 실패: {e}')
    return out


# ── 영업 숫자(문의·방문·계약) — 당근 홈의 「최근 7일 통계」 칸처럼 누르면 그래프 ─────
# 집계 기준은 매일 목표 점검과 같다: 문의=리드 시트 '상담 시간' 날짜(전체 플랫폼),
# 방문=그 문의 중 방문을 잡은 것(상태가 방문 예약 이후), 계약=공사현황 '공사 확정' 날짜(공사 취소 제외).
VISIT_STATUSES = ('방문 예약', '방문 완료', '견적 요청', '견적 제출', '공사 확정', '공사 취소')
GOALS = {'inquiry': 20, 'visit': 10}   # 사장님 매일 목표(평일 문의 20~25·방문 10+)
_SALES_TTL = 300
_sales_cache: Dict[str, tuple] = {}

import re as _re  # noqa: E402

_DATE_RE = _re.compile(r'(\d{4})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})')


def _ymd(v) -> Optional[date]:
    if v is None or v == '':
        return None
    if hasattr(v, 'year') and hasattr(v, 'month') and hasattr(v, 'day'):
        try:
            return date(v.year, v.month, v.day)
        except (TypeError, ValueError):
            return None
    m = _DATE_RE.search(str(v))
    if not m:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def _amount(v) -> int:
    try:
        return int(float(_re.sub(r'[^\d.\-]', '', str(v)) or 0))
    except ValueError:
        return 0


def _sales_counts(start: date, end: date) -> Dict[date, dict]:
    from dashboard.services.lead_service import get_lead_records
    from dashboard.services.project_service import load_data
    days: Dict[date, dict] = {}

    def slot(d):
        return days.setdefault(d, {'inquiry': 0, 'visit': 0, 'contract': 0, 'amount': 0})

    for r in get_lead_records() or []:
        d = _ymd(r.get('상담 시간'))
        if not d or not (start <= d <= end):
            continue
        s = slot(d)
        s['inquiry'] += 1
        if str(r.get('상태') or '').strip() in VISIT_STATUSES:
            s['visit'] += 1
    df = load_data()
    if df is not None and len(df) and '공사 확정' in df.columns:
        cols = [c for c in ('공사 확정', '총액 2', '수금 관련 특이사항') if c in df.columns]
        for r in df[cols].to_dict('records'):
            d = _ymd(r.get('공사 확정'))
            if not d or not (start <= d <= end):
                continue
            if _re.search(r'공사\s*취소', str(r.get('수금 관련 특이사항') or '')):
                continue
            s = slot(d)
            s['contract'] += 1
            s['amount'] += _amount(r.get('총액 2'))
    return days


def _is_workday(d: date) -> bool:
    try:
        from dashboard.services import leave
        return leave.is_workday(d)
    except Exception:
        return d.weekday() < 5


def sales(today: Optional[date] = None, span: int = 30) -> dict:
    """이번 달(오늘까지)·지난달 같은 기간 합계, 평일 하루 평균, 최근 span일 날짜별."""
    today = today or date.today()
    key = today.isoformat()
    hit = _sales_cache.get(key)
    if hit and time.time() - hit[0] < _SALES_TTL:
        return hit[1]
    month_start = today.replace(day=1)
    prev_end_month = month_start - timedelta(days=1)
    prev_start = prev_end_month.replace(day=1)
    prev_end = prev_start + timedelta(days=min(today.day, prev_end_month.day) - 1)
    span_start = today - timedelta(days=span - 1)
    counts = _sales_counts(min(prev_start, span_start), today)

    def total(a: date, b: date) -> dict:
        t = {'inquiry': 0, 'visit': 0, 'contract': 0, 'amount': 0}
        d = a
        while d <= b:
            for k, v in counts.get(d, {}).items():
                t[k] += v
            d += timedelta(days=1)
        return t

    this, prev = total(month_start, today), total(prev_start, prev_end)
    workdays = [month_start + timedelta(days=i) for i in range((today - month_start).days + 1)]
    workdays = [d for d in workdays if _is_workday(d)]
    wd_total = {k: sum(counts.get(d, {}).get(k, 0) for d in workdays) for k in ('inquiry', 'visit')}
    out = {
        'month': month_start.strftime('%Y-%m'), 'today': key,
        'prev_label': f'{prev_start.month}/{prev_start.day}~{prev_end.month}/{prev_end.day}',
        'this': this, 'prev': prev, 'goals': GOALS,
        'workday_avg': {k: (round(v / len(workdays), 1) if workdays else None) for k, v in wd_total.items()},
        'days': [dict({'date': (span_start + timedelta(days=i)).isoformat(),
                       'workday': _is_workday(span_start + timedelta(days=i))},
                      **counts.get(span_start + timedelta(days=i), {'inquiry': 0, 'visit': 0, 'contract': 0, 'amount': 0}))
                 for i in range(span)],
    }
    _sales_cache.clear()
    _sales_cache[key] = (time.time(), out)
    return out


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


_MAIL_STALE = 15 * 60          # 이 시간 안이면 옛 값을 먼저 보여 주고 뒤에서 새로 받는다
_mail_refreshing: set = set()


def _item(email: str, m: dict, msg: dict) -> dict:
    h = {x['name'].lower(): x['value'] for x in (msg.get('payload') or {}).get('headers') or []}
    name, addr = parseaddr(h.get('from', ''))
    return {
        'from': name or addr, 'subject': h.get('subject') or '(제목 없음)',
        'snippet': (msg.get('snippet') or '')[:120], 'when': _when(h.get('date', '')),
        'unread': 'UNREAD' in (msg.get('labelIds') or []),
        'link': f'https://mail.google.com/mail/u/?authuser={email}#inbox/{msg.get("threadId") or m["id"]}',
    }


def _fetch_mail(email: str, limit: int, service=None) -> dict:
    """구글에 두 번만 묻는다: ① 받은편지함 목록 ② 안 읽은 수 + 메일 몇 통을 한 번에 묶어서(batch)."""
    inbox = f'https://mail.google.com/mail/u/?authuser={email}#inbox'
    try:
        svc = service or _gmail(email)
        users = svc.users()
        lst = users.messages().list(userId='me', q='in:inbox', maxResults=limit).execute()
        ids = [m['id'] for m in lst.get('messages') or []]
        got: Dict[str, dict] = {}
        if hasattr(svc, 'new_batch_http_request'):
            def _cb(rid, resp, exc):
                if exc is None:
                    got[rid] = resp
            batch = svc.new_batch_http_request(callback=_cb)
            batch.add(users.labels().get(userId='me', id='INBOX'), request_id='__inbox__')
            for mid in ids:
                batch.add(users.messages().get(userId='me', id=mid, format='metadata',
                                               metadataHeaders=['From', 'Subject', 'Date']), request_id=mid)
            batch.execute()
        else:   # 묶음 요청을 못 쓰는 경우(시험용 가짜 등)
            got['__inbox__'] = users.labels().get(userId='me', id='INBOX').execute()
            for mid in ids:
                got[mid] = users.messages().get(userId='me', id=mid, format='metadata',
                                                metadataHeaders=['From', 'Subject', 'Date']).execute()
        items = [_item(email, {'id': mid}, got[mid]) for mid in ids if mid in got]
        return {'ok': True, 'unread': int((got.get('__inbox__') or {}).get('messagesUnread') or 0),
                'items': items, 'inbox': inbox}
    except Exception as e:
        logger.warning(f'[GW-HOME] 메일 불러오기 실패({email}): {type(e).__name__}: {e}')
        return {'ok': False, 'reason': 'error', 'inbox': inbox}


def _refresh_in_background(email: str, limit: int) -> None:
    with _mail_lock:
        if email in _mail_refreshing:
            return
        _mail_refreshing.add(email)

    def _run():
        try:
            out = _fetch_mail(email, limit)
            if out.get('ok'):
                with _mail_lock:
                    _mail_cache[email] = (time.time(), out)
        finally:
            with _mail_lock:
                _mail_refreshing.discard(email)
    threading.Thread(target=_run, daemon=True, name='gw-mail-refresh').start()


def mail_summary(email: str, limit: int = 5, service=None) -> dict:
    """안 읽은 수 + 받은편지함 최근 몇 통. 회사 메일만, 실패하면 {'ok': False}.
    1분 안이면 기억한 값, 15분 안이면 기억한 값을 바로 주고 뒤에서 새로 받음(홈이 메일 때문에 늦지 않게)."""
    email = (email or '').strip().lower()
    if not email.endswith(COMPANY_DOMAIN):
        return {'ok': False, 'reason': 'no_company_mail'}
    now = time.time()
    with _mail_lock:
        hit = _mail_cache.get(email)
    if hit and hit[1].get('ok'):
        age = now - hit[0]
        if age < _MAIL_TTL:
            return hit[1]
        if age < _MAIL_STALE and service is None:
            _refresh_in_background(email, limit)
            return hit[1]
    elif hit and now - hit[0] < _MAIL_TTL:      # 실패도 1분은 기억(구글 장애 때 매번 기다리지 않게)
        return hit[1]
    out = _fetch_mail(email, limit, service)
    with _mail_lock:
        _mail_cache[email] = (now, out)
    return out
