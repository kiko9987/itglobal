"""휴가·토요 휴무 → 구글 캘린더(11-25 부장님 '토요일 휴무랑 연차 휴가 구글 캘린더에 반영').

- 캘린더 하나('아이티글로벌 휴가' 같은 공유 캘린더)에 승인된 휴가·반차와 확정된 토요 휴무를 올린다. 회사 전체 공개.
- 올리는 내용은 #휴가_일정과 같다 — 이름과 휴가/오전·오후 반차/토요 휴무만. 휴가 종류(병가·경조사)·사유는 올리지 않는다.
- 그룹웨어 기록이 기준. 10분마다 '있어야 할 일정'을 만들어 캘린더와 맞춘다(없으면 넣고, 다르면 고치고, 그룹웨어에서
  사라졌으면 지움). 승인·취소·변경 신청·휴무 출근 등 어느 길로 바뀌어도 따로 손대지 않아도 따라온다.
- 이 기능이 올린 일정만 만진다(extendedProperties.private.itg_src). 캘린더에 사람이 직접 넣은 일정은 그대로 둔다.
- 기존 PM 공사 캘린더(calendar_service, 10-06부터 꺼 둠)와는 별개 — 서비스 계정으로 쓴다.

설정(.env)
- GW_LEAVE_CALENDAR_ID   공유 캘린더 ID(비면 끔). 캘린더 설정 > '캘린더 통합'의 캘린더 ID.
  그 캘린더를 서비스 계정(credentials.json 의 client_email)에 '일정 변경' 권한으로 공유해야 쓸 수 있다.
"""
from __future__ import annotations

import hashlib
import os
import threading
import time
from datetime import date, timedelta
from typing import Dict, List, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)
_LOG = '[휴가 캘린더]'

SRC = 'gw_leave'                 # 이 기능이 올린 일정 표시
PAST_DAYS = 90                   # 맞추는 범위 — 오늘 기준 앞뒤(이 밖의 일정은 건드리지 않음)
AHEAD_DAYS = 400
SAT_COLOR = '2'                  # 토요 휴무는 다른 색(구글 캘린더 colorId 2 = 세이지)
DESC = '그룹웨어 휴가·스케줄 근무에서 자동으로 올리는 일정입니다. 고칠 때는 그룹웨어에서 바꿔 주세요.'

_run_lock = threading.Lock()
_alerted_on: Optional[date] = None


def calendar_id() -> str:
    return (os.getenv('GW_LEAVE_CALENDAR_ID') or '').strip()


def enabled() -> bool:
    return bool(calendar_id())


def _service():
    """호출마다 새로 만든다(googleapiclient 는 스레드 공유 금지)."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds_file = (os.getenv('GOOGLE_CREDENTIALS_FILE') or os.getenv('GOOGLE_APPLICATION_CREDENTIALS')
                  or 'credentials.json')
    creds = service_account.Credentials.from_service_account_file(
        creds_file, scopes=['https://www.googleapis.com/auth/calendar.events'])
    return build('calendar', 'v3', credentials=creds, cache_discovery=False)


# ── 있어야 할 일정 ──────────────────────────────────────────────────────
def _window(today: date):
    return today - timedelta(days=PAST_DAYS), today + timedelta(days=AHEAD_DAYS)


def _event(summary: str, start: date, end: date, color: str = '') -> dict:
    ev = {'summary': summary, 'start': {'date': start.isoformat()},
          'end': {'date': (end + timedelta(days=1)).isoformat()},   # 종일 일정 끝은 다음 날(구글 규칙)
          'transparency': 'transparent', 'description': DESC}
    if color:
        ev['colorId'] = color
    return ev


def _eid(key: str) -> str:
    """일정 ID — 구글 규칙(소문자 a~v·숫자, 5자 이상)에 맞게 키에서 늘 같은 값으로."""
    return 'itg' + hashlib.md5(key.encode('utf-8')).hexdigest()


def desired(today: Optional[date] = None) -> Dict[str, dict]:
    """{키: 일정} — 승인된 휴가·반차 + 확정된 달의 토요 휴무(전체 휴무 날은 한 건)."""
    from dashboard.services import leave as lv
    from dashboard.services import sat_off as so
    today = today or date.today()
    ws, we = _window(today)
    out: Dict[str, dict] = {}
    with lv._conn() as conn:
        rows = conn.execute('SELECT id, name, ltype, start, "end" FROM leave_requests '
                            'WHERE status=? AND start<=? AND "end">=?',
                            (lv.APPROVED, we.isoformat(), ws.isoformat())).fetchall()
    for r in rows:
        s, e = lv._d(r['start']), lv._d(r['end'])
        if not s or not e:
            continue
        if r['ltype'] in lv.HALF_TYPES:
            what = '오전 반차' if r['ltype'] == lv.HALF_AM else '오후 반차'
        else:
            what = '휴가'
        out[f"lv:{r['id']}"] = _event(f"{r['name']} {what}", s, e)

    months, m = [], ws.replace(day=1)
    while m <= we:
        months.append(m.strftime('%Y-%m'))
        m = (m + timedelta(days=32)).replace(day=1)
    with lv._conn() as conn:
        q = ','.join('?' * len(months))
        done = [r[0] for r in conn.execute(f'SELECT month FROM sat_off_months WHERE status=? AND month IN ({q})',
                                           (so.CONFIRMED, *months))]
    for month in sorted(done):
        full = so.full_days(month)
        for day, names in so.picks(month).items():
            d = lv._d(day)
            if not d or not (ws <= d <= we):
                continue
            if day in full:
                note = full[day]
                out[f'sf:{day}'] = _event('토요 전체 휴무' + (f' ({note})' if note else ''), d, d, SAT_COLOR)
                continue
            for n in names:
                out[f'so:{day}:{n}'] = _event(f'{n} 토요 휴무', d, d, SAT_COLOR)
    return out


# ── 캘린더와 맞추기 ─────────────────────────────────────────────────────
def _existing(svc, cal: str, ws: date, we: date) -> Dict[str, dict]:
    out, token = {}, None
    while True:
        res = svc.events().list(calendarId=cal, privateExtendedProperty=f'itg_src={SRC}',
                                timeMin=f'{ws.isoformat()}T00:00:00+09:00',
                                timeMax=f'{(we + timedelta(days=1)).isoformat()}T00:00:00+09:00',
                                singleEvents=True, showDeleted=False, maxResults=2500, pageToken=token).execute()
        for ev in res.get('items', []):
            key = ((ev.get('extendedProperties') or {}).get('private') or {}).get('itg_key')
            if key:
                out[key] = ev
        token = res.get('nextPageToken')
        if not token:
            return out


def _same(have: dict, want: dict) -> bool:
    return (have.get('summary') == want['summary']
            and (have.get('start') or {}).get('date') == want['start']['date']
            and (have.get('end') or {}).get('date') == want['end']['date']
            and (have.get('colorId') or '') == want.get('colorId', '')
            and have.get('status') != 'cancelled')


def _body(key: str, want: dict) -> dict:
    return dict(want, id=_eid(key), status='confirmed',
                extendedProperties={'private': {'itg_src': SRC, 'itg_key': key}})


def _status(e) -> int:
    try:
        return int(getattr(getattr(e, 'resp', None), 'status', 0) or 0)
    except (TypeError, ValueError):
        return 0


def sync(today: Optional[date] = None, svc=None) -> dict:
    """그룹웨어 기록대로 캘린더를 맞춘다. 반환 {added, changed, removed, errors}."""
    global _alerted_on
    cal = calendar_id()
    if not cal:
        return {'skipped': 'off'}
    if not _run_lock.acquire(blocking=False):
        return {'skipped': 'running'}
    try:
        today = today or date.today()
        ws, we = _window(today)
        want = desired(today)
        svc = svc or _service()
        try:
            have = _existing(svc, cal, ws, we)
        except Exception as e:
            code = _status(e)
            msg = (f'{_LOG} 캘린더를 읽지 못했어요({code or e}). 캘린더 ID 와 서비스 계정 공유(일정 변경 권한)를 확인해 주세요.'
                   if code in (403, 404) else f'{_LOG} 캘린더 읽기 실패: {e}')
            if _alerted_on != today:              # 같은 날 10분마다 알림이 쌓이지 않게 하루 한 번만
                _alerted_on = today
                logger.error(msg, extra={'force_alert': True})
            else:
                logger.warning(msg)
            return {'error': str(e)}
        ev = svc.events()
        added = changed = removed = 0
        errors: List[str] = []
        for key, w in want.items():
            h = have.get(key)
            if h and _same(h, w):
                continue
            body = _body(key, w)
            try:
                if h:
                    ev.update(calendarId=cal, eventId=h['id'], body=body).execute()
                    changed += 1
                else:
                    try:
                        ev.insert(calendarId=cal, body=body).execute()
                    except Exception as e:
                        if _status(e) != 409:       # 같은 ID 가 지운 채로 남아 있으면 되살림
                            raise
                        ev.update(calendarId=cal, eventId=body['id'], body=body).execute()
                    added += 1
            except Exception as e:
                errors.append(f"{w['summary']} {w['start']['date']}: {e}")
            time.sleep(0.05)
        for key, h in have.items():
            if key in want:
                continue
            try:
                ev.delete(calendarId=cal, eventId=h['id']).execute()
                removed += 1
            except Exception as e:
                if _status(e) not in (404, 410):
                    errors.append(f"{h.get('summary')} 지우기: {e}")
            time.sleep(0.05)
        if added or changed or removed:
            logger.info(f'{_LOG} 넣음 {added} · 고침 {changed} · 지움 {removed}')
        if errors:
            msg = f'{_LOG} {len(errors)}건 반영 실패(10분 뒤 다시 시도): ' + ' / '.join(errors[:5])
            if _alerted_on != today:              # 계속 실패해도 알림은 하루 한 번
                _alerted_on = today
                logger.error(msg)
            else:
                logger.warning(msg)
        return {'added': added, 'changed': changed, 'removed': removed, 'errors': errors}
    finally:
        _run_lock.release()
