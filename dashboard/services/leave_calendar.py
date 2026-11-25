"""휴가·토요 휴무 → 구글 캘린더(11-25 부장님 '토요일 휴무랑 연차 휴가 구글 캘린더에 반영').

- 회사가 이미 쓰던 휴가 공유 캘린더(회사 전체 공개)에 승인된 휴가·반차와 확정된 토요 휴무를 올린다.
- 올리는 내용은 #휴가_일정과 같다 — 이름과 휴가/오전·오후 반차/스케줄휴가만. 휴가 종류(병가·경조사)·사유는 올리지 않는다.
- 제목은 이 캘린더에 쓰던 형식(11-25 부장님 '스케줄도 아이콘 추가해서 통일'): '🌴 [이름] 휴가'·'🌴 [이름] 오후 반차',
  토요일은 하루 한 건 '📅 [이름, 이름] 스케줄휴가'(전체 휴무는 '📅 [전체] 스케줄휴가 (메모)').
- 그룹웨어 기록이 기준. 10분마다 '있어야 할 일정'을 만들어 캘린더와 맞춘다(없으면 넣고, 다르면 고치고, 그룹웨어에서
  사라졌으면 지움). 승인·취소·변경 신청·휴무 출근 등 어느 길로 바뀌어도 따로 손대지 않아도 따라온다.
- 이 기능이 올린 일정만 만진다(extendedProperties.private.itg_src). 사람이 직접 넣은 일정은 그대로 둔다.
  단 START(11-25) 뒤 일정 중 손으로 넣은 같은 일정(같은 날 같은 제목, 스케줄휴가는 이름이 모두 들어 있는 것)은
  새로 만들지 않고 이어받는다 — 같은 일정이 두 번 보이지 않게. 이어받은 날에 이름이 모두 들어 있는 손 일정이 더 있으면
  지운다(구글 휴지통에 30일 남음). START 전 일정(옛 기록)은 손대지 않는다.
- 기존 PM 공사 캘린더(calendar_service, 10-06부터 꺼 둠)와는 별개 — 서비스 계정으로 쓴다.

설정(.env)
- GW_LEAVE_CALENDAR_ID   공유 캘린더 ID(비면 끔). 캘린더 설정 > '캘린더 통합'의 캘린더 ID.
  그 캘린더를 서비스 계정(credentials.json 의 client_email)에 '일정 변경' 권한으로 공유해야 쓸 수 있다.
- GW_LEAVE_CALENDAR_AS   회사 계정 메일(선택). 조직 정책상 바깥(서비스 계정)에 '일정 변경'을 못 줄 때 — 서비스 계정이
  도메인 전체 위임으로 이 사람 이름으로 쓴다. 관리 콘솔 위임 범위에 calendar.events 추가 + 이 사람이 캘린더 '일정 변경' 권한 필요.
"""
from __future__ import annotations

import hashlib
import os
import re
import threading
import time
from datetime import date, timedelta
from typing import Dict, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)
_LOG = '[휴가 캘린더]'

SRC = 'gw_leave'                 # 이 기능이 올린 일정 표시
START = date(2026, 11, 25)       # 이날부터 맞춘다 — 그 전은 손으로 넣어 온 기록이라 그대로 둠
PAST_DAYS = 90                   # 맞추는 범위 — 오늘 기준 앞뒤(이 밖의 일정은 건드리지 않음)
AHEAD_DAYS = 400
LEAVE_ICON, SAT_ICON, SAT_WORD = '🌴', '📅', '스케줄휴가'
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
    who = (os.getenv('GW_LEAVE_CALENDAR_AS') or '').strip()
    if who:      # 회사 계정 이름으로(도메인 전체 위임) — 조직 밖(서비스 계정)엔 '일정 변경' 공유가 막혀 있어서(11-25)
        creds = creds.with_subject(who)
    return build('calendar', 'v3', credentials=creds, cache_discovery=False)


# ── 있어야 할 일정 ──────────────────────────────────────────────────────
def _window(today: date):
    return max(today - timedelta(days=PAST_DAYS), START), today + timedelta(days=AHEAD_DAYS)


def _event(summary: str, start: date, end: date, names=()) -> dict:
    return {'summary': summary, 'start': {'date': start.isoformat()},
            'end': {'date': (end + timedelta(days=1)).isoformat()},   # 종일 일정 끝은 다음 날(구글 규칙)
            'transparency': 'transparent', 'description': DESC, '_names': list(names)}


def _eid(key: str) -> str:
    """일정 ID — 구글 규칙(소문자 a~v·숫자, 5자 이상)에 맞게 키에서 늘 같은 값으로."""
    return 'itg' + hashlib.md5(key.encode('utf-8')).hexdigest()


def leave_title(name: str, what: str) -> str:
    return f'{LEAVE_ICON} [{name}] {what}'


def sat_title(names, note: str = '', full: bool = False) -> str:
    who = '전체' if full else ', '.join(names)
    return f'{SAT_ICON} [{who}] {SAT_WORD}' + (f' ({note})' if note else '')


def desired(today: Optional[date] = None) -> Dict[str, dict]:
    """{키: 일정} — 승인된 휴가·반차 + 확정된 달의 토요 휴무(토요일마다 한 건)."""
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
        if not s or not e or s < ws:            # 범위 앞에서 시작한 휴가는 옛 기록 쪽에 있음
            continue
        if r['ltype'] in lv.HALF_TYPES:
            what = '오전 반차' if r['ltype'] == lv.HALF_AM else '오후 반차'
        else:
            what = '휴가'
        out[f"lv:{r['id']}"] = _event(leave_title(r['name'], what), s, e, [r['name']])

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
            if not d or not (ws <= d <= we) or not (names or day in full):
                continue
            out[f'so:{day}'] = _event(sat_title(names, full.get(day, ''), day in full), d, d, names)
    return out


# ── 캘린더와 맞추기 ─────────────────────────────────────────────────────
def _existing(svc, cal: str, ws: date, we: date):
    """(이 기능이 올린 일정 {키: 일정}, 손으로 넣은 종일 일정 {날짜: [일정]})."""
    tagged: Dict[str, dict] = {}
    manual: Dict[str, List[dict]] = {}
    token = None
    while True:
        res = svc.events().list(calendarId=cal,
                                timeMin=f'{ws.isoformat()}T00:00:00+09:00',
                                timeMax=f'{(we + timedelta(days=1)).isoformat()}T00:00:00+09:00',
                                singleEvents=True, showDeleted=False, maxResults=2500, pageToken=token).execute()
        for ev in res.get('items', []):
            priv = (ev.get('extendedProperties') or {}).get('private') or {}
            if priv.get('itg_src') == SRC and priv.get('itg_key'):
                tagged[priv['itg_key']] = ev
            elif ev.get('status') != 'cancelled' and (ev.get('start') or {}).get('date', '') >= ws.isoformat():
                manual.setdefault(ev['start']['date'], []).append(ev)
        token = res.get('nextPageToken')
        if not token:
            return tagged, manual


def _split_names(s: str) -> set:
    return {x for x in re.split(r'[,\s/·()\[\]]+', s or '') if x}


def _hand_names(summary: str) -> Optional[set]:
    """손으로 넣은 스케줄휴가 일정의 이름들 — '스케줄휴가)박민재,빈승정' → {'박민재','빈승정'}. 아니면 None."""
    s = (summary or '').strip()
    m = re.match(r'^📅\s*\[(.*?)\]\s*스케줄\s*휴가', s)
    if m:
        return _split_names(m.group(1))
    m = re.match(r'^스케줄\s*휴가\s*\)?\s*(.*)$', s)
    return _split_names(m.group(1)) if m else None


def _adoptable(key: str, want: dict, ev: dict) -> bool:
    if key.startswith('lv:'):
        return (ev.get('summary') or '').strip() == want['summary']
    names = _hand_names(ev.get('summary'))
    return bool(names) and names <= set(want['_names'])


def _same(have: dict, want: dict) -> bool:
    return (have.get('summary') == want['summary']
            and (have.get('start') or {}).get('date') == want['start']['date']
            and (have.get('end') or {}).get('date') == want['end']['date']
            and have.get('status') != 'cancelled')


def _body(key: str, want: dict, with_id: bool = True) -> dict:
    b = {k: v for k, v in want.items() if not k.startswith('_')}
    b.update(status='confirmed', extendedProperties={'private': {'itg_src': SRC, 'itg_key': key}})
    if with_id:
        b['id'] = _eid(key)
    return b


def _status(e) -> int:
    try:
        return int(getattr(getattr(e, 'resp', None), 'status', 0) or 0)
    except (TypeError, ValueError):
        return 0


def plan(today: Optional[date] = None, svc=None) -> List[Tuple[str, str, str]]:
    """미리 보기 — 쓰지 않고 할 일만 [(할 일, 날짜, 제목)]."""
    return sync(today, svc, dry=True).get('plan', [])


_COUNT = {'넣기': 'added', '이어받기': 'adopted', '고치기': 'changed'}


def sync(today: Optional[date] = None, svc=None, dry: bool = False) -> dict:
    """그룹웨어 기록대로 캘린더를 맞춘다. 반환 {added, adopted, changed, removed, errors} (dry 면 {plan})."""
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
            have, manual = _existing(svc, cal, ws, we)
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
        n = {'added': 0, 'adopted': 0, 'changed': 0, 'removed': 0}
        errors: List[str] = []
        todo: List[Tuple[str, str, str]] = []

        def do(kind: str, day: str, title: str, fn) -> None:
            todo.append((kind, day, title))
            if dry:
                return
            try:
                fn()
                n[_COUNT.get(kind, 'removed')] += 1
            except Exception as e:
                errors.append(f'{title} {day} {kind}: {e}')
            time.sleep(0.05)

        def ins(key: str, w: dict) -> None:
            body = _body(key, w)
            try:
                ev.insert(calendarId=cal, body=body).execute()
            except Exception as e:
                if _status(e) != 409:             # 같은 ID 가 지운 채로 남아 있으면 되살림
                    raise
                ev.update(calendarId=cal, eventId=body['id'], body=body).execute()

        def rm(eid: str) -> None:
            try:
                ev.delete(calendarId=cal, eventId=eid).execute()
            except Exception as e:
                if _status(e) not in (404, 410):
                    raise

        for key, w in want.items():
            day = w['start']['date']
            h = have.get(key)
            cands = [x for x in manual.get(day, []) if _adoptable(key, w, x)]
            if h:
                if not _same(h, w):
                    do('고치기', day, w['summary'],
                       lambda h=h, key=key, w=w: ev.update(calendarId=cal, eventId=h['id'], body=_body(key, w)).execute())
                for x in cands:                   # 나중에 그룹웨어가 맞춰져 손 일정이 겹치게 되면 지움
                    do('겹쳐서 지우기', day, x.get('summary') or '', lambda x=x: rm(x['id']))
                    manual[day].remove(x)
                continue
            if not cands:
                do('넣기', day, w['summary'], lambda key=key, w=w: ins(key, w))
                continue
            first = cands[0]
            do('이어받기', day, f"{first.get('summary')} → {w['summary']}",
               lambda first=first, key=key, w=w: ev.update(calendarId=cal, eventId=first['id'],
                                                           body=_body(key, w, with_id=False)).execute())
            for x in cands[1:]:                   # 같은 날 이름이 모두 들어간 손 일정이 더 있으면 겹치니 지움
                do('겹쳐서 지우기', day, x.get('summary') or '', lambda x=x: rm(x['id']))
            for x in cands:
                manual[day].remove(x)
        for key, h in have.items():
            if key not in want:
                do('지우기', (h.get('start') or {}).get('date', ''), h.get('summary') or '', lambda h=h: rm(h['id']))
        if dry:
            return {'plan': todo}
        if any(n.values()):
            logger.info(f"{_LOG} 넣음 {n['added']} · 이어받음 {n['adopted']} · 고침 {n['changed']} · 지움 {n['removed']}")
        if errors:
            msg = f'{_LOG} {len(errors)}건 반영 실패(10분 뒤 다시 시도): ' + ' / '.join(errors[:5])
            if _alerted_on != today:              # 계속 실패해도 알림은 하루 한 번
                _alerted_on = today
                logger.error(msg)
            else:
                logger.warning(msg)
        return dict(n, errors=errors)
    finally:
        _run_lock.release()
