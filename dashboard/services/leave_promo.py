"""연차 사용 촉진 (2026-11-18, 플렉스 '연차 촉진' 옮기기) — 근로기준법 제61조 절차.

법(2026-11-18 원문 확인, 국가법령 MCP):
  ① 1년 이상 근속자의 연차(60조 1·4항, 1년 미만 때 생긴 월차 제외):
     1. 쓸 수 있는 기간이 끝나기 6개월 전을 기준으로 10일 안에, 직원마다 안 쓴 일수를 알리고
        직원이 사용 시기를 정해 알려 주도록 서면으로 촉구
     2. 촉구를 받은 뒤 10일 안에 직원이 (전부 또는 일부) 사용 시기를 알리지 않으면, 끝나기 2개월 전까지
        회사가 안 쓴 휴가의 사용 시기를 정해 직원에게 서면으로 통보
  ② 1년 미만 근속자의 월차(60조 2항): 위와 같되 '최초 1년이 끝나기 3개월 전 기준 10일 안' 촉구,
     '1개월 전까지' 지정 통보. 촉구 뒤 새로 생긴 월차는 '1개월 전 기준 5일 안' 촉구, '10일 전까지' 지정 통보.
  이렇게 했는데도 안 쓴 연차는 소멸해도 보상하지 않는다.

플렉스 기록(flex_promotion_20261118.json)도 같은 날짜 규칙이다: 촉구일 = 소멸일−6개월 무렵, 계획 기한 = 촉구일+10일,
지정 기한 = 소멸일−2개월.

여기서 하는 일:
  - 매일(daily) 촉구할 사람을 찾는다 → 자동 보내기(LEAVE_PROMOTION_AUTO, 기본 켬)면 본인 DM + 그룹웨어 '내 휴가'
    촉구 문서, 아니면 경영지원에게 '보낼 사람' 알림.
  - 직원은 그룹웨어에서 [받았어요]·[사용 계획 내기](날짜를 고르면 보통 휴가 신청으로 들어가 JW→YG 승인).
  - 계획 기한이 지나도 안 쓴(신청 안 한) 연차가 남으면 경영지원이 [사용일 지정] → 승인된 휴가로 들어가고
    본인에게 지정 통보 DM.
  - 기록은 모두 이 표(leave_promotions)에 남는다(보낸 때·받은 때·계획·지정). 플렉스 기록도 옮겨 넣는다.
"""
from __future__ import annotations

import json
import os
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Tuple

from dashboard.services import leave as lv
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)
_LOG = '[연차촉진]'

ANNUAL, FIRST, FIRST_LATE = 'annual', 'first_year', 'first_year_late'
KIND_LABEL = {ANNUAL: '연차', FIRST: '1년 미만 월차', FIRST_LATE: '1년 미만 월차(촉구 뒤 생긴 몫)'}
PLAN_DAYS = 10
PLAN_REASON = '연차 사용 계획(촉진)'
DESIGNATE_REASON = lv.PROMO_DESIGNATE_REASON

NOTICED, PLANNED, DESIGNATED, CLOSED, CANCELLED = 'noticed', 'planned', 'designated', 'closed', 'cancelled'

_DDL = [
    '''CREATE TABLE IF NOT EXISTS leave_promotions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        slack_id TEXT,
        kind TEXT NOT NULL,
        grant_to TEXT NOT NULL,
        days REAL NOT NULL,
        noticed_at TEXT,
        noticed_by TEXT,
        plan_due TEXT,
        designate_due TEXT,
        ack_at TEXT,
        plan_at TEXT,
        plan_dates TEXT,
        plan_rids TEXT,
        designated_at TEXT,
        designated_by TEXT,
        designated_dates TEXT,
        designated_rids TEXT,
        status TEXT NOT NULL,
        source TEXT,
        note TEXT,
        created_at TEXT NOT NULL,
        UNIQUE(name, grant_to, kind)
    )''',
    'CREATE INDEX IF NOT EXISTS idx_leave_promo_name ON leave_promotions(name)',
]
_ready = False
_ready_db = None


@contextmanager
def _conn():
    """휴가 장부와 같은 DB(같은 연결 규칙). 표는 그 DB 에 처음 붙을 때 한 번 만든다(시험에서 DB 를 바꿔 끼워도 따라감)."""
    global _ready, _ready_db
    db = lv._db()
    with lv._conn() as c:
        if not _ready or _ready_db is not db:
            for ddl in _DDL:
                c.execute(ddl)
            cols = {r[1] for r in c.execute('PRAGMA table_info(leave_promotions)').fetchall()}
            if 'designate_dm' not in cols:      # 지정 통보 DM 결과 'sent'·'failed'(11-24 조용한 실패 점검), 옛 기록은 빈칸
                c.execute('ALTER TABLE leave_promotions ADD COLUMN designate_dm TEXT')
            c.commit()
            _ready, _ready_db = True, db
        yield c


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def auto() -> bool:
    """촉구를 자동으로 보내나 — 플렉스도 '스마트 촉진 시점'에 자동으로 보냈다. 끄려면 LEAVE_PROMOTION_AUTO=false."""
    return os.getenv('LEAVE_PROMOTION_AUTO', 'true').strip().lower() not in ('0', 'false', 'no', 'off')


# ── 날짜 규칙 ─────────────────────────────────────────────────────────
def schedule(kind: str, grant_to: date) -> Dict[str, date]:
    """촉구를 보낼 기간(notice_from~notice_to)과 지정 통보 기한(designate_due). 끝나는 날 = grant_to 다음 날."""
    end = grant_to + timedelta(days=1)
    if kind == ANNUAL:
        frm, span, due = lv.add_months(end, -6), 10, lv.add_months(end, -2)
    elif kind == FIRST:
        frm, span, due = lv.add_months(end, -3), 10, lv.add_months(end, -1)
    else:   # FIRST_LATE
        frm, span, due = lv.add_months(end, -1), 5, end - timedelta(days=10)
    return {'end': end, 'notice_from': frm, 'notice_to': frm + timedelta(days=span - 1), 'designate_due': due}


def _first_year_end(person: dict) -> Optional[date]:
    hire = lv._d(person.get('hire_date'))
    return lv.add_years(hire, 1) - timedelta(days=1) if hire else None


def period_left(name: str, grant_to: str) -> float:
    """그 기간(같은 날 소멸하는 부여들)에서 아직 쓰지도 신청하지도 않은 연차."""
    return round(sum(g['days'] - g['used'] for g in lv.grants(name) if g['valid_to'] == grant_to), 2)


def periods(person: dict, today: date) -> List[dict]:
    """지금 살아 있거나 앞으로 생긴 기간 — [{grant_to, kind, left}]. 1년 미만 월차 기간이면 kind=FIRST."""
    fy = _first_year_end(person)
    out = {}
    for g in lv.grants(person['name']):
        if g['valid_to'] < today.isoformat() or g['valid_to'] == lv.ANYTIME:
            continue
        out.setdefault(g['valid_to'], 0.0)
        out[g['valid_to']] += g['days'] - g['used']
    rows = []
    for to, left in sorted(out.items()):
        kind = FIRST if fy and to == fy.isoformat() else ANNUAL
        rows.append({'grant_to': to, 'kind': kind, 'left': round(left, 2)})
    return rows


def _late_days(name: str, grant_to: str, first_noticed: str) -> float:
    """1년 미만: 1차 촉구 뒤에 생긴 월차 중 남은 몫."""
    return round(sum(g['days'] - g['used'] for g in lv.grants(name)
                     if g['valid_to'] == grant_to and str(g['source'] or '').startswith('monthly:')
                     and g['valid_from'] > first_noticed[:10]), 2)


# ── 기록 ─────────────────────────────────────────────────────────────
def _row(r) -> dict:
    d = dict(r)
    for k in ('plan_dates', 'plan_rids', 'designated_dates', 'designated_rids'):
        try:
            d[k] = json.loads(d.get(k) or '[]')
        except ValueError:
            d[k] = []
    return d


def get(pid: int) -> Optional[dict]:
    with _conn() as c:
        r = c.execute('SELECT * FROM leave_promotions WHERE id=?', (int(pid),)).fetchone()
    return _row(r) if r else None


def records(name: str = '', include_flex: bool = True) -> List[dict]:
    sql, args = 'SELECT * FROM leave_promotions WHERE 1=1', []
    if name:
        sql += ' AND name=?'
        args.append(name)
    if not include_flex:
        sql += " AND COALESCE(source, '') NOT LIKE 'flex:%'"
    with _conn() as c:
        return [_row(r) for r in c.execute(sql + ' ORDER BY grant_to DESC, id DESC', args).fetchall()]


def _existing(name: str, grant_to: str, kind: str) -> Optional[dict]:
    with _conn() as c:
        r = c.execute('SELECT * FROM leave_promotions WHERE name=? AND grant_to=? AND kind=?',
                      (name, grant_to, kind)).fetchone()
    return _row(r) if r else None


def state(rec: dict, today: Optional[date] = None) -> Dict[str, str]:
    """화면에 보일 단계 — {key, label}. key: wait_plan·planned·need_designate·designated·late·done·cancelled."""
    today = today or date.today()
    t = today.isoformat()
    if rec['status'] == CANCELLED:
        return {'key': 'cancelled', 'label': '취소'}
    if rec['grant_to'] < t or rec['status'] == CLOSED:
        return {'key': 'done', 'label': '끝남'}
    if rec['status'] == DESIGNATED:
        return {'key': 'designated', 'label': '사용일 지정함'}
    left = period_left(rec['name'], rec['grant_to'])
    planned = rec['status'] == PLANNED
    if rec['plan_due'] and t <= rec['plan_due']:
        if planned:
            return {'key': 'planned', 'label': '계획 받음' + (f' · 아직 신청 안 한 연차 {lv.fmt_days(left)}' if left > 1e-9 else '')}
        return {'key': 'wait_plan', 'label': f"계획 기다림 (~{lv.kdate(lv._d(rec['plan_due']))})"}
    if left <= 1e-9:
        return {'key': 'planned' if planned else 'done', 'label': '계획 받음' if planned else '남은 연차 없음'}
    # 계획 기한이 지났는데 아직 신청 안 한 연차가 남음(계획을 일부만 냈거나, 낸 날을 취소·반려) → 회사가 지정해야 함
    head = f'계획 받음 · 미신청 {lv.fmt_days(left)} — ' if planned else ''
    if rec['designate_due'] and t <= rec['designate_due']:
        return {'key': 'need_designate', 'label': head + f"사용일 지정 필요 (~{lv.kdate(lv._d(rec['designate_due']))})"}
    return {'key': 'late', 'label': head + '지정 기한 지남'}


# ── 할 일 찾기 ───────────────────────────────────────────────────────
def candidates(today: Optional[date] = None, include_missed: bool = False) -> List[dict]:
    """오늘 촉구를 보내야 하는 사람 — [{name, slack_id, kind, grant_to, left, notice_from, notice_to, designate_due, missed}].
    include_missed: 보낼 기간을 이미 지났는데 기록이 없는 것도(경영지원 확인용 — 늦게 보내면 법 효과가 없을 수 있다)."""
    today = today or date.today()
    t = today.isoformat()
    out = []
    for p in lv.staff():
        for per in periods(p, today):
            kinds = [per['kind']]
            if per['kind'] == FIRST:
                kinds.append(FIRST_LATE)
            for kind in kinds:
                if _existing(p['name'], per['grant_to'], kind):
                    continue
                sch = schedule(kind, lv._d(per['grant_to']))
                left = per['left']
                if kind == FIRST_LATE:
                    first = _existing(p['name'], per['grant_to'], FIRST)
                    if not first or not first.get('noticed_at'):
                        continue
                    left = _late_days(p['name'], per['grant_to'], first['noticed_at'])
                if left <= 1e-9 or t < sch['notice_from'].isoformat():
                    continue
                missed = t > sch['notice_to'].isoformat()
                if missed and (not include_missed or t > per['grant_to']):
                    continue
                out.append({'name': p['name'], 'slack_id': p.get('slack_id') or '', 'kind': kind,
                            'grant_to': per['grant_to'], 'left': left,
                            'notice_from': sch['notice_from'].isoformat(), 'notice_to': sch['notice_to'].isoformat(),
                            'designate_due': sch['designate_due'].isoformat(), 'missed': missed})
    return out


def send_notice(name: str, kind: str, grant_to: str, actor: str = '자동', today: Optional[date] = None,
                uid: str = '') -> Tuple[Optional[dict], str]:
    """촉구 기록을 만든다(알림은 슬랙 쪽 notify_notice). 경영지원이 누를 땐 uid 확인."""
    today = today or date.today()
    if uid and uid not in lv.manager_ids():
        return None, '경영지원만 보낼 수 있어요.'
    st = lv.staff_by_name(name)
    if not st:
        return None, '휴가 명단에 없는 직원이에요.'
    if kind not in KIND_LABEL:
        return None, '촉진 종류가 틀렸어요.'
    if _existing(name, grant_to, kind):
        return None, '이미 촉구를 보냈어요.'
    if kind == FIRST_LATE:
        first = _existing(name, grant_to, FIRST)
        left = _late_days(name, grant_to, first['noticed_at']) if first and first.get('noticed_at') else 0.0
    else:
        left = period_left(name, grant_to)
    if left <= 1e-9:
        return None, '안 쓴 연차가 없어요.'
    sch = schedule(kind, lv._d(grant_to))
    plan_due = today + timedelta(days=PLAN_DAYS)
    with _conn() as c:
        try:
            cur = c.execute(
                'INSERT INTO leave_promotions (name, slack_id, kind, grant_to, days, noticed_at, noticed_by, plan_due, '
                'designate_due, status, source, created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)',
                (name, st.get('slack_id') or None, kind, grant_to, left, _now(), actor, plan_due.isoformat(),
                 sch['designate_due'].isoformat(), NOTICED, 'gw', _now()))
        except Exception as exc:     # 같은 줄을 동시에 — UNIQUE
            if 'UNIQUE' in str(exc):
                return None, '이미 촉구를 보냈어요.'
            raise
        rid = cur.lastrowid
        c.commit()
    logger.info(f'{_LOG} 촉구 {name} {kind} ~{grant_to} {left:g}일 ({actor})')
    return get(rid), ''


def _is_owner(rec: dict, uid: str) -> bool:
    if not uid:
        return False
    st = lv.staff_by_name(rec['name']) or {}
    return uid in ((rec.get('slack_id') or ''), (st.get('slack_id') or ''))


def ack(pid: int, uid: str) -> str:
    """직원 [받았어요] — 받은 시각(서면 촉구 수령 확인). 본인만."""
    rec = get(pid)
    if not rec or not _is_owner(rec, uid):
        return '본인 촉구만 확인할 수 있어요.'
    if not rec.get('ack_at'):
        with _conn() as c:
            c.execute('UPDATE leave_promotions SET ack_at=? WHERE id=? AND ack_at IS NULL', (_now(), rec['id']))
    return ''


# ── 날짜 묶기 ────────────────────────────────────────────────────────
def _parse_items(items) -> Tuple[List[Tuple[date, str]], str]:
    """[{date, half: ''|'am'|'pm'}] → [(날짜, 'all'|'am'|'pm')] (날짜순). 틀리면 오류."""
    out, seen = [], set()
    for it in items or []:
        d = lv._d(str((it or {}).get('date') or ''))
        if not d:
            continue
        h = str((it or {}).get('half') or '').strip()
        h = h if h in ('am', 'pm') else 'all'
        if (d, 'all') in seen or (h == 'all' and any(x[0] == d for x in seen)) or (d, h) in seen:
            return [], f'{lv.kdate(d)}이 두 번 들어 있어요.'
        if not lv.is_workday(d):
            return [], f'{lv.kdate(d, True)}은 쉬는 날이에요.'
        seen.add((d, h))
        out.append((d, h))
    if not out:
        return [], '날짜를 하나 이상 골라 주세요.'
    return sorted(out), ''


def _ranges(items: List[Tuple[date, str]]) -> List[Tuple[str, date, date]]:
    """종일 날짜는 쉬는 날만 끼고 이어지면 한 신청으로 묶는다. 반차는 따로."""
    out: List[list] = []
    for d, h in items:
        if h != 'all':
            out.append([lv.HALF_AM if h == 'am' else lv.HALF_PM, d, d])
            continue
        if out and out[-1][0] == lv.ANNUAL:
            nxt = out[-1][2] + timedelta(days=1)
            while nxt < d and not lv.is_workday(nxt):
                nxt += timedelta(days=1)
            if nxt == d:
                out[-1][2] = d
                continue
        out.append([lv.ANNUAL, d, d])
    return [tuple(x) for x in out]


def _total(items: List[Tuple[date, str]]) -> float:
    return sum(0.5 if h != 'all' else 1.0 for _, h in items)


def _check_dates(rec: dict, items, today: date, first_ok: date) -> Tuple[list, str]:
    parsed, err = _parse_items(items)
    if err:
        return [], err
    to = lv._d(rec['grant_to'])
    for d, _ in parsed:
        if d < first_ok or d > to:
            return [], f'{lv.kdate(first_ok)}부터 {lv.kdate(to)} 사이 날짜로 골라 주세요.'
    left = period_left(rec['name'], rec['grant_to'])
    if _total(parsed) > left + 1e-9:
        return [], f'안 쓴 연차가 {lv.fmt_days(left)}이라 {lv.fmt_days(_total(parsed))}를 고를 수 없어요.'
    for ltype, s, e in _ranges(parsed):
        errs = lv.validate(rec['name'], ltype, '', s, e, allow_admin=True)
        if errs:
            return [], next(iter(errs.values()))
    return parsed, ''


def _undo(rids: List[int], uid: str, actor: str) -> None:
    for rid in rids:
        row = lv.get(rid)
        if row and row['status'] in lv.HOLDING:
            with lv._conn() as c:
                c.execute('UPDATE leave_requests SET status=?, updated_at=? WHERE id=?', (lv.CANCELLED, _now(), rid))
                c.execute('DELETE FROM leave_uses WHERE request_id=?', (rid,))
                c.execute('INSERT INTO leave_steps (request_id, step, actor_id, actor_name, result, reason, at) '
                          'VALUES (?,?,?,?,?,?,?)', (rid, 'cancel', uid, actor, 'cancelled', '촉진 처리 되돌림', _now()))


def submit_plan(pid: int, uid: str, items, today: Optional[date] = None) -> Tuple[Optional[dict], List[dict], str]:
    """직원 사용 계획 → (기록, 만든 신청들, 오류). 날짜마다 보통 휴가 신청(승인 대기)으로 들어간다."""
    today = today or date.today()
    rec = get(pid)
    if not rec or not _is_owner(rec, uid):
        return None, [], '본인 촉구에만 계획을 낼 수 있어요.'
    if rec['status'] not in (NOTICED, PLANNED) or rec['grant_to'] < today.isoformat():
        return None, [], '지금은 계획을 낼 수 없어요.'
    parsed, err = _check_dates(rec, items, today, today)
    if err:
        return None, [], err
    made = []
    for ltype, s, e in _ranges(parsed):
        row = lv.create(rec['name'], uid, uid, rec['name'], ltype, '', s, e, PLAN_REASON)
        if not row:
            _undo([r['id'] for r in made], uid, rec['name'])
            return None, [], '그 사이 남은 연차가 바뀌었어요. 다시 해 주세요.'
        made.append(row)
    dates = [{'date': d.isoformat(), 'half': '' if h == 'all' else h} for d, h in parsed]
    with _conn() as c:
        c.execute('UPDATE leave_promotions SET plan_at=?, plan_dates=?, plan_rids=?, status=?, '
                  'ack_at=COALESCE(ack_at, ?) WHERE id=?',
                  (_now(), json.dumps(rec['plan_dates'] + dates), json.dumps(rec['plan_rids'] + [r['id'] for r in made]),
                   PLANNED, _now(), rec['id']))
    logger.info(f'{_LOG} 계획 #{pid} {rec["name"]} {len(dates)}일 → 신청 {[r["id"] for r in made]}')
    return get(pid), made, ''


def designate(pid: int, uid: str, actor: str, items, today: Optional[date] = None) -> Tuple[Optional[dict], List[dict], str]:
    """경영지원 사용일 지정 → 승인된 휴가로 들어간다(연차에서 뗌). 계획 기한이 지난 뒤, 내일부터 기간 안 날짜."""
    today = today or date.today()
    if uid not in lv.manager_ids():
        return None, [], '경영지원만 지정할 수 있어요.'
    rec = get(pid)
    if not rec or rec['status'] in (CANCELLED, CLOSED):
        return None, [], '지정할 수 있는 촉진 기록이 아니에요.'
    if rec['plan_due'] and today.isoformat() <= rec['plan_due']:
        return None, [], f"직원 계획 기한({lv.kdate(lv._d(rec['plan_due']))})이 지나야 지정할 수 있어요."
    parsed, err = _check_dates(rec, items, today, today + timedelta(days=1))
    if err:
        return None, [], err
    made = []
    for ltype, s, e in _ranges(parsed):
        row, errs = lv.admin_register(rec['name'], ltype, '', s, e, DESIGNATE_REASON, uid, actor)
        if errs:
            _undo([r['id'] for r in made], uid, actor)
            return None, [], next(iter(errs.values()))
        made.append(row)
    dates = [{'date': d.isoformat(), 'half': '' if h == 'all' else h} for d, h in parsed]
    with _conn() as c:
        c.execute('UPDATE leave_promotions SET designated_at=?, designated_by=?, designated_dates=?, designated_rids=?, '
                  'status=? WHERE id=?',
                  (_now(), actor, json.dumps(rec['designated_dates'] + dates),
                   json.dumps(rec['designated_rids'] + [r['id'] for r in made]), DESIGNATED, rec['id']))
    logger.info(f'{_LOG} 지정 #{pid} {rec["name"]} {len(dates)}일 ({actor})')
    return get(pid), made, ''


# ── 문구(서면 촉구·지정 통보) ─────────────────────────────────────────
def _when(d: dict) -> str:
    x = lv._d(d['date'])
    return lv.kdate(x, True) + {'am': ' 오전', 'pm': ' 오후'}.get(d.get('half') or '', '')


def notice_text(rec: dict) -> str:
    """서면 촉구 문서 — 그룹웨어 '내 휴가'와 DM 에 같은 글."""
    to = lv._d(rec['grant_to'])
    due = lv._d(rec['plan_due']) if rec.get('plan_due') else None
    des = lv._d(rec['designate_due']) if rec.get('designate_due') else None
    kind = KIND_LABEL.get(rec['kind'], '연차')
    lines = [
        f"{rec['name']} 님, {to.year}년 {lv.kdate(to)}까지 쓸 수 있는 {kind} 중 아직 쓰지 않은 연차가 {lv.fmt_days(rec['days'])} 있어요.",
        f"이 연차는 {lv.kdate(to)}이 지나면 소멸해요.",
    ]
    if due:
        lines.append(f"{lv.kdate(due)}까지 쓸 날짜를 정해 회사에 알려 주세요. 그룹웨어 [내 휴가]에서 [사용 계획 내기]를 누르면 돼요.")
    if des:
        lines.append(f"기한까지 알려 주지 않으면 회사가 쓸 날짜를 정해 {lv.kdate(des)}까지 서면으로 알려 드려요.")
    lines.append('이렇게 사용을 권했는데도 쓰지 않은 연차는 근로기준법 제61조에 따라 금전으로 보상하지 않아요.')
    return '\n'.join(lines)


def designate_text(rec: dict) -> str:
    dates = ', '.join(_when(d) for d in rec.get('designated_dates') or [])
    to = lv._d(rec['grant_to'])
    return (f"{rec['name']} 님이 사용 계획을 알려 주지 않은 연차의 사용 날짜를 회사가 아래와 같이 정했어요.\n"
            f"• 지정한 날: {dates}\n"
            f"• 대상: {to.year}년 {lv.kdate(to)}까지 쓸 수 있는 연차\n"
            '이 날은 연차 휴가로 처리돼요. 출근하지 않으셔도 돼요. 날짜를 바꿔야 하면 경영지원에 말씀해 주세요.')


# ── 화면용 ───────────────────────────────────────────────────────────
def view(rec: dict, today: Optional[date] = None) -> dict:
    st = state(rec, today)
    return {
        'id': rec['id'], 'name': rec['name'], 'kind': rec['kind'], 'kind_label': KIND_LABEL.get(rec['kind'], rec['kind']),
        'grant_to': rec['grant_to'], 'days': rec['days'], 'left': period_left(rec['name'], rec['grant_to']),
        'noticed_at': (rec.get('noticed_at') or '')[:16], 'noticed_by': rec.get('noticed_by') or '',
        'ack_at': (rec.get('ack_at') or '')[:16], 'plan_due': rec.get('plan_due') or '',
        'designate_due': rec.get('designate_due') or '', 'plan_at': (rec.get('plan_at') or '')[:16],
        'plan_dates': rec.get('plan_dates') or [], 'designated_at': (rec.get('designated_at') or '')[:16],
        'designated_by': rec.get('designated_by') or '', 'designated_dates': rec.get('designated_dates') or [],
        'source': 'flex' if str(rec.get('source') or '').startswith('flex:') else 'gw', 'note': rec.get('note') or '',
        'state': st['key'], 'state_label': st['label'],
        'designate_dm': rec.get('designate_dm') or '',      # 'failed' 면 촉진 탭에 '통보 DM 못 보냄'(11-24)
    }


def set_designate_dm(pid: int, ok: bool) -> None:
    """지정 통보(근로기준법 61조) 슬랙 DM 결과 — 못 보냈으면 촉진 탭이 '직접 전해 주세요'를 띄운다(11-24 조용한 실패 점검)."""
    with _conn() as c:
        c.execute('UPDATE leave_promotions SET designate_dm=? WHERE id=?', ('sent' if ok else 'failed', pid))
        c.commit()


def for_person(name: str, today: Optional[date] = None) -> List[dict]:
    """내 휴가 — 아직 끝나지 않은 촉진(촉구 문서·계획·지정)."""
    today = today or date.today()
    out = []
    for r in records(name):
        if r['grant_to'] < today.isoformat() or r['status'] == CANCELLED:
            continue
        v = view(r, today)
        v['notice_text'] = notice_text(r)
        v['designate_text'] = designate_text(r) if r.get('designated_dates') else ''
        v['can_plan'] = r['status'] in (NOTICED, PLANNED) and v['left'] > 1e-9 and bool(r.get('plan_due'))
        v['can_ack'] = r['status'] in (NOTICED, PLANNED, DESIGNATED) and not r.get('ack_at')
        out.append(v)
    return out


# ── 매일 ─────────────────────────────────────────────────────────────
def daily(today: Optional[date] = None, notify=None) -> dict:
    """매일 아침 — 보낼 촉구(자동이면 보냄)·계획 기한 이틀 전 직원 알림·지정 필요 경영지원 알림.
    notify(kind, payload): 슬랙 알림 함수(slack_leave.promo_notify). 휴가 기능을 켜기(LEAVE_ENABLED) 전엔 안 돈다."""
    today = today or date.today()
    out = {'sent': [], 'due': [], 'remind_plan': [], 'need_designate': [], 'late': []}
    if not lv.enabled():
        return out
    for c in candidates(today):
        if auto():
            rec, err = send_notice(c['name'], c['kind'], c['grant_to'], actor='자동', today=today)
            if rec:
                out['sent'].append(rec)
                if notify:
                    notify('notice', rec)
        else:
            out['due'].append(c)
    t = today.isoformat()
    for r in records(include_flex=True):
        if r['status'] in (CANCELLED, CLOSED, DESIGNATED) or r['grant_to'] < t:
            continue
        st = state(r, today)
        if st['key'] == 'wait_plan' and r['status'] == NOTICED and r['plan_due'] == (today + timedelta(days=2)).isoformat():
            out['remind_plan'].append(r)
            if notify:
                notify('remind_plan', r)
        elif st['key'] == 'need_designate':
            out['need_designate'].append(view(r, today))
        elif st['key'] == 'late':
            out['late'].append(view(r, today))
    if notify and (out['due'] or out['need_designate'] or out['late'] or out['sent']):
        notify('managers', out)
    return out


# ── 플렉스 기록 옮기기 ───────────────────────────────────────────────
_FLEX_STATUS = {'PENDING_APPROVAL': PLANNED, 'COMPLETE': PLANNED, 'SHOULD_MANAGER_WRITE': NOTICED,
                'CANNOT_WRITE': CLOSED, 'CANCELED': CANCELLED}


def import_flex(path: str, today: Optional[date] = None, apply: bool = False) -> List[dict]:
    """flex_promotion_YYYYMMDD.json → 기록(플렉스 문서 하나 = 한 줄, source 'flex:<문서 키>').
    같은 사람·기간에 플렉스 문서가 둘이면(재촉구) 마지막 것만. apply=False 면 넣을 줄만 돌려준다."""
    data = json.load(open(path, encoding='utf-8'))
    users = data.get('users') or {}
    pick: Dict[tuple, dict] = {}
    for d in data.get('docs') or []:
        name = users.get(d['u'])
        if not name:
            continue
        key = (name, d['to'])
        # 같은 기간 문서가 여럿(재촉구·취소)이면 취소 안 된 것 중 마지막 촉구
        rank = (not d.get('canceled'), d.get('boostedAt') or '')
        cur = pick.get(key)
        if cur is None or rank > (not cur.get('canceled'), cur.get('boostedAt') or ''):
            pick[key] = d
    rows = []
    for (name, to), d in sorted(pick.items()):
        st = lv.staff_by_name(name) or {}
        fy = _first_year_end(st) if st else None
        b = datetime.strptime(d['boostedAt'][:19], '%Y-%m-%dT%H:%M:%S') + timedelta(hours=9)   # UTC → 한국
        rows.append({
            'name': name, 'slack_id': st.get('slack_id'), 'kind': FIRST if fy and fy.isoformat() == to else ANNUAL,
            'grant_to': to, 'days': float(d['days'] or 0),
            'noticed_at': b.strftime('%Y-%m-%d %H:%M:%S'), 'noticed_by': '플렉스', 'plan_due': d.get('userBy'),
            'designate_due': d.get('mgrBy'),
            'plan_at': (datetime.utcfromtimestamp(d['writtenAt'] / 1000) + timedelta(hours=9)).strftime('%Y-%m-%d %H:%M:%S')
            if d.get('writtenAt') else None,
            'plan_dates': json.dumps([{'date': x, 'half': ''} for x in sorted(d.get('planned') or [])]),
            # 모르는 상태(직원 작성 기간 등)는 진행 중으로 — 끝난 것으로 두면 지정 알림이 안 간다
            'status': _FLEX_STATUS.get(d.get('status'), NOTICED) if not d.get('canceled') else CANCELLED,
            'source': f"flex:{d['task']}",
            'note': f"플렉스 {d.get('status')}" + (f" / 승인 {d['approval']}" if d.get('approval') else ''),
        })
    if apply:
        with _conn() as c:
            for r in rows:
                c.execute(
                    'INSERT OR IGNORE INTO leave_promotions (name, slack_id, kind, grant_to, days, noticed_at, noticed_by, '
                    'plan_due, designate_due, plan_at, plan_dates, plan_rids, designated_dates, designated_rids, status, '
                    "source, note, created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,'[]','[]','[]',?,?,?,?)",
                    (r['name'], r['slack_id'], r['kind'], r['grant_to'], r['days'], r['noticed_at'], r['noticed_by'],
                     r['plan_due'], r['designate_due'], r['plan_at'], r['plan_dates'], r['status'], r['source'],
                     r['note'], _now()))
            c.commit()
    return rows
