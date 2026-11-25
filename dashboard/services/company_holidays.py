"""회사 쉬는 날(2026-11-25 부장님 — 플렉스 「쉬는 날」 대체).

쉬는 날 = 법정 공휴일(holidays 라이브러리 + 노동절, 자동·고칠 수 없음)
        + 경영지원이 넣는 임시공휴일·회사 지정 휴일(company_holidays 표).
회사 달력의 '쉬는 날'은 모두 여기서 판단한다 — 휴가 일수(leave.count_days), 근무일(leave.is_workday),
부재중 리마인드·정합성 체크·스케줄 근무표가 각자 공휴일 라이브러리를 부르던 것을 하나로(어긋남 방지).

휴가 일수는 신청할 때 세어 저장한다(leave_requests.days, 연차는 leave_uses 로 떼어 둠). 그래서 쉬는 날을 넣거나
지우면 그 날짜에 걸린 휴가(승인·승인 대기)를 다시 센다 — 바뀐 몫만(전·후 차이) 더하고 빼고, 연차를 다시 떼고, 본인에게 DM.
"""
from __future__ import annotations

import threading
import time
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from typing import Dict, Iterable, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

KIND_LEGAL, KIND_TEMP, KIND_COMPANY = 'legal', 'temp', 'company'
KIND_LABELS = {KIND_LEGAL: '법정 공휴일', KIND_TEMP: '임시공휴일', KIND_COMPANY: '회사 지정'}
EDIT_KINDS = (KIND_TEMP, KIND_COMPANY)
NAME_MAX = 30
SPAN_MAX_DAYS = 31          # 한 번에 넣는 기간(여름 휴무 등) — 잘못 고른 먼 날짜 막기
YEAR_MIN, YEAR_MAX = 2020, 2100
LABOR_DAY = (5, 1)          # 노동절 — 회사 휴무(11-05 KiKO). holidays.KR 엔 없음
# 라이브러리의 법령 이름 → 화면 이름(11-25 부장님 '성탄절로' → 같은 날 '크리스마스'로 — 휴가 캘린더와 같게). 대체 공휴일 이름('… 대체 공휴일')도 같이 바뀜
LEGAL_RENAME = {'기독탄신일': '크리스마스'}

DDL = [
    '''CREATE TABLE IF NOT EXISTS company_holidays (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        start TEXT NOT NULL,
        "end" TEXT NOT NULL,
        name TEXT NOT NULL,
        kind TEXT NOT NULL,
        yearly INTEGER NOT NULL DEFAULT 0,
        note TEXT,
        created_by TEXT,
        created_by_name TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT,
        updated_by_name TEXT,
        deleted_at TEXT,
        deleted_by_name TEXT
    )''',
]

_CACHE_TTL = 60.0
_lock = threading.Lock()
_rows_cache: Dict[str, object] = {'at': 0.0, 'rows': None}
_legal_cache: Dict[int, Dict[date, str]] = {}
_local = threading.local()      # 미리 보기용 가정(이 요청 스레드에만) — {'add': [row], 'drop': {id}}


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _d(s) -> Optional[date]:
    try:
        return date.fromisoformat(str(s)[:10])
    except (TypeError, ValueError):
        return None


# ── 법정 공휴일 ──────────────────────────────────────────────────────────
def legal(year: int) -> Dict[date, str]:
    """{날짜: 이름} — 라이브러리(대체 공휴일·선거일 포함) + 노동절. 라이브러리가 없으면 노동절만."""
    got = _legal_cache.get(year)
    if got is not None:
        return got
    out: Dict[date, str] = {}
    try:
        import holidays
        for d, n in holidays.KR(years=year).items():
            n = str(n)
            for old, new in LEGAL_RENAME.items():
                n = n.replace(old, new)
            out[d] = n
    except Exception as e:      # 라이브러리 문제 — 노동절만이라도(휴가 일수가 공휴일을 세게 되므로 기록)
        logger.warning(f'[쉬는 날] 공휴일 라이브러리 실패 {year}: {e}')
        return {date(year, *LABOR_DAY): '노동절'}
    out.setdefault(date(year, *LABOR_DAY), '노동절')
    _legal_cache[year] = out
    return out


# ── 회사가 넣은 쉬는 날 ──────────────────────────────────────────────────
def _conn():
    from dashboard.services import leave as lv
    return lv._conn()


def _load_rows() -> List[dict]:
    with _conn() as conn:
        return [dict(r) for r in conn.execute(
            'SELECT * FROM company_holidays WHERE deleted_at IS NULL ORDER BY start, id').fetchall()]


def rows(fresh: bool = False) -> List[dict]:
    """지워지지 않은 회사 쉬는 날 줄 — 1분 기억(휴가 일수 셀 때 날마다 DB 를 읽지 않게). 읽기 실패면 마지막 값."""
    with _lock:
        cached, at = _rows_cache['rows'], float(_rows_cache['at'])
    if not fresh and cached is not None and time.time() - at < _CACHE_TTL:
        return list(cached)                                   # type: ignore[arg-type]
    try:
        got = _load_rows()
    except Exception as e:
        logger.warning(f'[쉬는 날] 회사 쉬는 날 읽기 실패 — 마지막 값으로: {e}')
        return list(cached or [])                             # type: ignore[arg-type]
    with _lock:
        _rows_cache['rows'], _rows_cache['at'] = got, time.time()
    return list(got)


def invalidate() -> None:
    with _lock:
        _rows_cache['rows'], _rows_cache['at'] = None, 0.0


def _active_rows() -> List[dict]:
    rs = rows()
    hyp = getattr(_local, 'hyp', None)
    if hyp:
        rs = [r for r in rs if r.get('id') not in hyp['drop']] + list(hyp['add'])
    return rs


def _row_dates(r: dict, year: int) -> List[date]:
    """그 줄이 그해에 쉬게 하는 날짜들. 매년 반복은 시작한 해부터 같은 월·일(2/29 는 없는 해엔 건너뜀)."""
    s, e = _d(r.get('start')), _d(r.get('end'))
    if not s or not e or e < s:
        return []
    if r.get('yearly'):
        if year < s.year:
            return []
        out, d = [], s
        while d <= e:
            try:
                out.append(date(year + (d.year - s.year), d.month, d.day))
            except ValueError:
                pass
            d += timedelta(days=1)
        return [x for x in out if x.year == year]
    out, d = [], max(s, date(year, 1, 1))
    while d <= e and d.year == year:
        out.append(d)
        d += timedelta(days=1)
    return out


def custom(year: int) -> Dict[date, Tuple[str, str, int]]:
    """{날짜: (이름, 구분, id)} — 회사가 넣은 쉬는 날(같은 날 두 줄이면 먼저 넣은 것)."""
    out: Dict[date, Tuple[str, str, int]] = {}
    for r in _active_rows():
        for d in _row_dates(r, year):
            out.setdefault(d, (r['name'], r['kind'], r.get('id') or 0))
    return out


def holiday_name(d: date) -> str:
    """쉬는 날이면 이름(법정이 먼저), 아니면 ''."""
    n = legal(d.year).get(d)
    if n:
        return n
    c = custom(d.year).get(d)
    return c[0] if c else ''


def is_holiday(d: date) -> bool:
    return bool(holiday_name(d))


def month_days(year: int, month: int) -> List[dict]:
    """그달 쉬는 날 [{day, name, kind}] — 달력에 빨간 날로."""
    out = []
    lg, cu = legal(year), custom(year)
    for d in sorted(set(lg) | set(cu)):
        if d.month == month:
            if d in lg:
                out.append({'day': d.isoformat(), 'name': lg[d], 'kind': KIND_LEGAL})
            else:
                out.append({'day': d.isoformat(), 'name': cu[d][0], 'kind': cu[d][1]})
    return out


# ── 화면 목록 ────────────────────────────────────────────────────────────
def _stem(name: str) -> str:
    """설날 전날·설날·설날 다음날 → 설날 (연휴를 한 줄로 묶을 때)."""
    for cut in (' 전날', ' 다음날'):
        if name.endswith(cut):
            return name[: -len(cut)]
    return name


def year_list(year: int) -> dict:
    """그해 쉬는 날 목록 + 요약. 법정 연휴(설날·추석)는 이어진 날을 한 줄로."""
    lg, cu = legal(year), custom(year)
    entries: List[dict] = []
    for d in sorted(lg):
        stem = _stem(lg[d])
        last = entries[-1] if entries else None
        if last and last['kind'] == KIND_LEGAL and last['name'] == stem and _d(last['end']) + timedelta(days=1) == d:
            last['end'] = d.isoformat()
            continue
        entries.append({'id': None, 'kind': KIND_LEGAL, 'name': stem, 'start': d.isoformat(), 'end': d.isoformat(),
                        'yearly': True, 'editable': False, 'note': ''})
    for r in _active_rows():
        ds = _row_dates(r, year)
        if not ds:
            continue
        entries.append({'id': r.get('id'), 'kind': r['kind'], 'name': r['name'], 'start': ds[0].isoformat(),
                        'end': ds[-1].isoformat(), 'yearly': bool(r.get('yearly')), 'editable': True,
                        'note': r.get('note') or '', 'by': r.get('updated_by_name') or r.get('created_by_name') or '',
                        'base_start': r['start'], 'base_end': r['end']})
    entries.sort(key=lambda e: (e['start'], e['kind'] != KIND_LEGAL))
    all_days = set(lg) | set(cu)
    for e in entries:
        s, t = _d(e['start']), _d(e['end'])
        e['days'] = (t - s).days + 1
        e['kind_label'] = KIND_LABELS.get(e['kind'], e['kind'])
    by_kind = {KIND_LEGAL: len(lg), KIND_TEMP: 0, KIND_COMPANY: 0}
    for d, (_n, k, _i) in cu.items():
        if d not in lg and k in by_kind:        # 법정 공휴일과 겹친 날은 법정 쪽으로만 센다
            by_kind[k] += 1
    return {'year': year, 'entries': entries, 'total': len(all_days),
            'weekdays': sum(1 for d in all_days if d.weekday() < 5), 'by_kind': by_kind}


# ── 넣기·고치기·지우기 ───────────────────────────────────────────────────
def _clean(data: dict) -> dict:
    """입력 확인 — 틀리면 ValueError(화면에 그대로 보여 줄 문장)."""
    name = ' '.join(str(data.get('name') or '').split())
    if not name:
        raise ValueError('이름을 적어 주세요.')
    if len(name) > NAME_MAX:
        raise ValueError(f'이름은 {NAME_MAX}자까지 적을 수 있어요.')
    kind = str(data.get('kind') or '')
    if kind not in EDIT_KINDS:
        raise ValueError('구분을 골라 주세요.')
    s = _d(data.get('start'))
    if not s:
        raise ValueError('날짜를 골라 주세요.')
    e = _d(data.get('end')) if data.get('end') else s
    if not e:
        raise ValueError('끝 날짜를 다시 골라 주세요.')
    if e < s:
        raise ValueError('끝 날짜가 시작 날짜보다 앞서요.')
    if (e - s).days + 1 > SPAN_MAX_DAYS:
        raise ValueError(f'한 번에 {SPAN_MAX_DAYS}일까지 넣을 수 있어요.')
    if not (YEAR_MIN <= s.year <= YEAR_MAX):
        raise ValueError('날짜를 다시 확인해 주세요.')
    yearly = bool(data.get('yearly')) and kind == KIND_COMPANY      # 임시공휴일은 그해만
    if yearly and s.year != e.year:
        raise ValueError('매년 반복은 한 해 안의 기간만 넣을 수 있어요.')
    note = ' '.join(str(data.get('note') or '').split())[:200]
    return {'name': name, 'kind': kind, 'start': s.isoformat(), 'end': e.isoformat(), 'yearly': int(yearly), 'note': note}


def get(hid: int) -> Optional[dict]:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM company_holidays WHERE id=? AND deleted_at IS NULL', (hid,)).fetchone()
    return dict(r) if r else None


@contextmanager
def _hypothetical(add: Iterable[dict] = (), drop: Iterable[int] = ()):
    """이 스레드에서만 쉬는 날을 바꾼 셈 치고 센다(다른 요청은 그대로)."""
    prev = getattr(_local, 'hyp', None)
    _local.hyp = {'add': list(add), 'drop': set(drop)}
    try:
        yield
    finally:
        _local.hyp = prev


def _changed_days(add: List[dict], drop: List[dict], today: date) -> List[date]:
    """바뀌는 날짜(작년~내후년) — 법정 공휴일과 겹치는 날은 원래 쉬는 날이라 빼고."""
    out = set()
    for y in range(today.year - 1, today.year + 3):
        lg = legal(y)
        before = set(custom(y))
        with _hypothetical(add, [r['id'] for r in drop if r.get('id')]):
            after = set(custom(y))
        out |= {d for d in before ^ after if d not in lg}
    return sorted(out)


def _affected(add: List[dict], drop: List[dict], today: Optional[date] = None) -> List[dict]:
    """다시 셀 휴가 — 바뀌는 날짜에 걸린 승인·승인 대기 신청. 전·후 일수 차이(delta)만 반영한다
    (저장된 일수가 다른 이유로 지금 셈과 달라도 쉬는 날 몫만 바꾸게)."""
    from dashboard.services import leave as lv
    today = today or date.today()
    days = _changed_days(add, drop, today)
    if not days:
        return []
    lo, hi = days[0].isoformat(), days[-1].isoformat()
    with lv._conn() as conn:
        reqs = [dict(r) for r in conn.execute(
            f'SELECT * FROM leave_requests WHERE status IN ({",".join("?" * len(lv.HOLDING))}) '
            f'AND start<=? AND "end">=? ORDER BY start, id', (*lv.HOLDING, hi, lo)).fetchall()]
    dayset = set(days)
    out = []
    for r in reqs:
        s, e = _d(r['start']), _d(r['end'])
        hit = [d for d in dayset if s <= d <= e]
        if not hit:
            continue
        before = lv.count_days(r['ltype'], s, e, r['name'])
        with _hypothetical(add, [x['id'] for x in drop if x.get('id')]):
            after = lv.count_days(r['ltype'], s, e, r['name'])
        delta = round(after - before, 2)
        if abs(delta) < 1e-9:
            continue
        new_days = max(0.0, round(float(r['days'] or 0) + delta, 2))
        out.append({'id': r['id'], 'name': r['name'], 'slack_id': r.get('slack_id') or '', 'ltype': r['ltype'],
                    'type_label': lv.type_label(r['ltype'], r.get('sub') or ''), 'status': r['status'],
                    'status_label': lv.STATUS_LABELS.get(r['status'], r['status']),
                    'start': r['start'], 'end': r['end'], 'period': lv.period_text(r),
                    'old_days': float(r['days'] or 0), 'new_days': new_days, 'delta': delta,
                    'deducts': lv.deducts(r['ltype'])})
    return out


def _plan(action: str, data: dict, hid: int = 0) -> Tuple[List[dict], List[dict], Optional[dict], Optional[dict]]:
    """(add, drop, 새 줄, 원래 줄) — 미리 보기·저장이 같은 계산을 쓰게."""
    old = get(hid) if hid else None
    if action in ('edit', 'delete') and not old:
        raise LookupError('이미 지워졌거나 없는 쉬는 날이에요. 새로 고쳐 주세요.')
    if old and old['kind'] not in EDIT_KINDS:
        raise ValueError('법정 공휴일은 고칠 수 없어요.')
    if action == 'delete':
        return [], [old], None, old
    new = _clean(data)
    new_row = dict(new, id=hid or -1)
    return [new_row], ([old] if old else []), new, old


def preview(action: str, data: dict, hid: int = 0, today: Optional[date] = None) -> dict:
    add, drop, new, _old = _plan(action, data, hid)
    aff = _affected(add, drop, today)
    return {'affected': aff, 'count': len(aff), 'clean': new}


def _josa(word: str, with_final: str, without: str) -> str:
    """받침 있으면 with_final(이·을), 없으면 without(가·를). 한글이 아니면 받침 없는 쪽."""
    ch = (word or ' ')[-1]
    if '가' <= ch <= '힣' and (ord(ch) - 0xAC00) % 28:
        return with_final
    return without


def _dm_text(a: dict, name: str, action: str) -> str:
    from dashboard.services import leave as lv
    what = f"{a['period']} {a['type_label']}"
    q = f'「{name}」'
    if action == 'delete':
        head = f"회사 쉬는 날 {q}{_josa(name, '이', '가')} 취소되어 휴가 일수를 다시 셌어요."
    else:
        head = f"{q}{_josa(name, '이', '가')} 회사 쉬는 날로 정해져 휴가 일수를 다시 셌어요."
    line = f"{what}: {lv.fmt_days(a['old_days'])} → {lv.fmt_days(a['new_days'])}"
    if a['deducts'] and a['delta'] < 0:
        tail = f"연차 {lv.fmt_days(-a['delta'])}이 돌아왔어요."
    elif a['deducts'] and a['delta'] > 0:
        tail = f"연차 {lv.fmt_days(a['delta'])}을 더 썼어요."
    else:
        tail = ''
    return '\n'.join(x for x in (head, line, tail) if x)


def _notify(affected: List[dict], name: str, action: str) -> int:
    """본인 DM — 실패는 기록만(쉬는 날·일수는 이미 바뀜). 보낸 수."""
    if not affected:
        return 0
    try:
        from dashboard.blueprints.slack_leave import _client
        c = _client()
    except Exception as e:
        logger.warning(f'[쉬는 날] 휴가 봇 준비 실패 — DM {len(affected)}건 못 보냄: {e}')
        return 0
    sent = 0
    for a in affected:
        if not a['slack_id']:
            logger.warning(f"[쉬는 날] 슬랙 ID 없음 — {a['name']} 휴가 #{a['id']} DM 못 보냄")
            continue
        try:
            ch = c.conversations_open(users=a['slack_id'])['channel']['id']
            c.chat_postMessage(channel=ch, text=_dm_text(a, name, action))
            sent += 1
        except Exception as e:
            logger.warning(f"[쉬는 날] DM 실패 {a['name']} #{a['id']}: {e}")
    return sent


def save(action: str, data: dict, actor_email: str, actor_name: str, hid: int = 0,
         expect: Optional[int] = None, today: Optional[date] = None, notify: bool = True) -> dict:
    """넣기(add)·고치기(edit)·지우기(delete) + 걸린 휴가 다시 세기를 한 번에. expect = 화면이 본 걸린 휴가 수
    (미리 본 뒤 그사이 휴가가 바뀌었으면 다시 보여 주게). 연차가 모자라 다시 뗄 수 없으면 아무것도 바꾸지 않는다."""
    from dashboard.services import leave as lv
    add, drop, new, old = _plan(action, data, hid)
    aff = _affected(add, drop, today)
    if expect is not None and int(expect) != len(aff):
        raise ValueError(f'그사이 걸린 휴가가 바뀌었어요({len(aff)}건). 다시 확인해 주세요.')
    now = _now()
    label = (new or old or {}).get('name', '')
    with lv._conn() as conn:
        if action == 'add':
            cur = conn.execute(
                'INSERT INTO company_holidays (start, "end", name, kind, yearly, note, created_by, created_by_name, '
                'created_at) VALUES (?,?,?,?,?,?,?,?,?)',
                (new['start'], new['end'], new['name'], new['kind'], new['yearly'], new['note'],
                 actor_email, actor_name, now))
            hid = cur.lastrowid
        elif action == 'edit':
            cur = conn.execute(
                'UPDATE company_holidays SET start=?, "end"=?, name=?, kind=?, yearly=?, note=?, updated_at=?, '
                'updated_by_name=? WHERE id=? AND deleted_at IS NULL',
                (new['start'], new['end'], new['name'], new['kind'], new['yearly'], new['note'], now, actor_name, hid))
            if cur.rowcount != 1:
                conn.rollback()
                raise LookupError('이미 지워졌거나 없는 쉬는 날이에요. 새로 고쳐 주세요.')
        else:
            cur = conn.execute('UPDATE company_holidays SET deleted_at=?, deleted_by_name=? WHERE id=? '
                               'AND deleted_at IS NULL', (now, actor_name, hid))
            if cur.rowcount != 1:
                conn.rollback()
                raise LookupError('이미 지워졌거나 없는 쉬는 날이에요. 새로 고쳐 주세요.')
        short = []
        for a in aff:
            row = conn.execute('SELECT * FROM leave_requests WHERE id=?', (a['id'],)).fetchone()
            if not row or row['status'] != a['status'] or abs(float(row['days'] or 0) - a['old_days']) > 1e-9:
                conn.rollback()
                raise ValueError(f"{a['name']}님 휴가가 그사이 바뀌었어요. 다시 확인해 주세요.")
            conn.execute('UPDATE leave_requests SET days=?, updated_at=? WHERE id=?', (a['new_days'], now, a['id']))
            if a['deducts']:
                conn.execute('DELETE FROM leave_uses WHERE request_id=?', (a['id'],))
                s, e = _d(a['start']), _d(a['end'])
                if a['new_days'] > 0 and not lv._allocate(conn, a['id'], a['name'], s, a['new_days'], e):
                    short.append(a['name'])
                    continue
            what = '추가' if action == 'add' else ('수정' if action == 'edit' else '취소')
            conn.execute('INSERT INTO leave_steps (request_id, step, actor_id, actor_name, result, reason, at) '
                         'VALUES (?,?,?,?,?,?,?)',
                         (a['id'], 'holiday', actor_email, actor_name, 'recounted',
                          f"회사 쉬는 날 「{label}」 {what} — {lv.fmt_days(a['old_days'])} → {lv.fmt_days(a['new_days'])}",
                          now))
        if short:
            conn.rollback()
            raise ValueError(f"{', '.join(sorted(set(short)))}님 연차가 모자라 다시 셀 수 없어요. "
                             '연차를 먼저 조정한 뒤 다시 해 주세요.')
        conn.commit()
    invalidate()
    logger.info(f'[쉬는 날] {action} #{hid} 「{label}」 by {actor_name} — 휴가 {len(aff)}건 다시 셈')
    sent = _notify(aff, label, action) if notify else 0
    return {'id': hid, 'affected': aff, 'count': len(aff), 'dm_sent': sent}
