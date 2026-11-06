# -*- coding: utf-8 -*-
"""시공팀 출퇴근·조출/연장/야간 수당 (2026-11-06) — 세콤 기록 + 현장 출퇴근 보고.

데이터:
  · 세콤매니저 「ERP 내보내기」 자동(매시 5분) → C:\\S1\\ERPExport_YYYYMMDD.txt (탭 구분, cp949)
    날짜·시간·사원번호·이름·모드(1 출근 / 4 퇴근) … → att_events 에 쌓는다(같은 줄은 한 번).
  · 현장으로 바로 출퇴근해 세콤 태그가 없는 날 → 본인이 슬랙으로 시각 보고(예전 카톡) →
    경영지원 승인 → att_reports. 세콤 기록이 있으면 세콤이 먼저, 보고는 빈 칸만 채운다.

계산 — 경영지원(SB) 정산표 「아이티글로벌 기준」(2026-07~10 표·세콤 근태조회로 맞춤):
  · 조출 = 07:30 이전 출근, 연장 = 17:30 이후 퇴근. 06시 이전·22시 이후는 야간으로 따로.
  · 각 칸(조출·조출 야간·연장·연장 야간)을 하루 단위 30분 반올림(15분부터 올림).
  · 수당 = (조출+연장−야간) 시간 × 통상시급 × 1.5 + 야간 시간 × 통상시급 × 2 (원 단위 반올림).
  · 연장 270분 이상인 날 = 연장식대(법인카드로 따로 결제 — 수당엔 안 넣고 건수만).
  · 휴일도 평일과 같이 조출·연장만(SB 표 그대로). 휴일근무 가산은 확인 뒤 따로.
확인 필요로 표시: 출근·퇴근 중 하나만 있는 날, 퇴근이 다음 날 새벽(태그 누락이 의심되는 날).
"""
import glob
import os
import re
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Tuple

DEFAULT_STAFF = ('김태현', '김종연')
DEFAULT_WAGE = 15165
EXPORT_DIR_DEFAULT = r'C:\S1'
EARLY_BASE, LATE_BASE = 7 * 60 + 30, 17 * 60 + 30    # 07:30 · 17:30
NIGHT_END, NIGHT_START = 6 * 60, 22 * 60             # 06:00 · 22:00
MEAL_MIN = 270
OVERNIGHT_LIMIT = 3 * 60                             # 다음 날 03:00 까지의 퇴근은 그날 연장으로
LATE_ARRIVAL = 9 * 60                                # 09시 넘어 출근한 날의 연장은 확인 필요

IN_MODE, OUT_MODE = 1, 4
PENDING, APPROVED, REJECTED = 'pending', 'approved', 'rejected'


def staff() -> Tuple[str, ...]:
    raw = os.getenv('ATTENDANCE_STAFF', '').strip()
    return tuple(x.strip() for x in raw.split(',') if x.strip()) or DEFAULT_STAFF


def hourly_wage() -> int:
    try:
        return int(os.getenv('ATTENDANCE_HOURLY_WAGE', '') or DEFAULT_WAGE)
    except ValueError:
        return DEFAULT_WAGE


def export_dir() -> str:
    return os.getenv('SECOM_EXPORT_DIR', '').strip() or EXPORT_DIR_DEFAULT


def enabled() -> bool:
    """시공팀 본인 DM(누락 보고 요청)·경영지원 월말 정산표 — 켜기 전엔 KiKO 에게만 미리보기."""
    return os.getenv('ATTENDANCE_ENABLED', '').strip().lower() in ('1', 'true', 'yes')


def manager_ids() -> Tuple[str, ...]:
    """보고 승인·월말 정산표 받는 사람. 켜기 전(시험)엔 ATTENDANCE_MANAGER_IDS 또는 KiKO."""
    raw = tuple(x.strip() for x in os.getenv('ATTENDANCE_MANAGER_IDS', '').split(',') if x.strip())
    if raw:
        return raw
    if not enabled():
        return ('U04UL2ZLJAX',)
    from dashboard.services.settlement_checkers import checker_ids
    return checker_ids()


# ── 장부 ──────────────────────────────────────────────────────────────
_DDL = [
    '''CREATE TABLE IF NOT EXISTS att_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sabun TEXT NOT NULL,
        name TEXT NOT NULL,
        ts TEXT NOT NULL,
        mode INTEGER NOT NULL,
        src TEXT NOT NULL DEFAULT 'secom',
        UNIQUE(sabun, ts, mode)
    )''',
    'CREATE INDEX IF NOT EXISTS idx_att_events_name_ts ON att_events(name, ts)',
    '''CREATE TABLE IF NOT EXISTS att_reports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        day TEXT NOT NULL,
        in_hm TEXT,
        out_hm TEXT,
        note TEXT,
        status TEXT NOT NULL,
        reporter_id TEXT,
        reporter_name TEXT,
        decided_by TEXT,
        decided_reason TEXT,
        created_at TEXT NOT NULL,
        decided_at TEXT,
        dm_refs TEXT
    )''',
    'CREATE INDEX IF NOT EXISTS idx_att_reports_name_day ON att_reports(name, day)',
    '''CREATE TABLE IF NOT EXISTS att_asks (
        name TEXT NOT NULL,
        day TEXT NOT NULL,
        asked_at TEXT NOT NULL,
        PRIMARY KEY(name, day)
    )''',
]
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


# ── 세콤 파일 읽기 ─────────────────────────────────────────────────────
def parse_export(text: str) -> List[dict]:
    """ERP 내보내기 한 파일 → [{sabun, name, ts, mode}]. 모르는 줄은 건너뛴다."""
    out = []
    for line in text.splitlines():
        f = line.split('\t')
        if len(f) < 5 or not re.fullmatch(r'\d{8}', f[0]) or not re.fullmatch(r'\d{6}', f[1]):
            continue
        try:
            mode = int(f[4])
        except ValueError:
            continue
        ts = f'{f[0][:4]}-{f[0][4:6]}-{f[0][6:]} {f[1][:2]}:{f[1][2:4]}:{f[1][4:]}'
        out.append({'sabun': f[2].strip(), 'name': f[3].strip(), 'ts': ts, 'mode': mode})
    return out


def add_events(rows: List[dict], src: str = 'secom') -> int:
    n = 0
    with _conn() as conn:
        for r in rows:
            if not r['sabun'] or not r['name']:
                continue
            cur = conn.execute('INSERT OR IGNORE INTO att_events (sabun, name, ts, mode, src) VALUES (?,?,?,?,?)',
                               (r['sabun'], r['name'], r['ts'], r['mode'], src))
            n += cur.rowcount
    return n


def ingest_dir(path: str = '', days: int = 3) -> int:
    """최근 며칠 내보내기 파일을 읽어 새 줄만 넣는다(여러 번 읽어도 같은 줄은 한 번). 새로 넣은 줄 수."""
    path = path or export_dir()
    if not os.path.isdir(path):
        return 0
    cutoff = datetime.now().timestamp() - days * 86400
    n = 0
    for f in glob.glob(os.path.join(path, 'ERPExport*.txt')):
        try:
            if os.path.getmtime(f) < cutoff:
                continue
            with open(f, 'rb') as fh:
                n += add_events(parse_export(fh.read().decode('cp949', errors='replace')))
        except OSError:
            continue
    return n


# ── 하루 기록 ──────────────────────────────────────────────────────────
def _hm(m: Optional[int]) -> str:
    if m is None:
        return ''
    nd, mm = divmod(m, 1440)
    s = f'{mm // 60:02d}:{mm % 60:02d}'
    return s + (' (다음 날)' if nd else '')


def _to_min(day: str, ts: str) -> int:
    d0 = datetime.strptime(day, '%Y-%m-%d')
    return int((datetime.strptime(ts[:16], '%Y-%m-%d %H:%M') - d0).total_seconds() // 60)


def _hm_to_min(hm: str) -> Optional[int]:
    m = re.fullmatch(r'(\d{1,2}):(\d{2})', (hm or '').strip())
    return int(m.group(1)) * 60 + int(m.group(2)) if m else None


def events_between(name: str, start: str, end: str) -> List[dict]:
    with _conn() as conn:
        return [dict(r) for r in conn.execute(
            'SELECT * FROM att_events WHERE name=? AND ts>=? AND ts<? ORDER BY ts', (name, start, end)).fetchall()]


def reports_of(name: str, day: str) -> List[dict]:
    with _conn() as conn:
        return [dict(r) for r in conn.execute(
            'SELECT * FROM att_reports WHERE name=? AND day=? ORDER BY id DESC', (name, day)).fetchall()]


def day_record(name: str, day: str) -> dict:
    """그날 출근·퇴근(분, 0시 기준; 다음 날 새벽 퇴근은 1440+). 세콤 먼저, 빈 칸은 승인된 보고로."""
    d = datetime.strptime(day, '%Y-%m-%d')
    nxt = (d + timedelta(days=1)).strftime('%Y-%m-%d')
    ev = events_between(name, f'{day} 00:00:00', f'{nxt} {OVERNIGHT_LIMIT // 60:02d}:00:00')
    ins = [_to_min(day, e['ts']) for e in ev if e['mode'] == IN_MODE and e['ts'] < f'{nxt}']
    outs = [_to_min(day, e['ts']) for e in ev if e['mode'] == OUT_MODE]
    # 다음 날 새벽 퇴근이 다음 날 출근과 같은 시각(태그 하나가 둘로 잡힌 경우)이면 퇴근으로 안 친다
    nxt_ins = {_to_min(day, e['ts']) for e in events_between(name, f'{nxt} 00:00:00', f'{nxt} 23:59:59')
               if e['mode'] == IN_MODE}
    outs = [o for o in outs if o < 1440 or o not in nxt_ins]
    rec = {'name': name, 'day': day, 'in': min(ins) if ins else None,
           'out': max(outs) if outs else None, 'in_src': 'secom' if ins else '', 'out_src': 'secom' if outs else '',
           'report': None}
    reps = reports_of(name, day)
    appr = next((r for r in reps if r['status'] == APPROVED), None)
    pend = next((r for r in reps if r['status'] == PENDING), None)
    if appr:
        if rec['in'] is None and _hm_to_min(appr['in_hm']) is not None:
            rec['in'], rec['in_src'] = _hm_to_min(appr['in_hm']), 'report'
        if rec['out'] is None and _hm_to_min(appr['out_hm']) is not None:
            o = _hm_to_min(appr['out_hm'])
            rec['out'], rec['out_src'] = (o + 1440 if rec['in'] is not None and o < rec['in'] else o), 'report'
        rec['report'] = appr
    elif pend:
        rec['report'] = pend
    return rec


def r30(x: int) -> int:
    """30분 단위 반올림(15분부터 올림)."""
    return int((max(0, x) + 15) // 30 * 30)


def calc(in_min: Optional[int], out_min: Optional[int]) -> dict:
    """하루 → 조출(야간 제외)·조출 야간·연장(야간 제외)·연장 야간(분, 30분 반올림)·식대."""
    c = {'early': 0, 'early_night': 0, 'late': 0, 'late_night': 0, 'meal': False}
    if in_min is not None and in_min < EARLY_BASE:
        c['early_night'] = r30(NIGHT_END - in_min) if in_min < NIGHT_END else 0
        c['early'] = r30(EARLY_BASE - max(in_min, NIGHT_END))
    if out_min is not None and out_min > LATE_BASE:
        c['late'] = r30(min(out_min, NIGHT_START) - LATE_BASE)
        c['late_night'] = r30(out_min - NIGHT_START) if out_min > NIGHT_START else 0
    c['meal'] = c['late'] + c['late_night'] >= MEAL_MIN
    return c


def needs_check(rec: dict) -> str:
    """확인 필요 사유(없으면 '')."""
    if rec['in'] is None and rec['out'] is None:
        return ''
    if rec['in'] is None:
        return '출근 기록 없음'
    if rec['out'] is None:
        return '퇴근 기록 없음'
    if rec['out'] >= 1440:
        return '퇴근이 다음 날 새벽'
    return ''


def month_days(ym: str) -> List[str]:
    y, m = int(ym[:4]), int(ym[5:7])
    d = date(y, m, 1)
    out = []
    while d.month == m:
        out.append(d.isoformat())
        d += timedelta(days=1)
    return out


def summary(ym: str, names: Optional[Tuple[str, ...]] = None) -> Dict[str, dict]:
    """월 정산 — 사람별 날짜 칸·합계·수당·확인 필요 날."""
    wage = hourly_wage()
    out = {}
    for name in names or staff():
        days, checks, meals = [], [], []
        tot = {'early': 0, 'early_night': 0, 'late': 0, 'late_night': 0}
        for day in month_days(ym):
            rec = day_record(name, day)
            why = needs_check(rec)
            if why:
                checks.append({'day': day, 'why': why, 'in': _hm(rec['in']), 'out': _hm(rec['out']),
                               'pending': bool(rec['report'] and rec['report']['status'] == PENDING)})
            c = calc(rec['in'], rec['out'])
            if not why and rec['in'] is not None and rec['in'] > LATE_ARRIVAL and c['late']:
                # 10/17 김태현 10:02~22:06 — SB 는 연장 270 대신 다르게 넣음. 사람이 정하게 표시만(계산은 그대로)
                checks.append({'day': day, 'why': '늦게 출근한 날의 연장', 'in': _hm(rec['in']),
                               'out': _hm(rec['out']), 'pending': False})
            if any(c[k] for k in tot):
                days.append({'day': day, **c, 'in': _hm(rec['in']), 'out': _hm(rec['out']),
                             'reported': 'report' in (rec['in_src'], rec['out_src'])})
                for k in tot:
                    tot[k] += c[k]
            if c['meal']:
                meals.append(day)
        night = tot['early_night'] + tot['late_night']
        normal_h = (tot['early'] + tot['late']) / 60
        night_h = night / 60
        pay_normal = int(normal_h * wage * 1.5 + 0.5)
        pay_night = int(night_h * wage * 2 + 0.5)
        out[name] = {
            'days': days, 'checks': checks, 'meals': meals,
            'early_total': tot['early'] + tot['early_night'], 'late_total': tot['late'] + tot['late_night'],
            'night': night, 'normal_h': normal_h, 'night_h': night_h, 'wage': wage,
            'pay_normal': pay_normal, 'pay_night': pay_night, 'pay': pay_normal + pay_night,
        }
    return out


def _won(n: int) -> str:
    return f'{n:,}원'


def summary_text(ym: str, s: Dict[str, dict]) -> str:
    """경영지원 DM — SB 정산표와 같은 숫자(조출·연장 합산, 일반·야간 누계, 수당)."""
    y, m = ym[:4], int(ym[5:7])
    lines = [f'*🕒 {y}년 {m}월 시공팀 조출·연장 정산*', f'통상시급: {_won(s[next(iter(s))]["wage"]) if s else "-"}', '']
    for name, v in s.items():
        lines.append(f"*{name}* — 예상 수당 *{_won(v['pay'])}*")
        lines.append(f"• 조출 {v['early_total']}분 · 연장 {v['late_total']}분 · 야간 {v['night']}분")
        lines.append(f"• 일반 {v['normal_h']:g}시간 × 1.5 = {_won(v['pay_normal'])}"
                     f" / 야간 {v['night_h']:g}시간 × 2 = {_won(v['pay_night'])}")
        if v['meals']:
            lines.append(f"• 연장식대(270분 이상) {len(v['meals'])}일: "
                         + ', '.join(str(int(d[8:])) + '일' for d in v['meals']))
        if v['checks']:
            lines.append(f"• ⚠️ 확인 필요 {len(v['checks'])}일 — 찍힌 쪽만 계산했어요. 빠진 시각은 본인 보고로 채워져요:")
            for c in v['checks']:
                lines.append(f"   - {int(c['day'][5:7])}/{int(c['day'][8:])} {c['why']}"
                             f" (출근 {c['in'] or '-'} · 퇴근 {c['out'] or '-'})"
                             + (' — 본인 보고 승인 대기' if c['pending'] else ''))
        lines.append('')
    return '\n'.join(lines).strip()


def detail_text(ym: str, s: Dict[str, dict]) -> str:
    """날짜별 칸(SB 표 모양) — 코드 블록."""
    rows = []
    for name, v in s.items():
        rows.append(f'[{name}] 날짜   출근         퇴근          조출(야간)  연장(야간)')
        for d in v['days']:
            rows.append(f"{d['day'][5:]}  {d['in'] or '-':<12} {d['out'] or '-':<13} "
                        f"{d['early']:>3}({d['early_night']:>3})  {d['late']:>3}({d['late_night']:>3})"
                        + ('  ✍보고' if d['reported'] else ''))
        rows.append('')
    return '```' + '\n'.join(rows).rstrip() + '```'


# ── 현장 출퇴근 보고 ───────────────────────────────────────────────────
def add_report(name: str, day: str, in_hm: str, out_hm: str, note: str, reporter_id: str, reporter_name: str) -> dict:
    """같은 날 대기 중 보고가 있으면 새 보고로 바꾼다."""
    with _conn() as conn:
        conn.execute('UPDATE att_reports SET status=?, decided_reason=?, decided_at=? '
                     'WHERE name=? AND day=? AND status=?', (REJECTED, '새 보고로 바뀜', _now(), name, day, PENDING))
        cur = conn.execute(
            'INSERT INTO att_reports (name, day, in_hm, out_hm, note, status, reporter_id, reporter_name, created_at) '
            'VALUES (?,?,?,?,?,?,?,?,?)', (name, day, in_hm or None, out_hm or None, note or None, PENDING,
                                          reporter_id, reporter_name, _now()))
        return dict(conn.execute('SELECT * FROM att_reports WHERE id=?', (cur.lastrowid,)).fetchone())


def get_report(rid: int) -> Optional[dict]:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM att_reports WHERE id=?', (rid,)).fetchone()
        return dict(r) if r else None


def decide_report(rid: int, ok: bool, actor: str, reason: str = '') -> Optional[dict]:
    """대기 중일 때만(두 번 눌러도 한 번)."""
    with _conn() as conn:
        cur = conn.execute('UPDATE att_reports SET status=?, decided_by=?, decided_reason=?, decided_at=? '
                           'WHERE id=? AND status=?',
                           (APPROVED if ok else REJECTED, actor, reason or None, _now(), rid, PENDING))
        if cur.rowcount != 1:
            return None
        return dict(conn.execute('SELECT * FROM att_reports WHERE id=?', (rid,)).fetchone())


def set_report_refs(rid: int, refs: str) -> None:
    with _conn() as conn:
        conn.execute('UPDATE att_reports SET dm_refs=? WHERE id=?', (refs, rid))


def validate_report(in_hm: str, out_hm: str) -> Dict[str, str]:
    errs = {}
    if not in_hm and not out_hm:
        errs['att_in'] = '출근이나 퇴근 시각 중 하나는 넣어 주세요.'
    return errs


# ── 누락 묻기 ──────────────────────────────────────────────────────────
def missing_days(name: str, today: Optional[date] = None, back: int = 3) -> List[dict]:
    """어제부터 back 일 — 출근·퇴근 중 하나만 있는 날, 또는 평일인데 둘 다 없는 날(보고 대기·승인분 제외)."""
    today = today or date.today()
    from dashboard.services import leave as lv
    out = []
    for i in range(1, back + 1):
        d = today - timedelta(days=i)
        rec = day_record(name, d.isoformat())
        if rec['report']:
            continue
        try:   # 휴가(대기·승인)인 날은 묻지 않는다
            if any(r['ltype'] not in lv.HALF_TYPES for r in lv.overlaps(name, d, d)):
                continue
        except Exception:
            pass
        why = needs_check(rec)
        if not why and rec['in'] is None and rec['out'] is None and lv.is_workday(d):
            why = '출근·퇴근 기록 없음'
        if why and why != '퇴근이 다음 날 새벽':
            out.append({'day': d.isoformat(), 'why': why, 'in': _hm(rec['in']), 'out': _hm(rec['out'])})
    return out


def mark_asked(name: str, day: str) -> bool:
    """처음 묻는 날이면 True(같은 날 두 번 안 묻기)."""
    with _conn() as conn:
        cur = conn.execute('INSERT OR IGNORE INTO att_asks (name, day, asked_at) VALUES (?,?,?)', (name, day, _now()))
        return cur.rowcount == 1


A_REPORT, CB_REPORT = 'att_report_open', 'att_report_submit'
A_APPROVE, A_REJECT = 'att_report_approve', 'att_report_reject'
