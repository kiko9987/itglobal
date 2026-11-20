"""현장·미화 직원 휴대폰 등록 (2026-11-15, KiKO: 플렉스 앱처럼 한 번 로그인하면 계속).

메일·슬랙을 잘 안 보는 직원(미화 여사님)용. 흐름:
  1) 인사카드 관리자가 PC 인사카드에서 [로그인 번호 만들기] → 8자리 숫자(7일·한 번만) → [슬랙 DM으로 보내기] 또는 1:1 로 전달.
     (11-20 부장님: 현장은 바빠서 10분 안에 넣기 어렵다 → 미리 만들어 두고 시간 날 때. 기간이 길어진 만큼 6→8자리)
  2) 직원 휴대폰의 홈 화면 앱(또는 브라우저) 로그인 화면 「번호로 로그인」 칸에 그 번호.
  3) 그 휴대폰에 '기기 열쇠' 쿠키(gw_device, 원문은 DB 에 sha256 만)를 심고 현장 세션을 연다.
     그 뒤로는 세션이 끝나도 열쇠로 다시 연다 — 마지막 사용 뒤 IDLE_DAYS 일 안 쓰면 끝.

끊기는 경우: 인사카드에서 [끊기] · 그 휴대폰에서 로그아웃 · 퇴사나 로그인 방식 변경(현장 아님) ·
인사카드 슬랙 ID 가 등록 때와 달라짐 · IDLE_DAYS 동안 안 씀.
숫자 맞히기 막기: 같은 주소에서 15분에 5번, 전체 15분에 30번 틀리면 잠시 막는다(숫자는 1억 가지, 7일 유효 —
하루 최대 2,880번 시도해도 살아 있는 번호 하나를 맞힐 확률은 7일에 0.02% 안팎). 단톡방에 올리지 말 것(번호=그 사람 로그인).
"""
from __future__ import annotations

import hashlib
import secrets
from contextlib import contextmanager
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple

from dashboard.services import hr
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

CODE_TTL_DAYS = 7             # 미리 만들어 두고 시간 날 때 넣게(11-20)
CODE_LEN = 8
IDLE_DAYS = 90                 # 마지막 사용 뒤 이만큼 안 쓰면 다시 등록
TOUCH_EVERY = timedelta(hours=1)   # 마지막 사용 시각·쿠키 기한은 한 시간에 한 번만 고친다
COOKIE = 'gw_device'
FAIL_WINDOW = timedelta(minutes=15)
FAIL_PER_IP, FAIL_ALL = 5, 30

_DDL = [
    '''CREATE TABLE IF NOT EXISTS gw_device_codes (
        code_hash TEXT PRIMARY KEY,
        person_id INTEGER NOT NULL,
        created_by TEXT,
        created_at TEXT NOT NULL,
        expires_at TEXT NOT NULL,
        used_at TEXT,
        used_ip TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS gw_devices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        token_hash TEXT NOT NULL UNIQUE,
        person_id INTEGER NOT NULL,
        slack_id TEXT,
        label TEXT,
        created_by TEXT,
        created_at TEXT NOT NULL,
        last_used_at TEXT NOT NULL,
        last_ip TEXT,
        revoked_at TEXT,
        revoked_by TEXT
    )''',
    'CREATE INDEX IF NOT EXISTS idx_gw_devices_person ON gw_devices(person_id)',
    '''CREATE TABLE IF NOT EXISTS gw_device_fails (
        ip TEXT,
        at TEXT NOT NULL
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


def _hash(s: str) -> str:
    return hashlib.sha256(s.encode()).hexdigest()


def _fmt(dt: datetime) -> str:
    return dt.strftime('%Y-%m-%d %H:%M:%S')


def _parse(s: str) -> Optional[datetime]:
    try:
        return datetime.strptime(str(s)[:19], '%Y-%m-%d %H:%M:%S')
    except (TypeError, ValueError):
        return None


def _field_ok(p: Optional[dict]) -> bool:
    return bool(p) and p.get('status') != hr.STATUS_RESIGNED and p.get('access') == hr.ACCESS_FIELD


def device_label(user_agent: str) -> str:
    """'아이폰 · 사파리'처럼 사람이 알아볼 이름 — 목록에서 어느 휴대폰인지 구분용."""
    ua = user_agent or ''
    if 'iPhone' in ua:
        dev = '아이폰'
    elif 'iPad' in ua:
        dev = '아이패드'
    elif 'Android' in ua:
        dev = '안드로이드'
    elif 'Windows' in ua:
        dev = '윈도 PC'
    elif 'Macintosh' in ua:
        dev = '맥'
    else:
        dev = '기기'
    if 'SamsungBrowser' in ua:
        br = '삼성 인터넷'
    elif 'Edg/' in ua:
        br = '엣지'
    elif 'CriOS' in ua or 'Chrome/' in ua:
        br = '크롬'
    elif 'Safari' in ua:
        br = '사파리'
    else:
        br = ''
    return f'{dev} · {br}' if br else dev


# ── 숫자 만들기(인사카드 관리자) ──────────────────────────────────────
def issue_code(person_id: int, by: str, now: Optional[datetime] = None) -> Dict:
    """6자리 숫자 — 그 사람의 아직 안 쓴 숫자는 지운다. 현장 재직자가 아니면 ValueError."""
    p = hr.get_person(int(person_id))
    if not _field_ok(p):
        raise ValueError('로그인 방식이 현장인 재직자만 로그인 번호를 만들 수 있어요.')
    if not (p.get('slack_id') or '').strip():
        raise ValueError('인사카드에 슬랙 ID 가 있어야 휴가를 신청할 수 있어요. 슬랙 ID 를 먼저 넣어 주세요.')
    now = now or datetime.now()
    with _conn() as conn:
        conn.execute('DELETE FROM gw_device_codes WHERE (person_id = ? AND used_at IS NULL) OR expires_at < ?',
                     (p['id'], _fmt(now - timedelta(days=1))))
        for _ in range(20):   # 지금 살아 있는 다른 숫자와 겹치지 않게
            code = f'{secrets.randbelow(10 ** CODE_LEN):0{CODE_LEN}d}'
            try:
                conn.execute('INSERT INTO gw_device_codes (code_hash, person_id, created_by, created_at, expires_at) '
                             'VALUES (?,?,?,?,?)', (_hash(code), p['id'], by, _fmt(now),
                                                    _fmt(now + timedelta(days=CODE_TTL_DAYS))))
                break
            except Exception:
                continue
        else:
            raise RuntimeError('번호를 만들지 못했어요. 다시 눌러 주세요.')
    logger.info(f'[GW-DEVICE] 등록 숫자 발급 #{p["id"]} by {by}')
    return {'code': code, 'expires_at': _fmt(now + timedelta(days=CODE_TTL_DAYS)), 'days': CODE_TTL_DAYS}


def pending_code(person_id: int, now: Optional[datetime] = None) -> Optional[dict]:
    """아직 안 쓴 번호(있으면) — {'created_at','expires_at'}. 번호 자체는 남기지 않아 다시 보여 줄 수 없다."""
    now = now or datetime.now()
    with _conn() as conn:
        r = conn.execute('SELECT created_at, expires_at FROM gw_device_codes WHERE person_id = ? AND used_at IS NULL '
                         'AND expires_at >= ? ORDER BY created_at DESC LIMIT 1', (int(person_id), _fmt(now))).fetchone()
    return dict(r) if r else None


def is_pending(person_id: int, code: str, now: Optional[datetime] = None) -> bool:
    """이 번호가 그 사람의 아직 안 쓴 번호인가 — 슬랙 DM 보내기 전에 확인."""
    now = now or datetime.now()
    code = ''.join(ch for ch in str(code or '') if ch.isdigit())
    with _conn() as conn:
        r = conn.execute('SELECT 1 FROM gw_device_codes WHERE code_hash = ? AND person_id = ? AND used_at IS NULL '
                         'AND expires_at >= ?', (_hash(code), int(person_id), _fmt(now))).fetchone()
    return bool(r)


# ── 숫자 넣기(직원 휴대폰) ────────────────────────────────────────────
def _locked(conn, ip: str, now: datetime) -> bool:
    since = _fmt(now - FAIL_WINDOW)
    conn.execute('DELETE FROM gw_device_fails WHERE at < ?', (_fmt(now - timedelta(days=1)),))
    mine = conn.execute('SELECT COUNT(*) FROM gw_device_fails WHERE ip = ? AND at >= ?', (ip, since)).fetchone()[0]
    every = conn.execute('SELECT COUNT(*) FROM gw_device_fails WHERE at >= ?', (since,)).fetchone()[0]
    return mine >= FAIL_PER_IP or every >= FAIL_ALL


def redeem(code: str, ip: str = '', user_agent: str = '',
           now: Optional[datetime] = None) -> Tuple[Optional[dict], str, str]:
    """(인사카드, 기기 열쇠 원문, 못 하면 문구). 성공이면 문구가 ''."""
    now = now or datetime.now()
    code = ''.join(ch for ch in str(code or '') if ch.isdigit())
    ip = (ip or '')[:64]
    with _conn() as conn:
        if _locked(conn, ip, now):
            logger.warning(f'[GW-DEVICE] 숫자 시도 잠김 ip={ip}')
            return None, '', '여러 번 틀려서 잠시 막혔어요. 15분 뒤에 다시 해 주세요.'
        row = None
        if len(code) in (6, CODE_LEN):
            cur = conn.execute('UPDATE gw_device_codes SET used_at = ?, used_ip = ? '
                               'WHERE code_hash = ? AND used_at IS NULL AND expires_at >= ?',
                               (_fmt(now), ip, _hash(code), _fmt(now)))
            if cur.rowcount == 1:
                row = conn.execute('SELECT person_id, created_by FROM gw_device_codes WHERE code_hash = ?',
                                   (_hash(code),)).fetchone()
        if not row:
            conn.execute('INSERT INTO gw_device_fails (ip, at) VALUES (?, ?)', (ip, _fmt(now)))
            return None, '', '번호가 맞지 않거나 기한이 지났어요. 경영지원에서 새 번호를 받아 주세요.'
    p = hr.get_person(row[0])
    if not _field_ok(p):
        return None, '', '이 번호로는 로그인할 수 없어요. 경영지원에 문의해 주세요.'
    token = secrets.token_urlsafe(32)
    with _conn() as conn:
        conn.execute('INSERT INTO gw_devices (token_hash, person_id, slack_id, label, created_by, created_at, '
                     'last_used_at, last_ip) VALUES (?,?,?,?,?,?,?,?)',
                     (_hash(token), p['id'], (p.get('slack_id') or '').strip(), device_label(user_agent)[:60],
                      row[1], _fmt(now), _fmt(now), ip))
    logger.info(f'[GW-DEVICE] 휴대폰 등록 #{p["id"]} ({device_label(user_agent)}) ip={ip}')
    return p, token, ''


# ── 열쇠 확인(요청마다) ──────────────────────────────────────────────
def check(token: str, now: Optional[datetime] = None) -> Optional[dict]:
    """기기 열쇠 → {device, person} 또는 None(없음·끊김·오래 안 씀·퇴사·현장 아님·슬랙 ID 바뀜)."""
    if not token or len(token) > 200:
        return None
    now = now or datetime.now()
    with _conn() as conn:
        r = conn.execute('SELECT * FROM gw_devices WHERE token_hash = ? AND revoked_at IS NULL',
                         (_hash(token),)).fetchone()
    if not r:
        return None
    d = dict(r)
    last = _parse(d['last_used_at'])
    if not last or now - last > timedelta(days=IDLE_DAYS):
        return None
    p = hr.get_person(d['person_id'])
    if not _field_ok(p) or (p.get('slack_id') or '').strip() != (d.get('slack_id') or ''):
        return None
    return {'device': d, 'person': p}


def needs_touch(device: dict, now: Optional[datetime] = None) -> bool:
    last = _parse(device.get('last_used_at'))
    return not last or (now or datetime.now()) - last >= TOUCH_EVERY


def touch(device_id: int, ip: str = '', now: Optional[datetime] = None) -> None:
    with _conn() as conn:
        conn.execute('UPDATE gw_devices SET last_used_at = ?, last_ip = ? WHERE id = ?',
                     (_fmt(now or datetime.now()), (ip or '')[:64], device_id))


def revoke(device_id: int, by: str, person_id: Optional[int] = None) -> bool:
    q, args = 'UPDATE gw_devices SET revoked_at = ?, revoked_by = ? WHERE id = ? AND revoked_at IS NULL', \
        [_fmt(datetime.now()), by, int(device_id)]
    if person_id is not None:
        q += ' AND person_id = ?'
        args.append(int(person_id))
    with _conn() as conn:
        ok = conn.execute(q, args).rowcount == 1
    if ok:
        logger.info(f'[GW-DEVICE] 휴대폰 끊기 device#{device_id} by {by}')
    return ok


def revoke_token(token: str, by: str) -> None:
    if not token:
        return
    with _conn() as conn:
        conn.execute('UPDATE gw_devices SET revoked_at = ?, revoked_by = ? WHERE token_hash = ? AND revoked_at IS NULL',
                     (_fmt(datetime.now()), by, _hash(token)))


def list_for(person_id: int, now: Optional[datetime] = None) -> List[dict]:
    """인사카드 '로그인된 휴대폰' — 끊긴 것·오래 안 쓴 것은 빼고 최근 사용 순."""
    now = now or datetime.now()
    since = _fmt(now - timedelta(days=IDLE_DAYS))
    with _conn() as conn:
        rows = conn.execute('SELECT id, label, created_at, created_by, last_used_at FROM gw_devices '
                            'WHERE person_id = ? AND revoked_at IS NULL AND last_used_at >= ? '
                            'ORDER BY last_used_at DESC', (int(person_id), since)).fetchall()
    return [dict(r) for r in rows]
