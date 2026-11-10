"""현장·미화 직원 슬랙 로그인 링크 (2026-11-10).

회사 구글 계정이 없는 현장·미화 직원(인사카드 접근=field)이 슬랙 휴가 봇의 [그룹웨어 열기]를 누르면
10분 동안 한 번만 쓸 수 있는 링크를 만든다. 링크로 들어오면 사용자 표(users)에 계정을 만들지 않고
인사카드 기준의 '현장 세션'을 연다 → 내 정보·휴가만 보이고 기존 PM 화면·API 는 막힌다(groupware 블루프린트 문지기).

- 토큰 원문은 DB 에 남기지 않는다(sha256 만). 쓰면 바로 used_at 을 채워 다시 못 쓴다.
- 링크는 누른 본인에게만 보이는 슬랙 창(모달)으로만 준다.
"""
from __future__ import annotations

import hashlib
import os
import secrets
from contextlib import contextmanager
from datetime import datetime, timedelta
from typing import Optional

from dashboard.services import hr
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

TTL_MINUTES = 10

_DDL = [
    '''CREATE TABLE IF NOT EXISTS gw_login_tokens (
        token_hash TEXT PRIMARY KEY,
        person_id INTEGER NOT NULL,
        slack_id TEXT,
        created_at TEXT NOT NULL,
        expires_at TEXT NOT NULL,
        used_at TEXT,
        used_ip TEXT
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


def _hash(token: str) -> str:
    return hashlib.sha256(token.encode()).hexdigest()


def _fmt(dt: datetime) -> str:
    return dt.strftime('%Y-%m-%d %H:%M:%S')


def public_url() -> str:
    return os.getenv('GW_PUBLIC_URL', 'https://pm.itg-aircon.com').rstrip('/')


def field_person_for_slack(slack_id: str) -> Optional[dict]:
    """링크를 받을 수 있는 사람인가 — 재직 + 접근 field. 아니면 None."""
    p = hr.find_by_slack(slack_id)
    if not p or p.get('access') != hr.ACCESS_FIELD:
        return None
    return p


def can_issue(slack_id: str) -> bool:
    """홈 탭에 [그룹웨어 열기]를 보일지 — 현장 직원이고, 새 화면을 쓸 수 있을 때(정식 오픈 전엔 시범 명단)."""
    try:
        from dashboard.services import groupware as gw
        p = field_person_for_slack(slack_id)
        return bool(p) and gw.can_use(hr.identity_email(p))
    except Exception as e:
        logger.warning(f'[GW-LOGIN] 버튼 표시 판단 실패 → 숨김: {e}')
        return False


def issue(slack_id: str, now: Optional[datetime] = None) -> str:
    """로그인 링크 URL. 받을 수 없는 사람이면 PermissionError."""
    if not can_issue(slack_id):
        raise PermissionError('그룹웨어 로그인 링크를 받을 수 있는 계정이 아닙니다.')
    p = field_person_for_slack(slack_id)
    now = now or datetime.now()
    token = secrets.token_urlsafe(32)
    with _conn() as conn:
        # 오래된 줄 정리(하루 지난 것)
        conn.execute('DELETE FROM gw_login_tokens WHERE expires_at < ?', (_fmt(now - timedelta(days=1)),))
        conn.execute('INSERT INTO gw_login_tokens (token_hash, person_id, slack_id, created_at, expires_at) '
                     'VALUES (?, ?, ?, ?, ?)',
                     (_hash(token), p['id'], slack_id, _fmt(now), _fmt(now + timedelta(minutes=TTL_MINUTES))))
    logger.info(f'[GW-LOGIN] 링크 발급 #{p["id"]} ({slack_id})')
    return f'{public_url()}/gw/login/{token}'


def consume(token: str, ip: str = '', now: Optional[datetime] = None) -> Optional[dict]:
    """링크를 쓴다 — 유효하면 인사카드, 아니면 None(만료·사용함·없음·퇴사·접근 변경)."""
    if not token or len(token) > 200:
        return None
    now = now or datetime.now()
    h = _hash(token)
    with _conn() as conn:
        cur = conn.execute('UPDATE gw_login_tokens SET used_at = ?, used_ip = ? '
                           'WHERE token_hash = ? AND used_at IS NULL AND expires_at >= ?',
                           (_fmt(now), ip[:64], h, _fmt(now)))
        if cur.rowcount != 1:
            return None
        row = conn.execute('SELECT person_id FROM gw_login_tokens WHERE token_hash = ?', (h,)).fetchone()
    p = hr.get_person(row[0]) if row else None
    if not p or p.get('status') == hr.STATUS_RESIGNED or p.get('access') != hr.ACCESS_FIELD:
        return None
    return p


def session_user(p: dict) -> dict:
    """현장 세션의 session['user'] — permission_level 'field' 는 기존 권한 검사(관리자·편집자)를 모두 통과 못 함."""
    return {
        'email': hr.identity_email(p),
        'name': p['name'],
        'permission_level': 'field',
        'is_active': True,
        'auth_method': 'slack_link',
        'picture': '',
        'gw_field': True,
        'gw_person_id': p['id'],
    }


def field_session_valid(user: dict) -> bool:
    """세션 검사 때 — 인사카드가 아직 재직·field 인지."""
    try:
        p = hr.get_person(int(user.get('gw_person_id') or 0))
    except Exception:
        return False
    return bool(p) and p.get('status') != hr.STATUS_RESIGNED and p.get('access') == hr.ACCESS_FIELD
