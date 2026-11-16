# -*- coding: utf-8 -*-
"""Redis 에만 있던 업무 기록의 DB 사본 + 'Redis 가 비었다' 감지 (2026-11-16 상태 머신 점검 묶음 D).

왜: 입금 문자 원문·지정(sms_intake:*), 계산서 요청 카드(invoice_card:*), 금액 수정·매출 0원 요청
(project_amount_req:* / project_cancel_money_req:*)은 Redis 에만 있었다. Redis 가 비면(07-24 flush 같은 사고,
AOF 없는 재생성) 버튼이 '이미 처리된 입금'으로 막히거나 계산서 첨부가 조용히 무시되고, 대기 요청이 목록에서
사라졌다. 그런데 Redis 가 '살아 있는데 비어 있는' 상태는 PING 으로는 알 수 없었다.

어떻게:
  - 쓰는 곳이 Redis 와 함께 put(), 지우는 곳이 delete() 를 부른다(users.db `redis_durable` 표).
  - Redis 에 세대 표시(EPOCH_KEY)를 두고 DB 에도 같은 값을 적어 둔다. Redis 에 표시가 없는데 DB 에는 있으면
    = Redis 가 비워진 것. 그때 DB 사본 중 아직 유효한 것을 Redis 에 되채운다(restore) → 기존 코드는 그대로 동작.
  - 평소엔 DB 사본을 읽지 않는다. 처리 끝나 Redis 에서 지운 기록을 사본이 되살리지 않게(두 번 기록 방지).

pytest 중에는 꺼진다(DURABLE_KV_TEST=1 일 때만) — 테스트가 운영 users.db 에 쓰지 않게.
"""
import os
import time
import uuid
from contextlib import contextmanager
from typing import Dict, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

EPOCH_KEY = 'sys:redis_epoch'
PROTECTED_PREFIXES = ('sms_intake:', 'invoice_card:', 'project_amount_req:', 'project_cancel_money_req:')

_DDL = [
    '''CREATE TABLE IF NOT EXISTS redis_durable (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL,
        expires_at INTEGER NOT NULL,
        updated_at INTEGER NOT NULL
    )''',
    '''CREATE TABLE IF NOT EXISTS sys_markers (
        name TEXT PRIMARY KEY,
        value TEXT,
        updated_at INTEGER
    )''',
]
_table_ready = False


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


def enabled() -> bool:
    if 'PYTEST_CURRENT_TEST' in os.environ or os.getenv('FLASK_ENV') == 'testing':
        return os.getenv('DURABLE_KV_TEST', '') == '1'
    return os.getenv('DURABLE_KV_ENABLED', '1').strip() not in ('0', 'false', 'off')


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


def _protected(key: str) -> bool:
    return any(key.startswith(p) for p in PROTECTED_PREFIXES) and not key.endswith(':proc')


def put(key: str, value, ttl_sec: int) -> None:
    """Redis 에 쓴 값의 사본. 실패해도 예외를 올리지 않는다(Redis 쓰기가 본 기록)."""
    if not enabled() or not _protected(key):
        return
    if isinstance(value, bytes):
        value = value.decode('utf-8')
    now = int(time.time())
    try:
        with _conn() as conn:
            conn.execute('INSERT INTO redis_durable (key, value, expires_at, updated_at) VALUES (?,?,?,?) '
                         'ON CONFLICT(key) DO UPDATE SET value=excluded.value, expires_at=excluded.expires_at, '
                         'updated_at=excluded.updated_at', (key, str(value), now + int(ttl_sec), now))
            conn.commit()
    except Exception as exc:
        logger.error(f'[DURABLE] 사본 저장 실패 ({key}): {exc}')


def delete(*keys: str) -> None:
    """Redis 에서 지운 키의 사본도 지운다 — 남으면 Redis 가 비었을 때 처리 끝난 기록이 되살아난다."""
    keys = [k for k in keys if k and _protected(k)]
    if not enabled() or not keys:
        return
    try:
        with _conn() as conn:
            conn.executemany('DELETE FROM redis_durable WHERE key=?', [(k,) for k in keys])
            conn.commit()
    except Exception as exc:
        logger.error(f'[DURABLE] 사본 삭제 실패 ({keys}): {exc} — Redis 가 비면 이 기록이 되살아날 수 있음',
                     extra={'force_alert': True})


def add_once(key: str, ttl_sec: int) -> bool:
    """Redis 를 못 쓸 때의 중복 표시(SET NX 대신) — 처음이면 True. 꺼져 있으면 True(막지 않음)."""
    if not enabled():
        return True
    now = int(time.time())
    try:
        with _conn() as conn:
            conn.execute('DELETE FROM redis_durable WHERE key=? AND expires_at < ?', (key, now))
            cur = conn.execute('INSERT OR IGNORE INTO redis_durable (key, value, expires_at, updated_at) '
                               'VALUES (?,?,?,?)', (key, '1', now + int(ttl_sec), now))
            conn.commit()
            return cur.rowcount == 1
    except Exception as exc:
        logger.warning(f'[DURABLE] 중복 표시 실패 ({key}): {exc}')
        return True


def _marker(conn, name: str) -> str:
    r = conn.execute('SELECT value FROM sys_markers WHERE name=?', (name,)).fetchone()
    return (r[0] if r else '') or ''


def _set_marker(conn, name: str, value: str) -> None:
    conn.execute('INSERT INTO sys_markers (name, value, updated_at) VALUES (?,?,?) '
                 'ON CONFLICT(name) DO UPDATE SET value=excluded.value, updated_at=excluded.updated_at',
                 (name, value, int(time.time())))


def check_and_restore(rc) -> Optional[Dict]:
    """Redis 세대 표시를 확인 — 비워졌으면 DB 사본을 되채우고 {'wiped': True, 'restored': n} 를 돌려준다.

    처음 설치(양쪽 다 없음)·평소(같음)는 None. Redis 표시가 DB 와 다르면(옛 백업으로 복구 등) 그 값을 기준으로
    맞추고 None. 여러 스레드가 동시에 불러도 SET NX 로 한 번만 되채운다."""
    if not enabled():
        return None
    cur = rc.get(EPOCH_KEY)
    if isinstance(cur, bytes):
        cur = cur.decode()
    with _conn() as conn:
        known = _marker(conn, 'redis_epoch')
        now = int(time.time())
        if cur:
            if cur != known:
                if known:   # 옛 백업(dump.rdb)으로 되돌렸거나 다른 Redis 에 붙음 — 최근 기록이 빠졌을 수 있다
                    logger.error('[DURABLE] Redis 세대 표시가 바뀜(옛 백업 복구·다른 Redis?) — 최근 입금 인입·계산서 '
                                 '요청 카드가 빠졌는지 확인 필요', extra={'force_alert': True})
                _set_marker(conn, 'redis_epoch', cur)
                conn.commit()
            if now % 50 == 0:      # 가끔 만료된 사본 정리
                conn.execute('DELETE FROM redis_durable WHERE expires_at < ?', (now,))
                conn.commit()
            return None
        token = uuid.uuid4().hex
        if not rc.set(EPOCH_KEY, token, nx=True):
            return None            # 다른 스레드가 방금 처리
        _set_marker(conn, 'redis_epoch', token)
        conn.commit()
        if not known:
            return None            # 처음 설치 — 비워진 게 아니다
        restored = 0
        rows = conn.execute('SELECT key, value, expires_at FROM redis_durable WHERE expires_at > ?',
                            (now,)).fetchall()
        for key, value, exp in rows:
            if not _protected(key):
                continue
            try:
                if rc.set(key, value, ex=max(1, int(exp) - now), nx=True):
                    restored += 1
            except Exception as exc:
                logger.error(f'[DURABLE] 되채우기 실패 ({key}): {exc}')
        logger.error(f'[DURABLE] Redis 가 비워진 것을 감지 — DB 사본 {restored}건 되채움', extra={'force_alert': True})
        return {'wiped': True, 'restored': restored}
