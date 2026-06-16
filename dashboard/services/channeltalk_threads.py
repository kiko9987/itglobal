"""
채널톡 chatId ↔ 슬랙 thread_ts 매핑 저장소.

저장 매체: Redis (이미 인프라 있음, 영속성은 Redis RDB로)
Fallback: 메모리 dict (Redis 연결 실패 시)
"""

import threading
from typing import Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_MEM_MAP: dict = {}     # chat_id -> thread_ts
_MEM_REV: dict = {}     # thread_ts -> chat_id
_MEM_PENDING: dict = {} # chat_id -> {thread_ts, created_at, reminded}
_LOCK = threading.Lock()

REDIS_KEY_FWD = 'channeltalk:thread_map'   # chat_id -> thread_ts
REDIS_KEY_REV = 'channeltalk:chat_map'     # thread_ts -> chat_id
REDIS_KEY_PENDING = 'channeltalk:pending'  # chat_id -> JSON({thread_ts, created_at, reminded})


def _get_redis():
    """Redis client (text 모드: decode_responses=True). 실패 시 None → 메모리 fallback."""
    try:
        from dashboard.utils.redis_client import get_redis_client
        client = get_redis_client()
        if client.ping():
            return client.redis
    except Exception as exc:
        logger.debug(f'[ChannelTalk/Threads] Redis 사용 불가, 메모리 fallback: {exc}')
    return None


def get_thread_ts(chat_id: str) -> Optional[str]:
    """chat_id → slack thread_ts (없으면 None)"""
    if not chat_id:
        return None
    r = _get_redis()
    if r is not None:
        try:
            val = r.hget(REDIS_KEY_FWD, chat_id)
            if val:
                return val
        except Exception as exc:
            logger.warning(f'[ChannelTalk/Threads] Redis HGET fwd 실패: {exc}')
    with _LOCK:
        return _MEM_MAP.get(chat_id)


def get_chat_id(thread_ts: str) -> Optional[str]:
    """slack thread_ts → 채널톡 chat_id (역방향 lookup, 슬랙 답글 forward에 사용)"""
    if not thread_ts:
        return None
    r = _get_redis()
    if r is not None:
        try:
            val = r.hget(REDIS_KEY_REV, thread_ts)
            if val:
                return val
        except Exception as exc:
            logger.warning(f'[ChannelTalk/Threads] Redis HGET rev 실패: {exc}')
    with _LOCK:
        return _MEM_REV.get(thread_ts)


def set_thread_ts(chat_id: str, thread_ts: str) -> None:
    """양방향 매핑 저장 (chat_id ↔ thread_ts)"""
    if not chat_id or not thread_ts:
        return
    r = _get_redis()
    if r is not None:
        try:
            r.hset(REDIS_KEY_FWD, chat_id, thread_ts)
            r.hset(REDIS_KEY_REV, thread_ts, chat_id)
            return
        except Exception as exc:
            logger.warning(f'[ChannelTalk/Threads] Redis HSET 실패: {exc}')
    with _LOCK:
        _MEM_MAP[chat_id] = thread_ts
        _MEM_REV[thread_ts] = chat_id


def remove_thread_ts(chat_id: str) -> None:
    """매핑 양방향 제거 (채팅 종료 시)"""
    if not chat_id:
        return
    thread_ts = get_thread_ts(chat_id)
    r = _get_redis()
    if r is not None:
        try:
            r.hdel(REDIS_KEY_FWD, chat_id)
            if thread_ts:
                r.hdel(REDIS_KEY_REV, thread_ts)
        except Exception:
            pass
    with _LOCK:
        _MEM_MAP.pop(chat_id, None)
        if thread_ts:
            _MEM_REV.pop(thread_ts, None)


# ─────────────────────────────────────────────────────────────
# 미배정 알림 — pending chats (5분 경과 + 미응답 감지용)
# ─────────────────────────────────────────────────────────────
import json
import time


def add_pending(chat_id: str, thread_ts: str) -> None:
    """새 채팅을 pending 등록 — scheduler가 5분 후 응답 없으면 알림"""
    if not chat_id or not thread_ts:
        return
    entry = json.dumps({
        'thread_ts': thread_ts,
        'created_at': int(time.time()),
        'reminded': False,
    })
    r = _get_redis()
    if r is not None:
        try:
            r.hset(REDIS_KEY_PENDING, chat_id, entry)
            return
        except Exception:
            pass
    with _LOCK:
        _MEM_PENDING[chat_id] = json.loads(entry)


def remove_pending(chat_id: str) -> None:
    """직원이 응답하면 pending에서 제거"""
    if not chat_id:
        return
    r = _get_redis()
    if r is not None:
        try:
            r.hdel(REDIS_KEY_PENDING, chat_id)
        except Exception:
            pass
    with _LOCK:
        _MEM_PENDING.pop(chat_id, None)


def list_pending() -> list:
    """전체 pending 채팅 목록 반환 — scheduler용

    Returns:
        [{'chat_id', 'thread_ts', 'created_at', 'reminded'}, ...]
    """
    out = []
    r = _get_redis()
    if r is not None:
        try:
            data = r.hgetall(REDIS_KEY_PENDING)
            for chat_id, entry_str in data.items():
                try:
                    entry = json.loads(entry_str)
                    out.append({'chat_id': chat_id, **entry})
                except Exception:
                    continue
            return out
        except Exception:
            pass
    with _LOCK:
        for chat_id, entry in _MEM_PENDING.items():
            out.append({'chat_id': chat_id, **entry})
    return out


def mark_reminded(chat_id: str) -> None:
    """알림 발송 완료 표시 (중복 알림 방지)"""
    if not chat_id:
        return
    r = _get_redis()
    if r is not None:
        try:
            cur = r.hget(REDIS_KEY_PENDING, chat_id)
            if cur:
                entry = json.loads(cur)
                entry['reminded'] = True
                r.hset(REDIS_KEY_PENDING, chat_id, json.dumps(entry))
            return
        except Exception:
            pass
    with _LOCK:
        if chat_id in _MEM_PENDING:
            _MEM_PENDING[chat_id]['reminded'] = True
