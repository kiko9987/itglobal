# -*- coding: utf-8 -*-
"""실시간 소켓 연결 뒤 세션 재확인 (2026-11-13 보안 점검 2차 #5).

소켓은 연결할 때 한 번만 세션을 본다(dashboard/__init__.py _socket_connect_auth). 그 뒤로는 HTTP 요청을 안 거쳐서,
이미 열린 탭은 4시간 세션 만료·비활성·퇴사·권한 변경(세션 무효화) 뒤에도 공사 행 방송을 계속 받았다.

연결할 때 (sid → 이메일·로그인 시각)을 적어 두고, 5분마다 HTTP 세션 검사(auth._validate_session)와 같은 기준으로
다시 본다. 어긋나면 그 연결만 끊는다 — 화면이 다시 연결을 시도하면 연결 검사에서 세션이 다시 확인된다.
DB 조회가 실패하면 끊지 않는다(DB 장애로 모든 화면 실시간이 끊기는 일 방지, 연결 검사는 그대로 막음).
"""
import os
import threading
from datetime import datetime, timedelta
from typing import Callable, Dict, List, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

SWEEP_SECONDS = 300

_lock = threading.Lock()
_conns: Dict[str, dict] = {}


def register(sid: str, email: str, login_time: Optional[str]) -> None:
    if not sid or not email:
        return
    with _lock:
        _conns[sid] = {'email': email, 'login_time': login_time or ''}


def unregister(sid: str) -> None:
    with _lock:
        _conns.pop(sid, None)


def tracked() -> Dict[str, dict]:
    with _lock:
        return dict(_conns)


def session_still_valid(email: str, login_time: str, now: Optional[datetime] = None, db=None) -> bool:
    """auth._validate_session 과 같은 기준(시간 초과·계정 삭제·비활성·퇴사·세션 무효화). 판단 못 하면 True."""
    now = now or datetime.now()
    hours = int(os.getenv('SESSION_TIMEOUT_HOURS', 4))
    lt = None
    if login_time:
        try:
            lt = datetime.fromisoformat(login_time)
        except (ValueError, TypeError):
            return False
        if now - lt > timedelta(hours=hours):
            return False
    try:
        if db is None:
            from dashboard.utils.user_database import get_user_database
            db = get_user_database()
        u = db.get_user_by_email(email)
    except Exception as exc:
        logger.warning(f"[SOCKET_GUARD] 사용자 조회 실패 → 유지: {exc}")
        return True
    if not u or not u.get('is_active', True) or u.get('is_resigned', False):
        return False
    inv = u.get('session_invalidate_after')
    if inv and lt is not None:
        try:
            if lt < datetime.fromisoformat(str(inv).replace(' ', 'T')):
                return False
        except (ValueError, TypeError):
            pass
    return True


def sweep(disconnect: Callable[[str], None], now: Optional[datetime] = None, db=None) -> List[str]:
    """세션이 끝난 연결을 끊는다. Returns 끊은 sid."""
    cut = []
    for sid, c in tracked().items():
        if session_still_valid(c['email'], c['login_time'], now=now, db=db):
            continue
        try:
            disconnect(sid)
        except Exception as exc:
            logger.warning(f"[SOCKET_GUARD] 연결 끊기 실패 ({sid}): {exc}")
        unregister(sid)
        cut.append(sid)
    if cut:
        logger.info(f"[SOCKET_GUARD] 세션 끝난 실시간 연결 {len(cut)}개 끊음")
    return cut


def start(socketio) -> None:
    """5분마다 sweep. socketio.server.disconnect 로 그 연결만 끊는다."""
    def _disc(sid):
        socketio.server.disconnect(sid, namespace='/')

    def _loop():
        while True:
            socketio.sleep(SWEEP_SECONDS)
            try:
                sweep(_disc)
            except Exception as exc:
                logger.warning(f"[SOCKET_GUARD] 점검 실패: {exc}")

    socketio.start_background_task(_loop)
    logger.info(f"[SOCKET_GUARD] 실시간 연결 세션 재확인 시작 ({SWEEP_SECONDS}초 간격)")
