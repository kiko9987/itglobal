# -*- coding: utf-8 -*-
"""실시간 소켓 연결 뒤 세션 재확인 (2026-11-13 보안 점검 2차 #5).

소켓은 연결할 때 한 번만 세션을 본다(dashboard/__init__.py _socket_connect_auth). 그 뒤로는 HTTP 요청을 안 거쳐서,
이미 열린 탭은 4시간 세션 만료·비활성·퇴사·권한 변경(세션 무효화) 뒤에도 공사 행 방송을 계속 받았다.

연결할 때 (sid → 이메일·로그인 시각)을 적어 두고, 5분마다 HTTP 세션 검사(auth._validate_session)와 같은 기준으로
다시 본다. 어긋나면 그 연결만 끊는다 — 화면이 다시 연결을 시도하면 연결 검사에서 세션이 다시 확인된다.
DB 조회가 실패하면 끊지 않는다(DB 장애로 모든 화면 실시간이 끊기는 일 방지, 연결 검사는 그대로 막음).

연결은 그것을 연 HTTP 세션(session['session_id'])에 묶는다(2026-11-15 5차 #5): 연장(touch)은 같은 세션 연결만,
로그아웃(revoke_session)은 그 세션 연결을 바로 끊는다. 관리자 세션 무효화 뒤엔 옛 세션이 연장을 못 해
login_time 이 session_invalidate_after 보다 옛 값으로 남으므로 sweep 이 끊는다.
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
_disconnect: Optional[Callable[[str], None]] = None   # start() 가 채움 — 로그아웃 때 바로 끊기용


def register(sid: str, email: str, login_time: Optional[str], session_id: Optional[str] = None) -> None:
    """session_id = 그 연결을 연 HTTP 세션(session['session_id']). 연장·로그아웃을 그 세션 기준으로 한다."""
    if not sid or not email:
        return
    with _lock:
        _conns[sid] = {'email': email, 'login_time': login_time or '', 'session_id': session_id or '',
                       'revoked': False}


def touch(email: str, login_time: Optional[str], session_id: Optional[str] = None) -> int:
    """화면이 세션을 연장하면(/api/auth/ping) 그 세션이 연 연결의 기준 시각도 같이 — 안 하면 연결한 지 4시간이 지난
    탭은 세션이 살아 있어도 끊겨, 새로고침 전까지 실시간 갱신이 조용히 멈췄다(11-15).

    같은 세션 연결만 연장한다(2026-11-15 5차 점검 #5). 예전엔 같은 메일 연결을 모두 연장해, 다른 탭·브라우저를
    쓰는 동안 로그아웃한 탭의 연결도 끊기지 않고 공사 취소·재개 행(금액·주소·연락처)을 계속 받았다.
    session_id 가 없으면 아무것도 연장하지 않는다(그 연결은 4시간 뒤 끊기고 화면이 다시 연결). Returns 갱신한 연결 수."""
    if not email or not login_time or not session_id:
        return 0
    n = 0
    with _lock:
        for c in _conns.values():
            if (c.get('session_id') == session_id and c.get('email') == email and not c.get('revoked')
                    and (c.get('login_time') or '') < login_time):
                c['login_time'] = login_time
                n += 1
    return n


def revoke_session(session_id: Optional[str]) -> List[str]:
    """로그아웃한 세션이 연 연결을 바로 끊는다(2026-11-15 5차 #5). 끊기 함수가 아직 없으면(테스트·시작 전)
    표시만 해 두고 다음 sweep 이 끊는다. Returns 해당 sid."""
    if not session_id:
        return []
    with _lock:
        sids = [sid for sid, c in _conns.items() if c.get('session_id') == session_id]
        for sid in sids:
            _conns[sid]['revoked'] = True
    disc = _disconnect
    if disc is not None:
        for sid in sids:
            try:
                disc(sid)
            except Exception as exc:
                logger.warning(f"[SOCKET_GUARD] 로그아웃 연결 끊기 실패 ({sid}): {exc}")
                continue
            unregister(sid)
    if sids:
        logger.info(f"[SOCKET_GUARD] 로그아웃한 세션의 실시간 연결 {len(sids)}개 끊음")
    return sids


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
        if not c.get('revoked') and session_still_valid(c['email'], c['login_time'], now=now, db=db):
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
    global _disconnect

    def _disc(sid):
        socketio.server.disconnect(sid, namespace='/')

    _disconnect = _disc

    def _loop():
        while True:
            socketio.sleep(SWEEP_SECONDS)
            try:
                sweep(_disc)
            except Exception as exc:
                logger.warning(f"[SOCKET_GUARD] 점검 실패: {exc}")

    socketio.start_background_task(_loop)
    logger.info(f"[SOCKET_GUARD] 실시간 연결 세션 재확인 시작 ({SWEEP_SECONDS}초 간격)")
