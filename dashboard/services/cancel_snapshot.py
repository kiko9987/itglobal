"""공사 취소 스냅샷 — 재개 때 되돌릴 원래 값(공사 확정일·수금 확인·총액 1).

원래 Redis(`project_cancel_snapshot:{code}`, 1년)에만 있어, Redis 를 잃으면(07-24 flush 같은 사고) 재개 시
확정일이 오늘로, 총액 1(매출)은 0 그대로 남았다(2026-10-30 점검). 저장할 때 감사 기록(users.db,
action='CANCEL_SNAPSHOT')에도 같은 JSON 을 남기고, Redis 에 없으면 그 기록으로 복원한다.
감사 기록은 그 공사의 마지막 재개(RESUME_PROJECT)보다 나중 것만 쓴다 — 지난 취소의 값이 섞이지 않게.
"""
from __future__ import annotations

import json
from typing import Any, Dict, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_KEY = 'project_cancel_snapshot:'
_TTL = 60 * 60 * 24 * 365
ACTION = 'CANCEL_SNAPSHOT'


def _rc():
    from dashboard.utils.redis_client import get_redis_client
    return get_redis_client().redis


def payment_confirmed(raw: Any) -> bool:
    return (raw is True
            or (isinstance(raw, str) and raw.strip().upper() in ('TRUE', 'Y', 'YES', '1'))
            or raw == 1)


def save(code: str, snap: Dict[str, Any], by: str = '') -> None:
    """Redis + 감사 기록에 함께 저장. 둘 다 실패해도 취소 자체는 진행(경고만)."""
    body = json.dumps(snap, ensure_ascii=False)
    try:
        _rc().set(f'{_KEY}{code}', body, ex=_TTL)
    except Exception as exc:
        logger.warning(f'[취소 스냅샷] Redis 저장 실패 ({code}): {exc}')
    try:
        _audit_save(code, body, by)
    except Exception as exc:
        logger.warning(f'[취소 스냅샷] 감사 기록 저장 실패 ({code}): {exc}')


def _audit_save(code: str, body: str, by: str) -> None:
    from dashboard.utils.user_database import get_audit_repository
    get_audit_repository().log_action(
        user_email=by or 'system', action=ACTION, project_code=code,
        details='공사 취소 전 값 (재개 시 복원용)', new_value=body, ip_address='system')


def _from_audit(code: str) -> Optional[str]:
    from dashboard.utils.user_database import get_user_database
    with get_user_database()._get_connection() as conn:
        row = conn.execute(
            "SELECT new_value FROM audit_logs WHERE project_code = ? AND action = ? "
            "AND id > COALESCE((SELECT MAX(id) FROM audit_logs WHERE project_code = ? "
            "AND action = 'RESUME_PROJECT'), 0) ORDER BY id DESC LIMIT 1",
            (code, ACTION, code)).fetchone()
    return (row[0] or None) if row else None


def load_raw(code: str) -> Optional[str]:
    """저장된 스냅샷 문자열(JSON 또는 옛 '확정일' 문자열). Redis → 감사 기록 순. 없으면 None."""
    try:
        v = _rc().get(f'{_KEY}{code}')
        if v:
            return v.decode() if isinstance(v, bytes) else v
    except Exception as exc:
        logger.warning(f'[취소 스냅샷] Redis 조회 실패 ({code}): {exc}')
    try:
        v = _from_audit(code)
        if v:
            logger.warning(f'[취소 스냅샷] Redis 에 없어 감사 기록으로 복원 ({code})')
        return v
    except Exception as exc:
        logger.warning(f'[취소 스냅샷] 감사 기록 조회 실패 ({code}): {exc}')
        return None


def delete(code: str) -> None:
    """재개 후 Redis 정리. 감사 기록은 남는다(재개 기록보다 앞이라 다시 쓰이지 않음)."""
    try:
        _rc().delete(f'{_KEY}{code}')
    except Exception:
        pass
