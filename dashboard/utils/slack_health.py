"""슬랙 API 응답 실패 감지·관리자 알림.

용도: `views.open`, `chat.postMessage`, `views.update` 등의 실패가
5분 이내 3건 이상 발생하면 관리자에게 슬랙 DM 발송.

사용법:
    from dashboard.utils.slack_health import record_slack_api_failure

    try:
        client.views_open(trigger_id=..., view=...)
    except Exception as exc:
        record_slack_api_failure('views.open', exc, context={'trigger_id': ...})
        raise

Redis 키:
    slack_api_failures:{api_name} - 최근 5분간 실패 timestamp 리스트 (ZSET)
    slack_api_alert_cooldown - 관리자 DM 쿨다운 (5분)
"""
import json
import logging
import os
import time
import urllib.request
from typing import Optional

logger = logging.getLogger(__name__)

FAILURE_WINDOW_SEC = 300           # 실패 카운트 윈도우 (5분)
FAILURE_THRESHOLD = 3              # 5분 이내 3건 이상 시 알림
ALERT_COOLDOWN_SEC = 300           # 알림 쿨다운 (5분)
_COOLDOWN_KEY = 'slack_api_alert_cooldown'


def _rc():
    """Redis 클라이언트 지연 로딩 (import cycle 방지)."""
    from dashboard.utils.redis_client import get_redis_client
    return get_redis_client()


def record_slack_api_failure(
    api_name: str,
    error: Exception,
    context: Optional[dict] = None,
) -> None:
    """슬랙 API 호출 실패를 기록. 임계값 초과 시 관리자에게 슬랙 DM.

    Args:
        api_name: 'views.open', 'chat.postMessage', 'views.update' 등
        error: 발생한 예외
        context: 추가 컨텍스트 (op_id, trigger_id, channel 등) — 로그·DM 에 포함
    """
    now = time.time()
    zkey = f'slack_api_failures:{api_name}'
    try:
        rc = _rc()
        # 실패 timestamp 등록 (score = ts, member = unique id)
        rc.zadd(zkey, {f'{now:.3f}': now})
        rc.expire(zkey, FAILURE_WINDOW_SEC * 2)
        # 윈도우 밖의 오래된 데이터 정리
        rc.zremrangebyscore(zkey, 0, now - FAILURE_WINDOW_SEC)
        # 현재 윈도우 내 카운트
        count = rc.zcount(zkey, now - FAILURE_WINDOW_SEC, now)
    except Exception:
        # Redis 문제 자체가 알림 대상이므로 여기서는 실패해도 무시
        return

    if count >= FAILURE_THRESHOLD:
        _maybe_send_admin_alert(api_name, count, error, context or {})


def _maybe_send_admin_alert(
    api_name: str,
    count: int,
    error: Exception,
    context: dict,
) -> None:
    """5분 쿨다운 준수하며 관리자에게 슬랙 DM."""
    try:
        rc = _rc()
        # SET NX EX — 쿨다운 안이면 알림 skip
        got = rc.set(_COOLDOWN_KEY, '1', nx=True, ex=ALERT_COOLDOWN_SEC)
        if not got:
            return
    except Exception:
        return

    token = os.getenv('SLACK_BOT_TOKEN', '').strip()
    admin_id = os.getenv('SLACK_ADMIN_CHANNEL', '').strip()
    if not token or not admin_id:
        logger.warning(
            f'[SLACK_HEALTH] {api_name} 실패 {count}건 감지, '
            f'SLACK_BOT_TOKEN/ADMIN_CHANNEL 미설정으로 알림 skip'
        )
        return

    ctx_lines = []
    for k, v in list(context.items())[:5]:
        ctx_lines.append(f'  {k}: `{str(v)[:80]}`')
    ctx_str = '\n'.join(ctx_lines) if ctx_lines else '  (컨텍스트 없음)'

    text = (
        f':warning: *슬랙 API 실패 급증*\n'
        f'API: `{api_name}`\n'
        f'최근 5분 실패: *{count}건* (임계값 {FAILURE_THRESHOLD})\n'
        f'최근 오류: `{str(error)[:200]}`\n'
        f'컨텍스트:\n{ctx_str}\n'
        f'_5분간 반복 알림 억제됨_'
    )
    payload = json.dumps({'channel': admin_id, 'text': text}).encode('utf-8')
    req = urllib.request.Request(
        'https://slack.com/api/chat.postMessage',
        data=payload,
        headers={
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json; charset=utf-8',
        },
    )
    try:
        urllib.request.urlopen(req, timeout=5).read()
        logger.info(f'[SLACK_HEALTH] {api_name} 실패 급증 알림 DM 전송')
    except Exception as exc:
        logger.warning(f'[SLACK_HEALTH] 관리자 DM 실패: {exc}')


def get_stats() -> dict:
    """관리자 모니터링용 실패 통계 조회."""
    try:
        rc = _rc()
        now = time.time()
        result = {}
        for key in rc.scan_iter('slack_api_failures:*'):
            key_str = key.decode() if isinstance(key, bytes) else key
            api = key_str.split(':', 1)[1]
            cnt = rc.zcount(key_str, now - FAILURE_WINDOW_SEC, now)
            if cnt > 0:
                result[api] = cnt
        return result
    except Exception as exc:
        return {'error': str(exc)}
