"""Google Sheets write-behind 지속 큐 (2026-07-09).

목적: 매니저가 관리 사이트에서 편집·취소·재개 등 mutate 요청 → 캐시에 즉시 반영 +
      즉시 응답. 실제 Google Sheets write 는 큐로 넘겨 백그라운드 워커가 처리.
      Google 지연·재시작에 영향 받지 않는 안정성 확보.

큐 구조 (Redis List):
  - `sheet_write_queue`         : 처리 대기 op JSON 리스트 (FIFO)
  - `sheet_write_processing`    : 현재 처리 중 op (프로세스 죽으면 여기 남음)
  - `sheet_write_failed`        : 3회 재시도 후 실패한 op (데드레터, 관리자 확인)

각 op:
  {
    'op_id': str,           # UUID
    'op_type': str,         # 등록된 핸들러 이름
    'payload': dict,        # 핸들러 인자
    'meta': dict,           # user_email, ip 등 감사용
    'created_at': float,    # time.time()
    'attempts': int,
    'last_error': str,
  }

사용:
  1) 핸들러 등록: `@register('project_edit')` 함수 정의
  2) 큐잉: `enqueue('project_edit', payload={...}, meta={...})`
  3) 워커 시작: `start_worker()` — Flask 부팅 시 자동 호출됨
"""
from __future__ import annotations

import json
import threading
import time
import uuid
from typing import Any, Callable, Dict, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)


QUEUE_KEY = 'sheet_write_queue'
PROCESSING_KEY = 'sheet_write_processing'
FAILED_KEY = 'sheet_write_failed'
MAX_ATTEMPTS = 3

# 큐 depth 알림 임계값 (2026-07-09)
DEPTH_ALERT_THRESHOLD = 100
DEPTH_ALERT_COOLDOWN_SEC = 300  # 5분마다 최대 1회 알림
_last_depth_alert_ts = 0.0

# 처리 통계 (매 시간 rollup)
_stats_lock = threading.Lock()
_processed_since_rollup = 0
_last_rollup_ts = time.time()
ROLLUP_INTERVAL_SEC = 3600

# op_type → handler(payload: dict) -> None
_registry: Dict[str, Callable[[dict], None]] = {}
_worker_started = False
_worker_lock = threading.Lock()


def register(op_type: str) -> Callable:
    """핸들러 등록 데코레이터."""
    def deco(fn: Callable[[dict], None]) -> Callable[[dict], None]:
        _registry[op_type] = fn
        logger.debug(f'[QUEUE] handler 등록: {op_type}')
        return fn
    return deco


def _rc():
    from dashboard.utils.redis_client import get_redis_client
    return get_redis_client().redis


def _maybe_alert_depth() -> None:
    """큐 depth 가 임계값 초과 시 관리자 DM (쿨다운 있음)."""
    global _last_depth_alert_ts
    try:
        depth = _rc().llen(QUEUE_KEY)
        if depth < DEPTH_ALERT_THRESHOLD:
            return
        now = time.time()
        if now - _last_depth_alert_ts < DEPTH_ALERT_COOLDOWN_SEC:
            return
        _last_depth_alert_ts = now
        # 관리자 슬랙 DM
        import os as _os
        import urllib.request as _urllib_req
        token = _os.getenv('SLACK_BOT_TOKEN', '').strip()
        admin_id = _os.getenv('SLACK_ADMIN_CHANNEL', '').strip()
        if not token or not admin_id:
            return
        text = (
            f':warning: *시트 write 큐 정체 감지*\n'
            f'pending: {depth}건 (임계값 {DEPTH_ALERT_THRESHOLD})\n'
            f'Google Sheets 지연 또는 워커 병목 가능성. `/api/admin/sheet-write-queue` 확인 필요.'
        )
        req = _urllib_req.Request(
            'https://slack.com/api/chat.postMessage',
            data=json.dumps({'channel': admin_id, 'text': text}).encode('utf-8'),
            headers={
                'Content-Type': 'application/json; charset=utf-8',
                'Authorization': f'Bearer {token}',
            },
        )
        _urllib_req.urlopen(req, timeout=5).read()
        logger.warning(f'[QUEUE] depth 알림 발송: {depth}건')
    except Exception as _exc:
        logger.debug(f'[QUEUE] depth 알림 실패: {_exc}')


def _bump_processed_and_maybe_rollup() -> None:
    """처리 카운트 증가. 1시간마다 통계 rollup 로그."""
    global _processed_since_rollup, _last_rollup_ts
    with _stats_lock:
        _processed_since_rollup += 1
        now = time.time()
        elapsed = now - _last_rollup_ts
        if elapsed >= ROLLUP_INTERVAL_SEC:
            count = _processed_since_rollup
            per_min = count / (elapsed / 60) if elapsed > 0 else 0
            logger.info(
                f'[QUEUE/STATS] 지난 {int(elapsed/60)}분 처리: {count}건 '
                f'({per_min:.1f}건/분). 현재 pending: {_rc().llen(QUEUE_KEY)}건'
            )
            _processed_since_rollup = 0
            _last_rollup_ts = now


def enqueue(op_type: str, payload: dict, meta: Optional[dict] = None) -> str:
    """큐에 write op 등록. 등록 즉시 op_id 반환. 워커가 비동기 처리."""
    if op_type not in _registry:
        logger.warning(f'[QUEUE] 미등록 op_type: {op_type} — enqueue 진행 (부팅 순서 문제 가능)')
    op = {
        'op_id': str(uuid.uuid4()),
        'op_type': op_type,
        'payload': payload,
        'meta': meta or {},
        'created_at': time.time(),
        'attempts': 0,
        'last_error': '',
    }
    try:
        _rc().rpush(QUEUE_KEY, json.dumps(op, ensure_ascii=False))
        logger.info(f'[QUEUE] enqueue: {op_type} (op_id={op["op_id"][:8]}…)')
        _maybe_alert_depth()  # depth 임계값 알림
    except Exception as exc:
        logger.error(f'[QUEUE] enqueue 실패 ({op_type}): {exc}', exc_info=True)
    return op['op_id']


def get_stats() -> dict:
    """큐 상태 조회 (관리자 모니터링용)."""
    try:
        rc = _rc()
        return {
            'pending': rc.llen(QUEUE_KEY),
            'processing': rc.llen(PROCESSING_KEY),
            'failed': rc.llen(FAILED_KEY),
        }
    except Exception as exc:
        return {'error': str(exc)}


def peek_failed(limit: int = 20) -> list:
    """실패 데드레터 조회 (관리자 확인용)."""
    try:
        rc = _rc()
        raw = rc.lrange(FAILED_KEY, 0, limit - 1)
        return [json.loads(x) for x in raw if x]
    except Exception:
        return []


def retry_failed(op_id: str) -> bool:
    """실패 op 재큐잉. 관리자가 데드레터에서 재시도할 때."""
    rc = _rc()
    for raw in rc.lrange(FAILED_KEY, 0, -1):
        try:
            op = json.loads(raw)
        except Exception:
            continue
        if op.get('op_id') == op_id:
            op['attempts'] = 0
            op['last_error'] = ''
            rc.lrem(FAILED_KEY, 1, raw)
            rc.rpush(QUEUE_KEY, json.dumps(op, ensure_ascii=False))
            logger.info(f'[QUEUE] failed→queue 재시도: op_id={op_id[:8]}…')
            return True
    return False


def _notify_admin_deadletter(op: dict, last_exc: Exception) -> None:
    """데드레터 발생 시 관리자에게 슬랙 DM + 편집자 본인 브라우저 알림.

    - 관리자(KiKO): SLACK_ADMIN_CHANNEL 로 슬랙 DM (기술 상세 포함)
    - 편집자 본인: SocketIO broadcast (프론트가 user_email 매칭해 시스템 알림 표시)
      매니저는 슬랙 DM 받아도 할 게 없어서 화면 알림만 노출
    """
    op_type = op.get('op_type', '?')
    op_id = op.get('op_id', '')[:8]
    tag = op.get('payload', {}).get('tag', '') or op.get('payload', {}).get('project_code', '')
    user_email = (op.get('meta') or {}).get('user_email', '')

    # 1) 관리자 슬랙 DM (기존 유지)
    try:
        import os as _os
        import urllib.request as _urllib_req
        token = _os.getenv('SLACK_BOT_TOKEN', '').strip()
        admin_id = _os.getenv('SLACK_ADMIN_CHANNEL', '').strip()
        if token and admin_id:
            admin_text = (
                f':rotating_light: *시트 write 큐 데드레터*\n'
                f'op_type: `{op_type}`\n'
                f'op_id: `{op_id}…`\n'
                f'tag: `{tag}`\n'
                f'attempts: {op.get("attempts", 0)}\n'
                f'user: `{user_email or "unknown"}`\n'
                f'error: `{str(last_exc)[:400]}`\n'
                f'재시도: `POST /api/admin/sheet-write-queue/retry/{op.get("op_id", "")}`'
            )
            req = _urllib_req.Request(
                'https://slack.com/api/chat.postMessage',
                data=json.dumps({'channel': admin_id, 'text': admin_text}).encode('utf-8'),
                headers={
                    'Content-Type': 'application/json; charset=utf-8',
                    'Authorization': f'Bearer {token}',
                },
            )
            _urllib_req.urlopen(req, timeout=5).read()
    except Exception as _exc:
        logger.debug(f'[QUEUE] 관리자 알림 실패: {_exc}')

    # 2) 편집자 본인 SocketIO 브라우저 알림 (편집 계열 op 만)
    editor_ops = {'project_update_sheet', 'project_cancel_sheet', 'project_resume_sheet'}
    if user_email and op_type in editor_ops:
        try:
            from dashboard import socketio as _sio  # 부팅 완료 후 활성
            _sio.emit('error_notification', {
                'target_user_email': user_email,
                'severity': 'warning',
                'title': '편집이 시트에 반영되지 않았습니다',
                'message': (
                    f'프로젝트 {tag}: 3회 재시도 후 저장 실패. '
                    '화면에는 저장돼 보이지만 실제 시트에는 반영 안 됐습니다. '
                    '잠시 후 다시 시도해 주시거나 관리자에게 알려주세요.'
                ),
                'project_code': tag,
                'op_id': op.get('op_id', ''),
            })
        except Exception as _exc:
            logger.debug(f'[QUEUE] SocketIO 편집자 알림 실패: {_exc}')


def _drain_processing_on_startup():
    """부팅 시 processing 리스트에 남아있던 op 를 큐로 되돌림.

    이전 프로세스가 op 를 꺼내 처리 중 강제 종료된 경우 복구.
    """
    try:
        rc = _rc()
        moved = 0
        while True:
            raw = rc.rpoplpush(PROCESSING_KEY, QUEUE_KEY)
            if not raw:
                break
            moved += 1
        if moved:
            logger.warning(f'[QUEUE] 부팅 시 processing → queue 복구: {moved}건')
    except Exception as exc:
        logger.warning(f'[QUEUE] 부팅 시 processing 드레인 실패: {exc}')


def _worker_loop():
    """워커 메인 루프. LPOP + RPUSH 로 낮은 오버헤드 폴링.

    2026-07-09 BRPOPLPUSH 소켓 idle timeout 이슈로 lpop 폴링으로 전환.
    처리 중 크래시 시 op 유실 가능성 있지만 캐시는 이미 반영돼 있어 사용자 영향 없음.
    다음 refresh 에서 시트-캐시 불일치 감지되면 관리자가 큐 상태 endpoint 로 확인 가능.
    """
    import time as _time
    logger.info('[QUEUE] 워커 시작 (LPOP 폴링 모드)')
    while True:
        raw = None
        try:
            # Non-blocking pop. 없으면 짧게 sleep.
            raw = _rc().lpop(QUEUE_KEY)
            if raw is None:
                _time.sleep(1.0)
                continue
            # processing 리스트로 이동 — 강제 종료 시 부팅 시 복구용
            try:
                _rc().rpush(PROCESSING_KEY, raw)
            except Exception:
                pass
            try:
                op = json.loads(raw)
            except Exception as parse_exc:
                logger.error(f'[QUEUE] op 파싱 실패, drop: {parse_exc}')
                _rc().lrem(PROCESSING_KEY, 1, raw)
                continue

            op_type = op.get('op_type', '')
            op_id = op.get('op_id', '')[:8]
            handler = _registry.get(op_type)
            if not handler:
                logger.error(f'[QUEUE] 핸들러 없음: {op_type} (op_id={op_id}…) — 데드레터로 이동')
                _rc().rpush(FAILED_KEY, json.dumps(op, ensure_ascii=False))
                _rc().lrem(PROCESSING_KEY, 1, raw)
                continue

            try:
                handler(op['payload'])
                # 성공 → processing 제거
                _rc().lrem(PROCESSING_KEY, 1, raw)
                logger.info(f'[QUEUE] {op_type} 처리 완료 (op_id={op_id}…)')
                _bump_processed_and_maybe_rollup()
            except Exception as exc:
                op['attempts'] = int(op.get('attempts', 0)) + 1
                op['last_error'] = str(exc)[:500]
                _rc().lrem(PROCESSING_KEY, 1, raw)
                if op['attempts'] < MAX_ATTEMPTS:
                    # 지수 백오프 후 재큐잉
                    wait = 2.0 ** op['attempts']
                    logger.warning(
                        f'[QUEUE] {op_type} 실패 {op["attempts"]}/{MAX_ATTEMPTS} (op_id={op_id}…), '
                        f'{wait}s 후 재시도: {exc}'
                    )
                    threading.Timer(
                        wait,
                        lambda o=op: _rc().rpush(QUEUE_KEY, json.dumps(o, ensure_ascii=False)),
                    ).start()
                else:
                    logger.error(
                        f'[QUEUE] {op_type} {MAX_ATTEMPTS}회 실패 → 데드레터 '
                        f'(op_id={op_id}…): {exc}',
                        exc_info=True,
                    )
                    _rc().rpush(FAILED_KEY, json.dumps(op, ensure_ascii=False))
                    _notify_admin_deadletter(op, exc)
        except Exception as loop_exc:
            logger.error(f'[QUEUE] 워커 루프 예외: {loop_exc}', exc_info=True)
            if raw:
                try:
                    _rc().lrem(PROCESSING_KEY, 1, raw)
                except Exception:
                    pass
            time.sleep(1)


def start_worker():
    """워커 스레드 시작 (idempotent). Flask 부팅 시 1회 호출."""
    global _worker_started
    with _worker_lock:
        if _worker_started:
            return
        _drain_processing_on_startup()
        thread = threading.Thread(target=_worker_loop, daemon=True, name='SheetWriteWorker')
        thread.start()
        _worker_started = True
        logger.info('[QUEUE] 워커 스레드 등록 완료')
