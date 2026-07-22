"""
Slack 봇 블루프린트 (ITG 관리 봇)
- /slack/events : 슬랙이 우리 서버를 호출하는 단일 진입점
- 슬래시 명령, 인터랙티브 컴포넌트, 이벤트(file_shared, message 등) 모두 처리

환경변수:
  SLACK_BOT_TOKEN      - xoxb-... (Bot User OAuth Token)
  SLACK_SIGNING_SECRET - 슬랙이 우리 서버 호출 시 서명 검증용
  SLACK_BOT_ENABLED    - 'true' 일 때만 활성화 (기본 false 안전)
"""

import os
import re
import secrets
import time
import textwrap
import threading
import logging
import json
import urllib.request
from datetime import date, datetime
from typing import Optional
from flask import Blueprint, request, jsonify

from dashboard.utils.logging_config import get_logger
from dashboard.blueprints.slack_helpers import (
    _format_date_for_sheet,
    _format_visit_date_range,
    _split_visit_date_range,
    _v,
    _v_multi,
    _to_initial,
    _slack_user_to_initial,
    _slack_user_to_korean_name,
    _human_duration,
    SALES_INITIALS,
    slack_truncate,
)

logger = get_logger(__name__)

slack_bp = Blueprint('slack_bot', __name__, url_prefix='/slack')


# ─────────────────────────────────────────────────────────────
# 활성화 여부 + slack_bolt App 초기화
# ─────────────────────────────────────────────────────────────
_BOT_ENABLED = os.getenv('SLACK_BOT_ENABLED', 'false').lower() == 'true'
_BOT_TOKEN = os.getenv('SLACK_BOT_TOKEN', '')
_SIGNING_SECRET = os.getenv('SLACK_SIGNING_SECRET', '')

# 공사 현황 알림 봇 (별도 토큰/secret) — /공사확정 슬래시 + 모달 처리
_PROJECT_BOT_TOKEN = os.getenv('SLACK_PROJECT_BOT_TOKEN', '')
_PROJECT_SIGNING_SECRET = os.getenv('SLACK_PROJECT_SIGNING_SECRET', '')

# 방문 일정 알림 봇 (별도 토큰/secret) — #방문_일정 카드 + 날짜 수정/취소 액션
_VISIT_BOT_TOKEN = os.getenv('SLACK_VISIT_BOT_TOKEN', '')
_VISIT_SIGNING_SECRET = os.getenv('SLACK_VISIT_SIGNING_SECRET', '')

# A/S 사후 관리 봇 (별도 토큰/secret) — /as 슬래시 + 3단계 모달 흐름
_AS_BOT_TOKEN = os.getenv('SLACK_AS_BOT_TOKEN', '')
_AS_SIGNING_SECRET = os.getenv('SLACK_AS_SIGNING_SECRET', '')

# 세금계산서 관리 알림 봇 (별도 토큰/secret) — #영업_관리 카드 발송 + 스레드 첨부 자동 완료
_INVOICE_BOT_TOKEN = os.getenv('SLACK_INVOICE_BOT_TOKEN', '')
_INVOICE_SIGNING_SECRET = os.getenv('SLACK_INVOICE_SIGNING_SECRET', '')

_slack_app = None
_slack_handler = None
_project_slack_app = None
_project_slack_handler = None
_visit_slack_app = None
_visit_slack_handler = None
_as_slack_app = None
_as_slack_handler = None
_invoice_slack_app = None
_invoice_slack_handler = None

def _init_slack_app():
    """slack_bolt App 지연 초기화 (환경변수 누락 시 안전하게 비활성화)"""
    global _slack_app, _slack_handler

    if not _BOT_ENABLED:
        logger.info("[SLACK] SLACK_BOT_ENABLED=false — 봇 비활성화")
        return False

    if not _BOT_TOKEN or _BOT_TOKEN.startswith('여기에') or 'your' in _BOT_TOKEN.lower():
        logger.warning("[SLACK] SLACK_BOT_TOKEN 미설정 — 봇 비활성화")
        return False

    if not _SIGNING_SECRET or _SIGNING_SECRET.startswith('여기에') or 'your' in _SIGNING_SECRET.lower():
        logger.warning("[SLACK] SLACK_SIGNING_SECRET 미설정 — 봇 비활성화")
        return False

    try:
        from slack_bolt import App
        from slack_bolt.adapter.flask import SlackRequestHandler

        # Bolt 자체 디버그 로그 활성화 (메시지 라우팅 추적)
        import logging
        logging.getLogger("slack_bolt").setLevel(logging.DEBUG)
        logging.getLogger("slack_bolt.App").setLevel(logging.DEBUG)

        _slack_app = App(
            token=_BOT_TOKEN,
            signing_secret=_SIGNING_SECRET,
            # process_before_response=True : Flask 환경에서 필요
            process_before_response=True,
        )
        _slack_handler = SlackRequestHandler(_slack_app)

        # 핸들러 등록
        _register_handlers(_slack_app)

        # 토큰 유효성 health check — 시작 시 즉시 인지 (2026-07-10)
        _verify_bot_token(_slack_app.client, '메인 봇')

        logger.info("[SLACK] 봇 초기화 완료 ✅")
        return True

    except Exception as exc:
        logger.error(f"[SLACK] 봇 초기화 실패: {exc}", exc_info=True)
        return False


def _verify_bot_token(client, bot_label: str) -> bool:
    """auth_test 호출로 봇 토큰 유효성 즉시 확인 (2026-07-10).

    - 성공: team/user/user_id 로그 출력, True 반환
    - 실패 (invalid_auth, account_inactive, token_revoked 등): 명확한 경고
    - 네트워크 오류: warning 로그 후 True 반환 (부팅 자체는 계속)
    """
    try:
        res = client.auth_test()
        if res.get('ok'):
            logger.info(
                f'[SLACK/HEALTH] {bot_label} 토큰 유효 ✓ '
                f'team={res.get("team", "?")} bot={res.get("user", "?")} '
                f'user_id={res.get("user_id", "?")}'
            )
            return True
        err = res.get('error', 'unknown')
        logger.error(f'[SLACK/HEALTH] {bot_label} auth_test 실패: error={err}')
        return False
    except Exception as exc:
        # 네트워크 등 일시 오류는 warning 만 (부팅 계속)
        err_code = ''
        try:
            if hasattr(exc, 'response') and hasattr(exc.response, 'get'):
                err_code = exc.response.get('error', '')
        except Exception:
            pass
        if err_code in ('invalid_auth', 'account_inactive', 'token_revoked'):
            logger.error(
                f'[SLACK/HEALTH] {bot_label} 토큰 검증 실패 — 즉시 조치 필요: {err_code}'
            )
        else:
            logger.warning(f'[SLACK/HEALTH] {bot_label} auth_test 예외 (계속 진행): {exc}')
        return False


def _init_project_slack_app():
    """공사 현황 알림 봇 — 별도 Bolt App 인스턴스. /공사확정 슬래시 + 모달 처리."""
    global _project_slack_app, _project_slack_handler

    if not _BOT_ENABLED:
        return False
    if not _PROJECT_BOT_TOKEN:
        logger.warning("[SLACK/공사봇] SLACK_PROJECT_BOT_TOKEN 미설정 — 비활성화")
        return False
    if not _PROJECT_SIGNING_SECRET:
        logger.warning("[SLACK/공사봇] SLACK_PROJECT_SIGNING_SECRET 미설정 — 비활성화")
        return False

    try:
        from slack_bolt import App
        from slack_bolt.adapter.flask import SlackRequestHandler

        _project_slack_app = App(
            token=_PROJECT_BOT_TOKEN,
            signing_secret=_PROJECT_SIGNING_SECRET,
            process_before_response=True,
        )
        _project_slack_handler = SlackRequestHandler(_project_slack_app)

        _register_project_handlers(_project_slack_app)
        _verify_bot_token(_project_slack_app.client, '공사봇')
        logger.info("[SLACK/공사봇] 초기화 완료 ✅")
        return True
    except Exception as exc:
        logger.error(f"[SLACK/공사봇] 초기화 실패: {exc}", exc_info=True)
        return False


def _init_visit_slack_app():
    """방문 일정 알림 봇 — 별도 Bolt App 인스턴스. #방문_일정 카드 발송 + 액션 처리."""
    global _visit_slack_app, _visit_slack_handler

    if not _BOT_ENABLED:
        return False
    if not _VISIT_BOT_TOKEN:
        logger.warning("[SLACK/방문봇] SLACK_VISIT_BOT_TOKEN 미설정 — 비활성화")
        return False
    if not _VISIT_SIGNING_SECRET:
        logger.warning("[SLACK/방문봇] SLACK_VISIT_SIGNING_SECRET 미설정 — 비활성화")
        return False

    try:
        from slack_bolt import App
        from slack_bolt.adapter.flask import SlackRequestHandler

        _visit_slack_app = App(
            token=_VISIT_BOT_TOKEN,
            signing_secret=_VISIT_SIGNING_SECRET,
            process_before_response=True,
        )
        _visit_slack_handler = SlackRequestHandler(_visit_slack_app)

        _register_visit_handlers(_visit_slack_app)
        _verify_bot_token(_visit_slack_app.client, '방문봇')
        logger.info("[SLACK/방문봇] 초기화 완료 ✅")
        return True
    except Exception as exc:
        logger.error(f"[SLACK/방문봇] 초기화 실패: {exc}", exc_info=True)
        return False


def _init_as_slack_app():
    """A/S 사후 관리 봇 — 별도 Bolt App 인스턴스. /as 슬래시 + 3단계 흐름."""
    global _as_slack_app, _as_slack_handler

    if not _BOT_ENABLED:
        return False
    if not _AS_BOT_TOKEN:
        logger.warning("[SLACK/AS봇] SLACK_AS_BOT_TOKEN 미설정 — 비활성화")
        return False
    if not _AS_SIGNING_SECRET:
        logger.warning("[SLACK/AS봇] SLACK_AS_SIGNING_SECRET 미설정 — 비활성화")
        return False

    try:
        from slack_bolt import App
        from slack_bolt.adapter.flask import SlackRequestHandler

        _as_slack_app = App(
            token=_AS_BOT_TOKEN,
            signing_secret=_AS_SIGNING_SECRET,
            process_before_response=True,
        )
        _as_slack_handler = SlackRequestHandler(_as_slack_app)

        _register_as_handlers(_as_slack_app)
        _verify_bot_token(_as_slack_app.client, 'A/S봇')
        logger.info("[SLACK/AS봇] 초기화 완료 ✅")
        return True
    except Exception as exc:
        logger.error(f"[SLACK/AS봇] 초기화 실패: {exc}", exc_info=True)
        return False


def _init_invoice_slack_app():
    """세금계산서 관리 알림 봇 — 별도 Bolt App. #영업_관리 카드 발송 + 스레드 첨부 자동 완료."""
    global _invoice_slack_app, _invoice_slack_handler

    if not _BOT_ENABLED:
        return False
    if not _INVOICE_BOT_TOKEN:
        logger.warning("[SLACK/계산서봇] SLACK_INVOICE_BOT_TOKEN 미설정 — 비활성화")
        return False
    if not _INVOICE_SIGNING_SECRET:
        logger.warning("[SLACK/계산서봇] SLACK_INVOICE_SIGNING_SECRET 미설정 — 비활성화")
        return False

    try:
        from slack_bolt import App
        from slack_bolt.adapter.flask import SlackRequestHandler

        _invoice_slack_app = App(
            token=_INVOICE_BOT_TOKEN,
            signing_secret=_INVOICE_SIGNING_SECRET,
            process_before_response=True,
        )
        _invoice_slack_handler = SlackRequestHandler(_invoice_slack_app)

        _register_invoice_handlers(_invoice_slack_app)
        _verify_bot_token(_invoice_slack_app.client, '계산서봇')
        logger.info("[SLACK/계산서봇] 초기화 완료 ✅")
        return True
    except Exception as exc:
        logger.error(f"[SLACK/계산서봇] 초기화 실패: {exc}", exc_info=True)
        return False


def _try_acquire_action_lock(lead_no: str, action: str, ttl: int = 5) -> bool:
    """동시/더블 클릭 방지 락 — 첫 클릭만 통과. 락 못 잡으면 False.
    Redis 다운 시 보수적으로 True 반환 (기능 끊지 않음)."""
    if not lead_no:
        return True
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        return bool(rc.set(
            f'visit_action_lock:{lead_no}:{action}', '1', nx=True, ex=ttl,
        ))
    except Exception:
        return True


def _register_visit_handlers(app):
    """방문 일정 알림 봇 핸들러 — [✏️ 방문일 수정] / [✅ 방문 완료] / [🗑️ 방문 취소]"""

    # file_shared 이벤트 no-op (2026-07-22): Slack 이 파일 업로드 시 message.file_share
    # 이벤트와 별개로 file_shared 도 발송. 우리는 message.file_share 로만 처리 →
    # file_shared 는 Unhandled request 로그 노이즈. 명시적으로 ack 처리.
    @app.event("file_shared")
    def _noop_file_shared_visit(ack):
        ack()

    @app.action("visit_modify_date")
    def handle_visit_modify_date(ack, body, client):
        ack()
        # background로 분리 — ack 즉시 응답 + 3초 안에 views_open 호출
        def _bg():
            try:
                lead_no = body["actions"][0].get("value") or ''
                # 동시 클릭 방지 락 (5초)
                if not _try_acquire_action_lock(lead_no, 'modify_date'):
                    logger.info(f'[SLACK/방문봇] visit_modify_date 중복 클릭 skip ({lead_no})')
                    return
                channel = body["channel"]["id"]
                message_ts = body["message"]["ts"]
                trigger_id = body["trigger_id"]
                # 카드에서 현재 방문일 파싱 — 단일(2026-07-08) 또는 범위(2026-07-01~03/07-03/2027-01-02)
                cur_start, cur_end = '', ''
                try:
                    msg_text = body["message"].get("text", "")
                    m = re.search(
                        r'방문일\s*:\s*(\d{4}-\d{2}-\d{2}(?:~(?:\d{2}|\d{2}-\d{2}|\d{4}-\d{2}-\d{2}))?)',
                        msg_text,
                    )
                    if m:
                        cur_start, cur_end = _split_visit_date_range(m.group(1))
                except Exception:
                    pass
                metadata = json.dumps({
                    "lead_no": lead_no, "channel": channel, "message_ts": message_ts,
                }, ensure_ascii=False)
                dp_start = {"type": "datepicker", "action_id": "value"}
                if cur_start:
                    dp_start["initial_date"] = cur_start
                dp_end = {"type": "datepicker", "action_id": "value"}
                if cur_end:
                    dp_end["initial_date"] = cur_end
                client.views_open(trigger_id=trigger_id, view={
                    "type": "modal",
                    "callback_id": "submit_visit_modify",
                    "title": {"type": "plain_text", "text": "방문일 수정"},
                    "submit": {"type": "plain_text", "text": "수정"},
                    "close": {"type": "plain_text", "text": "취소"},
                    "private_metadata": metadata,
                    "blocks": [
                        {"type": "section", "text": {"type": "mrkdwn",
                            "text": f"*{lead_no}* 의 방문 예정일을 변경합니다."}},
                        {
                            "type": "input", "block_id": "visit_date",
                            "label": {"type": "plain_text", "text": "새 방문 예정일 (시작)"},
                            "element": dp_start,
                        },
                        {
                            "type": "input", "block_id": "visit_date_end", "optional": True,
                            "label": {"type": "plain_text", "text": "방문 예정일 (종료)"},
                            "hint": {"type": "plain_text",
                                     "text": "방문 일자가 범위 일 때만 입력. (예: 7/1~7/3)"},
                            "element": dp_end,
                        },
                    ],
                })
            except Exception as exc:
                logger.error(f"[SLACK/방문봇] visit_modify_date 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.view("submit_visit_modify")
    def handle_submit_visit_modify(ack, body, client, view):
        ack()
        def _bg():
            try:
                _process_visit_date_modify(client, body, view)
            except Exception as exc:
                logger.error(f"[SLACK/방문봇] submit_visit_modify 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    # ── [✏️ 정보 수정] 신규 (2026-07-15) ─────────────────────
    # 확장 모달 — 방문 유형·이름·연락처·주소·상담내용까지 편집 가능.
    # 유형 변경 시 자동 ETC↔정규 리드 전환 (커밋 3~4에서 추가).
    @app.action("visit_edit_info")
    def handle_visit_edit_info(ack, body, client):
        ack()
        def _bg():
            try:
                lead_no = body["actions"][0].get("value") or ''
                if not _try_acquire_action_lock(lead_no, 'edit_info'):
                    logger.info(f'[SLACK/방문봇] visit_edit_info 중복 클릭 skip ({lead_no})')
                    return
                channel = body["channel"]["id"]
                message_ts = body["message"]["ts"]
                trigger_id = body["trigger_id"]
                _open_visit_edit_modal(
                    client, lead_no=lead_no, channel=channel,
                    message_ts=message_ts, trigger_id=trigger_id,
                )
            except Exception as exc:
                logger.error(f"[SLACK/방문봇] visit_edit_info 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.view("submit_visit_edit")
    def handle_submit_visit_edit(ack, body, client, view):
        # 필수 검증 (거래처/소개는 연락처 필수, 기타는 선택)
        state = view.get('state', {}).get('values', {})
        _new_platform = _v(state, 'platform') or ''
        _new_phone = (_v(state, 'phone') or '').strip()
        if _new_platform in ('거래처', '소개') and not _new_phone:
            ack(response_action='errors', errors={
                'phone': '거래처/소개는 연락처가 필수입니다.',
            })
            return

        # 유형 변경 감지 → 확인 view 로 update
        metadata = json.loads(view.get('private_metadata') or '{}')
        _original_platform = metadata.get('original_platform', '')
        if _new_platform and _new_platform != _original_platform:
            try:
                confirm_view = _build_visit_edit_confirm_view(metadata, state)
                ack(response_action='update', view=confirm_view)
                return
            except Exception as exc:
                logger.error(f"[SLACK/방문봇] confirm view build 실패: {exc}",
                             exc_info=True)
                # fallback — 그냥 진행
        ack()
        def _bg():
            try:
                _process_visit_edit(client, body, view)
            except Exception as exc:
                logger.error(f"[SLACK/방문봇] submit_visit_edit 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.view("submit_visit_edit_confirm")
    def handle_submit_visit_edit_confirm(ack, body, client, view):
        ack()
        def _bg():
            try:
                _process_visit_edit_confirmed(client, body, view)
            except Exception as exc:
                logger.error(
                    f"[SLACK/방문봇] submit_visit_edit_confirm 실패: {exc}",
                    exc_info=True,
                )
        threading.Thread(target=_bg, daemon=True).start()

    @app.action("visit_cancel")
    def handle_visit_cancel(ack, body, client):
        ack()
        # 취소 사유 입력 모달 오픈 (즉시). 모달 submit 시 실제 처리 (2026-07-19).
        try:
            lead_no = body["actions"][0].get("value") or ''
            channel = body["channel"]["id"]
            message_ts = body["message"]["ts"]
            trigger_id = body["trigger_id"]
            _open_visit_cancel_reason_modal(
                client, lead_no, channel, message_ts, trigger_id,
            )
        except Exception as exc:
            logger.error(f"[SLACK/방문봇] visit_cancel 모달 open 실패: {exc}",
                         exc_info=True)

    @app.view("submit_visit_cancel_reason")
    def handle_submit_visit_cancel_reason(ack, body, client, view):
        ack()
        def _bg():
            try:
                _process_visit_cancel_confirmed(client, body, view)
            except Exception as exc:
                logger.error(f"[SLACK/방문봇] submit_visit_cancel_reason 실패: {exc}",
                             exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.action("visit_complete")
    def handle_visit_complete(ack, body, client):
        ack()
        def _bg():
            try:
                lead_no = body["actions"][0].get("value") or ''
                # 5초 락 (동일 프로세스 내 초근접 중복 방어)
                if not _try_acquire_action_lock(lead_no, 'complete'):
                    logger.info(f'[SLACK/방문봇] visit_complete 중복 클릭 skip ({lead_no})')
                    return
                # 이미 완료 처리된 lead 재클릭 방어 (2026-07-21) — 락 TTL 지난 후
                # 재클릭 시 List 워크플로우가 이미 삭제된 항목 재삭제 시도 → 오류.
                # visit_auto_completed flag (30일 TTL) 로 확인. Redis 다운 시엔 통과.
                try:
                    from dashboard.utils.redis_client import get_redis_client
                    if get_redis_client().redis.get(f'visit_auto_completed:{lead_no}'):
                        logger.info(
                            f'[SLACK/방문봇] visit_complete 이미 완료됨 - 재실행 skip ({lead_no})'
                        )
                        return
                except Exception:
                    pass
                _process_visit_complete(client, body)
            except Exception as exc:
                logger.error(f"[SLACK/방문봇] visit_complete 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.action("visit_uncancel")
    def handle_visit_uncancel(ack, body, client):
        ack()
        def _bg():
            try:
                lead_no = body["actions"][0].get("value") or ''
                if not _try_acquire_action_lock(lead_no, 'uncancel'):
                    logger.info(f'[SLACK/방문봇] visit_uncancel 중복 클릭 skip ({lead_no})')
                    return
                _process_visit_uncancel(client, body)
            except Exception as exc:
                logger.error(f"[SLACK/방문봇] visit_uncancel 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    # 방문 카드 thread 메시지 — 사진 첨부(드라이브 업로드) + 상호명 답글(폴더명 갱신)
    @app.event("message")
    def handle_visit_message(event, client):
        thread_ts = event.get("thread_ts")
        subtype = event.get("subtype")
        bot_id = event.get("bot_id")
        text_preview = (event.get("text") or '')[:30]
        logger.info(
            f"[SLACK/방문봇] message event: thread_ts={thread_ts} "
            f"subtype={subtype} bot_id={bot_id} text={text_preview!r}"
        )
        if not thread_ts:
            return
        if bot_id:  # 봇 메시지는 무시 (echo 방지)
            return

        # 1) 사진/파일 첨부 → 드라이브 업로드
        if subtype == "file_share" and event.get("files"):
            threading.Thread(
                target=_process_visit_thread_files,
                args=(client, event), daemon=True,
            ).start()
            return

        # 2) 상호 / 상호명 prefix 답글 → 폴더명 갱신
        text = (event.get("text") or '').strip()
        if text.startswith('상호 ') or text.startswith('상호명 '):
            logger.info(f"[SLACK/방문봇] 상호명 답글 감지 → shop_name 갱신 트리거")
            threading.Thread(
                target=_process_visit_shop_name_update,
                args=(client, event), daemon=True,
            ).start()

    # 방문 일정 조정 캔버스 (JW 전용) 파서 (2026-07-15)
    # 권한 = 박정우(JW)·박용구(YG)·고광일(KiKO) 만 실행 (2026-07-19)
    _VISIT_CMD_ALLOWED_EMAILS = {
        'jw@itg-aircon.com', 'yg@itg-aircon.com', 'kiko@itg-aircon.com',
    }

    def _check_visit_cmd_permission(command) -> Optional[str]:
        """권한 없으면 안내 문구 반환, 있으면 None."""
        user_id = command.get('user_id', '')
        try:
            info = _visit_slack_app.client.users_info(user=user_id)
            email = (info.get('user') or {}).get('profile', {}).get('email', '')
        except Exception as exc:
            logger.warning(f"[일정cmd] 권한 조회 실패 ({user_id}): {exc}")
            email = ''
        if email not in _VISIT_CMD_ALLOWED_EMAILS:
            return (':lock: 이 명령은 박정우·박용구·고광일만 실행할 수 있습니다.')
        return None

    @app.command("/일정확인")
    def handle_visit_assignment_dryrun(ack, command, respond):
        ack()
        deny = _check_visit_cmd_permission(command)
        if deny:
            respond({'response_type': 'ephemeral', 'text': deny})
            return
        def _bg():
            try:
                from dashboard.services.visit_assignment_sync import dry_run
                result = dry_run()
                respond(_format_assignment_result(result, committed=False))
            except Exception as exc:
                logger.error(f"[일정확인] 예외: {exc}", exc_info=True)
                respond({'response_type': 'ephemeral',
                         'text': f':x: 일정 확인 실패: {exc}'})
        threading.Thread(target=_bg, daemon=True).start()

    @app.command("/일정확정")
    def handle_visit_assignment_commit(ack, command, respond):
        ack()
        deny = _check_visit_cmd_permission(command)
        if deny:
            respond({'response_type': 'ephemeral', 'text': deny})
            return
        def _bg():
            try:
                from dashboard.services.visit_assignment_sync import commit
                result = commit()
                respond(_format_assignment_result(result, committed=True))
            except Exception as exc:
                logger.error(f"[일정확정] 예외: {exc}", exc_info=True)
                respond({'response_type': 'ephemeral',
                         'text': f':x: 일정 확정 실패: {exc}'})
        threading.Thread(target=_bg, daemon=True).start()


def _format_assignment_result(result: dict, committed: bool) -> dict:
    """visit_assignment_sync 결과 → 슬랙 ephemeral 응답 포맷."""
    if not result.get('ok'):
        return {'response_type': 'ephemeral',
                'text': f':x: 실패: {result.get("reason", "unknown")}'}
    if committed:
        lines = [
            f":white_check_mark: *일정 확정 완료* — 시트 {result.get('updated_count', 0)}건",
        ]
        li_ok = result.get('list_updated', 0)
        li_fail = result.get('list_failed', 0)
        if li_ok or li_fail:
            lines.append(f":clipboard: Slack List 담당자 {li_ok}건 update"
                         + (f" (실패 {li_fail})" if li_fail else ""))
        dm = result.get('dm') or {}
        if dm.get('target_date'):
            lines.append(
                f":envelope: {dm['target_date']} 방문 담당자 DM {dm.get('visit_mgr_sent', 0)}명"
                + (f" (실패 {dm.get('visit_mgr_failed', 0)})"
                   if dm.get('visit_mgr_failed') else "")
                + (f" · 온라인 당번 {dm['online_duty_sent']}명"
                   if dm.get('online_duty_sent') else "")
                + (f" · 배정 해제 알림 {dm['deassign_sent']}명"
                   if dm.get('deassign_sent') else "")
            )
        if result.get('online_duty'):
            lines.append(f"_온라인 당번:_ {'·'.join(result['online_duty'])}")
        if result.get('off_duty'):
            lines.append(f"_휴무:_ {'·'.join(result['off_duty'])}")
        if result.get('updated'):
            lines.append('_시트 업데이트 리드:_ ' + ', '.join(result['updated'][:20]))
        if result.get('failed_count'):
            lines.append(f":warning: 시트 실패 {result['failed_count']}건")
            for ln, err in result.get('failed', [])[:5]:
                lines.append(f'  - {ln}: {err[:80]}')
        lines.append('_방문 캔버스 rebuild 백그라운드 진행 중_')
        if (result.get('dm') or {}).get('visit_mgr_sent'):
            lines.append(
                ':warning: _재실행 시 담당자에게 중복 DM 발송됩니다. '
                '캔버스 수정 후에만 다시 실행하세요._'
            )
        return {'response_type': 'ephemeral', 'text': '\n'.join(lines)}

    rows = result.get('rows', [])
    matched = [r for r in rows if r['matched']]
    unmatched = [r for r in rows if not r['matched']]
    changed = [r for r in matched if r['changed']]
    unchanged = [r for r in matched if not r['changed']]
    header = f":clipboard: *일정 확인 (dry-run)* — 총 {len(rows)}건 파싱"
    if result.get('target_date'):
        header += f" · DM 대상 {result['target_date']}"
    lines = [
        header,
        f"   ✓ 시트 매칭 {len(matched)}건 (변경 {len(changed)}, 유지 {len(unchanged)})",
        f"   ✗ 매칭 실패 {len(unmatched)}건",
    ]
    if result.get('online_duty'):
        lines.append(f"   :headphones: 온라인 당번 : {'·'.join(result['online_duty'])}")
    if result.get('off_duty'):
        lines.append(f"   :palm_tree: 휴무 : {'·'.join(result['off_duty'])}")
    lines.append('')
    if changed:
        lines.append('*변경 대상:*')
        for r in changed[:20]:
            lines.append(
                f"  `{r['lead_no']}` {r['phone']} : {r['current']} → *{r['assign_names']}*"
            )
        if len(changed) > 20:
            lines.append(f'  ... 외 {len(changed) - 20}건')
        lines.append('')
    if unmatched:
        lines.append('*매칭 실패:*')
        for r in unmatched[:10]:
            lines.append(f"  {r['phone']} — 이 연락처 시트에 없음")
        if len(unmatched) > 10:
            lines.append(f'  ... 외 {len(unmatched) - 10}건')
    lines.append('')
    lines.append('확정하려면 `/일정확정` 실행.')
    return {'response_type': 'ephemeral', 'text': '\n'.join(lines)}


def _register_project_handlers(app):
    """공사 현황 알림 봇 핸들러 — /공사확정 + submit_project"""

    @app.event("file_shared")  # no-op — message.file_share 로 실질 처리 (2026-07-22)
    def _noop_file_shared_project(ack):
        ack()

    @app.command("/공사확정")
    def handle_project_command(ack, command, client):
        ack()
        trigger_id = command.get("trigger_id", "")
        channel = command.get("channel_id", "")
        user_id = command.get("user_id", "")
        if not trigger_id:
            return
        try:
            _open_project_modal(client, trigger_id, channel, user_id)
        except Exception as exc:
            logger.error(f"[SLACK/공사확정] 모달 열기 실패: {exc}", exc_info=True)

    @app.view("submit_project")
    def handle_submit_project(ack, body, client, view):
        # 날짜 순서 검증 (2026-07-10): 공사 종료가 공사 시작보다 앞이면 반려
        state = view.get("state", {}).get("values", {})
        start_date = _v(state, "start_date") or ''
        end_date = _v(state, "end_date") or ''
        if start_date and end_date and end_date < start_date:
            ack(response_action="errors", errors={
                "end_date": f"공사 종료일({end_date})은 시작일({start_date})보다 이후여야 합니다.",
            })
            return
        ack()
        _run_bg_with_notify(
            client, body, '공사 확정',
            lambda: _process_project_submission(client, body, view),
        )

    @app.options("value")
    def handle_external_options(ack, body):
        """external_select 옵션 응답. block_id로 분기.

        현재는 company_name(사업자명) 한 곳만 사용.
        """
        block_id = body.get("block_id", "")
        query = (body.get("value") or "").strip()
        logger.info(
            f"[SLACK/공사확정/options] 요청 수신: block_id={block_id!r}, query={query!r}"
        )
        if block_id == "company_name":
            try:
                options = _search_company_names(query.lower())
                logger.info(f"[SLACK/공사확정/options] {len(options)}개 반환")
                ack(options=options)
            except Exception as exc:
                logger.warning(f"[SLACK/공사확정] 사업자명 검색 실패: {exc}", exc_info=True)
                ack(options=[])
        else:
            ack(options=[])

    # ─────────────────────────────────────────────────────────
    # 계산서 발행 요청 흐름 (공사 확정 카드 → 모달 → #영업_관리 카드 → 발행 완료)
    # ─────────────────────────────────────────────────────────
    @app.action("invoice_request_open")
    def handle_invoice_request_open(ack, body, client):
        ack()
        def _bg():
            try:
                _open_invoice_modal(client, body)
            except Exception as exc:
                logger.error(f"[SLACK/계산서] 모달 열기 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.view("submit_invoice")
    def handle_submit_invoice(ack, body, client, view):
        # UX 개선 (2026-07-16 사고): 검증 (Drive API + 시트 조회) 이 3초 넘어가
        # modal 이 안 닫히는 문제 → ack() 즉시 호출 + 검증·발송을 BG 스레드로.
        # 검증 실패 시 매니저에게 chat.postEphemeral (DM fallback) 로 반려 안내.
        ack()
        def _bg():
            try:
                _process_invoice_submit_bg(client, body, view)
            except Exception as exc:
                logger.error(f"[SLACK/계산서] submit 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.action("invoice_complete")
    def handle_invoice_complete(ack, body, client):
        ack()
        def _bg():
            try:
                _process_invoice_complete(client, body)
            except Exception as exc:
                logger.error(f"[SLACK/계산서] complete 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    # ─────────────────────────────────────────────────────────
    # 사업자등록증 스레드 첨부 → Google Drive 저장 (2026-07-08)
    # ─────────────────────────────────────────────────────────
    # 공사 확정 카드 스레드에 매니저가 파일(이미지/PDF)을 첨부하면 봇이 감지해
    # 프로젝트 폴더의 '사업자등록증/' 하위 폴더로 자동 저장.
    # 첫 파일은 '사업자등록증.{ext}', 재첨부 시 기존 것을 '사업자등록증_{N}.{ext}'로 밀고
    # 새 파일을 다시 '사업자등록증.{ext}' 로 저장 (최신본이 항상 canonical name).
    # 계산서 요청 시 이 canonical 파일 존재 여부로 검증.
    @app.event("message")
    def handle_thread_message(event, client):
        # 모든 message 이벤트가 들어옴 — subtype 필터 + 진단 로그
        subtype = event.get("subtype") or ""
        has_files = bool(event.get("files"))
        thread_ts = event.get("thread_ts")
        channel = event.get('channel')
        ts = event.get('ts')
        logger.info(
            f"[LICENSE/EVT] message 수신: subtype={subtype!r}, "
            f"has_files={has_files}, thread_ts={thread_ts!r}, "
            f"channel={channel}"
        )

        # 스레드 파일 첨부만 처리 + 봇 자신 메시지 skip (bot_message subtype 등).
        # 계산서 스레드 첨부·삭제 감지는 계산서봇(_register_invoice_handlers)이 담당.
        if not thread_ts or not has_files:
            return
        if subtype == 'bot_message' or event.get('bot_id'):
            return

        # 2026-07-10 UX 개선 — 파일 첨부 즉시 시각 피드백.
        #   기존엔 Drive 저장·검증까지 3-10초 걸리는 동안 매니저 관점에선 봇이 아무
        #   반응 없어 보이던 UX 사고. 즉시 :hourglass_flowing_sand: reaction 붙여
        #   "봇이 인지·처리 중" 인지 → 완료 시 ✅ or ❌ 로 교체.
        _sand = 'hourglass_flowing_sand'
        if channel and ts:
            try:
                client.reactions_add(channel=channel, timestamp=ts, name=_sand)
            except Exception:
                pass  # 이미 붙어있거나 권한 이슈 — 무시

        def _safe_react(name: str) -> None:
            """hourglass 제거 후 최종 상태 reaction 부착. 실패는 무시."""
            if not (channel and ts):
                return
            try:
                client.reactions_remove(channel=channel, timestamp=ts, name=_sand)
            except Exception:
                pass
            try:
                client.reactions_add(channel=channel, timestamp=ts, name=name)
            except Exception:
                pass

        def _bg():
            try:
                from dashboard.services.business_license_handler import handle_thread_file_share as _h
                result = _h(event, _PROJECT_BOT_TOKEN)
                if not result:
                    logger.info("[LICENSE] 프로젝트 카드 스레드 아님 → skip")
                    # 프로젝트 스레드 아니면 reaction 정리만 (오해 방지)
                    if channel and ts:
                        try:
                            client.reactions_remove(channel=channel, timestamp=ts, name=_sand)
                        except Exception:
                            pass
                    return
                saved = result.get('saved') or []
                skipped = result.get('skipped') or []

                # reaction 최종 상태 반영
                if saved and not skipped:
                    _safe_react('white_check_mark')  # 전건 성공
                elif saved and skipped:
                    _safe_react('warning')  # 부분 성공
                else:
                    _safe_react('x')  # 실패

                lines = []
                if saved:
                    lines.append(f":white_check_mark: 사업자등록증 저장 완료 — `{result['code']}`")
                    for fn in saved:
                        lines.append(f"  • {fn}")
                # OCR 결과 (2026-07-13): 법인명·상호 자동 추출 + 사업자명 자동 반영.
                #   saved      : 사업자명 비어있어서 자동 저장
                #   match      : 기존값 == OCR (안내 생략)
                #   mismatch   : 기존값 ≠ OCR (덮어쓰지 않음, 매니저 확인 유도)
                #   error / '' : 실패 or OCR 매치 못함
                _biz = (result.get('business_name') or '').strip()
                _biz_status = result.get('biz_update_status') or ''
                _biz_existing = (result.get('biz_update_existing') or '').strip()
                if _biz and _biz_status == 'saved':
                    lines.append(
                        f":memo: OCR 자동 등록 — 사업자명: *{_biz}* (시트에 저장됨)"
                    )
                elif _biz and _biz_status == 'mismatch':
                    lines.append(
                        f":memo: 사업자등록증 OCR 결과와 시트값이 달라요. "
                        f"어느 쪽이 맞는지 확인해주세요.\n"
                        f"  • 시트값: *{_biz_existing}*\n"
                        f"  • OCR 결과: *{_biz}*"
                    )
                elif _biz and _biz_status == 'error':
                    lines.append(
                        f":memo: OCR 결과 — 사업자명 추정: *{_biz}*  "
                        f"_(자동 저장 실패 — 관리 페이지에서 확인 후 수동 입력하세요)_"
                    )
                # match / '' 인 경우 조용히 skip
                if skipped:
                    lines.append(f":warning: 저장 안 됨:")
                    for s in skipped:
                        lines.append(f"  • {s}")
                if lines:
                    try:
                        client.chat_postMessage(
                            channel=result['channel'],
                            thread_ts=result['thread_ts'],
                            text='\n'.join(lines),
                        )
                    except Exception as exc:
                        logger.warning(f"[LICENSE] 답글 발송 실패: {exc}")
            except Exception as exc:
                logger.error(f"[LICENSE] 파일 처리 예외: {exc}", exc_info=True)
                _safe_react('x')  # 예외 시 실패 표시
        threading.Thread(target=_bg, daemon=True).start()

    # ─────────────────────────────────────────────────────────
    # 공사 확정 카드 [✏️ 내용 수정] / [❌ 공사 취소] / [↩️ 취소 되돌리기]
    # ─────────────────────────────────────────────────────────
    @app.action("project_edit_open")
    def handle_project_edit_open(ack, body, client):
        ack()
        def _bg():
            try:
                _open_project_edit_modal(client, body)
            except Exception as exc:
                logger.error(f"[SLACK/공사수정] 모달 열기 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.view("submit_project_edit")
    def handle_submit_project_edit(ack, body, client, view):
        # 필수 사유 검증 — 빈값이면 modal errors로 응답
        values = view.get("state", {}).get("values", {})
        reason = ''
        try:
            reason = (values.get("reason", {}).get("value", {}) or {}).get("value", '') or ''
        except Exception:
            reason = ''
        if not reason.strip():
            ack(response_action="errors", errors={"reason": "수정 사유를 반드시 입력해야 합니다."})
            return
        # 날짜 순서 검증 (2026-07-10)
        start_date = _v(values, "start_date") or ''
        end_date = _v(values, "end_date") or ''
        if start_date and end_date and end_date < start_date:
            ack(response_action="errors", errors={
                "end_date": f"공사 종료일({end_date})은 시작일({start_date})보다 이후여야 합니다.",
            })
            return
        ack()
        _run_bg_with_notify(
            client, body, '공사 정보 수정',
            lambda: _process_project_edit_submission(client, body, view),
        )

    @app.action("project_cancel_confirm")
    def handle_project_cancel(ack, body, client):
        ack()
        _run_bg_with_notify(
            client, body, '공사 취소',
            lambda: _process_project_cancel(client, body),
        )

    @app.action("project_uncancel")
    def handle_project_uncancel(ack, body, client):
        ack()
        _run_bg_with_notify(
            client, body, '공사 취소 되돌리기',
            lambda: _process_project_uncancel(client, body),
        )
        threading.Thread(target=_bg, daemon=True).start()

    logger.info(
        "[SLACK/공사봇] 핸들러 등록 완료: /공사확정, submit_project, "
        "options(company_name), invoice_request_open, submit_invoice, invoice_complete, "
        "message.file_share(사업자등록증), "
        "project_edit_open, submit_project_edit, project_cancel_confirm, project_uncancel"
    )


# ─────────────────────────────────────────────────────────────
# 슬랙 이벤트 핸들러
# ─────────────────────────────────────────────────────────────
def _register_handlers(app):
    """슬래시 명령, 인터랙티브, 이벤트 핸들러 등록"""

    @app.event("file_shared")  # no-op (2026-07-22)
    def _noop_file_shared_main(ack):
        ack()

    # ① 슬래시 명령: /상태 (사이트 헬스체크)
    @app.command("/상태")
    def handle_status(ack, command, respond):
        ack()
        try:
            from dashboard.services.lead_service import load_leads_data
            df = load_leads_data()
            lead_count = len(df) if df is not None else 0
            respond({
                "response_type": "ephemeral",
                "blocks": [
                    {
                        "type": "header",
                        "text": {"type": "plain_text", "text": "📊 ITG 시스템 상태"},
                    },
                    {
                        "type": "section",
                        "fields": [
                            {"type": "mrkdwn", "text": f"*리드 데이터:*\n{lead_count}건"},
                            {"type": "mrkdwn", "text": f"*사이트:*\n✅ 정상"},
                            {"type": "mrkdwn", "text": f"*봇:*\n✅ 동작 중"},
                        ],
                    },
                ],
            })
        except Exception as exc:
            logger.error(f"[SLACK] /상태 실패: {exc}", exc_info=True)
            respond({"text": f"❌ 상태 조회 실패: {exc}"})

    # ② 봇 멘션 이벤트 (예: @ITG관리봇 안녕)
    @app.event("app_mention")
    def handle_mention(event, say):
        user = event.get("user", "")
        text = event.get("text", "")
        say(f"<@{user}> 부르셨나요? `/상태`, `/전화`, `/청소` 명령을 사용해보세요.")

    # ③ DM + 채널톡 thread 답글 통합 처리
    @app.event("message")
    def handle_message(event, say, client):
        # 디버그 — 들어온 이벤트 무조건 로깅
        logger.info(
            f"[SLACK/msg] type={event.get('type')} subtype={event.get('subtype')} "
            f"channel_type={event.get('channel_type')} thread_ts={event.get('thread_ts')} "
            f"bot_id={event.get('bot_id')} user={event.get('user')} "
            f"text={(event.get('text') or '')[:40]!r}"
        )

        # 봇 자신의 메시지는 무시
        if event.get("bot_id") or event.get("subtype") == "bot_message":
            logger.debug("[SLACK/msg] bot/bot_message → skip")
            return

        channel_type = event.get("channel_type")

        # ③-1. DM: 안내 메시지
        if channel_type == "im":
            text = event.get("text", "")
            say(f"메시지 받았습니다: _{text}_\n슬래시 명령 `/상태`, `/전화`, `/청소`도 사용 가능합니다.")
            return

        # ③-2. 채널 thread 답글 — 채널톡 thread면 채널톡으로 forward
        if channel_type in ("channel", "group"):
            thread_ts = event.get("thread_ts")
            if not thread_ts:
                return  # thread가 아닌 일반 채널 메시지는 무시

            try:
                from dashboard.services.channeltalk_threads import get_chat_id
                from dashboard.services.channeltalk_api import (
                    send_manager_message,
                    assign_user_chat,
                )
                logger.info(f"[ChannelTalk→] thread 답글 수신 (thread_ts={thread_ts})")

                chat_id = get_chat_id(thread_ts)
                if not chat_id:
                    logger.info(f"[ChannelTalk→] 채널톡 매핑 없음 — 무시 (thread_ts={thread_ts})")
                    return

                text = (event.get("text") or "").strip()
                if not text:
                    return

                manager_id = os.getenv("CHANNELTALK_OPERATOR_ID", "").strip()
                if not manager_id:
                    logger.warning("[ChannelTalk→] CHANNELTALK_OPERATOR_ID 미설정 — 전송 불가")
                    return

                # 채널톡은 봇 명의로 메시지 발신 — 배정 없이도 동작 확인됨
                # echo loop 방지: 이 메시지가 webhook으로 되돌아올 때 skip하도록 캐시
                from dashboard.blueprints.channeltalk import mark_our_sent
                mark_our_sent(chat_id, text)
                resp = send_manager_message(chat_id, manager_id, text)
                logger.info(f"[ChannelTalk→] 메시지 발신: text={text[:40]!r}, resp_ok={resp is not None}")
                # 이메일 자동 치환 재시도로 성공한 케이스 안내 (2026-07-10 CT2)
                if resp and resp.get('_email_auto_escaped'):
                    try:
                        user_id = event.get('user', '')
                        if user_id:
                            client.chat_postEphemeral(
                                channel=event["channel"],
                                user=user_id,
                                thread_ts=thread_ts,
                                text=(
                                    ':information_source: *이메일 주소가 감지되어 전각 골뱅이(＠) 로 자동 치환해 전송했습니다.*\n'
                                    '_채널톡이 이메일 형식을 자동 차단하는 경우가 있어 우회한 것입니다._\n'
                                    '_고객 화면에는 정상적인 이메일로 보이니 안심하세요._'
                                ),
                            )
                    except Exception:
                        pass

                # 직원 응답했으니 미배정 알림 큐에서 제거
                from dashboard.services.channeltalk_threads import remove_pending
                remove_pending(chat_id)
                if resp:
                    # 1) 답글에 ✅ — 본인 전송 성공 표시
                    try:
                        client.reactions_add(
                            channel=event["channel"],
                            timestamp=event["ts"],
                            name="white_check_mark",
                        )
                    except Exception:
                        pass
                    # 2) 원본 카드(thread_ts)에 ✅ — 다른 영업도 "처리됨" 한눈에 확인
                    try:
                        client.reactions_add(
                            channel=event["channel"],
                            timestamp=thread_ts,
                            name="white_check_mark",
                        )
                    except Exception:
                        pass
                    logger.info(f"[ChannelTalk→] 슬랙→채널톡 전송 완료 (chat_id={chat_id})")
                else:
                    try:
                        client.reactions_add(
                            channel=event["channel"],
                            timestamp=event["ts"],
                            name="x",
                        )
                    except Exception:
                        pass
                    logger.warning(f"[ChannelTalk→] 슬랙→채널톡 전송 실패 (chat_id={chat_id})")
                    # 매니저에게 명시적 ephemeral 안내 (2026-07-10 CT1)
                    # 리액션 X 만으로는 놓칠 수 있음. 답변 안 갔음을 확실히 인지시킴.
                    try:
                        user_id = event.get('user', '')
                        if user_id:
                            client.chat_postEphemeral(
                                channel=event["channel"],
                                user=user_id,
                                thread_ts=thread_ts,
                                text=(
                                    ':warning: *답변이 고객에게 전송되지 않았습니다.*\n'
                                    '채널톡 API 오류로 실패했습니다. 잠시 후 답글에 같은 내용을 다시 입력해 주세요.\n'
                                    '_반복 실패 시 관리자에게 문의하세요._'
                                ),
                            )
                    except Exception as ephemeral_exc:
                        logger.debug(f'[ChannelTalk→] ephemeral 실패 (무시): {ephemeral_exc}')
            except Exception as exc:
                logger.error(f"[ChannelTalk→] thread 답글 처리 예외: {exc}", exc_info=True)

    # ④ 인입 알림 메시지의 [방문 요청] 버튼
    # ⓑ [📋 상담하기] 통합 버튼 — 인입 카드 모든 처리 흐름의 단일 진입점
    # 2026-07-10 UX 개선 — modal 열기 handler 를 백그라운드 스레드로 이관.
    #   기존엔 ack() 후 handler 안에서 직접 _open_xxx_modal (views_open) 호출 →
    #   Slack API 지연 + Waitress 큐 대기가 겹치면 3초 초과 → 슬랙에 세모 느낌표.
    #   process_before_response=True 모드에서 Bolt 는 handler 반환 시점에 200 응답 →
    #   백그라운드 스레드로 넘기면 handler 즉시 반환 → 세모 원천 차단.
    #   trigger_id 는 3초 유효하지만 스레드는 즉시 시작하므로 실측상 여유 있음.
    @app.action("button_consult")
    def handle_button_consult(ack, body, client):
        ack()
        def _bg():
            try:
                _open_consult_modal(client, body, from_slash=False)
            except Exception as exc:
                logger.error(f"[SLACK] button_consult 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    # 채널톡 카드 [🔗 기존 lead 연결] — 같은 사람이 다른 채널로도 인입했을 때
    @app.action("link_existing_lead")
    def handle_link_existing_lead(ack, body, client):
        ack()
        def _bg():
            try:
                chat_id = body["actions"][0]["value"]
                channel = body["channel"]["id"]
                message_ts = body["message"]["ts"]
                _open_link_lead_modal(client, body, chat_id, channel, message_ts)
            except Exception as exc:
                logger.error(f"[SLACK] link_existing_lead 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.options("link_lead_search")
    def handle_link_lead_options(ack, body):
        """external_select 검색 — 매니저가 입력한 query로 시트 lead 매칭."""
        try:
            query = (body.get("value") or "").strip()
            block_id = (body.get("block_id") or "")
            logger.info(
                f"[SLACK/options] action=link_lead_search block_id={block_id!r} query={query!r}"
            )
            options = _search_leads_for_options(query, limit=30)
            logger.info(f"[SLACK/options] 반환 {len(options)}건")
            ack(options=options)
        except Exception as exc:
            logger.error(f"[SLACK] link lead options 실패: {exc}", exc_info=True)
            try:
                ack(options=[])
            except Exception:
                pass

    @app.view("submit_link_lead")
    def handle_submit_link_lead(ack, body, view, client):
        # 2026-07-11 UX 개선 — 링크 처리(시트 write + slack post 여러 개)가 3초 초과 →
        #   Slack UI 에 "연결하는 데 문제가 발생했습니다" 표시. 실제 처리는 성공했지만
        #   매니저 눈에는 실패로 보임. 검증만 handler 안에서 하고 실제 통합 처리는
        #   background thread 로 이관 + ack() 즉시.
        try:
            metadata = json.loads(view["private_metadata"])
            chat_id = metadata.get("chat_id", "")
            channel = metadata.get("channel", "")
            message_ts = metadata.get("message_ts", "")
            user_id = (body.get("user") or {}).get("id", "")
            state = view["state"]["values"]
            # external_select 결과 — selected_option.value = lead_no
            sel = state.get("target_lead_no", {}).get("link_lead_search", {}).get("selected_option")
            target_lead_no = (sel or {}).get("value", "").strip().upper() if sel else ""
            if not re.match(r"^L-\d{5}$", target_lead_no):
                ack(response_action="errors", errors={
                    "target_lead_no": "검색해서 lead를 선택해주세요"
                })
                return
            ack()
        except Exception as exc:
            logger.error(f"[SLACK] submit_link_lead 검증 실패: {exc}", exc_info=True)
            try:
                ack()
            except Exception:
                pass
            return

        def _bg():
            try:
                target_lead = _find_lead_by_no(target_lead_no)
                if not target_lead:
                    if channel and user_id:
                        try:
                            client.chat_postEphemeral(
                                channel=channel, user=user_id,
                                text=f":warning: `{target_lead_no}` 시트에 없는 lead 입니다. 다시 시도해주세요.",
                            )
                        except Exception:
                            pass
                    return
                _link_chat_to_existing_lead(
                    client, chat_id, target_lead_no, channel, message_ts,
                    slack_user_id=user_id,
                )
            except Exception as exc:
                logger.error(f"[SLACK] submit_link_lead 백그라운드 실패: {exc}", exc_info=True)
                if channel and user_id:
                    try:
                        client.chat_postEphemeral(
                            channel=channel, user=user_id,
                            text=f":warning: 링크 처리 중 오류: {exc}",
                        )
                    except Exception:
                        pass
        threading.Thread(target=_bg, daemon=True).start()

    # ⓓ /방문 슬래시 명령 — 거래처/기타 방문 직접 등록
    # 2026-07-12 /방문 슬래시 명령어 제거 — 사용 안 함.
    #   상담 모달은 리드 카드 [상담하기] 버튼으로만 진입.

    # ⓒ 통합 상담 모달 제출
    @app.view("submit_consult")
    def handle_submit_consult(ack, body, client, view):
        # 방문 예약일 때만 필수 필드 검증 (유선 상담·문의 드랍 등은 옵션 유지)
        state = view["state"]["values"]
        status = _v(state, "status")
        if status == '방문 예약':
            errors = {}
            visit_date = (_v(state, "visit_date") or '').strip()
            name = (_v(state, "name") or '').strip()
            contact = (_v(state, "contact") or '').strip()
            visit_address = (_v(state, "visit_address") or '').strip()

            def _is_empty(v):
                return not v or v == '-'

            if _is_empty(visit_date):
                errors["visit_date"] = "방문 예약 시 방문 예정일을 선택해주세요."
            if _is_empty(name):
                errors["name"] = "방문 예약 시 이름/상호를 입력해주세요."
            if _is_empty(contact):
                errors["contact"] = "방문 예약 시 연락처를 입력해주세요."
            if _is_empty(visit_address):
                errors["visit_address"] = "방문 예약 시 방문 주소를 입력해주세요."

            if errors:
                ack(response_action="errors", errors=errors)
                return
        ack()
        def _bg():
            try:
                _process_consult_submission(client, body, view)
            except Exception as exc:
                logger.error(f"[SLACK] submit_consult 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.action("value")
    def handle_status_change_dispatch(ack, body, client):
        """상담 모달의 처리 유형 변경 시 모달 재렌더링 (라벨/필수 동기화).

        action_id='value' 로 여러 필드가 공유하므로 block_id='status' 인 경우만 처리.
        """
        ack()
        try:
            actions = body.get("actions") or []
            if not actions or actions[0].get("block_id") != "status":
                return
            view = body.get("view") or {}
            state = view.get("state", {}).get("values", {})

            # 현재 상태에서 prefilled 재구성 (사용자가 입력한 값 보존)
            def _cur(bid):
                return (_v(state, bid) or '').strip() if bid in state else ''

            new_status = _cur("status") or '유선 상담'
            prefilled = {
                'visit_type': _cur("visit_type") or '온라인',
                'status': new_status,
                'visit_date': _cur("visit_date"),
                'visit_date_end': _cur("visit_date_end"),
                'name': _cur("name"),
                'contact': _cur("contact"),
                'email': _cur("email"),
                'visit_address': _cur("visit_address"),
                'consultation': _cur("consultation"),
            }
            metadata = view.get("private_metadata", "") or ""
            # info_blocks (인입 정보 section + divider) 유지 — Slack 이 재렌더 후
            # section 에 block_id 자동 부여하는 경우 있어서 block_id 여부로 필터 안 함
            info_blocks = [b for b in view.get("blocks", [])
                          if b.get("type") in ("section", "divider")]

            new_view = _build_consult_view(info_blocks, metadata, prefilled)
            client.views_update(view_id=view["id"], view=new_view)
        except Exception as exc:
            logger.warning(f"[SLACK/상담] 처리 유형 변경 재렌더 실패: {exc}")

    @app.action("button_visit")
    def handle_button_visit(ack, body, client):
        ack()
        def _bg():
            try:
                _open_inquiry_modal(client, body, action='visit')
            except Exception as exc:
                logger.error(f"[SLACK] button_visit 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    # ⑦ 인입 알림 메시지의 [가격 문의] 버튼
    @app.action("button_price")
    def handle_button_price(ack, body, client):
        ack()
        def _bg():
            try:
                _open_inquiry_modal(client, body, action='price')
            except Exception as exc:
                logger.error(f"[SLACK] button_price 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    # ⑧ 방문 요청 모달 제출
    @app.view("submit_visit")
    def handle_submit_visit(ack, body, client, view):
        ack()
        try:
            _process_visit_submission(client, body, view)
        except Exception as exc:
            logger.error(f"[SLACK] submit_visit 실패: {exc}", exc_info=True)

    # ⑨ 가격 문의 모달 제출
    @app.view("submit_price")
    def handle_submit_price(ack, body, client, view):
        ack()
        try:
            _process_price_submission(client, body, view)
        except Exception as exc:
            logger.error(f"[SLACK] submit_price 실패: {exc}", exc_info=True)

    # ⑩ /전화 슬래시 명령 — 전화 문의 등록 모달
    @app.command("/수금")
    def handle_payment_command(ack, command, client):
        """수금 관리 봇 — /수금 [코드 또는 '요약' 또는 '미수금']"""
        ack()
        text = command.get("text", "").strip()
        channel = command.get("channel_id", "")
        user_id = command.get("user_id", "")

        def _bg():
            try:
                from dashboard.services.payment_sync import (
                    search_project, daily_payment_summary, build_overdue_message,
                )
                if not text or text.lower() in ('도움', 'help', '안내'):
                    msg = (
                        "*수금 관리 봇 사용법*\n"
                        "• `/수금 G3491-YG` — 특정 프로젝트 history 조회\n"
                        "• `/수금 요약` — 오늘 발송 일일 요약\n"
                        "• `/수금 미수금` — 30일 이상 경과 미수금 리스트\n"
                        "• `/수금 미수금 60` — N일 이상 경과 (커스텀)"
                    )
                elif text.lower() == '요약':
                    msg = daily_payment_summary() or "오늘 발송 이력 없음"
                elif text.lower().startswith('미수금'):
                    parts = text.split()
                    days = 30
                    if len(parts) > 1:
                        try:
                            days = int(parts[1])
                        except Exception:
                            pass
                    msg = build_overdue_message(days=days) or f"{days}일 이상 경과한 미수금 없음"
                else:
                    msg = search_project(text) or f"`{text}` 검색 결과 없음"
                client.chat_postEphemeral(channel=channel, user=user_id, text=msg)
            except Exception as exc:
                logger.error(f"[SLACK] /수금 처리 실패: {exc}", exc_info=True)
        import threading
        threading.Thread(target=_bg, daemon=True).start()

    # 2026-07-12 전화 문의 등록 기능 전체 제거 — 사용 안 함.
    #   전화 문의는 별도 슬랙 워크플로 앱 '전화문의 등록하기' 로 이관.
    #   제거: /전화 슬래시 · button_phone 버튼 · phone_inquiry_shortcut App Shortcut
    #        · submit_phone view · _open_phone_modal · _post_phone_setup_message
    #        · _process_phone_submission 헬퍼

    # #방문_일정 카드의 [✏️ 방문일 수정] / [🗑️ 방문 취소] 액션은
    # 방문 일정 알림 봇(_visit_slack_app)이 처리 — _register_visit_handlers 참조

    # ⑬ /청소 슬래시 명령 — 채널 메시지 일괄 청소 (봇이 보낸 메시지만)
    # 권한 — SLACK_ADMIN_CHANNEL(=admin user ID) 만 실행 가능
    @app.command("/청소")
    def handle_sweep_command(ack, command, client, respond):
        ack()
        user_id = command.get("user_id", "")
        admin_uid = os.getenv('SLACK_ADMIN_CHANNEL', '').strip()
        if admin_uid and admin_uid.startswith('U') and user_id != admin_uid:
            respond({
                "response_type": "ephemeral",
                "text": ":no_entry: `/청소` 명령은 관리자만 실행할 수 있습니다.",
            })
            return

        text = command.get("text", "").strip()
        channel = command.get("channel_id", "")

        parsed = _parse_sweep_args(text)
        if not parsed["valid"]:
            respond({"response_type": "ephemeral", "text": parsed["error"]})
            return

        if parsed["mode"] == "all":
            mode_desc = "*전체* 메시지"
        elif parsed["mode"] == "count":
            mode_desc = f"최근 *{parsed['value']}개* 메시지"
        else:
            mode_desc = f"최근 *{_human_duration(parsed['value'])}* 이내 메시지"

        private_meta = json.dumps({
            "channel": channel,
            "mode": parsed["mode"],
            "value": parsed.get("value", 0),
        })

        respond({
            "response_type": "ephemeral",
            "text": f"🧹 {mode_desc}를 청소합니다.",
            "blocks": [
                {"type": "section", "text": {"type": "mrkdwn", "text": (
                    f"🧹 {mode_desc}를 청소합니다.\n"
                    "_• 봇이 보낸 메시지만 삭제됩니다 (Slack 정책)_\n"
                    "_• 1초당 1개 속도 (rate limit) — 100개 약 2분_"
                )}},
                {"type": "actions", "elements": [
                    {"type": "button", "text": {"type": "plain_text", "text": "✅ 시작"},
                     "style": "danger", "action_id": "sweep_confirm",
                     "value": private_meta},
                    {"type": "button", "text": {"type": "plain_text", "text": "취소"},
                     "action_id": "sweep_cancel"},
                ]},
            ],
        })

    # ⑭ [시작] 버튼 - 청소 백그라운드 실행
    @app.action("sweep_confirm")
    def handle_sweep_confirm(ack, body, client, respond):
        ack()
        try:
            meta = json.loads(body["actions"][0].get("value", "{}"))
        except Exception:
            respond({"response_type": "ephemeral", "text": "❌ 인자 디코딩 실패"})
            return

        response_url = body.get("response_url", "")
        channel = meta.get("channel", "")
        mode = meta.get("mode", "")
        value = meta.get("value", 0)

        respond({
            "response_type": "ephemeral",
            "replace_original": True,
            "text": "🧹 청소 시작... (50개마다 진행 보고)",
        })

        def _bg():
            try:
                _run_sweep(client, channel, response_url, mode, value)
            except Exception as exc:
                logger.error(f"[SWEEP] 실패: {exc}", exc_info=True)
                _sweep_update(response_url, f"❌ 청소 실패: {type(exc).__name__}: {exc}")
        threading.Thread(target=_bg, daemon=True).start()

    # ⑮ [취소] 버튼
    @app.action("sweep_cancel")
    def handle_sweep_cancel(ack, body, respond):
        ack()
        respond({
            "response_type": "ephemeral",
            "replace_original": True,
            "text": "_청소 취소됨._",
        })

    # /공사확정 + submit_project는 별도 공사 봇이 처리 (_init_project_slack_app)
    # /as + 사후 관리 흐름은 별도 A/S 봇이 처리 (_init_as_slack_app)

    logger.info(
        "[SLACK] 메인 봇 핸들러 등록 완료: /상태, /청소, app_mention, message(DM), "
        "button_visit, button_price, submit_visit, submit_price, "
        "sweep_confirm, sweep_cancel"
    )


def _register_invoice_handlers(app):
    """계산서 봇 핸들러."""

    @app.event("file_shared")  # no-op (2026-07-22)
    def _noop_file_shared_invoice(ack):
        ack()


    """세금계산서 관리 알림 봇 핸들러.

    - message event: #영업_관리 채널 스레드 첨부 감지 → 카드 자동 완료 update
    - invoice_complete action (backward compat): 이전 발송된 카드의 [✅ 발행 완료]
    """

    @app.event("message")
    def handle_invoice_thread_message(event, client):
        subtype = event.get("subtype") or ""
        has_files = bool(event.get("files"))
        thread_ts = event.get("thread_ts")
        channel = event.get('channel')

        # 파일 삭제 이벤트 (subtype=message_deleted) — 스레드에 안내 메시지
        if subtype == 'message_deleted':
            prev = event.get('previous_message', {}) or {}
            prev_thread_ts = prev.get('thread_ts', '')
            if prev_thread_ts and (prev.get('files') or []):
                try:
                    from dashboard.utils.redis_client import get_redis_client
                    rc = get_redis_client().redis
                    if rc.get(f'invoice_card:{channel}:{prev_thread_ts}'):
                        client.chat_postMessage(
                            channel=channel, thread_ts=prev_thread_ts,
                            text=(':warning: 세금계산서 첨부 파일이 삭제됐어요. '
                                  '확인이 필요하면 재첨부 부탁드립니다.'),
                        )
                        logger.info(f"[SLACK/계산서] 파일 삭제 알림 발송: thread={prev_thread_ts}")
                except Exception as del_exc:
                    logger.warning(f"[SLACK/계산서] 삭제 알림 처리 실패: {del_exc}")
            return

        # 스레드 파일 첨부만 처리 + 봇 자신 메시지 skip
        if not thread_ts or not has_files:
            return
        if subtype == 'bot_message' or event.get('bot_id'):
            return

        # Redis 에서 계산서 카드 metadata 조회
        try:
            from dashboard.utils.redis_client import get_redis_client
            rc = get_redis_client().redis
            meta_raw = rc.get(f'invoice_card:{channel}:{thread_ts}')
            if not meta_raw:
                return  # 계산서 스레드 아님
            meta = json.loads(
                meta_raw.decode() if isinstance(meta_raw, bytes) else meta_raw
            )
        except Exception as exc:
            logger.warning(f"[SLACK/계산서] 스레드 metadata 조회 실패: {exc}")
            return

        def _bg():
            try:
                _auto_complete_invoice_card(client, channel, thread_ts, event, meta)
            except Exception as exc:
                logger.error(f"[SLACK/계산서] 자동 완료 예외: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.action("invoice_complete")
    def handle_invoice_complete_action(ack, body, client):
        """Backward compat — 이전 발송된 카드의 [✅ 발행 완료] 버튼 처리."""
        ack()
        def _bg():
            try:
                _process_invoice_complete(client, body)
            except Exception as exc:
                logger.error(f"[SLACK/계산서] complete 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()


def _register_as_handlers(app):
    """A/S 사후 관리 봇 핸들러 — /as + 3단계 모달 흐름."""

    @app.event("file_shared")  # no-op (2026-07-22)
    def _noop_file_shared_as(ack):
        ack()

    @app.command("/as")
    def handle_as_command(ack, command, client):
        ack()
        trigger_id = command.get("trigger_id", "")
        user_id = command.get("user_id", "")
        if not trigger_id:
            return
        try:
            _open_as_request_modal(client, trigger_id, user_id)
        except Exception as exc:
            logger.error(f"[SLACK/AS] 요청 모달 열기 실패: {exc}", exc_info=True)

    @app.view("submit_as_request")
    def handle_submit_as_request(ack, body, client, view):
        ack()
        def _bg():
            try:
                _process_as_request_submission(client, body, view)
            except Exception as exc:
                logger.error(f"[SLACK/AS] submit_as_request 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.action("as_accept_open")
    def handle_as_accept_open(ack, body, client):
        ack()
        def _bg():
            try:
                _open_as_accept_modal(client, body)
            except Exception as exc:
                logger.error(f"[SLACK/AS] 접수 모달 열기 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.view("submit_as_accept")
    def handle_submit_as_accept(ack, body, client, view):
        # 방문 유형이 내부/외주면 담당자 이름 필수 검증
        values = view.get("state", {}).get("values", {})
        visitor_type = ''
        try:
            opt = values.get("visitor_type", {}).get("value", {}).get("selected_option", {})
            visitor_type = (opt or {}).get("value", '') or ''
        except Exception:
            pass
        visitor_name = ''
        try:
            visitor_name = (values.get("visitor_name", {}).get("value", {}) or {}).get("value", '') or ''
        except Exception:
            pass
        if visitor_type in ('내부', '외주') and not visitor_name.strip():
            ack(response_action="errors", errors={
                "visitor_name": "내부/외주 방문 시 방문 예정자 이름을 입력해야 합니다.",
            })
            return
        ack()
        def _bg():
            try:
                _process_as_accept_submission(client, body, view)
            except Exception as exc:
                logger.error(f"[SLACK/AS] submit_as_accept 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.action("as_complete_open")
    def handle_as_complete_open(ack, body, client):
        ack()
        def _bg():
            try:
                _open_as_complete_modal(client, body)
            except Exception as exc:
                logger.error(f"[SLACK/AS] 완료 모달 열기 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.view("submit_as_complete")
    def handle_submit_as_complete(ack, body, client, view):
        ack()
        def _bg():
            try:
                _process_as_complete_submission(client, body, view)
            except Exception as exc:
                logger.error(f"[SLACK/AS] submit_as_complete 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.options("value")
    def handle_as_bot_options(ack, body):
        """external_select 옵션 응답 (A/S 봇). block_id=as_project_code."""
        block_id = body.get("block_id", "")
        query = (body.get("value") or "").strip()
        if block_id == "as_project_code":
            try:
                from dashboard.services.as_service import search_confirmed_projects
                matched = search_confirmed_projects(query, limit=100)
                options = []
                for p in matched:
                    biz = (p.get('biz') or '').strip()
                    if not biz or biz == '-':
                        biz_disp = '사업자 비어 있음'
                    else:
                        biz_disp = biz
                    label = f'{p["code"]} : {biz_disp}'
                    options.append({
                        "text": {"type": "plain_text", "text": label[:75]},
                        "value": p["code"][:75],
                    })
                ack(options=options)
            except Exception as exc:
                logger.warning(f"[SLACK/AS/options] 실패: {exc}", exc_info=True)
                ack(options=[])
        else:
            ack(options=[])

    @app.action("value")
    def handle_as_block_action(ack, body, client):
        """모달 내 external_select 선택 → 프로젝트 정보 pre-fill 갱신."""
        ack()
        if not body.get("view"):
            return
        action = (body.get("actions") or [{}])[0]
        if action.get("block_id") != "as_project_code":
            return
        def _bg():
            try:
                _update_as_modal_with_project(client, body, action)
            except Exception as exc:
                logger.error(f"[SLACK/AS] 모달 갱신 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    logger.info(
        "[SLACK/AS봇] 핸들러 등록 완료: /as, submit_as_request, "
        "as_accept_open, submit_as_accept, as_complete_open, submit_as_complete, "
        "options(as_project_code)"
    )


# ─────────────────────────────────────────────────────────────
# /청소 — 채널 메시지 일괄 청소 헬퍼
# ─────────────────────────────────────────────────────────────
_BOT_INFO = {"user_id": "", "bot_id": ""}
_BOT_INFO_LOCK = threading.Lock()


def _get_bot_info(client) -> dict:
    """auth.test로 봇의 user_id/bot_id 조회. 한 번만 호출 후 캐시."""
    with _BOT_INFO_LOCK:
        if _BOT_INFO["user_id"]:
            return _BOT_INFO
        try:
            res = client.auth_test()
            _BOT_INFO["user_id"] = res.get("user_id", "")
            _BOT_INFO["bot_id"] = res.get("bot_id", "")
            logger.info(f"[SWEEP] 봇 ID 캐시: user_id={_BOT_INFO['user_id']}, bot_id={_BOT_INFO['bot_id']}")
        except Exception as exc:
            logger.warning(f"[SWEEP] auth.test 실패: {exc}")
    return _BOT_INFO


def _parse_sweep_args(text: str) -> dict:
    """/청소 인자 파싱.

    반환:
        {"valid": True, "mode": "count", "value": 100}
        {"valid": True, "mode": "duration", "value": 86400}  # 초
        {"valid": True, "mode": "all"}
        {"valid": False, "error": "..."}
    """
    import re as _re
    text = text.strip().lower()

    if not text or text in ("help", "도움말", "?"):
        return {"valid": False, "error": (
            "*사용법*\n"
            "`/청소 100` — 최근 100개 메시지 청소\n"
            "`/청소 24h` — 24시간 이내 메시지 청소\n"
            "`/청소 7d` — 7일 이내 메시지 청소\n"
            "`/청소 all` — 전체 청소 (위험)\n\n"
            "_※ 봇이 보낸 메시지만 삭제됩니다 (Slack 정책)_"
        )}

    if text == "all":
        return {"valid": True, "mode": "all"}

    # 시간 단위 (60m, 24h, 7d)
    m = _re.match(r'^(\d+)([mhd])$', text)
    if m:
        n = int(m.group(1))
        unit_secs = {"m": 60, "h": 3600, "d": 86400}[m.group(2)]
        return {"valid": True, "mode": "duration", "value": n * unit_secs}

    # 숫자 (최근 N개)
    m = _re.match(r'^(\d+)$', text)
    if m:
        n = int(m.group(1))
        if n <= 0 or n > 10000:
            return {"valid": False, "error": "개수는 1~10000 사이"}
        return {"valid": True, "mode": "count", "value": n}

    return {"valid": False, "error": f"인식 못 함: `{text}`. `/청소 help`로 사용법 확인"}


def _sweep_update(response_url: str, text: str):
    """response_url로 ephemeral 갱신 (실패해도 무시)."""
    if not response_url:
        return
    try:
        req = urllib.request.Request(
            response_url,
            data=json.dumps({
                "response_type": "ephemeral",
                "replace_original": True,
                "text": text,
            }).encode("utf-8"),
            headers={"Content-Type": "application/json"},
        )
        urllib.request.urlopen(req, timeout=5)
    except Exception as exc:
        logger.warning(f"[SWEEP] response_url 갱신 실패: {exc}")


# ─────────────────────────────────────────────────────────────
# A/S 사후 관리 헬퍼 (2026-07-09)
# ─────────────────────────────────────────────────────────────
def _as_status_emoji(status: str) -> str:
    if status == '접수 완료':
        return '📥'
    if status == '처리 완료':
        return '✅'
    return '🔔'


def _build_as_card_text(data: dict, view_state: str = 'requested', proj: Optional[dict] = None) -> str:
    """A/S 카드 본문 텍스트. view_state: requested / accepted / completed.

    공사 확정 카드와 동등한 정보량으로 렌더 — 유입 구분·발주처 담당자/연락처/이메일·
    도급 구분·시공자·공사 금액·공사 시작 추가.

    proj: 프로젝트 상세 (호출자가 미리 조회 후 전달 가능 — 중복 API 호출 방지).
    """
    # 프로젝트 상세 조회 (호출자가 전달 안 했으면 여기서 조회)
    code = str(data.get('프로젝트 코드', '') or '').strip()
    if proj is None and code and code != '-':
        try:
            from dashboard.services.as_service import get_project_details
            proj = get_project_details(code) or {}
        except Exception:
            proj = None

    def _pick(key_data: str, key_proj: str, default: str = '-') -> str:
        v = data.get(key_data)
        if v not in (None, '', '-'):
            return str(v)
        if proj:
            v2 = proj.get(key_proj)
            if v2 not in (None, '', '-'):
                return str(v2)
        return default

    as_no = data.get('No', '')
    lines = []
    if view_state == 'requested':
        lines.append(f"🔔 *[A/S 요청]*  `{as_no}`")
    elif view_state == 'accepted':
        lines.append(f"📥 *[A/S 접수 완료]*  `{as_no}`")
    else:
        lines.append(f"✅ *[A/S 처리 완료]*  `{as_no}`")
    lines.append("--------------------------------------------")

    inflow = (proj or {}).get('inflow', '-') if proj else '-'
    biz = (proj or {}).get('biz', '-') if proj else '-'

    lines.append(f"🔗 프로젝트 코드 : `{code or '-'}`")
    lines.append(f"📥 유입 구분 : {inflow or '-'}")
    lines.append(f"🏢 사업자명 : {biz or '-'}")
    lines.append(f"📍 현장 주소 : {_pick('현장주소', 'address')}")
    lines.append(f"👤 발주처 담당자 : {(proj or {}).get('client_manager', '-') or '-'}")
    lines.append(f"📞 발주처 연락처 : {(proj or {}).get('client_phone', '-') or '-'}")
    lines.append(f"✉️ 발주처 이메일 : {(proj or {}).get('client_email', '-') or '-'}")
    lines.append(f"📋 공사 내용 : {_pick('공사내용', 'work_content')}")
    lines.append(f"🛠️ 도급 구분 : {(proj or {}).get('contract_type', '-') or '-'}")
    lines.append(f"👷 시공자 : {(proj or {}).get('contractor', '-') or '-'}")
    lines.append(f"💲 공사 금액 : {(proj or {}).get('amount', '-') or '-'}")
    lines.append(f"📅 공사 시작 : {(proj or {}).get('work_start', '-') or '-'}")
    lines.append(f"📅 공사 종료 : {_pick('공사 종료일', 'work_end')}")
    lines.append("--------------------------------------------")
    lines.append(f"📝 A/S 요청 내용 : {data.get('요청 내용', '-') or '-'}")
    lines.append(f"👤 요청자 : {data.get('요청자', '-') or '-'}")
    if view_state in ('accepted', 'completed'):
        lines.append("--------------------------------------------")
        lines.append(f"👷 방문 예정자 : {data.get('방문 예정자', '-') or '-'}")
        lines.append(f"📅 방문 예정일 : {data.get('방문 예정일', '-') or '-'}")
        lines.append(f"✅ 접수자 : {data.get('접수자', '-') or '-'}  {data.get('접수 일자', '')}")
    if view_state == 'completed':
        lines.append("--------------------------------------------")
        lines.append(f"🎯 처리 내용 : {data.get('처리 내용', '-') or '-'}")
    lines.append("--------------------------------------------")
    return "⠀\n" + "\n".join(lines)


def _build_as_blocks(data: dict, view_state: str = 'requested') -> list:
    # 프로젝트 상세 한 번만 조회 (card text + button value 양쪽에서 재사용)
    proj = None
    code = str(data.get('프로젝트 코드', '') or '').strip()
    if code and code != '-':
        try:
            from dashboard.services.as_service import get_project_details
            proj = get_project_details(code) or {}
        except Exception:
            proj = None
    text = _build_as_card_text(data, view_state=view_state, proj=proj)
    # section 하단 구분선(-----)과 버튼 사이 여백 제거 (2026-07-09 UX).
    blocks: list = [
        {"type": "section", "text": {"type": "mrkdwn", "text": text}},
    ]
    as_no = data.get('No', '')
    # button value 에 시공자 이름 함께 저장 — 모달 오픈 시 시트 API 재조회 skip
    # → trigger_id 3초 만료 방지 (2026-07-13 사용자 관측: 모달 늦게 열림).
    contractor = (proj.get('contractor', '') or '').strip() if proj else ''
    accept_value = json.dumps({'as_no': as_no, 'contractor': contractor}, ensure_ascii=False)
    if view_state == 'requested':
        blocks.append({
            "type": "actions",
            "elements": [{
                "type": "button",
                "text": {"type": "plain_text", "text": "🛠️ A/S 접수하기", "emoji": True},
                "style": "primary",
                "action_id": "as_accept_open",
                "value": accept_value,
            }],
        })
    elif view_state == 'accepted':
        blocks.append({
            "type": "actions",
            "elements": [{
                "type": "button",
                "text": {"type": "plain_text", "text": "🎯 처리 완료하기", "emoji": True},
                "style": "primary",
                "action_id": "as_complete_open",
                "value": as_no,
            }],
        })
    # completed: no buttons
    blocks.append({"type": "context", "elements": [{"type": "mrkdwn", "text": "⠀"}]})
    return blocks


def _as_request_view_blocks(
    initial_project_option: Optional[dict] = None,
    project_details: Optional[dict] = None,
    initial_request_content: str = '',
) -> list:
    """요청 모달 blocks — 프로젝트 선택 전/후 공용."""
    project_element = {
        "type": "external_select", "action_id": "value",
        "min_query_length": 1,
        "placeholder": {"type": "plain_text", "text": "예: G3745 / R3845 (1글자부터 검색)"},
    }
    if initial_project_option:
        project_element["initial_option"] = initial_project_option

    blocks: list = [
        {
            "type": "input", "block_id": "as_project_code",
            "label": {"type": "plain_text", "text": "프로젝트 코드 (검색해서 선택)"},
            "element": project_element,
            "dispatch_action": True,  # 선택 즉시 block_actions 발동해 상세 pre-fill
        },
    ]
    if project_details:
        info = (
            f"*📥 유입 구분 :* {project_details.get('inflow','-') or '-'}\n"
            f"*🏢 사업자명 :* {project_details.get('biz','-') or '-'}\n"
            f"*📍 현장 주소 :* {project_details.get('address','-') or '-'}\n"
            f"*👤 발주처 담당자 :* {project_details.get('client_manager','-') or '-'}\n"
            f"*📞 발주처 연락처 :* {project_details.get('client_phone','-') or '-'}\n"
            f"*✉️ 발주처 이메일 :* {project_details.get('client_email','-') or '-'}\n"
            f"*📋 공사 내용 :* {project_details.get('work_content','-') or '-'}\n"
            f"*🛠️ 도급 구분 :* {project_details.get('contract_type','-') or '-'}\n"
            f"*👷 시공자 :* {project_details.get('contractor','-') or '-'}\n"
            f"*💲 공사 금액 :* {project_details.get('amount','-') or '-'}\n"
            f"*📅 공사 시작 :* {project_details.get('work_start','-') or '-'}\n"
            f"*📅 공사 종료 :* {project_details.get('work_end','-') or '-'}"
        )
        # 상단 여백 (⠀ context) + 정보 섹션 + 하단 여백 divider
        blocks.append({"type": "context", "elements": [{"type": "mrkdwn", "text": "⠀"}]})
        blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": info}})
        blocks.append({"type": "context", "elements": [{"type": "mrkdwn", "text": "⠀"}]})
        blocks.append({"type": "divider"})

    request_element = {
        "type": "plain_text_input", "action_id": "value", "multiline": True,
        "placeholder": {"type": "plain_text", "text": "예: 실외기 소음 발생, 점검 필요"},
    }
    if initial_request_content:
        request_element["initial_value"] = initial_request_content
    blocks.append({
        "type": "input", "block_id": "request_content",
        "label": {"type": "plain_text", "text": "A/S 요청 내용"},
        "element": request_element,
    })
    return blocks


def _open_as_request_modal(client, trigger_id: str, user_id: str) -> None:
    """`/as` 슬래시 → 요청 모달."""
    metadata = json.dumps({"user_id": user_id}, ensure_ascii=False)
    view = {
        "type": "modal",
        "callback_id": "submit_as_request",
        "private_metadata": metadata,
        "title": {"type": "plain_text", "text": "A/S 요청"},
        "submit": {"type": "plain_text", "text": "제출"},
        "close": {"type": "plain_text", "text": "취소"},
        "blocks": _as_request_view_blocks(),
    }
    client.views_open(trigger_id=trigger_id, view=view)


def _update_as_modal_with_project(client, body, action) -> None:
    """external_select 선택 → 프로젝트 상세 pre-fill 후 views.update."""
    from dashboard.services.as_service import get_project_details

    selected_option = action.get("selected_option") or {}
    selected_code = selected_option.get("value", '').strip()
    if not selected_code:
        return

    view = body["view"]
    view_id = view.get("id", '')
    view_hash = view.get("hash", '')
    metadata = view.get("private_metadata", '') or json.dumps({}, ensure_ascii=False)

    # 기존 A/S 요청 내용 보존
    current_content = ''
    try:
        current_content = (
            (view.get("state", {}) or {}).get("values", {})
            .get("request_content", {}).get("value", {})
            .get("value", '') or ''
        )
    except Exception:
        current_content = ''

    details = get_project_details(selected_code) or {}
    new_view = {
        "type": "modal",
        "callback_id": "submit_as_request",
        "private_metadata": metadata,
        "title": {"type": "plain_text", "text": "A/S 요청"},
        "submit": {"type": "plain_text", "text": "제출"},
        "close": {"type": "plain_text", "text": "취소"},
        "blocks": _as_request_view_blocks(
            initial_project_option=selected_option,
            project_details=details,
            initial_request_content=current_content,
        ),
    }
    try:
        client.views_update(view_id=view_id, hash=view_hash, view=new_view)
    except Exception as exc:
        logger.warning(f"[SLACK/AS] views_update 실패: {exc}")


def _process_as_request_submission(client, body, view) -> None:
    """요청 제출 → 프로젝트 정보 조회 → 시트 append → 카드 발송."""
    from dashboard.services.as_service import get_project_details, create_as_row

    values = view["state"]["values"]
    project_code = ''
    try:
        opt = values.get("as_project_code", {}).get("value", {}).get("selected_option", {})
        project_code = (opt or {}).get("value", "") or ''
    except Exception:
        pass
    request_content = ''
    try:
        request_content = (values.get("request_content", {}).get("value", {}) or {}).get("value", '') or ''
    except Exception:
        pass
    request_content = request_content.strip()
    project_code = project_code.strip()

    user_id = body.get("user", {}).get("id", "")
    requester_initial = _slack_user_to_initial(client, user_id) or '-'

    if not project_code:
        logger.warning('[SLACK/AS] 프로젝트 코드 누락')
        return

    details = get_project_details(project_code) or {}
    as_no, row_num = create_as_row(
        project_code=project_code,
        address=details.get('address', ''),
        work_content=details.get('work_content', ''),
        work_end=details.get('work_end', ''),
        request_content=request_content,
        requester=requester_initial,
    )

    channel = os.getenv('SLACK_AS_CHANNEL', '').strip()
    if not channel:
        logger.warning('[SLACK/AS] SLACK_AS_CHANNEL 미설정 — 카드 발송 skip')
        return

    card_data = {
        'No': as_no,
        '프로젝트 코드': project_code,
        '현장주소': details.get('address', ''),
        '공사내용': details.get('work_content', ''),
        '공사 종료일': details.get('work_end', ''),
        '요청 내용': request_content,
        '요청자': requester_initial,
    }
    text = f"[A/S 요청] {as_no} {project_code}"
    blocks = _build_as_blocks(card_data, view_state='requested')

    try:
        client.conversations_join(channel=channel)
    except Exception:
        pass
    resp = client.chat_postMessage(channel=channel, text=text, blocks=blocks, unfurl_links=False)
    if resp.get('ok'):
        ts = resp.get('ts', '')
        try:
            from dashboard.utils.redis_client import get_redis_client
            rc = get_redis_client().redis
            rc.set(f'as_card_msg:{as_no}', f'{channel}|{ts}', ex=60 * 60 * 24 * 365)
        except Exception as exc:
            logger.warning(f'[SLACK/AS] card 매핑 저장 실패: {exc}')
        logger.info(f'[SLACK/AS] 요청 카드 발송 완료: {as_no} ts={ts}')


def _open_as_accept_modal(client, body) -> None:
    """[✅ A/S 접수하기] 클릭 → 접수 모달.

    방문 유형(서비스 기사/내부/외주) 선택 후 담당자 이름을 별도 칸에 입력.
    서비스 기사 방문 시 담당자 이름 칸은 비워두면 되고, 그 외에는 필수.

    2026-07-13: 외주 케이스 대비 원본 시공자 이름을 이름 필드에 pre-fill.
    매니저가 '외주' 선택 시 시공자 이름을 다시 타이핑할 필요 없음. 내부는 지우고
    담당자 이름으로 변경.
    """
    trigger_id = body["trigger_id"]
    # button value 는 JSON `{as_no, contractor}` 또는 fallback 로 as_no 문자열 (구 카드)
    raw_val = (body["actions"][0].get("value") or '').strip()
    as_no, contractor = '', ''
    try:
        payload = json.loads(raw_val) if raw_val.startswith('{') else {}
        as_no = (payload.get('as_no', '') or '').strip()
        contractor = (payload.get('contractor', '') or '').strip()
    except Exception:
        pass
    if not as_no:
        as_no = raw_val  # 구 카드 fallback
    channel = body.get("channel", {}).get("id", "")
    message_ts = body.get("message", {}).get("ts", "")

    metadata = json.dumps({
        "as_no": as_no, "channel": channel, "message_ts": message_ts,
    }, ensure_ascii=False)

    # 구 카드로부터 열린 경우 contractor 가 payload 에 없음 → 시트 fallback 조회
    # (trigger_id 3초 만료 위험 있으나 구 카드에만 해당)
    if not contractor:
        try:
            from dashboard.services.as_service import get_as_data, get_project_details
            as_data = get_as_data(as_no) or {}
            code = (as_data.get('프로젝트 코드', '') or '').strip()
            if code and code != '-':
                proj = get_project_details(code) or {}
                _c = (proj.get('contractor', '') or '').strip()
                if _c and _c != '-':
                    contractor = _c
        except Exception as exc:
            logger.warning(f'[SLACK/AS] 시공자 pre-fill 조회 실패 (무시): {exc}')

    visitor_type_options = [
        {"text": {"type": "plain_text", "text": "서비스 기사"}, "value": "서비스 기사"},
        {"text": {"type": "plain_text", "text": "내부 (아이티)"}, "value": "내부"},
        {"text": {"type": "plain_text", "text": "외주 (시공자)"}, "value": "외주"},
    ]

    view = {
        "type": "modal",
        "callback_id": "submit_as_accept",
        "private_metadata": metadata,
        "title": {"type": "plain_text", "text": "A/S 접수"},
        "submit": {"type": "plain_text", "text": "접수 확정"},
        "close": {"type": "plain_text", "text": "취소"},
        "blocks": [
            {
                "type": "input", "block_id": "visitor_type",
                "label": {"type": "plain_text", "text": "방문 예정자"},
                "element": {
                    "type": "static_select", "action_id": "value",
                    "placeholder": {"type": "plain_text", "text": "선택"},
                    "options": visitor_type_options,
                },
            },
            {
                "type": "input", "block_id": "visitor_name", "optional": True,
                "label": {"type": "plain_text", "text": "방문 예정자 이름 (내부/외주 방문 시 필수)"},
                "hint": {"type": "plain_text",
                         "text": "외주 선택 시 시공자 이름 자동 채워짐. 내부는 담당자 이름으로 변경."},
                "element": {
                    "type": "plain_text_input", "action_id": "value",
                    "placeholder": {"type": "plain_text", "text": "예: 김철수"},
                    **({"initial_value": contractor} if contractor and contractor != '-' else {}),
                },
            },
            {
                "type": "input", "block_id": "visit_date_start",
                "label": {"type": "plain_text", "text": "방문 예정일 (시작)"},
                "element": {"type": "datepicker", "action_id": "value"},
            },
            {
                "type": "input", "block_id": "visit_date_end", "optional": True,
                "label": {"type": "plain_text", "text": "방문 예정일 (종료)"},
                "hint": {"type": "plain_text",
                         "text": "여러 날 방문 (예: 7/1~7/3) 일 때만 입력. 단일이면 비워두세요."},
                "element": {"type": "datepicker", "action_id": "value"},
            },
        ],
    }
    client.views_open(trigger_id=trigger_id, view=view)


def _process_as_accept_submission(client, body, view) -> None:
    """접수 제출 → 시트 갱신 → 카드 chat.update (State 2)."""
    from dashboard.services.as_service import (
        update_as_row, get_as_data,
        COL_ACCEPTER, COL_ACCEPT_DATE, COL_VISITOR, COL_VISIT_DATE, COL_STATUS,
        STATUS_ACCEPTED,
    )
    from dashboard.blueprints.slack_helpers import _format_visit_date_range

    metadata = json.loads(view.get("private_metadata") or "{}")
    as_no = metadata.get("as_no", '')
    channel = metadata.get("channel", '')
    message_ts = metadata.get("message_ts", '')
    if not as_no:
        return

    values = view["state"]["values"]
    visitor_type = ''
    try:
        opt = values.get("visitor_type", {}).get("value", {}).get("selected_option", {})
        visitor_type = (opt or {}).get("value", '') or ''
    except Exception:
        pass
    visitor_name = ''
    try:
        visitor_name = (values.get("visitor_name", {}).get("value", {}) or {}).get("value", '') or ''
    except Exception:
        pass
    visitor_name = visitor_name.strip()
    # 서비스 기사 → '서비스 기사' 그대로. 내부/외주 → 입력한 담당자 이름을 그대로 사용.
    if visitor_type == '서비스 기사':
        visitor = '서비스 기사'
    else:
        visitor = visitor_name or visitor_type or '-'
    date_start = (values.get("visit_date_start", {}).get("value", {}) or {}).get("selected_date", '') or ''
    date_end = (values.get("visit_date_end", {}).get("value", {}) or {}).get("selected_date", '') or ''
    visit_date = _format_visit_date_range(date_start, date_end)

    user_id = body.get("user", {}).get("id", "")
    accepter = _slack_user_to_initial(client, user_id) or '-'
    accept_dt = datetime.now().strftime('%Y.%m.%d. %H:%M')

    ok = update_as_row(as_no, {
        COL_ACCEPTER: accepter,
        COL_ACCEPT_DATE: accept_dt,
        COL_VISITOR: visitor,
        COL_VISIT_DATE: visit_date,
        COL_STATUS: STATUS_ACCEPTED,
    })
    if not ok:
        logger.warning(f'[SLACK/AS] 시트 갱신 실패 ({as_no})')

    # 카드 chat.update — 시트 재조회로 완전한 데이터 사용
    # 2026-07-20 AS-0006 관측: Google Sheets eventual consistency 로 재조회 시 방금
    # 저장한 J/K/L 값이 아직 반영 안 되어 카드에 '-' 로 표시되는 이슈. 방금 update
    # 한 값을 직접 덮어써서 지연 우회.
    data = get_as_data(as_no) or {}
    data['접수자'] = accepter
    data['접수 일자'] = accept_dt
    data['방문 예정자'] = visitor
    data['방문 예정일'] = visit_date
    data['진행 상태'] = STATUS_ACCEPTED
    text = f"[A/S 접수 완료] {as_no}"
    blocks = _build_as_blocks(data, view_state='accepted')
    try:
        client.chat_update(channel=channel, ts=message_ts, text=text, blocks=blocks)
        logger.info(f'[SLACK/AS] 접수 완료: {as_no} by {accepter}')
    except Exception as exc:
        logger.error(f'[SLACK/AS] chat.update 실패 ({as_no}): {exc}', exc_info=True)


def _open_as_complete_modal(client, body) -> None:
    """[🎯 처리 완료] 클릭 → 처리 완료 모달 (처리 내용)."""
    trigger_id = body["trigger_id"]
    as_no = (body["actions"][0].get("value") or '').strip()
    channel = body.get("channel", {}).get("id", "")
    message_ts = body.get("message", {}).get("ts", "")

    metadata = json.dumps({
        "as_no": as_no, "channel": channel, "message_ts": message_ts,
    }, ensure_ascii=False)

    view = {
        "type": "modal",
        "callback_id": "submit_as_complete",
        "private_metadata": metadata,
        "title": {"type": "plain_text", "text": "A/S 처리 완료"},
        "submit": {"type": "plain_text", "text": "처리 완료"},
        "close": {"type": "plain_text", "text": "취소"},
        "blocks": [
            {
                "type": "input", "block_id": "resolution",
                "label": {"type": "plain_text", "text": "처리 내용"},
                "element": {
                    "type": "plain_text_input", "action_id": "value", "multiline": True,
                    "placeholder": {"type": "plain_text", "text": "예: 실외기 팬 교체, 소음 해소 확인"},
                },
            },
        ],
    }
    client.views_open(trigger_id=trigger_id, view=view)


def _process_as_complete_submission(client, body, view) -> None:
    """처리 완료 제출 → 시트 갱신 → 카드 chat.update (State 3)."""
    from dashboard.services.as_service import (
        update_as_row, get_as_data,
        COL_STATUS, COL_RESOLUTION, STATUS_COMPLETED,
    )

    metadata = json.loads(view.get("private_metadata") or "{}")
    as_no = metadata.get("as_no", '')
    channel = metadata.get("channel", '')
    message_ts = metadata.get("message_ts", '')
    if not as_no:
        return

    values = view["state"]["values"]
    resolution = (values.get("resolution", {}).get("value", {}) or {}).get("value", '') or ''
    resolution = resolution.strip()
    if not resolution:
        logger.warning(f'[SLACK/AS] 처리 내용 누락 ({as_no})')
        return

    ok = update_as_row(as_no, {
        COL_STATUS: STATUS_COMPLETED,
        COL_RESOLUTION: resolution,
    })
    if not ok:
        logger.warning(f'[SLACK/AS] 완료 갱신 실패 ({as_no})')

    # 2026-07-20 eventual consistency 우회 — 방금 update 한 값 직접 반영
    data = get_as_data(as_no) or {}
    data['처리 내용'] = resolution
    data['진행 상태'] = STATUS_COMPLETED
    text = f"[A/S 처리 완료] {as_no}"
    blocks = _build_as_blocks(data, view_state='completed')
    try:
        client.chat_update(channel=channel, ts=message_ts, text=text, blocks=blocks)
        logger.info(f'[SLACK/AS] 처리 완료: {as_no}')
    except Exception as exc:
        logger.error(f'[SLACK/AS] chat.update 실패 ({as_no}): {exc}', exc_info=True)


def _run_sweep(client, channel: str, response_url: str, mode: str, value: int):
    """채널 청소 백그라운드 워커.

    - conversations.history로 페이지네이션
    - 봇 user_id 또는 bot_id 매칭 시 chat.delete (1초당 1개)
    - 50개마다 진행 보고, 완료 시 결과 보고
    """
    import time as _time
    bot = _get_bot_info(client)
    bot_uid = bot.get("user_id", "")
    bot_bid = bot.get("bot_id", "")
    if not bot_uid and not bot_bid:
        _sweep_update(response_url, "❌ 봇 정보 확인 실패 (auth.test 실패)")
        return

    deleted = 0
    skipped_not_ours = 0
    delete_failed = 0
    cursor = None
    oldest = ""
    target_count = value if mode == "count" else None
    if mode == "duration":
        oldest = str(_time.time() - value)

    while True:
        try:
            params = {"channel": channel, "limit": 100}
            if cursor:
                params["cursor"] = cursor
            if oldest:
                params["oldest"] = oldest
            res = client.conversations_history(**params)
        except Exception as exc:
            _sweep_update(response_url, f"❌ history 조회 실패: {exc}")
            return

        msgs = res.get("messages", []) or []
        for m in msgs:
            if target_count is not None and deleted >= target_count:
                break

            is_ours = (m.get("user") == bot_uid) or (m.get("bot_id") == bot_bid)
            if not is_ours:
                skipped_not_ours += 1
                continue

            ts = m.get("ts")
            if not ts:
                continue

            try:
                client.chat_delete(channel=channel, ts=ts)
                deleted += 1
                _time.sleep(1.1)  # Slack tier-3 rate limit (50/min) 안전 대기
            except Exception as exc:
                delete_failed += 1
                logger.warning(f"[SWEEP] chat.delete 실패 ({ts}): {exc}")
                _time.sleep(1.1)

            if deleted > 0 and deleted % 50 == 0:
                _sweep_update(response_url, f"🧹 진행 중... 삭제 {deleted}개")

        if target_count is not None and deleted >= target_count:
            break
        cursor = (res.get("response_metadata") or {}).get("next_cursor", "")
        if not cursor:
            break

    _sweep_update(
        response_url,
        (
            f"✅ 청소 완료\n"
            f"• 삭제: {deleted}개\n"
            f"• 봇 메시지 아님 (스킵): {skipped_not_ours}개\n"
            f"• 삭제 실패: {delete_failed}개"
        ),
    )


# ─────────────────────────────────────────────────────────────
# /공사확정 — 슬랙 모달로 공사 확정 등록 (모바일 친화)
# ─────────────────────────────────────────────────────────────
_PROJECT_COMPANY_OPTIONS = ["글로벌", "글로벌그룹", "플랜트"]
_PROJECT_SOURCE_OPTIONS = ["거래처", "온라인", "당근", "소개", "숨고"]


def _search_company_names(query: str) -> list:
    """시트의 사업자명 unique 목록에서 query 부분 매칭. 슬랙 옵션 형식으로 반환.

    Slack 제약:
    - 최대 100 options per response
    - option text/value 최대 75자
    - min_query_length=1 — 빈 query는 빈 결과
    """
    if not query:
        return []
    from dashboard.services.project_service import load_data
    df = load_data()
    if df is None or df.empty or '사업자명' not in df.columns:
        return []
    # 시트의 사업자명 unique (캐시돼 있어 빠름)
    names = df['사업자명'].dropna().astype(str).str.strip().unique().tolist()
    names = [n for n in names if n and n != '-']
    # 부분 매칭 (대소문자 무관)
    q = query.lower()
    matched = [n for n in names if q in n.lower()]
    matched = sorted(set(matched))[:100]  # 슬랙 100개 제한
    return [
        {"text": {"type": "plain_text", "text": n[:75]}, "value": n[:75]}
        for n in matched
    ]


def _open_project_modal(client, trigger_id: str, channel: str, user_id: str):
    """공사 확정 등록 모달 — 핵심 11개 필드."""
    def _select_options(values):
        return [
            {"text": {"type": "plain_text", "text": v}, "value": v}
            for v in values
        ]

    metadata = json.dumps({"channel": channel, "user_id": user_id})
    view = {
        "type": "modal",
        "callback_id": "submit_project",
        "private_metadata": metadata,
        "title": {"type": "plain_text", "text": "공사 확정 등록"},
        "submit": {"type": "plain_text", "text": "등록"},
        "close": {"type": "plain_text", "text": "취소"},
        "blocks": [
            {
                "type": "input", "block_id": "company",
                "label": {"type": "plain_text", "text": "사업자"},
                "element": {
                    "type": "static_select", "action_id": "value",
                    "placeholder": {"type": "plain_text", "text": "선택"},
                    "options": _select_options(_PROJECT_COMPANY_OPTIONS),
                },
            },
            {
                "type": "input", "block_id": "source",
                "label": {"type": "plain_text", "text": "유입 구분"},
                "element": {
                    "type": "static_select", "action_id": "value",
                    "placeholder": {"type": "plain_text", "text": "선택"},
                    "options": _select_options(_PROJECT_SOURCE_OPTIONS),
                },
            },
            {
                "type": "input", "block_id": "company_name",
                "optional": True,
                "label": {"type": "plain_text", "text": "사업자명 (고객사)"},
                "element": {
                    "type": "external_select",
                    "action_id": "value",
                    "min_query_length": 1,
                    "placeholder": {
                        "type": "plain_text",
                        "text": "예: 삼성 / 한국 / 김밥 (신규 거래처면 비워두세요)",
                    },
                },
            },
            {
                "type": "input", "block_id": "address",
                "label": {"type": "plain_text", "text": "현장 주소"},
                "element": {
                    "type": "plain_text_input", "action_id": "value",
                    "placeholder": {"type": "plain_text", "text": "예: 서울 강남구 테헤란로 152"},
                },
            },
            {
                "type": "input", "block_id": "customer", "optional": True,
                "label": {"type": "plain_text", "text": "발주처 담당자"},
                "element": {"type": "plain_text_input", "action_id": "value"},
            },
            {
                "type": "input", "block_id": "contact", "optional": True,
                "label": {"type": "plain_text", "text": "발주처 연락처"},
                "element": {
                    "type": "plain_text_input", "action_id": "value",
                    "placeholder": {"type": "plain_text", "text": "예: 010-1234-5678"},
                },
            },
            {
                "type": "input", "block_id": "start_date",
                "label": {"type": "plain_text", "text": "공사 시작"},
                "element": {"type": "datepicker", "action_id": "value"},
            },
            {
                "type": "input", "block_id": "end_date",
                "label": {"type": "plain_text", "text": "공사 종료"},
                "element": {"type": "datepicker", "action_id": "value"},
            },
            {
                "type": "input", "block_id": "content",
                "label": {"type": "plain_text", "text": "공사 내용"},
                "element": {
                    "type": "plain_text_input", "action_id": "value", "multiline": True,
                    "placeholder": {"type": "plain_text", "text": "예: LG 천장형 4way 2대 설치"},
                },
            },
            {
                "type": "input", "block_id": "amount",
                "label": {"type": "plain_text", "text": "공사 금액 (VAT 별도, 숫자만)"},
                "element": {
                    "type": "plain_text_input", "action_id": "value",
                    "placeholder": {"type": "plain_text", "text": "예: 4600000"},
                },
            },
            {
                "type": "input", "block_id": "vat", "optional": True,
                "label": {"type": "plain_text", "text": "부가세"},
                "element": {
                    "type": "checkboxes", "action_id": "value",
                    "options": [
                        {"text": {"type": "plain_text", "text": "VAT 별도 (10% 추가)"},
                         "value": "true"},
                    ],
                },
            },
        ],
    }
    client.views_open(trigger_id=trigger_id, view=view)


def _process_project_submission(client, body, view):
    """공사 확정 모달 제출 → 시트 등록 + Calendar + #공사_확정 알림."""
    metadata = json.loads(view.get("private_metadata") or "{}")
    channel = metadata.get("channel", "")
    user_id = metadata.get("user_id") or body["user"]["id"]
    state = view["state"]["values"]

    # 모달 입력
    company = _v(state, "company")
    source = _v(state, "source")
    company_name = (_v(state, "company_name") or '').strip() or '-'  # 선택 입력
    address = (_v(state, "address") or '').strip()
    customer = (_v(state, "customer") or '').strip()
    contact = (_v(state, "contact") or '').strip()
    start_date = _v(state, "start_date")
    end_date = _v(state, "end_date")
    content = (_v(state, "content") or '').strip()
    amount_raw = (_v(state, "amount") or '').strip()
    vat_separate = bool(_v_multi(state, "vat"))

    # 영업 담당자: 슬랙 사용자 → 한국 이름
    # 공사 봇 토큰엔 users:read.email 스코프가 없을 수 있어 메인 봇 client로 매핑
    manager_name = ''
    if _slack_app:
        try:
            manager_name = _slack_user_to_korean_name(_slack_app.client, user_id)
        except Exception as exc:
            logger.warning(f"[SLACK/공사확정] 메인 봇으로 사용자 매핑 실패: {exc}")
    if not manager_name:
        # 폴백: 공사 봇 client로 시도
        manager_name = _slack_user_to_korean_name(client, user_id) or '미지정'

    # 금액 정규화 (콤마/원 제거)
    amount_digits = ''.join(ch for ch in amount_raw if ch.isdigit())

    data = {
        '사업자': company,
        '담당자': manager_name,
        '유입 구분': source,
        '사업자명': company_name,
        '현장 주소': address,
        '발주처 담당자': customer,
        '발주처 연락처': contact,
        '공사 시작': start_date,
        '공사 종료': end_date,
        '공사 내용': content,
        '총액 1': amount_digits or '0',
        '부가세': vat_separate,
    }

    try:
        code = _slack_create_project(data)
        msg = (
            f":white_check_mark: *{company_name or '공사'}* 등록 완료 — "
            f"`{code}`\n_담당자: {manager_name} · 시작: {start_date} · "
            f"금액: {int(amount_digits or 0):,}원_"
        )
        client.chat_postEphemeral(
            channel=channel or user_id, user=user_id, text=msg,
        )
    except Exception as exc:
        logger.error(f"[SLACK/공사확정] 등록 실패: {exc}", exc_info=True)
        client.chat_postEphemeral(
            channel=channel or user_id, user=user_id,
            text=f":x: 등록 실패: {type(exc).__name__}: {exc}",
        )


def _slack_create_project(data: dict) -> str:
    """슬랙 진입점 — 시트 등록 + 후처리. 성공 시 프로젝트 코드 반환.

    Flask request context 없이 동작 (직접 service 함수 호출).
    """
    from dashboard.services.project_service import (
        get_sheets_manager, load_data, _auto_project_code,
        invalidate_project_cache,
    )
    from dashboard.blueprints.projects import (
        _prepare_project_defaults, _build_row_values, _build_project_response_data,
    )
    from dashboard.services.calendar_service import create_project_calendar_event
    from dashboard.services.project_slack_notifier import send_project_created_notification

    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    if not sheet_id:
        raise Exception('GOOGLE_SHEET_ID 미설정')

    df = load_data()
    next_row = (len(df) + 2) if df is not None and not df.empty else 2

    company = str(data.get('사업자', '')).strip()
    owner = str(data.get('담당자', '')).strip()
    code = _auto_project_code(df, company, owner)
    if not code:
        raise Exception(f'프로젝트 코드 생성 실패 (사업자={company}, 담당자={owner})')
    data['프로젝트 코드'] = code

    manager = get_sheets_manager()
    _prepare_project_defaults(data, next_row)
    values = _build_row_values(data, manager, next_row)
    result = manager.append_row(sheet_id, values)
    if not result:
        raise Exception('시트 등록 실패')

    # 후처리 (실패해도 등록은 성공으로 간주)
    try:
        invalidate_project_cache(code)
    except Exception as exc:
        logger.warning(f"[SLACK/공사확정] 캐시 무효화 실패: {exc}")

    try:
        project_data = _build_project_response_data(code, data)
        create_project_calendar_event(project_data)
    except Exception as exc:
        logger.warning(f"[SLACK/공사확정] Calendar 등록 실패: {exc}")

    try:
        send_project_created_notification(data, code)
    except Exception as exc:
        logger.warning(f"[SLACK/공사확정] #공사_확정 알림 실패: {exc}")

    return code


# ─────────────────────────────────────────────────────────────
# Flask endpoint — 슬랙이 호출하는 단일 진입점
# ─────────────────────────────────────────────────────────────
def _run_bg_with_notify(client, body, action_label: str, work_fn) -> None:
    """배경 스레드 실행 유틸. 실패 시 매니저에게 ephemeral 안내 (2026-07-10).

    각 handler 안의 `def _bg(): try: ... except: logger.error(...)` 패턴 대체.
    매니저 관점: 취소·편집 등 명시적 액션 후 응답 없으면 성공/실패 판단 어려움 → 실패 시
    확실한 안내로 재시도 유도.

    Args:
        client: slack client
        body: slack action body (user.id, channel.id 있음)
        action_label: '공사 확정', '공사 취소' 등 사용자에게 노출할 액션 이름
        work_fn: 실제 작업 함수 (인자 없음)
    """
    def _run():
        try:
            work_fn()
        except Exception as exc:
            import uuid as _uuid_e
            error_id = str(_uuid_e.uuid4())[:8]
            logger.error(
                f'[SLACK/BG] {action_label} 실패 (error_id={error_id}): {exc}',
                exc_info=True,
            )
            # 매니저에게 ephemeral 안내
            try:
                user_id = (body.get('user') or {}).get('id', '')
                channel = (body.get('channel') or {}).get('id', '') or \
                          (body.get('container') or {}).get('channel_id', '')
                if user_id and channel:
                    client.chat_postEphemeral(
                        channel=channel, user=user_id,
                        text=(
                            f':x: *{action_label}* 처리 중 오류가 발생했습니다.\n'
                            f'잠시 후 다시 시도해 주세요.\n'
                            f'오류 ID: `{error_id}` (관리자 문의 시 전달)'
                        ),
                    )
            except Exception as notify_exc:
                logger.debug(f'[SLACK/BG] 매니저 알림 실패 (무시): {notify_exc}')
    threading.Thread(target=_run, daemon=True).start()


def _is_slack_retry_duplicate() -> bool:
    """Slack 이 3초 timeout 후 재전송한 요청인지 감지.

    Slack Events API 는 200 응답 못 받으면 3회 재전송. 실제로는 처리 성공했는데
    응답 지연이면 idempotency 를 위해 중복 처리 skip 필요.

    Headers:
    - X-Slack-Retry-Num: 재시도 회수 (1, 2, 3)
    - X-Slack-Retry-Reason: 'http_timeout' 등
    """
    retry_num = request.headers.get('X-Slack-Retry-Num', '')
    retry_reason = request.headers.get('X-Slack-Retry-Reason', '')
    if not retry_num:
        return False
    # 재시도 헤더 있으면 원본 event_id 로 dedup key 만들어 Redis 캐시 확인
    try:
        body = request.get_json(silent=True) or {}
    except Exception:
        body = {}
    event_id = body.get('event_id') or ''
    # slash command / interactive payload 는 event_id 없음 — 처리
    if not event_id:
        # slash command 재시도 케이스는 상황상 매우 드묾. 로그만 남기고 계속 진행
        logger.info(f'[SLACK/RETRY] event_id 없는 재시도 (reason={retry_reason}) — 계속 진행')
        return False
    try:
        from dashboard.utils.redis_client import get_redis_client as _grc
        rc = _grc().redis
        dedup_key = f'slack_event_seen:{event_id}'
        # NX + 1시간 TTL — 처음 보는 event 면 마킹하고 처리 진행
        first_seen = rc.set(dedup_key, retry_num, nx=True, ex=3600)
        if not first_seen:
            logger.warning(
                f'[SLACK/RETRY] 중복 event 무시: event_id={event_id[:12]}… '
                f'retry={retry_num} reason={retry_reason}'
            )
            return True
    except Exception as exc:
        logger.warning(f'[SLACK/RETRY] Redis dedup 실패 (계속 진행): {exc}')
    return False


@slack_bp.route("/events", methods=["POST"])
def slack_events():
    """슬랙 → 우리 서버 webhook (메인 봇: 모든 이벤트/명령/인터랙션 통합 endpoint)"""
    if _slack_handler is None:
        if not _init_slack_app():
            return jsonify({"error": "Slack bot not configured"}), 503

    # Slack retry idempotency (2026-07-10)
    if _is_slack_retry_duplicate():
        return jsonify({"ok": True, "dedup": True}), 200

    return _slack_handler.handle(request)


@slack_bp.route("/project-events", methods=["POST"])
def slack_project_events():
    """슬랙 → 공사 현황 알림 봇 전용 endpoint (/공사확정 슬래시 + 모달)"""
    if _project_slack_handler is None:
        if not _init_project_slack_app():
            return jsonify({"error": "Project Slack bot not configured"}), 503

    if _is_slack_retry_duplicate():
        return jsonify({"ok": True, "dedup": True}), 200

    return _project_slack_handler.handle(request)


@slack_bp.route("/as-events", methods=["POST"])
def slack_as_events():
    """슬랙 → A/S 사후 관리 봇 전용 endpoint (/as 슬래시 + 3단계 모달)"""
    if _as_slack_handler is None:
        if not _init_as_slack_app():
            return jsonify({"error": "AS Slack bot not configured"}), 503

    if _is_slack_retry_duplicate():
        return jsonify({"ok": True, "dedup": True}), 200

    return _as_slack_handler.handle(request)


@slack_bp.route("/visit-events", methods=["POST"])
def slack_visit_events():
    """슬랙 → 방문 일정 알림 봇 전용 endpoint (날짜 수정/취소 액션)"""
    if _visit_slack_handler is None:
        if not _init_visit_slack_app():
            return jsonify({"error": "Visit Slack bot not configured"}), 503

    if _is_slack_retry_duplicate():
        return jsonify({"ok": True, "dedup": True}), 200

    return _visit_slack_handler.handle(request)


@slack_bp.route("/invoice-events", methods=["POST"])
def slack_invoice_events():
    """슬랙 → 세금계산서 관리 알림 봇 전용 endpoint (스레드 첨부 자동 완료)"""
    if _invoice_slack_handler is None:
        if not _init_invoice_slack_app():
            return jsonify({"error": "Invoice Slack bot not configured"}), 503

    if _is_slack_retry_duplicate():
        return jsonify({"ok": True, "dedup": True}), 200

    return _invoice_slack_handler.handle(request)


@slack_bp.route("/list-assignee", methods=["POST"])
def slack_list_assignee():
    """슬랙 List [담당자] 컬럼 변경 → 리드 시트 '영업 담당자' 반영.

    Slack Workflow Builder의 웹훅 액션이 이 URL을 호출.
    페이로드(JSON):
      {
        "lead_no": "L-03116",
        "assignee": "고광일"        # 한국 이름 또는 슬랙 user_id
      }

    보안: Workflow 웹훅 URL 자체가 시크릿 역할. 추가로 SLACK_LIST_WEBHOOK_SECRET
    환경변수 설정 시 X-Auth 헤더로 이중 검증 (선택).
    """
    try:
        # 선택 검증 — .env에 SLACK_LIST_WEBHOOK_SECRET 있으면 헤더 확인
        expected = os.getenv('SLACK_LIST_WEBHOOK_SECRET', '').strip()
        if expected:
            if request.headers.get('X-Auth', '') != expected:
                logger.warning("[SLACK/LIST] list-assignee: 인증 헤더 불일치")
                return jsonify({"ok": False, "error": "unauthorized"}), 401

        data = request.get_json(silent=True) or {}
        lead_no = str(data.get('lead_no') or '').strip()
        assignee_raw = str(data.get('assignee') or '').strip()
        logger.info(f"[SLACK/LIST] 담당자 배정 수신: lead={lead_no}, assignee={assignee_raw!r}")

        if not lead_no:
            return jsonify({"ok": False, "error": "lead_no 누락"}), 400

        # 슬랙 user_id(U01234...) 형식이면 한국 이름으로 변환
        assignee_name = assignee_raw
        if assignee_raw.startswith('U') and len(assignee_raw) <= 15 and \
                assignee_raw[1:].replace('0', '').isalnum():
            try:
                if _slack_app is None:
                    _init_slack_app()
                if _slack_app is not None:
                    resolved = _slack_user_to_korean_name(_slack_app.client, assignee_raw)
                    if resolved:
                        assignee_name = resolved
            except Exception as exc:
                logger.warning(f"[SLACK/LIST] user_id → 이름 변환 실패: {exc}")

        # 반영 — 빈값/'미정'이면 '-' 로 초기화
        if not assignee_name or assignee_name in ('미정', '-'):
            new_value = '-'
        else:
            new_value = assignee_name

        try:
            # 정규 리드(L-XXXXX) 는 시트, ETC-xxx 는 Redis 로 자동 분기
            _update_lead_dispatch(lead_no, {'영업 담당자': new_value})
            logger.info(f"[SLACK/LIST] 담당자 반영: {lead_no} → {new_value!r}")
            return jsonify({"ok": True, "lead_no": lead_no, "assignee": new_value})
        except Exception as exc:
            logger.error(f"[SLACK/LIST] 담당자 반영 실패 ({lead_no}): {exc}", exc_info=True)
            return jsonify({"ok": False, "error": str(exc)}), 500

    except Exception as exc:
        logger.error(f"[SLACK/LIST] list-assignee 처리 오류: {exc}", exc_info=True)
        return jsonify({"ok": False, "error": str(exc)}), 500


@slack_bp.route("/normalize-visit-dates", methods=["GET", "POST"])
def slack_normalize_visit_dates():
    """관리자 트리거 — 리드 시트 방문 예정일 형식 재정규화.

    - 앞의 ' escape prefix 제거 (셀 서식 텍스트라 리터럴로 저장됨)
    - 공백 포함 범위 ('2026-07-15 ~ 2026-07-17') → 표준 축약 ('2026-07-15~17')

    ?dry_run=1 (기본) / ?dry_run=0 실제 실행
    ?etc_only=1 → 기타 리드만 (기본 false, 전체 리드)
    """
    try:
        dry_run = request.args.get('dry_run', '1') != '0'
        etc_only = request.args.get('etc_only', 'false').lower() == 'true'
        result = _normalize_visit_dates(dry_run=dry_run, etc_only=etc_only)
        return jsonify({
            "ok": True, "dry_run": dry_run, "etc_only": etc_only, **result,
        })
    except Exception as exc:
        logger.error(f"[NORMALIZE/방문일] 실패: {exc}", exc_info=True)
        return jsonify({"ok": False, "error": str(exc)}), 500


def _normalize_visit_dates(dry_run: bool = True, etc_only: bool = False) -> dict:
    """리드 시트 E열 (방문 예정일) 재정규화."""
    stats = {'scanned': 0, 'changed': 0, 'skipped': 0, 'errors': 0}
    try:
        from dashboard.services.lead_service import (
            load_leads_data, _get_sheet_config, get_sheets_manager,
            invalidate_leads_cache,
        )
        from dashboard.blueprints.slack_helpers import _format_visit_date_range
        df = load_leads_data(force_refresh=True)
        if df is None or df.empty:
            return {'error': '시트 로드 실패', **stats}
        cfg = _get_sheet_config()
        manager = get_sheets_manager()

        updates = []
        for idx, row in df.iterrows():
            stats['scanned'] += 1
            lead_no = str(row.get('리드 No', '') or '').strip()
            if not lead_no:
                continue
            if etc_only and not lead_no.startswith('ETC-'):
                continue

            raw = str(row.get('방문 예정일', '') or '')
            if not raw or raw == '-':
                continue

            # 정규화
            new_val = raw
            # 1) 앞의 ' 제거
            if new_val.startswith("'"):
                new_val = new_val[1:]
            # 2) 공백 포함 범위 → 표준
            if '~' in new_val:
                parts = [p.strip() for p in new_val.split('~')]
                parts = [p for p in parts if p]
                if len(parts) == 2:
                    new_val = _format_visit_date_range(parts[0], parts[1])
                elif len(parts) == 1:
                    new_val = parts[0]

            if new_val == raw:
                stats['skipped'] += 1
                continue

            sheet_row = int(idx) + 2  # 헤더 1 + 0-based
            updates.append((sheet_row, lead_no, raw, new_val))
            stats['changed'] += 1

        if dry_run or not updates:
            for sr, ln, old, new in updates[:20]:  # 로그 상한
                logger.info(
                    f"[NORMALIZE/방문일/DRY] row {sr} {ln}: {old!r} → {new!r}"
                )
            return stats

        # batchUpdate
        batch = {
            'valueInputOption': 'USER_ENTERED',
            'data': [
                {'range': f"'{cfg['sheet_name']}'!E{sr}", 'values': [[new]]}
                for sr, _, _, new in updates
            ],
        }
        try:
            manager.service.spreadsheets().values().batchUpdate(
                spreadsheetId=cfg['sheet_id'], body=batch,
            ).execute()
            invalidate_leads_cache()
            for sr, ln, old, new in updates:
                logger.info(f"[NORMALIZE/방문일] row {sr} {ln}: {old!r} → {new!r}")
        except Exception as exc:
            logger.error(f"[NORMALIZE/방문일] batchUpdate 실패: {exc}", exc_info=True)
            stats['errors'] += 1
    except Exception as exc:
        logger.error(f"[NORMALIZE/방문일] 스캔 실패: {exc}", exc_info=True)
        stats['errors'] += 1

    return stats


@slack_bp.route("/migrate-etc-to-sheet", methods=["GET", "POST"])
def slack_migrate_etc_to_sheet():
    """관리자 트리거 — 기존 Redis ETC pseudo-lead metadata 를 시트로 이관.

    시나리오 D 전환 후 옛 Redis 저장분 (etc_visit:*) 처리용.
    ?dry_run=1 (기본, 카운트만) or ?dry_run=0 (실제 실행)
    """
    try:
        dry_run = request.args.get('dry_run', '1') != '0'
        result = _migrate_etc_redis_to_sheet(dry_run=dry_run)
        return jsonify({"ok": True, "dry_run": dry_run, **result})
    except Exception as exc:
        logger.error(f"[MIGRATE/ETC] 실패: {exc}", exc_info=True)
        return jsonify({"ok": False, "error": str(exc)}), 500


def _migrate_etc_redis_to_sheet(dry_run: bool = True) -> dict:
    """Redis 의 etc_visit:* hash → 시트 append + Redis 삭제."""
    stats = {'scanned': 0, 'migrated': 0, 'errors': 0}
    try:
        from dashboard.utils.redis_client import get_redis_client
        from dashboard.services.lead_service import (
            _get_sheet_config, get_sheets_manager, LEAD_COLUMN_ORDER,
            invalidate_leads_cache,
        )
        rc = get_redis_client().redis
        keys = list(rc.scan_iter(match='etc_visit:*'))
        stats['scanned'] = len(keys)
        if not keys:
            logger.info("[MIGRATE/ETC] Redis 에 etc_visit:* 없음 — 이관 대상 zero")
            return stats

        cfg = _get_sheet_config()
        if not cfg:
            return {'error': 'ONLINE_LEADS_SHEET_ID 미설정', **stats}
        manager = get_sheets_manager()

        for key in keys:
            try:
                key_str = key.decode() if isinstance(key, bytes) else key
                etc_lead_no = key_str.split(':', 1)[1]

                raw = rc.hgetall(key) or {}
                data = {
                    (k.decode() if isinstance(k, bytes) else k):
                    (v.decode() if isinstance(v, bytes) else v)
                    for k, v in raw.items()
                }

                _vd = data.get('방문 예정일', '')
                row_dict = {
                    '리드 No': etc_lead_no,
                    '상담 시간': data.get('상담 시간', ''),
                    '플랫폼': '기타',
                    '상태': data.get('상태', '방문 예약'),
                    '방문 예정일': (f"'{_vd}" if _vd and '~' not in _vd else _vd),
                    '고객 연락처': data.get('고객 연락처', '-'),
                    '이메일': data.get('이메일', '-'),
                    '고객명': data.get('고객명', '-'),
                    '방문 주소': data.get('방문 주소', '-'),
                    '문의 내용': data.get('문의 내용', '-') or '-',
                    '상담 내용': data.get('상담 내용', ''),
                    '키워드': data.get('키워드', '-'),
                    '온라인 상담자': data.get('온라인 상담자', '-'),
                    '영업 담당자': data.get('영업 담당자', '-'),
                    '마지막 연락일': '-',
                    '폴더 ID': '',
                }
                row = [row_dict.get(col, '') for col in LEAD_COLUMN_ORDER]

                if dry_run:
                    logger.info(f"[MIGRATE/ETC/DRY] would migrate {etc_lead_no}")
                    stats['migrated'] += 1
                    continue

                # manager.append_row 는 '공사 현황' 시트 하드코딩이라
                # 리드 시트에는 못 씀. values().append() 직접 호출.
                manager.service.spreadsheets().values().append(
                    spreadsheetId=cfg['sheet_id'],
                    range=f"'{cfg['sheet_name']}'!A:P",
                    valueInputOption='USER_ENTERED',
                    insertDataOption='INSERT_ROWS',
                    body={'values': [row]},
                ).execute()
                rc.delete(key)
                stats['migrated'] += 1
                logger.info(f"[MIGRATE/ETC] {etc_lead_no} 시트 이관 완료")
            except Exception as exc:
                logger.error(f"[MIGRATE/ETC] {key} 이관 실패: {exc}",
                             exc_info=True)
                stats['errors'] += 1

        if not dry_run and stats['migrated'] > 0:
            invalidate_leads_cache()
    except Exception as exc:
        logger.error(f"[MIGRATE/ETC] 스캔 실패: {exc}", exc_info=True)
        stats['errors'] += 1

    return stats


@slack_bp.route("/migrate-visit-buttons", methods=["GET", "POST"])
def slack_migrate_visit_buttons():
    """관리자 트리거 — #방문_일정 채널의 기존 카드 버튼을
    [✏️ 방문일 수정] → [✏️ 정보 수정] 로 일괄 교체 (chat.update).

    ?dry_run=1 (기본) → 실제 update 안 하고 카운트만
    ?dry_run=0 → 실제 실행
    ?days=30 (기본) → 최근 N일 카드만 스캔
    """
    try:
        dry_run = request.args.get('dry_run', '1') != '0'
        try:
            days = int(request.args.get('days', '30'))
        except ValueError:
            days = 30
        result = _migrate_visit_card_buttons(days=days, dry_run=dry_run)
        return jsonify({"ok": True, "dry_run": dry_run, "days": days, **result})
    except Exception as exc:
        logger.error(f"[MIGRATE/방문버튼] 실패: {exc}", exc_info=True)
        return jsonify({"ok": False, "error": str(exc)}), 500


def _migrate_visit_card_buttons(days: int = 30, dry_run: bool = True) -> dict:
    """#방문_일정 채널의 방문 카드 스캔 → 새 버튼 blocks 로 chat.update."""
    channel = os.getenv('SLACK_VISIT_CHANNEL', '').strip()
    if not channel:
        return {'error': 'SLACK_VISIT_CHANNEL 미설정'}

    # 방문 봇 client — _visit_slack_app 이 이미 있으면 그대로, 없으면 WebClient
    # 직접 생성 (fallback). 진단용 로그 포함.
    client = None
    try:
        global _visit_slack_app
        if _visit_slack_app is None:
            try:
                _init_visit_slack_app()
            except Exception as exc:
                logger.warning(f"[MIGRATE/방문버튼] _init_visit_slack_app 실패: {exc}")
        if _visit_slack_app is not None:
            client = _visit_slack_app.client
    except Exception as exc:
        logger.warning(f"[MIGRATE/방문버튼] _visit_slack_app 접근 실패: {exc}")

    if client is None:
        # Fallback — 봇 토큰으로 WebClient 직접 생성
        bot_token = os.getenv('SLACK_VISIT_BOT_TOKEN', '').strip()
        if not bot_token:
            return {'error': 'SLACK_VISIT_BOT_TOKEN 미설정 + visit app fallback 실패'}
        try:
            from slack_sdk import WebClient
            client = WebClient(token=bot_token)
            logger.info("[MIGRATE/방문버튼] fallback: WebClient 직접 생성")
        except Exception as exc:
            return {'error': f'WebClient 생성 실패: {exc}'}
    logger.info(f"[MIGRATE/방문버튼] 시작: channel={channel} days={days} dry_run={dry_run}")

    stats = {
        'scanned': 0, 'visit_cards': 0, 'already_new': 0,
        'no_actions_skip': 0, 'no_lead_no_skip': 0, 'updated': 0, 'errors': 0,
        'too_old_skip': 0,
    }
    # SDK 의 oldest 파라미터가 서버 환경에서 이상 동작 (msg_count=0) — 회피.
    # 파라미터 없이 최근 메시지 페이지네이션 후 client-side 로 필터.
    oldest_ts_num = time.time() - days * 86400
    cursor = None
    pages = 0
    while pages < 20:  # 안전장치 (최대 20 페이지 = 200*20 = 4000 메시지)
        pages += 1
        kwargs = {'channel': channel, 'limit': 200}
        if cursor:
            kwargs['cursor'] = cursor
        try:
            resp = client.conversations_history(**kwargs)
        except Exception as exc:
            logger.error(f"[MIGRATE/방문버튼] history 실패 (page {pages}): {exc}")
            stats['errors'] += 1
            break

        _msgs = resp.get('messages', []) or []
        logger.info(
            f"[MIGRATE/방문버튼] page {pages}: msg_count={len(_msgs)} "
            f"has_more={resp.get('has_more', False)}"
        )

        # 이 페이지에서 가장 오래된 ts 가 oldest 보다 이전이면 다음 페이지 skip
        _reached_old = False

        for msg in _msgs:
            # client-side 시각 필터
            try:
                _msg_ts = float(msg.get('ts', '0'))
                if _msg_ts < oldest_ts_num:
                    stats['too_old_skip'] += 1
                    _reached_old = True
                    continue
            except (ValueError, TypeError):
                pass
            stats['scanned'] += 1
            blocks = msg.get('blocks') or []
            if not blocks:
                continue

            # 방문 카드 판별 — section 헤더에 '새 방문 일정' 포함
            header_text = ''
            for blk in blocks:
                if blk.get('type') == 'section':
                    bt = (blk.get('text') or {}).get('text', '') or ''
                    if '새 방문 일정' in bt:
                        header_text = bt
                        break
            if not header_text:
                continue
            stats['visit_cards'] += 1

            # actions 없으면 이미 완료/취소된 카드 → 스킵
            actions_blk = next(
                (b for b in blocks if b.get('type') == 'actions'), None,
            )
            if not actions_blk:
                stats['no_actions_skip'] += 1
                continue

            # 이미 새 버튼 (visit_edit_info) 이면 스킵
            elements = actions_blk.get('elements', []) or []
            if any(e.get('action_id') == 'visit_edit_info' for e in elements):
                stats['already_new'] += 1
                continue

            # lead_no 파싱 (헤더 or 본문)
            lead_no = ''
            m = re.search(r'(L-\d{5}|ETC-[a-f0-9]{6})', header_text)
            if m:
                lead_no = m.group(0)
            else:
                for blk in blocks:
                    if blk.get('type') == 'section':
                        bt = (blk.get('text') or {}).get('text', '') or ''
                        m = re.search(r'(L-\d{5}|ETC-[a-f0-9]{6})', bt)
                        if m:
                            lead_no = m.group(0)
                            break
            if not lead_no:
                stats['no_lead_no_skip'] += 1
                continue

            # 최신 lead 값 로드 (없으면 옛 카드 텍스트에서 파싱 fallback)
            lead = _find_lead_by_no(lead_no) or {}

            # 원본 카드에서 필드값 파싱 (fallback 용)
            def _parse_field(pattern):
                mp = re.search(pattern, header_text)
                return mp.group(1).strip() if mp else ''
            orig_visit_date = _parse_field(r'방문일\s*:\s*([^\n>]+)')
            orig_name = _parse_field(r'이름[^:]*:\s*([^\n>]+)')
            orig_contact = _parse_field(r'연락처\s*:\s*([^\n>]+)')
            orig_address = _parse_field(r'방문 주소\s*:\s*([^\n>]+)')
            orig_initial = _parse_field(r'등록자\s*:\s*([^\n>]+)')

            # 상담 내용은 여러 줄 → SEP 사이 텍스트 뽑기
            orig_consultation = ''
            m_con = re.search(
                r'상담 내용\s*:\s*\n((?:>[^\n]*\n)+?)>-{5,}',
                header_text,
            )
            if m_con:
                orig_consultation = '\n'.join(
                    ln.lstrip('>').strip() for ln in m_con.group(1).split('\n') if ln.strip()
                )

            # 필드값: lead 우선, fallback 원본 파싱
            visit_date = (
                str(lead.get('방문 예정일', '') or '').strip().lstrip("'")
                or orig_visit_date
            )
            name = str(lead.get('고객명', '') or '').strip() or orig_name
            contact = str(lead.get('고객 연락처', '') or '').strip() or orig_contact
            address = str(lead.get('방문 주소', '') or '').strip() or orig_address
            consultation = (
                str(lead.get('상담 내용', '') or '').strip() or
                str(lead.get('문의 내용', '') or '').strip() or
                orig_consultation
            )
            if consultation == '-':
                consultation = ''
            initial = orig_initial or '-'

            # category_display
            platform = str(lead.get('플랫폼', '') or '').strip()
            if platform in ('거래처', '기타', '소개'):
                category_display = platform
            else:
                category_display = f"온라인({platform})" if platform else '온라인'

            # 새 blocks 생성
            try:
                new_body_text, new_blocks = _build_visit_notice_blocks(
                    lead_no=lead_no, category_display=category_display,
                    initial=initial, visit_date=visit_date,
                    name=name, contact=contact,
                    visit_address=address, consultation=consultation,
                )
            except Exception as exc:
                logger.error(
                    f"[MIGRATE/방문버튼] blocks 생성 실패 ({lead_no}): {exc}",
                )
                stats['errors'] += 1
                continue

            if dry_run:
                logger.info(
                    f"[MIGRATE/방문버튼/DRY] would update ts={msg.get('ts')} "
                    f"lead={lead_no}"
                )
                stats['updated'] += 1
                continue

            try:
                client.chat_update(
                    channel=channel, ts=msg.get('ts'),
                    text=new_body_text, blocks=new_blocks,
                )
                stats['updated'] += 1
                logger.info(f"[MIGRATE/방문버튼] update: {lead_no}")
                time.sleep(0.15)  # Slack rate limit 여유
            except Exception as exc:
                logger.error(
                    f"[MIGRATE/방문버튼] chat.update 실패 ({lead_no}): {exc}",
                )
                stats['errors'] += 1

        # 이 페이지에서 이미 oldest 이전 메시지가 나왔으면 이후 페이지도 다 오래됨 → 중단
        if _reached_old:
            logger.info(f"[MIGRATE/방문버튼] {days}일 경계 도달 → 중단")
            break
        cursor = (resp.get('response_metadata') or {}).get('next_cursor', '')
        if not cursor:
            break

    return stats


@slack_bp.route("/sync-karrot", methods=["GET", "POST"])
def slack_sync_karrot_now():
    """관리자 트리거 — 당근 시트 즉시 동기화 (테스트/긴급용)"""
    try:
        from dashboard.services.sync_scheduler import trigger_karrot_sync_now
        result = trigger_karrot_sync_now()
        return jsonify({"ok": True, "result": result})
    except Exception as exc:
        logger.error(f"[SLACK] sync-karrot 트리거 실패: {exc}", exc_info=True)
        return jsonify({"ok": False, "error": str(exc)}), 500


@slack_bp.route("/sync-homepage", methods=["GET", "POST"])
def slack_sync_homepage_now():
    """관리자 트리거 — 홈페이지 메일 즉시 동기화"""
    try:
        from dashboard.services.sync_scheduler import trigger_homepage_sync_now
        result = trigger_homepage_sync_now()
        return jsonify({"ok": True, "result": result})
    except Exception as exc:
        logger.error(f"[SLACK] sync-homepage 트리거 실패: {exc}", exc_info=True)
        return jsonify({"ok": False, "error": str(exc)}), 500


@slack_bp.route("/workflow-phone-trigger", methods=["GET", "POST"])
def slack_workflow_phone_trigger():
    """슬랙 워크플로 form → 시트 추가 직후 즉시 호출 — 봇 보정 흐름을 즉시 실행.

    워크플로 빌더 "웹 요청 보내기" step에서 이 URL 호출:
    https://pm.itg-aircon.com/slack/workflow-phone-trigger
    body는 비워도 됨 (전체 시트 보정 폴링이라 행 정보 불필요).
    """
    def _bg():
        try:
            from dashboard.services.lead_sync import sync_workflow_phone_leads
            sync_workflow_phone_leads()
        except Exception as exc:
            logger.error(f"[SLACK] 워크플로 즉시 트리거 실패: {exc}", exc_info=True)
    threading.Thread(target=_bg, daemon=True).start()
    return jsonify({"ok": True}), 200


@slack_bp.route("/health", methods=["GET"])
def slack_health():
    """봇 상태 체크 (브라우저로 열어서 확인용)"""
    return jsonify({
        "enabled": _BOT_ENABLED,
        "token_set": bool(_BOT_TOKEN and not _BOT_TOKEN.startswith('여기에') and 'your' not in _BOT_TOKEN.lower()),
        "signing_secret_set": bool(_SIGNING_SECRET and not _SIGNING_SECRET.startswith('여기에') and 'your' not in _SIGNING_SECRET.lower()),
        "app_initialized": _slack_app is not None,
    })


# ─────────────────────────────────────────────────────────────
# 인입 알림 — [방문 요청] / [가격 문의] 모달 + 제출 처리
# ─────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────
# 기타 방문 pseudo-lead (2026-07-15 시나리오 D)
# ─────────────────────────────────────────────────────────────
# "기타" 방문(사후관리/A/S/수금 등)도 시트에 정상 등록. 리드 No 만 L- 대신
# ETC-xxxxxx (랜덤 hex) — L 번호 소모 방지.
# 시트가 유일 진실 → Redis metadata 저장 없음, 모든 조회·update 는 시트로.
_ETC_LEAD_PREFIX = 'ETC-'


def _is_etc_lead(lead_no: str) -> bool:
    """lead_no 형식 판별 — 대시보드 필터·카드 표시 등에서 사용."""
    return bool(lead_no) and str(lead_no).startswith(_ETC_LEAD_PREFIX)


def _etc_new_lead_no() -> str:
    """랜덤 hex ID. 16^6 = 16M 공간이라 실질 충돌 없음.

    시나리오 D 에서 Redis 중복 체크 제거 — 시트가 유일 진실이지만 시트
    조회는 sync loop 안에서만 하고 여기서는 랜덤만 반환. 충돌 확률
    극히 낮으므로 실무 안전.
    """
    return f"{_ETC_LEAD_PREFIX}{secrets.token_hex(3)}"


def _update_lead_dispatch(lead_no: str, updates: dict) -> None:
    """정규 리드·ETC- 모두 시트 update (시나리오 D)."""
    from dashboard.services.lead_service import update_lead
    update_lead(lead_no, updates)


def _find_lead_by_no(lead_no: str):
    """리드 No 로 메인 시트 행 dict 반환. 정규·ETC- 모두 시트에서 조회."""
    try:
        from dashboard.services.lead_service import load_leads_data
        df = load_leads_data()
        if df is None or df.empty:
            return None
        target = lead_no.strip()
        matches = df[df['리드 No'].astype(str).str.strip() == target]
        if matches.empty:
            return None
        return matches.iloc[0].to_dict()
    except Exception as exc:
        logger.error(f"[SLACK] 리드 조회 실패 ({lead_no}): {exc}")
        return None


# 재상담 이력 append/parse (2026-07-20) — 시트 K열은 여러 회차를 누적 저장.
# 형식: "[MM.DD HH:MM 이니셜 · status] 내용\n─────────\n[…] 내용"
_CONSULT_DIVIDER = '─────────'
_CONSULT_ENTRY_RE = re.compile(
    r'^\[\s*(?P<time>\d{2}\.\d{2}\s+\d{2}:\d{2})\s+'
    r'(?P<ini>\S+)\s*·\s*(?P<status>[^\]]+)\]\s*(?P<content>.*)$',
    re.DOTALL,
)


def _format_consultation_entry(consultation: str, initial: str, status: str) -> str:
    """새 상담 내용을 [시간 이니셜 · status] 헤더 붙여 저장 형식으로."""
    ts = datetime.now().strftime('%m.%d %H:%M')
    ini = (initial or '-').strip() or '-'
    st = (status or '-').strip() or '-'
    body = (consultation or '').strip()
    return f'[{ts} {ini} · {st}] {body}'


def _append_consultation(old: str, new_entry: str) -> str:
    """옛 상담 내용에 새 entry 를 divider 로 이어붙임. 옛 값 없으면 새 값만."""
    old = (old or '').strip()
    if old in ('', '-'):
        return new_entry
    return f'{old}\n{_CONSULT_DIVIDER}\n{new_entry}'


def _parse_consultation_entries(text: str) -> list:
    """저장된 상담 내용 → 회차별 dict 리스트.

    각 회차: {'time':..., 'ini':..., 'status':..., 'content':...}
    옛 형식(헤더 없음) 은 ini/status 빈 값으로 content 만 채워 반환.
    """
    entries = []
    if not text:
        return entries
    for chunk in text.split(_CONSULT_DIVIDER):
        chunk = chunk.strip()
        if not chunk:
            continue
        m = _CONSULT_ENTRY_RE.match(chunk)
        if m:
            entries.append({
                'time': m.group('time').strip(),
                'ini': m.group('ini').strip(),
                'status': m.group('status').strip(),
                'content': m.group('content').strip(),
            })
        else:
            entries.append({'time': '', 'ini': '', 'status': '', 'content': chunk})
    return entries


def _split_lead_content(content_text: str) -> dict:
    """
    상담 내용에 합쳐진 '장소: ... / 기기: ... / 문의: ...' 를 분리.
    실패 시 raw 텍스트만 반환.
    """
    place, device, inquiry = '', '', ''
    if not content_text:
        return {'place': place, 'device': device, 'inquiry': content_text}
    for part in content_text.split(' / '):
        if part.startswith('장소: '):
            place = part[4:].strip()
        elif part.startswith('기기: '):
            device = part[4:].strip()
        elif part.startswith('문의: '):
            inquiry = part[4:].strip()
    if not inquiry:
        inquiry = content_text
    return {'place': place, 'device': device, 'inquiry': inquiry}


def _open_inquiry_modal(client, body, action: str):
    """
    [방문 요청] 또는 [가격 문의] 버튼 클릭 → 모달 팝업

    슬랙 trigger_id는 3초 만료 → 시트 로드(3000+행)가 그 안에 안 끝남.
    해결: 즉시 placeholder 모달 → 데이터 로드 → views_update로 실제 모달로 교체.

    action: 'visit' or 'price'
    """
    lead_no = body["actions"][0]["value"]
    trigger_id = body["trigger_id"]
    channel = body["channel"]["id"]
    message_ts = body["message"]["ts"]
    user_id = body["user"]["id"]

    callback_id = "submit_visit" if action == 'visit' else "submit_price"
    title = "방문 요청" if action == 'visit' else "가격 문의"

    metadata = json.dumps({
        "lead_no": lead_no,
        "channel": channel,
        "message_ts": message_ts,
    }, ensure_ascii=False)

    # 1단계: 즉시 placeholder 모달 띄움 (trigger_id 3초 만료 회피)
    placeholder_view = {
        "type": "modal",
        "callback_id": callback_id,
        "title": {"type": "plain_text", "text": title},
        "close": {"type": "plain_text", "text": "취소"},
        "private_metadata": metadata,
        "blocks": [
            {"type": "section", "text": {"type": "mrkdwn",
                                          "text": f":hourglass_flowing_sand: `{lead_no}` 리드 정보 로딩 중...\n잠시만 기다려주세요."}},
        ],
    }
    try:
        resp = client.views_open(trigger_id=trigger_id, view=placeholder_view)
        view_id = resp["view"]["id"]
    except Exception as exc:
        logger.error(f"[SLACK] placeholder views_open 실패 ({lead_no}): {exc}", exc_info=True)
        return

    # 2단계: 시트 로드 + 실제 모달로 update
    lead = _find_lead_by_no(lead_no)
    if not lead:
        try:
            client.views_update(view_id=view_id, view={
                "type": "modal",
                "callback_id": callback_id,
                "title": {"type": "plain_text", "text": title},
                "close": {"type": "plain_text", "text": "닫기"},
                "blocks": [
                    {"type": "section", "text": {"type": "mrkdwn",
                                                  "text": f":x: `{lead_no}` 리드를 메인 시트에서 찾을 수 없습니다.\n시트가 갱신되었는지 확인하세요."}},
                ],
            })
        except Exception as exc:
            logger.error(f"[SLACK] 에러 모달 update 실패: {exc}", exc_info=True)
        return

    # 상담 내용에서 장소/기기/문의 분리 (옛 Apps Script 형식 지원)
    parts = _split_lead_content(str(lead.get('문의 내용', '') or lead.get('상담 내용', '')))
    name = str(lead.get('고객명') or '').strip()
    phone = str(lead.get('고객 연락처') or '').strip()
    email = str(lead.get('이메일') or '').strip()
    # 장소: split 결과만 있음 (시트 컬럼 없음). 값 없으면 UI 에서 생략.
    place = parts['place'].strip()
    # 기기: 시트 '키워드' 컬럼 우선 (실제 저장 값), fallback split 결과.
    #   _meta_device 는 인메모리 전용이라 시트 재조회 시 사라짐 → 키워드 컬럼이 안전.
    device = str(lead.get('키워드') or '').strip() or parts['device'].strip()
    inquiry = parts['inquiry'] or str(lead.get('문의 내용') or lead.get('상담 내용') or '').strip() or '-'
    # 이전 상담 이력 (재상담 시 참고용) — 값 있을 때만 표시
    prev_consultation = str(lead.get('상담 내용') or '').strip()
    address = str(lead.get('방문 주소') or '').strip()
    consult_time = str(lead.get('상담 시간') or '').strip() or '-'

    # 모달 상단 정보 — 값 있는 필드만 표시 (당근 리드처럼 이메일·장소·기기 등이
    # 없는 케이스에서 '-' 로 노출되는 시각적 노이즈 제거, 2026-07-20).
    def _dash(v): return v if v and v != '-' else ''
    _info_lines = [f"*접수번호:* `{lead_no}`", f"*문의시간 :* {consult_time}"]
    if _dash(name):    _info_lines.append(f"*이름 / 상호 :* {name}")
    if _dash(phone):   _info_lines.append(f"*연락처 :* {phone}")
    if _dash(email):   _info_lines.append(f"*이메일 :* {email}")
    if _dash(place):   _info_lines.append(f"*설치 희망 장소 :* {place}")
    if _dash(device):  _info_lines.append(f"*설치 희망 기기 :* {device}")
    info_blocks = [
        {"type": "section", "text": {"type": "mrkdwn", "text": "\n".join(_info_lines)}},
        # 문의 내용은 3000자 제한 대응 — 넘치면 자동 truncate + 안내
        {"type": "section", "text": {"type": "mrkdwn",
                                      "text": slack_truncate(f"*문의 내용 :*\n{inquiry}")}},
    ]
    # 이전 상담 내용 (재상담 시 참고) — 값 있을 때만 별도 섹션
    if prev_consultation and prev_consultation != '-':
        info_blocks.append({
            "type": "section", "text": {"type": "mrkdwn",
                "text": slack_truncate(f"*상담 내용 :*\n{prev_consultation}")},
        })
    info_blocks.append({"type": "divider"})

    # 입력 블록 — action에 따라 다름 (callback_id, title은 1단계에서 정의됨)
    today_iso = date.today().isoformat()
    if action == 'visit':
        input_blocks = [
            {
                "type": "input",
                "block_id": "visit_date",
                "label": {"type": "plain_text", "text": "방문일"},
                "element": {
                    "type": "datepicker",
                    "action_id": "value",
                    "initial_date": today_iso,
                },
            },
            {
                "type": "input",
                "block_id": "visit_address",
                "label": {"type": "plain_text", "text": "방문 주소"},
                "element": {
                    "type": "plain_text_input",
                    "action_id": "value",
                    "initial_value": address[:150] if address else "",
                },
            },
            {
                "type": "input",
                "block_id": "consultation",
                "label": {"type": "plain_text", "text": "상담 내용 / 특이사항"},
                "element": {
                    "type": "plain_text_input",
                    "action_id": "value",
                    "multiline": True,
                    "placeholder": {"type": "plain_text",
                                    "text": "상담 내용, 특이사항을 자유롭게 입력하세요"},
                },
                "optional": True,
            },
        ]
    else:  # price
        input_blocks = [
            {
                "type": "input",
                "block_id": "estimate",
                "label": {"type": "plain_text", "text": "가견적 요청 (O/X)"},
                "element": {
                    "type": "static_select",
                    "action_id": "value",
                    "placeholder": {"type": "plain_text", "text": "선택하세요"},
                    "options": [
                        {"text": {"type": "plain_text", "text": "O (요청 보냄)"},
                         "value": "yes"},
                        {"text": {"type": "plain_text", "text": "X (요청 안 보냄)"},
                         "value": "no"},
                    ],
                },
            },
            {
                "type": "input",
                "block_id": "consultation",
                "label": {"type": "plain_text", "text": "상담 내용"},
                "element": {
                    "type": "plain_text_input",
                    "action_id": "value",
                    "multiline": True,
                    "placeholder": {"type": "plain_text",
                                    "text": "고객과 나눈 상담 내용을 입력하세요"},
                },
            },
        ]

    full_view = {
        "type": "modal",
        "callback_id": callback_id,
        "title": {"type": "plain_text", "text": title},
        "submit": {"type": "plain_text", "text": "등록"},
        "close": {"type": "plain_text", "text": "취소"},
        "private_metadata": metadata,
        "blocks": info_blocks + input_blocks,
    }
    # 3단계: placeholder를 실제 모달로 교체
    try:
        client.views_update(view_id=view_id, view=full_view)
    except Exception as exc:
        logger.error(f"[SLACK] 모달 views_update 실패 ({lead_no}): {exc}", exc_info=True)


# ─────────────────────────────────────────────────────────────
# 통합 상담 모달 — 인입 카드 [📋 상담하기] / /방문 슬래시 공통 진입점
# 두 차원으로 분류: 방문 유형(어디서 받음) + 처리 유형(결과)
# ─────────────────────────────────────────────────────────────
_CONSULT_VISIT_TYPE_OPTIONS = [
    # 방문 유형 (시트 플랫폼 컬럼) — label 깔끔하게
    ('온라인', '온라인'),
    ('거래처', '거래처'),
    ('기타', '기타'),
]
_CONSULT_STATUS_OPTIONS = [
    # 처리 유형 (시트 상태 컬럼) — 순서: 유선 상담 → 방문 예약 → 견적 제출 → 문의 드랍 → 부재중
    # label/value 모두 "방문 예약"으로 통일 (시트 상태값과 일치)
    ('유선 상담', '유선 상담'),
    ('방문 예약', '방문 예약'),
    ('견적 제출', '견적 제출'),
    ('문의 드랍', '문의 드랍'),
    ('부재중', '부재중'),
]


def _search_leads_for_options(query: str, limit: int = 20) -> list:
    """external_select용 lead 검색 — 이름/연락처/lead_no/주소 매칭.
    각 옵션 라벨: "L-XXXXX | 이름 | 연락처 | 플랫폼" — 매니저가 식별 가능하게.
    """
    try:
        from dashboard.services.lead_service import get_lead_records
    except Exception:
        return []
    leads = get_lead_records() or []
    q = query.strip().lower()
    q_digits = re.sub(r'\D', '', q)
    # 채팅 lead 제외 — 같은 채팅방 재인입 시 스레드 유지되므로 다른 lead 연결 불필요
    _CHAT_PLATFORMS = {'카카오톡', '채널톡'}

    # 빈 검색 — 최근 lead N건 반환 (lead_no 내림차순)
    if not q:
        recent = []
        for lead in leads:
            lead_no = str(lead.get('리드 No') or '').strip()
            if not lead_no.startswith('L-'):
                continue
            platform = str(lead.get('플랫폼') or '').strip() or '-'
            if platform in _CHAT_PLATFORMS:
                continue  # 채팅 lead 제외
            name = str(lead.get('고객명') or '').strip()
            phone = str(lead.get('고객 연락처') or '').strip()
            label = f"{lead_no} | {name or '-'} | {phone or '-'} | {platform}"
            if len(label) > 75:
                label = label[:72] + '...'
            try:
                sort_key = int(lead_no.split('-')[1])
            except Exception:
                sort_key = 0
            recent.append((sort_key, lead_no, label))
        recent.sort(reverse=True)  # 최신순 (큰 lead_no 먼저)
        return [
            {"text": {"type": "plain_text", "text": label}, "value": lead_no}
            for _, lead_no, label in recent[:limit]
        ]

    matched = []
    for lead in leads:
        lead_no = str(lead.get('리드 No') or '').strip()
        if not lead_no:
            continue
        platform = str(lead.get('플랫폼') or '').strip() or '-'
        if platform in _CHAT_PLATFORMS:
            continue  # 채팅 lead 제외
        name = str(lead.get('고객명') or '').strip()
        phone = str(lead.get('고객 연락처') or '').strip()
        address = str(lead.get('방문 주소') or '').strip()
        phone_digits = re.sub(r'\D', '', phone)
        # 매칭 점수 — 정확 일치 우선
        score = 0
        if q and q.upper() in lead_no.upper():
            score = 100  # lead_no 정확 매칭
        elif q and q in name.lower():
            score = 90
        elif q_digits and q_digits in phone_digits:
            score = 80
        elif q and q in address.lower():
            score = 50
        if score > 0:
            label = f"{lead_no} | {name or '-'} | {phone or '-'} | {platform}"
            if len(label) > 75:
                label = label[:72] + '...'
            try:
                sort_key = int(lead_no.split('-')[1])
            except Exception:
                sort_key = 0
            matched.append((score, sort_key, lead_no, label))
    # 점수 내림차순 + lead_no 내림차순 (최신순)
    matched.sort(key=lambda x: (-x[0], -x[1]))
    return [
        {"text": {"type": "plain_text", "text": label}, "value": lead_no}
        for _, _, lead_no, label in matched[:limit]
    ]


def _open_link_lead_modal(client, body, chat_id: str, channel: str, message_ts: str):
    """채널톡 카드의 [🔗 기존 lead 연결] 모달 — 같은 사람이 다른 채널로도 인입한 경우.
    매니저가 lead_no를 입력하면 채팅 정보를 기존 lead에 통합 (피드백 컬럼에 메모 추가).
    """
    trigger_id = body["trigger_id"]
    # Redis pending lead에서 채팅 정보 가져와 모달에 표시
    chat_info_text = ''
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        pending_raw = rc.get(f'channeltalk_pending_lead:{chat_id}')
        if pending_raw:
            pending = json.loads(
                pending_raw.decode('utf-8') if isinstance(pending_raw, bytes) else pending_raw
            )
            chat_info_text = (
                f"*카톡 채팅 정보*\n"
                f"• 닉네임: `{pending.get('user_name', '-')}`\n"
                f"• 첫 메시지: {pending.get('first_message', '-')[:80]}"
            )
    except Exception:
        pass

    metadata = json.dumps({
        "chat_id": chat_id, "channel": channel, "message_ts": message_ts,
    }, ensure_ascii=False)

    blocks = []
    if chat_info_text:
        blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": chat_info_text}})
        blocks.append({"type": "divider"})
    blocks.append({
        "type": "input",
        "block_id": "target_lead_no",
        "label": {"type": "plain_text", "text": "통합할 기존 Lead 선택"},
        "element": {
            "type": "external_select",
            "action_id": "link_lead_search",
            "placeholder": {"type": "plain_text", "text": "클릭하면 최근 lead 표시 / 검색도 가능"},
            "min_query_length": 0,
        },
        "hint": {"type": "plain_text",
                 "text": "기본: 최근 30건 / 검색: 이름·연락처·Lead No 입력"},
    })

    view = {
        "type": "modal",
        "callback_id": "submit_link_lead",
        "title": {"type": "plain_text", "text": "기존 Lead에 연결"},
        "submit": {"type": "plain_text", "text": "연결"},
        "close": {"type": "plain_text", "text": "취소"},
        "private_metadata": metadata,
        "blocks": blocks,
    }
    try:
        client.views_open(trigger_id=trigger_id, view=view)
    except Exception as exc:
        logger.error(f"[SLACK/link] 모달 열기 실패: {exc}", exc_info=True)


def _link_chat_to_existing_lead(client, chat_id: str, target_lead_no: str,
                                 channel: str, message_ts: str,
                                 slack_user_id: str = '') -> None:
    """채널톡 채팅을 기존 lead에 통합.
    - 채팅 lead(chat_lead_no) 시트 업데이트:
      · 상태='문의 드랍'
      · 상담 내용에 `→ {target_lead_no} 로 통합` 마킹
      · 키워드/온라인 상담자는 target lead 값 복사 (통계 일관성)
    - target lead 상담 내용은 건드리지 않음
      (매니저가 이후 상담 결과 입력 시 덮어써지므로 마킹은 의미 없음)
    - Redis pending lead 삭제 (있는 경우)
    - `linked_chat:{chat_id}` 마커 저장 (30일) — 재시도 방어
    - 슬랙 thread 안내 + 원본 카드 ✅ reaction
    """
    try:
        from dashboard.utils.redis_client import get_redis_client
        from dashboard.services.lead_service import update_lead, get_lead_by_no
        from dashboard.blueprints.channeltalk import _get_chat_lead_no
        from datetime import datetime
        rc = get_redis_client().redis
        linked_key = f'linked_chat:{chat_id}'

        # === 재시도 방어 ===
        existing_linked = rc.get(linked_key)
        if isinstance(existing_linked, bytes):
            existing_linked = existing_linked.decode('utf-8')
        if existing_linked:
            existing_linked = str(existing_linked).strip()
            if existing_linked == target_lead_no:
                msg = f":information_source: 이 채팅은 이미 `{target_lead_no}` 로 통합돼있어요. 재처리는 스킵합니다."
            else:
                msg = (
                    f":warning: 이 채팅은 이미 `{existing_linked}` 로 통합돼있어요. "
                    f"다른 lead로 재통합하려면 관리자 문의 필요."
                )
            if channel and slack_user_id:
                try:
                    client.chat_postEphemeral(channel=channel, user=slack_user_id, text=msg)
                except Exception:
                    pass
            logger.info(f"[SLACK/link] 재시도 skip — chat_id={chat_id} 이미 {existing_linked} 로 통합됨")
            return

        # === target lead 정보 조회 ===
        target_lead = get_lead_by_no(target_lead_no) or {}

        # === chat lead 정보 조회 (chat_id → chat_lead_no) ===
        chat_lead_no = _get_chat_lead_no(chat_id)
        chat_lead = get_lead_by_no(chat_lead_no) if chat_lead_no else {}
        chat_lead = chat_lead or {}

        # === pending 데이터 삭제 (있으면) — 재감지 방지 ===
        pending_key = f'channeltalk_pending_lead:{chat_id}'

        # === 채팅 lead 자체 시트 업데이트 ===
        # 상태='문의 드랍', 상담 내용에 통합 마킹, 키워드/온라인 상담자 target 값 복사
        if chat_lead_no and chat_lead_no != target_lead_no:
            try:
                chat_old_feedback = (chat_lead.get('상담 내용') or chat_lead.get('피드백') or '').strip()
                chat_new_feedback = (
                    (chat_old_feedback + '\n\n' if chat_old_feedback else '')
                    + f'→ {target_lead_no} 로 통합'
                )
                chat_update: dict = {
                    '상태': '문의 드랍',
                    '상담 내용': chat_new_feedback,
                }
                # 키워드/온라인 상담자는 target lead 값 있을 때만 복사 (공백 덮어쓰기 방지)
                target_kw = (target_lead.get('키워드') or '').strip()
                target_op = (target_lead.get('온라인 상담자') or '').strip()
                if target_kw and target_kw != '-':
                    chat_update['키워드'] = target_kw
                if target_op:
                    chat_update['온라인 상담자'] = target_op
                update_lead(chat_lead_no, chat_update)
                logger.info(
                    f"[SLACK/link] 채팅 lead 통합 마킹: {chat_lead_no} → {target_lead_no} "
                    f"(키워드/상담자 복사={bool(target_kw)}/{bool(target_op)})"
                )
            except Exception as exc:
                logger.warning(f"[SLACK/link] 채팅 lead 통합 마킹 실패: {exc}")

        # === Redis pending 삭제 (있으면) ===
        rc.delete(pending_key)

        # === linked_chat 마커 저장 (30일) — 재시도 방어 ===
        try:
            rc.set(linked_key, target_lead_no, ex=60 * 60 * 24 * 30)
        except Exception as exc:
            logger.debug(f"[SLACK/link] linked_chat 마커 저장 실패: {exc}")

        # === 슬랙 thread 안내 ===
        if channel and message_ts:
            try:
                thread_msg = f":link: 기존 lead `{target_lead_no}` 에 통합 완료."
                if chat_lead_no:
                    thread_msg += f" 채팅 리드 `{chat_lead_no}` 는 '문의 드랍' 처리."
                client.chat_postMessage(
                    channel=channel, thread_ts=message_ts,
                    text=thread_msg,
                )
            except Exception:
                pass

        # === 원본 lead 카드에 ✅ reaction (시각적 처리 완료 표시) ===
        try:
            card_info = rc.get(f'lead_card_msg:{target_lead_no}')
            if card_info:
                card_info_s = card_info.decode('utf-8') if isinstance(card_info, bytes) else card_info
                if '|' in card_info_s:
                    target_channel, target_ts = card_info_s.split('|', 1)
                    try:
                        client.reactions_add(
                            channel=target_channel, timestamp=target_ts,
                            name='white_check_mark',
                        )
                    except Exception as exc:
                        logger.debug(f"[SLACK/link] reaction 추가 skip ({target_lead_no}): {exc}")
        except Exception as exc:
            logger.warning(f"[SLACK/link] 원본 카드 reaction 실패: {exc}")

        logger.info(
            f"[SLACK/link] chat_id={chat_id} → {target_lead_no} 통합 완료 "
            f"(chat_lead={chat_lead_no or '없음'})"
        )
    except Exception as exc:
        logger.error(f"[SLACK/link] 통합 처리 실패: {exc}", exc_info=True)


def _open_consult_modal(client, body, from_slash: bool = False):
    """통합 상담 모달 — 인입 카드 [📋 상담하기] 또는 /방문 슬래시에서 호출.

    인입 카드: lead_no 자동 채움 + 카테고리=방문 예약 prefill
    슬래시: lead_no 없음, 카테고리 자유 선택 (거래처/기타 방문 등록용)
    """
    trigger_id = body["trigger_id"]
    user_id = body["user"]["id"]
    lead_no = ''
    chat_id = ''  # 채널톡 카드 button value가 chat_id인 경우
    channel = ''
    message_ts = ''
    original_text = ''  # 모달 제출 후 카드 chat.update 시 옛 본문 보존용

    if from_slash:
        channel = body.get("channel_id", "")
    else:
        btn_value = body["actions"][0]["value"]
        # value가 L-XXXXX 양식이면 lead_no, 아니면 chat_id (채널톡 B 옵션 카드)
        if re.match(r"^L-\d{5}$", btn_value):
            lead_no = btn_value
        elif btn_value and btn_value != '-':
            chat_id = btn_value
        channel = body["channel"]["id"]
        message_ts = body["message"]["ts"]
        # 옛 카드 본문 — 모달 제출 후 회색 박스 변환에 사용
        # message.text는 짧은 fallback이라 실제 카드는 blocks에서 추출
        msg = body.get("message", {})
        block_texts = []
        for b in msg.get("blocks", []):
            if b.get("type") == "section":
                t = (b.get("text") or {}).get("text", "")
                if t:
                    block_texts.append(t)
        original_text = '\n'.join(block_texts) if block_texts else (msg.get("text", "") or '')
        # 재상담 accumulate 방어 (2026-07-21 L-03307 3차 상담 사고):
        # 이미 chat.update 된 카드는 blocks[section] = "헤더 + ```원본```" 형태.
        # 이걸 그대로 다음 회색 처리에 넘기면 ```원본``` 안에 (이전헤더+더이전원본)
        # 이 계속 accumulate. 코드 블록 있으면 그 안만 진짜 원본으로 추출.
        if original_text:
            _m_code = re.search(r'```\s*\n(.*?)\n\s*```', original_text, re.DOTALL)
            if _m_code:
                original_text = _m_code.group(1).strip()

    metadata = json.dumps({
        "lead_no": lead_no,
        "chat_id": chat_id,
        "channel": channel,
        "message_ts": message_ts,
        "original_text": original_text,
    }, ensure_ascii=False)

    # 2026-07-12 mobile 대응 — placeholder + views_update 조합이 슬랙 mobile 앱
    #   에서 반영 안 되는 이슈. 처음부터 full view 로 views_open. Lead 조회는 캐시
    #   사용 (force_refresh=False) 로 빠르게. trigger_id 3초 유효 시간 안에 완료.
    # lead_no 있으면 시트 조회 (인입 케이스 prefill) — 캐시 우선
    lead = _find_lead_by_no(lead_no) if lead_no else None

    # 자동 매칭 — lead_no 못 찾으면 슬랙 카드 메시지에서 이메일/연락처 파싱 후 매칭
    # (매니저가 시트 정리 시 lead_no 변경한 경우 — 슬랙 카드의 옛 lead_no가 stale)
    matched_lead_no = ''
    if lead_no and lead is None:
        try:
            card_text = (body.get("message") or {}).get("text", "") or ''
            # 이메일 / 연락처 추출
            email_m = re.search(r'[\w.+-]+@[\w-]+\.[\w.-]+', card_text)
            phone_m = re.search(r'\b(0\d{1,2}[- ]?\d{3,4}[- ]?\d{4})\b', card_text)
            email = email_m.group(0).strip().lower() if email_m else ''
            phone_digits = re.sub(r'\D', '', phone_m.group(1)) if phone_m else ''
            # 시트에서 매칭 — 캐시 사용 (mobile 대응 위해 force_refresh 제거)
            from dashboard.services.lead_service import load_leads_data
            df = load_leads_data()
            if df is not None and not df.empty:
                if email:
                    em_norm = df['이메일'].astype(str).str.strip().str.lower()
                    matches = df[em_norm == email]
                    if not matches.empty:
                        matched_lead_no = str(matches.iloc[0].get('리드 No') or '').strip()
                if not matched_lead_no and phone_digits:
                    ph_norm = df['고객 연락처'].astype(str).str.replace(r'\D', '', regex=True)
                    matches = df[ph_norm == phone_digits]
                    if not matches.empty:
                        matched_lead_no = str(matches.iloc[0].get('리드 No') or '').strip()
            if matched_lead_no:
                logger.info(
                    f"[SLACK/상담] {lead_no} 시트 없음 → "
                    f"이메일/연락처로 자동 매칭: {matched_lead_no}"
                )
                old_lead_no = lead_no
                lead = _find_lead_by_no(matched_lead_no)
                lead_no = matched_lead_no  # 모달 metadata도 업데이트
                # metadata 재구성 (original_text 유지)
                metadata = json.dumps({
                    "lead_no": lead_no, "chat_id": chat_id,
                    "channel": channel, "message_ts": message_ts,
                    "original_text": original_text,
                }, ensure_ascii=False)
        except Exception as exc:
            logger.warning(f"[SLACK/상담] 자동 매칭 실패: {exc}")
            old_lead_no = ''
    else:
        old_lead_no = ''

    info_blocks = _build_consult_info_blocks(lead, lead_no)
    # 자동 매칭 됐으면 상단에 안내 추가
    if old_lead_no and matched_lead_no:
        info_blocks.insert(0, {
            "type": "section",
            "text": {"type": "mrkdwn",
                     "text": f":arrows_counterclockwise: *자동 매칭됨* — `{old_lead_no}` "
                             f"시트에 없어 이메일/연락처로 `{matched_lead_no}` 매칭"},
        })

    # 채널톡 카드 케이스 — chat_id 있으면 Redis pending lead 정보로 prefill
    channeltalk_info = None
    if chat_id:
        try:
            from dashboard.utils.redis_client import get_redis_client
            rc = get_redis_client().redis
            pending_raw = rc.get(f'channeltalk_pending_lead:{chat_id}')
            if pending_raw:
                channeltalk_info = json.loads(
                    pending_raw.decode('utf-8') if isinstance(pending_raw, bytes) else pending_raw
                )
        except Exception as exc:
            logger.warning(f"[SLACK/상담] 채널톡 정보 조회 실패: {exc}")

    # 초기 prefill — 인입은 온라인, 슬래시는 거래처.
    # 재제출(시트에 상태/방문 예정일 이미 있음) → 시트 값으로 prefill
    # 첫 상담(시트 빈 상태) → 처리 유형은 '유선 상담' 기본, 방문 예정일은 빈값
    default_visit_type = '온라인' if (lead_no or chat_id) else '거래처'
    sheet_status = (str(lead.get('상태') or '').strip() if lead else '')
    sheet_visit_date_raw = (str(lead.get('방문 예정일') or '').strip() if lead else '')
    # 시트 escape prefix(') 제거 + ISO 양식만 허용 (datepicker initial_date 검증)
    if sheet_visit_date_raw.startswith("'"):
        sheet_visit_date_raw = sheet_visit_date_raw[1:]
    sheet_visit_date = sheet_visit_date_raw if re.fullmatch(r'\d{4}-\d{2}-\d{2}', sheet_visit_date_raw) else ''
    prefilled = {
        'visit_type': default_visit_type,
        'status': sheet_status if sheet_status else '유선 상담',
        'visit_date': sheet_visit_date,
        'name': (
            (str(lead.get('고객명') or '').strip() if lead else '')
            or (channeltalk_info.get('user_name', '') if channeltalk_info else '')
        ),
        'contact': (str(lead.get('고객 연락처') or '').strip() if lead else ''),
        'email': (str(lead.get('이메일') or '').strip() if lead else ''),
        'visit_address': (str(lead.get('방문 주소') or '').strip() if lead else ''),
        # 옛 상담 내용은 카드에 이미 표시 — 모달은 통화 후 추가 메모만 받음 (피드백 컬럼에 저장)
        'consultation': '',
    }
    full_view = _build_consult_view(info_blocks, metadata, prefilled)
    # 재상담 여부 판단 — 시트 상태가 이미 처리된 값 (유선 상담/방문 예약/견적 제출/
    # 문의 드랍/부재중) 이면 재상담. '인입' 이거나 빈 값은 첫 상담. (2026-07-20)
    _processed_statuses = {'유선 상담', '방문 예약', '견적 제출', '문의 드랍', '부재중', '방문 취소'}
    _modal_title = '재상담 처리' if sheet_status in _processed_statuses else '상담 처리'
    # full_view 의 title 을 재상담 여부에 맞게 덮어씀
    full_view['title'] = {'type': 'plain_text', 'text': _modal_title}
    # 2026-07-12 datepicker 표시 원인 확인 위한 임시 revert — 이전 placeholder +
    #   views_update 방식으로 되돌림. mobile 표시 vs datepicker 로케일 트레이드오프.
    placeholder = {
        "type": "modal",
        "callback_id": "submit_consult",
        "title": {"type": "plain_text", "text": _modal_title},
        "close": {"type": "plain_text", "text": "취소"},
        "private_metadata": metadata,
        "blocks": [{
            "type": "section",
            "text": {"type": "mrkdwn",
                     "text": ":hourglass_flowing_sand: 모달 준비 중..."},
        }],
    }
    try:
        resp = client.views_open(trigger_id=trigger_id, view=placeholder)
        view_id = resp["view"]["id"]
        client.views_update(view_id=view_id, view=full_view)
    except Exception as exc:
        logger.error(f"[SLACK/상담] 모달 open 실패: {exc}", exc_info=True)


def _build_consult_info_blocks(lead: dict | None, lead_no: str) -> list:
    """상담 모달 상단 인입 정보 블록 — lead 있으면 카드형 정보, 없고 lead_no만 있으면 경고.

    2026-07-20: 값 없는 필드(이메일 등) 는 UI 노이즈 제거 위해 생략. 재상담 케이스
    대응으로 이전 상담 내용(K열) 값 있으면 별도 섹션 추가.
    """
    if lead:
        parts = _split_lead_content(str(lead.get('문의 내용', '') or lead.get('상담 내용', '')))
        name = str(lead.get('고객명') or '').strip()
        phone = str(lead.get('고객 연락처') or '').strip()
        email = str(lead.get('이메일') or '').strip()
        consult_time = str(lead.get('상담 시간') or '').strip() or '-'
        inquiry = parts.get('inquiry') or str(lead.get('문의 내용') or lead.get('상담 내용') or '').strip() or '-'
        prev_consultation = str(lead.get('상담 내용') or '').strip()
        # 설치 희망 기기: 시트 '키워드' L열 우선 (재조회에도 유지), fallback split 결과
        device = str(lead.get('키워드') or '').strip() or parts.get('device', '').strip()

        def _dash(v): return v if v and v != '-' else ''
        info_lines = [f"*접수번호:* `{lead_no}`", f"*문의시간:* {consult_time}"]
        if _dash(name):   info_lines.append(f"*이름 / 상호:* {name}")
        if _dash(phone):  info_lines.append(f"*연락처:* {phone}")
        if _dash(email):  info_lines.append(f"*이메일:* {email}")
        if _dash(device): info_lines.append(f"*설치 희망 기기:* {device}")
        info_lines.append(f"*문의 내용:* {inquiry[:300]}")

        blocks = [
            {"type": "section", "text": {"type": "mrkdwn", "text": "\n".join(info_lines)}},
        ]
        # 재상담 시 이전 상담 내용 참고용 — 값 있을 때만 별도 섹션
        if _dash(prev_consultation) and prev_consultation != inquiry:
            blocks.append({
                "type": "section", "text": {"type": "mrkdwn",
                    "text": slack_truncate(f"*상담 내용:*\n{prev_consultation}")},
            })
        blocks.append({"type": "divider"})
        return blocks
    if lead_no:
        return [
            {"type": "section", "text": {"type": "mrkdwn",
                                          "text": f":warning: `{lead_no}` 리드를 시트에서 찾지 못했습니다."}},
            {"type": "divider"},
        ]
    return []


def _build_consult_view(info_blocks: list, metadata: str, prefilled: dict) -> dict:
    """상담 모달 view 빌더 — prefilled에 따라 입력 블록 구성.

    처리 유형이 '방문 예약'일 때만 visit_date 블록을 포함 (활성화).
    """
    visit_type = prefilled.get('visit_type') or '온라인'
    status = prefilled.get('status') or '유선 상담'
    is_visit = (status == '방문 예약')

    initial_visit_type = next(
        ({"text": {"type": "plain_text", "text": label}, "value": v}
         for v, label in _CONSULT_VISIT_TYPE_OPTIONS if v == visit_type),
        None,
    )
    visit_type_element = {
        "type": "static_select", "action_id": "value",
        "placeholder": {"type": "plain_text", "text": "방문 유형 선택"},
        "options": [
            {"text": {"type": "plain_text", "text": label}, "value": v}
            for v, label in _CONSULT_VISIT_TYPE_OPTIONS
        ],
    }
    if initial_visit_type:
        visit_type_element["initial_option"] = initial_visit_type

    initial_status = next(
        ({"text": {"type": "plain_text", "text": label}, "value": v}
         for v, label in _CONSULT_STATUS_OPTIONS if v == status),
        None,
    )
    status_element = {
        "type": "static_select", "action_id": "value",
        "placeholder": {"type": "plain_text", "text": "상담 유형 선택"},
        "options": [
            {"text": {"type": "plain_text", "text": label}, "value": v}
            for v, label in _CONSULT_STATUS_OPTIONS
        ],
    }
    if initial_status:
        status_element["initial_option"] = initial_status

    def _text_input(block_id, label, optional=True, multiline=False, placeholder=None):
        # 방문 예약 시 name/contact/visit_address 도 필수 처리.
        # 옵션 표시는 슬랙이 optional=True 시 라벨 옆에 회색 '(옵션)' 자동 추가 →
        # 라벨 문자열에 별도 표기하지 않음 (중복 방지).
        force_required = is_visit and block_id in ('name', 'contact', 'visit_address')
        effective_optional = optional and not force_required
        label_text = label
        elem = {"type": "plain_text_input", "action_id": "value"}
        if multiline:
            elem["multiline"] = True
        if placeholder:
            elem["placeholder"] = {"type": "plain_text", "text": placeholder}
        val = (prefilled.get(block_id) or '').strip()
        if val and val != '-':
            elem["initial_value"] = val[:300]
        return {
            "type": "input", "block_id": block_id, "optional": effective_optional,
            "label": {"type": "plain_text", "text": label_text},
            "element": elem,
        }

    # 2026-07-12 datepicker placeholder 명시 — 로케일 렌더링 안정화
    vd_element = {
        "type": "datepicker",
        "action_id": "value",
        "placeholder": {"type": "plain_text", "text": "날짜 선택"},
    }
    vd_initial = (prefilled.get('visit_date') or '').strip()
    if vd_initial:
        vd_element["initial_date"] = vd_initial

    # 인입 카드(lead_no 또는 chat_id 있음) 진입은 자동 분류 → 방문 유형 필드 숨김
    # /방문 슬래시 진입(둘 다 없음)에서만 방문 유형 dropdown 표시
    try:
        _meta = json.loads(metadata) if metadata else {}
    except Exception:
        _meta = {}
    is_lead_card = bool(_meta.get('lead_no') or _meta.get('chat_id'))

    # 처리 유형 변경 시 모달 자체를 다시 그려서 필수/옵션 라벨을 동기화
    status_input_block = {
        "type": "input", "block_id": "status",
        "dispatch_action": True,
        "label": {"type": "plain_text", "text": "상담 유형"},
        "element": status_element,
    }

    # visit_date — 방문 예약일 때만 표시 (상담 유형 바꾸면 dispatch_action 으로 재렌더링)
    vd_block = {
        "type": "input", "block_id": "visit_date",
        "label": {"type": "plain_text", "text": "방문 예정일 (시작)"},
        "element": vd_element,
    }

    input_blocks = []
    if not is_lead_card:
        input_blocks.append({
            "type": "input", "block_id": "visit_type",
            "label": {"type": "plain_text", "text": "방문 유형"},
            "element": visit_type_element,
        })
    input_blocks.append(status_input_block)
    if is_visit:
        input_blocks.extend([
            vd_block,
            {
                "type": "input", "block_id": "visit_date_end", "optional": True,
                "label": {"type": "plain_text", "text": "방문 예정일 (종료)"},
                "hint": {"type": "plain_text",
                         "text": "방문 일자가 범위 일 때만 입력. (예: 7/1~7/3)"},
                "element": {
                    "type": "datepicker",
                    "action_id": "value",
                    "placeholder": {"type": "plain_text", "text": "날짜 선택"},
                },
            },
            # 본인 방문 필수 (2026-07-17) — JW 가 담당자 배정 시 참고
            {
                "type": "input", "block_id": "assign_self",
                "label": {"type": "plain_text", "text": "본인 방문 필수"},
                "hint": {"type": "plain_text",
                         "text": "등록자 본인이 꼭 가야 하는 현장이면 '예'. JW 담당자 배정 참고용."},
                "element": {
                    "type": "radio_buttons",
                    "action_id": "value",
                    "initial_option": {
                        "text": {"type": "plain_text", "text": "아니오"},
                        "value": "no",
                    },
                    "options": [
                        {"text": {"type": "plain_text", "text": "아니오"}, "value": "no"},
                        {"text": {"type": "plain_text", "text": "예 (본인 방문)"}, "value": "yes"},
                    ],
                },
            },
        ])

    input_blocks.extend([
        _text_input("name", "이름 / 상호"),
        _text_input("contact", "연락처", placeholder="010-1234-5678"),
        _text_input("email", "이메일", placeholder="example@domain.com"),
        _text_input("visit_address", "방문 주소", multiline=True),
        _text_input("consultation", "상담 내역",
                    optional=False, multiline=True,
                    placeholder="통화 후 추가 정보, 방문 시 참고할 사항 등을 남겨주세요."),
    ])

    return {
        "type": "modal",
        "callback_id": "submit_consult",
        "title": {"type": "plain_text", "text": "상담 처리"},
        "submit": {"type": "plain_text", "text": "등록"},
        "close": {"type": "plain_text", "text": "취소"},
        "private_metadata": metadata,
        "blocks": info_blocks + input_blocks,
    }


def _process_consult_submission(client, body, view):
    """통합 상담 모달 제출 → 처리 유형별 분기 (방문/견적/유선/문의 드랍/거래처/기타)"""
    metadata = json.loads(view.get("private_metadata") or "{}")
    lead_no = metadata.get("lead_no", "")
    chat_id = metadata.get("chat_id", "")  # 채널톡 카드 케이스
    channel = metadata.get("channel", "")
    message_ts = metadata.get("message_ts", "")
    user_id = body["user"]["id"]

    # 채널톡 chat_id 있으면 Redis 락 + pending lead 정리 (중복 등록 방지)
    if chat_id:
        try:
            from dashboard.utils.redis_client import get_redis_client
            rc = get_redis_client().redis
            # SETNX 락 — 60초 TTL (모달 제출이 끝날 시간 충분)
            lock_key = f'channeltalk_lead_lock:{chat_id}'
            if not rc.set(lock_key, '1', nx=True, ex=60):
                logger.info(
                    f"[SLACK/상담] 채널톡 chat_id={chat_id} 이미 처리 중 — 중복 제출 무시"
                )
                return
            # pending lead 데이터 삭제 (이 모달 제출이 정상 lead 등록 흐름)
            rc.delete(f'channeltalk_pending_lead:{chat_id}')
        except Exception as exc:
            logger.warning(f"[SLACK/상담] chat_id Redis 정리 실패: {exc}")

    # 인입 lead 락 — 동시 매니저 제출 시 데이터 손실 방지
    if lead_no:
        try:
            from dashboard.utils.redis_client import get_redis_client
            rc = get_redis_client().redis
            lock_key = f'consult_submit_lock:{lead_no}'
            if not rc.set(lock_key, '1', nx=True, ex=30):
                logger.info(
                    f"[SLACK/상담] {lead_no} 다른 매니저 처리 중 — 중복 제출 무시"
                )
                # 슬랙 thread에 안내
                if channel and message_ts:
                    try:
                        client.chat_postMessage(
                            channel=channel, thread_ts=message_ts,
                            text=f":warning: 다른 매니저가 `{lead_no}`를 동시에 처리 중이라 이번 제출은 무시했습니다. "
                                 f"30초 후 다시 시도해주세요."
                        )
                    except Exception:
                        pass
                return
        except Exception as exc:
            logger.warning(f"[SLACK/상담] 락 획득 실패: {exc}")

    state = view["state"]["values"]
    visit_type = _v(state, "visit_type") or '온라인'  # 온라인 / 거래처 / 기타
    status = _v(state, "status")  # 방문 예약 / 견적 제출 / 유선 상담 / 문의 드랍
    visit_date_raw = (_v(state, "visit_date") or '').strip()
    visit_date_end_raw = (_v(state, "visit_date_end") or '').strip()
    # 범위 표시 양식 적용 (같은 달: "MM-DD~DD" / 다른 달: "MM-DD~MM-DD")
    visit_date_display = _format_visit_date_range(visit_date_raw, visit_date_end_raw)
    visit_date_for_sheet = _format_date_for_sheet(visit_date_display) if visit_date_display else ''
    # 슬랙 카드 발송용 raw 표시 — 범위 양식 또는 단일
    visit_date_raw = visit_date_display
    name = (_v(state, "name") or '').strip()
    contact = (_v(state, "contact") or '').strip()
    email = (_v(state, "email") or '').strip()
    visit_address = (_v(state, "visit_address") or '').strip()
    consultation = (_v(state, "consultation") or '').strip()

    # 본인 방문 필수 라디오 (2026-07-17) — JW 담당자 배정 참고용
    _assign_state = (state.get('assign_self', {}).get('value', {}) or {}).get('selected_option') or {}
    assign_self_yes = (_assign_state.get('value') == 'yes')

    is_visit = (status == '방문 예약')
    is_estimate = (status == '견적 제출')

    # 방문 예약 + 본인 방문 필수 → 상담 내용 앞에 태그 프리픽스
    # (시트 저장 → 방문 카드/캔버스/List 자동 반영, 매니저는 이니셜로 자신임을 확인)
    if is_visit and assign_self_yes:
        _register_initial = _slack_user_to_initial(client, user_id) or '-'
        _tag = f':raising_hand: 본인 방문 필수({_register_initial})'
        if consultation:
            consultation = f'{_tag} — {consultation}'
        else:
            consultation = _tag

    # 두 차원 매핑 (시트 컬럼)
    category = visit_type   # 플랫폼 컬럼 = 방문 유형
    sheet_status = status   # 상태 컬럼 = 처리 유형

    # 재상담 이력 append 대비 (2026-07-20) — 시트에 저장되는 최종 상담 내용 (누적).
    # lead_no 케이스에서 옛 값 조회 후 append. 신규 lead 케이스는 그대로.
    # 카드 회색 헤더도 이 값을 파싱해 회차별 (n차) 렌더.
    full_consultation = consultation

    # ─────────────────────────────────────────────
    # 1) 인입 리드 케이스 — 기존 lead 시트 업데이트
    # ─────────────────────────────────────────────
    if lead_no:
        try:
            from dashboard.services.lead_service import update_lead
            update_data = {'상태': sheet_status}
            if is_visit and visit_date_for_sheet:
                update_data['방문 예정일'] = visit_date_for_sheet
            # 본인 방문 여부 (O열) — 방문 예약 한정. 거래처 워크플로우 저장값과
            # 동일 형식으로 통일 → JW 담당자 배정·필터링 시 소스 관계없이 동작 (2026-07-21).
            if is_visit:
                update_data['본인 방문 여부'] = (
                    '본인 방문 필수' if assign_self_yes else '아무나 방문 가능'
                )
            if name:
                update_data['고객명'] = name
            if contact:
                from dashboard.services.lead_helpers import normalize_phone
                update_data['고객 연락처'] = normalize_phone(contact) or contact
            if email:
                update_data['이메일'] = email
            if visit_address:
                update_data['방문 주소'] = visit_address
            if consultation:
                # 재상담 이력 append (2026-07-20) — 옛 K열 값에 [시간 이니셜 · status]
                # 헤더 붙인 새 entry 를 divider 로 이어붙임. 카드 렌더는 이 값 파싱.
                _cur_lead = _find_lead_by_no(lead_no) or {}
                _old_consult = str(_cur_lead.get('상담 내용') or '').strip()
                _initial_now = _slack_user_to_initial(client, user_id) or '-'
                _new_entry = _format_consultation_entry(
                    consultation, _initial_now, sheet_status,
                )
                full_consultation = _append_consultation(_old_consult, _new_entry)
                update_data['상담 내용'] = full_consultation
            # 상담하기 누른 매니저 → L열(온라인 상담자) — 드롭다운 값과 매칭되는 한국 이름
            counselor = _slack_user_to_korean_name(client, user_id)
            if counselor:
                update_data['온라인 상담자'] = counselor
            update_lead(lead_no, update_data)
        except Exception as exc:
            logger.error(f"[SLACK/상담] 시트 업데이트 실패 ({lead_no}): {exc}", exc_info=True)

        # 슬랙 List webhook — 방문 예약 한정 (유선 상담/견적/드랍은 list 미등록)
        if is_visit:
            lead = _find_lead_by_no(lead_no) or {}
            _post_to_slack_list(
                client, lead,
                modal_fields={
                    'visit_date': visit_date_raw,
                    'visit_address': visit_address,
                    'consultation': consultation,
                    'estimate': '',
                },
                channel=channel, message_ts=message_ts,
                action='visit',
            )

    # ─────────────────────────────────────────────
    # 2) 신규 리드 케이스 (거래처/기타, 슬래시 진입) — 시트에 새 lead 등록
    # ─────────────────────────────────────────────
    elif category in ('거래처', '기타'):
        # 거래처 / 기타 — 신규 lead 아님 (기존 매니저 추가 공사 또는 현장 용건)
        # → 시트 등록 X, lead_no 발번 X
        # 방문 예약이면 슬랙 List만 등록 (다음날 일정 정리용)
        from dashboard.services.lead_helpers import normalize_phone
        contact = normalize_phone(contact) or contact or '-'
        if is_visit:
            synthetic_lead = {
                '리드 No': '',
                '상담 시간': datetime.now().strftime('%Y.%m.%d. %H:%M'),
                '플랫폼': category,
                '고객명': name or '-',
                '고객 연락처': contact,
                '이메일': email or '-',
                '방문 주소': visit_address or '-',
                '상담 내용': consultation or '-',
                '키워드': '-',
            }
            try:
                _post_to_slack_list(
                    client, synthetic_lead,
                    modal_fields={
                        'visit_date': visit_date_raw,
                        'visit_address': visit_address,
                        'consultation': consultation,
                        'estimate': '',
                    },
                    channel=channel, message_ts=message_ts, action='visit',
                )
            except Exception as exc:
                logger.error(f"[SLACK/상담] 거래처/기타 List 등록 실패: {exc}", exc_info=True)

    else:
        # 슬래시 진입의 예외 케이스 (visit_type='온라인'인데 lead_no 없음 등) — 옛 신규 lead 등록 흐름
        try:
            from dashboard.services.lead_sync import _append_leads_to_main
            from dashboard.services.lead_helpers import normalize_phone
            counselor = _slack_user_to_korean_name(client, user_id) or '-'
            now = datetime.now()
            new_lead = {
                '리드 No': '',
                '상담 시간': now.strftime('%Y.%m.%d. %H:%M'),
                '플랫폼': category,
                '상태': sheet_status,
                '방문 예정일': visit_date_for_sheet or '-',
                '고객 연락처': normalize_phone(contact) or contact or '-',
                '이메일': email or '-',
                '고객명': name or '-',
                '방문 주소': visit_address or '-',
                '문의 내용': '',                 # 슬래시 신규 등록 — 인입 원본 없음
                '상담 내용': consultation or '-', # 매니저 입력 (옛 피드백 자리)
                '키워드': '-',
                '온라인 상담자': counselor,
                '영업 담당자': '',
                '마지막 연락일': '',
                '_meta_consult_dt': now,
            }
            lead_nos = _append_leads_to_main([new_lead])
            lead_no = lead_nos[0] if lead_nos else ''
        except Exception as exc:
            logger.error(f"[SLACK/상담] 신규 lead 등록 실패: {exc}", exc_info=True)

    # ─────────────────────────────────────────────
    # 3a) #방문_일정 채널에 방문 케이스 메시지 발송 (방문 예약 시만)
    # ─────────────────────────────────────────────
    visit_notice_channel, visit_notice_ts = '', ''
    if is_visit:
        # 인입 lead의 플랫폼 (홈페이지/당근/카카오톡/전화) — 헤더에 부가 표시
        lead_platform = ''
        if lead_no:
            existing_lead = _find_lead_by_no(lead_no) or {}
            lead_platform = str(existing_lead.get('플랫폼', '')).strip()
        visit_notice_channel, visit_notice_ts = _post_visit_notice(
            client, lead_no=lead_no, category=category, user_id=user_id,
            visit_date=visit_date_raw, name=name, contact=contact,
            visit_address=visit_address, consultation=consultation,
            platform=lead_platform,
        )

    # ─────────────────────────────────────────────
    # 3b) 원본 카드 thread reply + ✅ reaction (인입 카드 케이스만)
    #     — 헤더만 다르고 본문은 #방문_일정 채널 양식과 동일 (혼동 방지)
    # ─────────────────────────────────────────────
    if message_ts:
        SEP = '--------------------------------------------'
        ini = _slack_user_to_initial(client, user_id) or '-'
        # 모든 라인을 `>` blockquote로 통일 — 복사 시 줄바꿈 보존
        reply_lines = [
            f">:white_check_mark: *상담 완료 - {status}* — `{lead_no}`",
            f">{SEP}",
            f">등록자 : {ini}",
        ]
        if is_visit and visit_date_raw:
            reply_lines.append(f">방문일 : {visit_date_raw}")
        if name:
            reply_lines.append(f">이름 / 상호 : {name}")
        if contact:
            reply_lines.append(f">연락처 : {contact}")
        if visit_address:
            reply_lines.append(f">방문 주소 : {visit_address}")
        if consultation:
            reply_lines.append(f">상담 내용 :")
            for raw in consultation[:500].split('\n'):
                wrapped = textwrap.fill(
                    raw, width=60, break_long_words=True, break_on_hyphens=False,
                ) or raw
                for ln in wrapped.split('\n'):
                    reply_lines.append(f">{ln}")
        reply_lines.append(f">{SEP}")
        reply_text = '\n'.join(reply_lines)
        # 같은 lead 재제출 시 옛 reply가 있으면 그 메시지를 chat.update로 갱신
        old_reply_ts = ''
        try:
            from dashboard.utils.redis_client import get_redis_client
            rc = get_redis_client().redis
            reply_key = f"consult_reply:{lead_no}"
            cached = rc.get(reply_key)
            if cached:
                old_reply_ts = (
                    cached.decode() if isinstance(cached, bytes) else cached
                )
        except Exception as exc:
            logger.debug(f"[SLACK/상담] reply 캐시 조회 실패: {exc}")

        # 순서 — chat.update(회색 박스) → reaction → thread reply
        # 옛 공사현황 봇과 같은 순서 — slack UI가 reply count 표시 안 되는 케이스 회피
        # 1) 원본 카드 본문 회색 박스 변환 (부재중은 배지만 표시 + 원본 유지, 2026-07-17)
        original_text = metadata.get("original_text", "") if isinstance(metadata, dict) else ''
        _initial_for_card = _slack_user_to_initial(client, user_id) or '-'
        _now_for_card = datetime.now().strftime('%m.%d %H:%M')

        if original_text and status == '부재중':
            # 부재중 — 원본 카드 유지 + 상단에 부재중 배지 (재클릭 시 시각·횟수 갱신)
            try:
                from dashboard.utils.redis_client import get_redis_client
                _rc = get_redis_client().redis
                _count_key = f'consult_missed_count:{lead_no}'
                _count = int(_rc.incr(_count_key) or 1)
                _rc.expire(_count_key, 60 * 60 * 24 * 90)
            except Exception:
                _count = 1

            _badge_text = (
                f':repeat: *부재중* — 마지막 시도: `{_now_for_card}` ({_initial_for_card}) · '
                f'총 *{_count}회*'
            )
            try:
                # 기존 카드 blocks fetch → 상단 배지 replace (기존 배지 있으면 갈아끼우기)
                _rp = client.conversations_replies(channel=channel, ts=message_ts, limit=1, inclusive=True)
                _root = ((_rp.get('messages') or [{}])[0]) if _rp else {}
                _existing_blocks = list(_root.get('blocks') or [])
                _has_badge = (
                    _existing_blocks
                    and _existing_blocks[0].get('type') == 'context'
                    and any(
                        '부재중' in ((el.get('text') or '') if isinstance(el, dict) else '')
                        for el in (_existing_blocks[0].get('elements') or [])
                    )
                )
                _badge_block = {
                    'type': 'context',
                    'elements': [{'type': 'mrkdwn', 'text': _badge_text}],
                }
                if _has_badge:
                    _existing_blocks[0] = _badge_block
                else:
                    _existing_blocks.insert(0, _badge_block)
                client.chat_update(
                    channel=channel, ts=message_ts,
                    text=_root.get('text', '') or '',
                    blocks=_existing_blocks,
                )
            except Exception as exc:
                logger.warning(f"[SLACK/상담] 부재중 배지 갱신 실패 ({lead_no}): {exc}")

        elif original_text:
            try:
                cancel_time = _now_for_card
                initial = _initial_for_card
                cleaned_lines = [ln.lstrip('>').lstrip() for ln in original_text.split('\n')]
                cleaned_lines = [ln.replace('*', '') for ln in cleaned_lines]
                clean_text = '\n'.join(cleaned_lines)
                clean_text = re.sub(r'^[\s⠀]+|[\s⠀]+$', '', clean_text)
                # shortcode → unicode (:bell: → 🔔 등) — 코드 블록 안 이모지 렌더 (2026-07-22)
                try:
                    from dashboard.blueprints.slack_helpers import _normalize_shortcodes_to_unicode
                    clean_text = _normalize_shortcodes_to_unicode(clean_text)
                except Exception:
                    pass
                header_lines = [
                    "⠀",
                    f":white_check_mark: *상담 완료 - {status}*",
                    f"처리자 : {initial}",
                    f"처리 시간 : {cancel_time}",
                ]
                # 재상담 이력 회차별 렌더 (2026-07-20) — full_consultation 은 위쪽에서
                # append 된 최종 값. 재상담(2회차 이상) 만 (n차) 라벨. 첫 상담(1개 회차)
                # 은 라벨 없이 그냥 '상담 내용 :' 으로 표시 (노이즈 방지).
                _entries = _parse_consultation_entries(full_consultation) if full_consultation else []
                if _entries and len(_entries) >= 2:
                    total = len(_entries)
                    for i, e in enumerate(_entries):
                        idx = i + 1
                        _c = e.get('content', '').strip() or '-'
                        _ini_tag = e.get('ini', '').strip()
                        # 마지막(최신) 회차는 이니셜 생략 (헤더 처리자와 동일)
                        if idx == total or not _ini_tag:
                            header_lines.append(f"상담 내용 ({idx}차) : {_c}")
                        else:
                            header_lines.append(f"상담 내용 ({idx}차) : {_c} ({_ini_tag})")
                elif _entries:
                    # 첫 상담 — 회차 라벨 없이 그냥 표시
                    header_lines.append(f"상담 내용 : {_entries[0].get('content','').strip() or '-'}")
                elif consultation:
                    header_lines.append(f"상담 내용 : {consultation}")
                new_text = '\n'.join(header_lines) + f"\n\n```\n{clean_text}\n```"
                new_blocks = [
                    {"type": "section", "text": {"type": "mrkdwn", "text": new_text}},
                ]
                # 재상담 버튼 — 회색 카드에서도 재편집 진입점 유지 (2026-07-17 사용자 요청).
                # 방문 예약은 별도 방문 카드에 [정보 수정] 있어 여기 재상담 버튼 불필요.
                # 2026-07-20: button_consult (통합 상담 모달) 로 스위칭 — 상담 유형
                # 드롭다운(유선/방문/견적/드랍/부재중) + 이전 상담 이력 미리보기 포함.
                if status != '방문 예약' and lead_no:
                    new_blocks.append({
                        'type': 'actions',
                        'elements': [{
                            'type': 'button',
                            'text': {'type': 'plain_text', 'text': '✏️ 재상담', 'emoji': True},
                            'value': lead_no,
                            'action_id': 'button_consult',
                        }],
                    })
                client.chat_update(
                    channel=channel, ts=message_ts, text=new_text, blocks=new_blocks,
                )
            except Exception as exc:
                logger.warning(f"[SLACK/상담] 카드 회색 처리 실패 ({lead_no}): {exc}")

        # 2) 원본 카드 ✅ reaction
        try:
            client.reactions_add(
                channel=channel, timestamp=message_ts, name="white_check_mark",
            )
        except Exception:
            pass

        # 3) thread reply 발송 (slack UI가 reply count 표시 갱신하도록 마지막에)
        # 2026-07-16: 방문 예약은 방문_일정 채널 root 카드 permalink 링크만 짧게 발송
        #             (중복 노이즈 방지). 유선 상담/견적/드랍/부재중은 상세 답글 유지.
        if is_visit and visit_notice_channel and visit_notice_ts:
            try:
                perm = client.chat_getPermalink(
                    channel=visit_notice_channel, message_ts=visit_notice_ts,
                )
                permalink = (perm or {}).get('permalink', '')
            except Exception as exc:
                logger.debug(f"[SLACK/상담] 방문 카드 permalink 조회 실패 ({lead_no}): {exc}")
                permalink = ''
            ini = _slack_user_to_initial(client, user_id) or '-'
            if permalink:
                reply_text = (
                    f":white_check_mark: *방문 예약 등록* — `{lead_no}` by `{ini}`\n"
                    f":round_pushpin: <{permalink}|#방문_일정 카드에서 상세 보기>"
                )
            else:
                reply_text = (
                    f":white_check_mark: *방문 예약 등록* — `{lead_no}` by `{ini}`\n"
                    f"_(#방문_일정 채널에 카드 발송됨)_"
                )
        # (is_visit 아니면 위쪽에서 조립된 상세 reply_text 그대로 사용)

        reply_sent = False
        if old_reply_ts:
            if is_visit:
                # 방문 예약은 방문 카드 unfurl embed 필요 — chat.update 로는
                # Slack 이 unfurl 재생성 안 함 (unfurl_links 파라미터 미지원).
                # delete + repost 로 새 chat.postMessage 발송해 unfurl 트리거.
                try:
                    client.chat_delete(channel=channel, ts=old_reply_ts)
                except Exception as exc:
                    logger.warning(
                        f"[SLACK/상담] 옛 reply 삭제 실패 — 새 reply 발송: {exc}"
                    )
            else:
                try:
                    client.chat_update(
                        channel=channel, ts=old_reply_ts, text=reply_text,
                    )
                    reply_sent = True
                except Exception as exc:
                    logger.warning(
                        f"[SLACK/상담] 옛 reply update 실패 — 새 reply 발송: {exc}"
                    )
        if not reply_sent:
            try:
                resp = client.chat_postMessage(
                    channel=channel, thread_ts=message_ts, text=reply_text,
                )
                if resp and resp.get('ok') and resp.get('ts'):
                    try:
                        rc.set(
                            reply_key, resp['ts'], ex=60 * 60 * 24 * 90,
                        )
                    except Exception:
                        pass
            except Exception as exc:
                logger.error(f"[SLACK/상담] thread reply 실패: {exc}", exc_info=True)
    else:
        # 슬래시 진입 케이스 — ephemeral 확인 메시지
        try:
            client.chat_postEphemeral(
                channel=channel or user_id, user=user_id,
                text=f":white_check_mark: *{status}* 등록 완료 — `{lead_no}` (카테고리: {category})",
            )
        except Exception:
            pass


def _post_visit_notice(client, lead_no: str, category: str, user_id: str,
                       visit_date: str, name: str, contact: str,
                       visit_address: str, consultation: str,
                       user_name: str = '', platform: str = '',
                       addr_note: Optional[dict] = None) -> tuple:
    """#방문_일정 채널에 방문 케이스 알림 발송 (통합 모달 + 전화 모달 + 워크플로 공용).

    헤더 양식:
      - platform 있으면: ":bell: 새 방문 일정 — {category}({platform}) `lead_no`"
      - 없으면:          ":bell: 새 방문 일정 — {category}  `lead_no`"
    본문 첫 줄에 "등록자 : 이니셜" 표시.

    발송 봇: 별도 방문 일정 알림 봇(_visit_slack_app) 우선. 미설정 시 메인 봇 fallback.
    """
    visit_channel = os.getenv('SLACK_VISIT_CHANNEL', '').strip()
    if not visit_channel:
        return ('', '')

    # 방문 일정 봇 client 우선 사용 (액션 매칭을 위해 카드도 visit bot 명의로)
    if _visit_slack_handler is None:
        _init_visit_slack_app()
    if _visit_slack_app is not None:
        client = _visit_slack_app.client

    # 등록자 이니셜
    initial = _slack_user_to_initial(client, user_id) if user_id else _to_initial(user_name)

    # 카테고리(플랫폼) 표시
    # 소개 건은 '거래처 (소개)' 로 표시 (2026-07-15 사용자 요청 — 소개 = 거래처 하위)
    if category == '소개' or platform == '소개':
        category_display = '거래처 (소개)'
    elif platform and platform != category:
        category_display = f"{category} ({platform})"
    else:
        category_display = category

    # 본인 방문 필수 배지 (2026-07-17) — 리드 시트 O열 값 감지
    _self_visit_by = ''
    if lead_no:
        _lead_ctx = _find_lead_by_no(lead_no) or {}
        _o = str(_lead_ctx.get('본인 방문 여부') or '').strip()
        if '본인 방문 필수' in _o:
            # 카드 신청자 이름 = 온라인 상담자 (M열, 워크플로 시작자 기록)
            _requester = str(_lead_ctx.get('온라인 상담자') or '').strip().lstrip('@')
            _self_visit_by = _requester if _requester and _requester != '-' else '본인'

    # 재상담 append 형식이면 최신 회차 content 만 카드에 표시 (헤더 제거).
    # K열이 '[MM.DD HH:MM 이니셜 · status] 내용 ─── [...]' 로 누적되므로 헤더 노출 방지.
    _entries = _parse_consultation_entries(consultation) if consultation else []
    if _entries:
        consultation = _entries[-1].get('content', '') or consultation

    body_text, blocks = _build_visit_notice_blocks(
        lead_no=lead_no, category_display=category_display, initial=initial,
        visit_date=visit_date, name=name, contact=contact,
        visit_address=visit_address, consultation=consultation,
        self_visit_by=_self_visit_by, addr_note=addr_note,
    )
    # 재제출이면 기존 방문 카드 메시지를 chat.update — 중복 발송 방지
    redis_key = f"visit_notice_msg:{lead_no}" if lead_no else ''
    existing_ts = ''
    if redis_key:
        try:
            from dashboard.utils.redis_client import get_redis_client
            rc = get_redis_client().redis
            stored = rc.get(redis_key)
            if stored:
                stored = stored.decode('utf-8') if isinstance(stored, bytes) else stored
                if '|' in stored:
                    stored_channel, existing_ts = stored.split('|', 1)
                    if stored_channel != visit_channel:
                        existing_ts = ''
        except Exception as exc:
            logger.warning(f"[SLACK/방문] 기존 메시지 ts 조회 실패 ({lead_no}): {exc}")

    if existing_ts:
        try:
            client.chat_update(
                channel=visit_channel, ts=existing_ts,
                text=body_text, blocks=blocks,
            )
            return (visit_channel, existing_ts)
        except Exception as exc:
            # 2026-07-17: message_not_found = 매니저가 슬랙에서 카드 message 자체를 삭제.
            # 이 경우 fallback 신규 발송은 오히려 노이즈 (매니저가 지운 걸 다시 발송하는 꼴).
            # Redis 매핑만 정리하고 return.
            _err_str = str(exc)
            if 'message_not_found' in _err_str:
                logger.info(
                    f"[SLACK/방문] 옛 카드 삭제 감지 ({lead_no}, ts={existing_ts}) → "
                    f"재발송 skip (매니저 의도 존중)"
                )
                try:
                    if redis_key:
                        rc.delete(redis_key)
                except Exception:
                    pass
                return ('', '')
            logger.warning(f"[SLACK/방문] chat.update 실패 ({lead_no}, ts={existing_ts}): {exc} — 신규 발송 fallback")

    try:
        resp = client.chat_postMessage(
            channel=visit_channel, text=body_text,
            blocks=blocks, unfurl_links=False,
        )
        ts = resp.get('ts', '') if resp else ''
        if redis_key and ts:
            try:
                from dashboard.utils.redis_client import get_redis_client
                rc = get_redis_client().redis
                rc.set(redis_key, f"{visit_channel}|{ts}", ex=60 * 60 * 24 * 180)  # 180일
            except Exception as exc:
                logger.warning(f"[SLACK/방문] ts 저장 실패 ({lead_no}): {exc}")
        # 주소 정규화 배지 있으면 등록자에게 ephemeral 발송 (신규 카드 한정).
        # chat.update 케이스는 이미 이전에 발송했으므로 중복 방지 위해 skip.
        if addr_note:
            _post_addr_note_ephemeral(
                client, visit_channel=visit_channel, lead_no=lead_no,
                user_id=user_id, user_name=user_name, addr_note=addr_note,
            )
        return (visit_channel, ts)
    except Exception as exc:
        logger.warning(f"[SLACK/방문] #방문_일정 발송 실패: {exc}")
        return ('', '')


def _post_addr_note_ephemeral(client, visit_channel: str, lead_no: str,
                               user_id: str, user_name: str,
                               addr_note: dict) -> None:
    """방문 카드 발송 직후 등록자에게 주소 정규화 결과 ephemeral 발송.

    거래처/기타/소개 워크플로 lead 전용 — 매니저가 raw 주소 붙여넣었을 때
    자동 정정된 결과 or 검증 실패 사실을 본인에게만 알려서 확인·재입력 유도.

    user_id 없으면 user_name → users.db email → users_lookupByEmail 로 조회.
    lookup 실패 or 채널·kind 미유효 시 조용히 skip.
    """
    if not addr_note or not isinstance(addr_note, dict):
        return
    if not visit_channel:
        return
    _kind = addr_note.get('kind', '')
    if _kind not in ('normalized', 'failed'):
        return

    # 등록자 slack user_id 확보
    target_uid = (user_id or '').strip()
    if not target_uid and user_name:
        try:
            from dashboard.utils.user_database import UserDatabase
            db = UserDatabase()
            email = ''
            for u in db.get_all_users():
                if (u.get('name') or '').strip() == user_name.strip():
                    email = (u.get('email') or '').strip()
                    break
            if email:
                u_resp = client.users_lookupByEmail(email=email)
                target_uid = ((u_resp.get('user') or {}) if u_resp else {}).get('id', '')
        except Exception as exc:
            logger.warning(
                f'[SLACK/방문] ephemeral 대상 lookup 실패 '
                f'({lead_no}, name={user_name}): {exc}'
            )
            return
    if not target_uid:
        return

    _orig = (addr_note.get('original') or '').strip()
    _norm = (addr_note.get('normalized') or '').strip()
    if _kind == 'normalized':
        text = (
            f":mag: 방금 등록한 `{lead_no}` 방문 주소가 "
            f"카카오 API 로 자동 정정됐어요.\n\n"
            f"  원본: {_orig}\n"
            f"  정정: {_norm}\n\n"
            "정정된 주소가 맞는지 위 카드에서 확인 부탁드립니다.\n"
            "잘못 매핑됐다면 [✏️ 정보 수정] 으로 다시 입력해주세요."
        )
    else:  # failed
        text = (
            f":warning: 방금 등록한 `{lead_no}` 방문 주소를 "
            f"카카오 API 가 인식하지 못했습니다.\n"
            "위 카드에서 [✏️ 정보 수정] 으로 정확한 주소를 다시 입력해주세요.\n\n"
            f"  입력값: {_orig}"
        )
    try:
        client.chat_postEphemeral(
            channel=visit_channel, user=target_uid, text=text,
        )
    except Exception as exc:
        logger.warning(
            f'[SLACK/방문] 주소 정정 ephemeral 발송 실패 ({lead_no}): {exc}'
        )


_WORD_JOINER = '⁠'  # Slack mrkdwn word-boundary 우회용 (폭 0, 복사·검색 무시)


def _wrap_diff_chunk(chunk: str, prev_ch: str, next_ch: str, marker: str) -> str:
    """diff 청크에 mrkdwn marker(*, `) 감싸기.

    한글/영문/숫자가 marker 에 딱 붙으면 Slack 이 리터럴로 렌더하므로
    필요 시 Word Joiner(U+2060) 삽입해 word boundary 를 만듦.
    카카오 원본 형식 (공백 없음) 을 시각적으로 유지.
    """
    left = _WORD_JOINER if prev_ch and prev_ch.isalnum() else ''
    right = _WORD_JOINER if next_ch and next_ch.isalnum() else ''
    return f'{left}{marker}{chunk}{marker}{right}'


def _highlight_addr_diff(original: str, converted: str) -> tuple:
    """원본↔변환 diff 청크를 길이에 따라 다른 스타일로 감싼 (orig, conv) 튜플 반환.

    - 청크 길이 ≤ 2 (자모/숫자 오탈자): 인라인 코드(`chunk`) — monospace 회색 배경
    - 청크 길이 ≥ 3 (지역명·건물명 추가): 볼드(*chunk*)

    자모 1자 diff (벨→밸) 는 프로포셔널 폰트에서 안 보이므로 monospace 필수.
    지역명·건물명 추가는 라벨 부담이 커 볼드로 완화.
    청크가 한글/영문/숫자 사이에 낀 경우 Word Joiner 로 word boundary 확보.

    주의: blockquote(>) 컨텍스트 전용. 회색 코드블록(```) 안에선 mrkdwn 리터럴이라 사용 X.
    """
    if not original or not converted or original == converted:
        return original, converted
    from difflib import SequenceMatcher
    sm = SequenceMatcher(None, original, converted, autojunk=False)
    orig_parts = []
    conv_parts = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        o = original[i1:i2]
        n = converted[j1:j2]
        if tag == 'equal':
            orig_parts.append(o)
            conv_parts.append(n)
            continue
        marker = '`' if max(len(o), len(n)) <= 2 else '*'
        if o:
            prev_ch = original[i1 - 1] if i1 > 0 else ''
            next_ch = original[i2] if i2 < len(original) else ''
            orig_parts.append(_wrap_diff_chunk(o, prev_ch, next_ch, marker))
        if n:
            prev_ch = converted[j1 - 1] if j1 > 0 else ''
            next_ch = converted[j2] if j2 < len(converted) else ''
            conv_parts.append(_wrap_diff_chunk(n, prev_ch, next_ch, marker))
    return ''.join(orig_parts), ''.join(conv_parts)


def _build_visit_notice_blocks(lead_no: str, category_display: str, initial: str,
                                visit_date: str, name: str, contact: str,
                                visit_address: str, consultation: str,
                                self_visit_by: str = '',
                                addr_note: Optional[dict] = None) -> tuple:
    """방문 일정 카드 양식 빌더 — (text, blocks) 반환.

    [✏️ 방문일 수정] + [🗑️ 방문 취소] 액션 버튼 포함. 카드 발송/복원 양쪽에서 재사용.

    self_visit_by (2026-07-17): 값 있으면 헤더 아래 '🙋 본인 방문 필수 (name)' 배지.
    addr_note (2026-07-20): 거래처/기타/소개 워크플로 lead 의 주소 정규화 배지.
      {'kind': 'normalized', 'original': ..., 'normalized': ...} → 자동 정정 표시
      {'kind': 'failed',     'original': ..., 'normalized': ''}  → 확인 실패 안내
      사업자등록증 배지와 동일 패턴 — 본문 구분선 바깥 하단 컨텍스트 라인.
    """
    SEP = '--------------------------------------------'
    # lead_no 없으면 (거래처/기타) 헤더에 표시 안 함
    header_suffix = f"  `{lead_no}`" if lead_no else ''
    # 등록자 라인 — 본인 방문 필수 시 인라인 배지 (2026-07-17)
    _register_line = f">등록자 : {initial or '-'}"
    if self_visit_by:
        _register_line += ' (:raising_hand: *본인 방문 필수*)'
    # 방문 주소 렌더 — 온라인 lead 카드와 동일 스타일로 통일 (2026-07-22)
    # addr_note 로 정규화 결과 판별:
    #   normalized + 원본 != 변환 → 원본/변환 두 줄
    #   failed → 방문 주소 + [주소 확인 필요] 배지
    #   그 외 (배지 없음) → 방문 주소 한 줄
    _addr_lines = []
    if addr_note and isinstance(addr_note, dict):
        _kind = addr_note.get('kind', '')
        _orig = (addr_note.get('original') or '').strip()
        if _kind == 'normalized' and _orig and _orig != (visit_address or '').strip():
            _orig_hl, _conv_hl = _highlight_addr_diff(_orig, (visit_address or '').strip())
            _addr_lines.append(f">*원본 주소* : {_orig_hl}")
            _addr_lines.append(f">*변환 주소* : {_conv_hl}")
        elif _kind == 'failed':
            _addr_lines.append(
                f">방문 주소 : {visit_address or '-'}  :warning: *[주소 확인 필요]*"
            )
    if not _addr_lines:
        _addr_lines.append(f">방문 주소 : {visit_address or '-'}")
    lines = [
        "⠀",
        f">:bell: *새 방문 일정* — {category_display}{header_suffix}",
        f">{SEP}",
        _register_line,
        f">방문일 : {visit_date or '-'}",
        f">이름 / 상호 : {name or '-'}",
        f">연락처 : {contact or '-'}",
        *_addr_lines,
    ]
    if consultation:
        # 재상담 이력 있으면 최신 회차 content 만 표시 (2026-07-22 L-03295 사고)
        # K열 append 형식 `[MM.DD HH:MM 이니셜 · 상태] content ─── [...]` 그대로 렌더되던 이슈.
        try:
            _entries = _parse_consultation_entries(consultation)
            if _entries:
                consultation = (_entries[-1].get('content') or '').strip() or consultation
        except Exception:
            pass
        lines.append(f">상담 내용 :")
        for raw in consultation[:500].split('\n'):
            wrapped = textwrap.fill(
                raw, width=60, break_long_words=True, break_on_hyphens=False,
            ) or raw
            for ln in wrapped.split('\n'):
                lines.append(f">{ln}")
    lines.append(f">{SEP}")
    # 하단 주소 정규화 배지 제거 (2026-07-22): 본문 상단 "원본/변환" 두 줄로 통합
    # → 하단 :mag: 이모지가 오히려 시선 끌어 본문 놓치는 이슈 해소.
    body_text = '\n'.join(lines)
    blocks = [
        {"type": "section", "text": {"type": "mrkdwn", "text": body_text}},
        {
            "type": "actions",
            "elements": [
                {
                    "type": "button",
                    "text": {"type": "plain_text", "text": "✏️ 정보 수정", "emoji": True},
                    "value": lead_no,
                    "action_id": "visit_edit_info",
                },
                {
                    "type": "button",
                    "text": {"type": "plain_text", "text": "✅ 방문 완료", "emoji": True},
                    "style": "primary",
                    "value": lead_no,
                    "action_id": "visit_complete",
                    "confirm": {
                        "title": {"type": "plain_text", "text": "방문 완료 처리"},
                        "text": {"type": "plain_text",
                                 "text": "이 방문을 완료 처리하시겠습니까?\n(슬랙 리스트에서 삭제됩니다)"},
                        "confirm": {"type": "plain_text", "text": "완료 확정"},
                        "deny": {"type": "plain_text", "text": "취소"},
                    },
                },
                {
                    "type": "button",
                    "text": {"type": "plain_text", "text": "🗑️ 방문 취소", "emoji": True},
                    "style": "danger",
                    "value": lead_no,
                    "action_id": "visit_cancel",
                    # 2026-07-19: confirm 팝업 대신 사유 입력 모달로 대체.
                    # 하루~일주일 미루기는 [정보 수정] 으로, 한 달 이상·완전 취소는 이 버튼.
                },
            ],
        },
        {"type": "context", "elements": [{"type": "mrkdwn", "text": "⠀"}]},
    ]
    return body_text, blocks


def _trigger_visit_list_webhook(env_key: str, lead_no: str, channel: str,
                                  message_ts: str, new_visit_date: str = '') -> None:
    """슬랙 워크플로우 webhook 호출 — list 행 삭제/업데이트.

    env_key: SLACK_VISIT_CANCEL_WEBHOOK_URL 또는 SLACK_VISIT_MODIFY_WEBHOOK_URL
    new_visit_date: 날짜 수정 시 새 날짜 (ISO YYYY-MM-DD). 빈값이면 payload 미포함.
    """
    url = os.getenv(env_key, '').strip()
    if not url:
        logger.debug(f"[SLACK/방문 list] {env_key} 미설정 — 호출 스킵")
        return

    lead = _find_lead_by_no(lead_no) or {}

    # 메시지 permalink
    message_link = ''
    if channel and message_ts:
        try:
            link_client = (_visit_slack_app.client if _visit_slack_app
                           else _slack_app.client)
            resp = link_client.chat_getPermalink(
                channel=channel, message_ts=message_ts,
            )
            if resp and resp.get('ok'):
                message_link = resp.get('permalink', '')
        except Exception:
            pass

    def _strip_escape(s: str) -> str:
        s = (s or '').strip()
        return s[1:] if s.startswith("'") else s

    # visit_type — 리스트 드롭다운(온라인/거래처/기타) 매핑
    # 플랫폼(홈페이지/카카오톡/당근/…) 값 그대로 보내면 드롭다운 매칭 실패 → 빈값(`-`) 저장됨
    _lead_platform = str(lead.get('플랫폼', '') or '').strip()
    if _lead_platform in ('거래처', '소개'):
        _visit_type_category = '거래처'
    elif _lead_platform == '기타':
        _visit_type_category = '기타'
    else:
        _visit_type_category = '온라인'
    # 방문 예정일 분리 — start/end ISO 변수도 함께 전달
    _vd_raw = str(lead.get('방문 예정일', '') or '').strip()
    _vd_start, _vd_end = _split_visit_date_range(_vd_raw)
    payload = {
        'lead_no': lead_no or '-',
        'platform': _lead_platform or '-',
        'visit_type': _visit_type_category,
        'email': str(lead.get('이메일', '') or '').strip(),
        'details': '',
        'contact': str(lead.get('고객 연락처', '') or '').strip(),
        'message_link': message_link,
        'payload': lead_no,
        'consultation': str(lead.get('상담 내용', '') or '').strip(),
        'estimate_request': '',
        'visit_date': _strip_escape(str(lead.get('방문 예정일', '') or '')),
        'visit_date_start': _vd_start or '-',
        'visit_date_end': _vd_end or '-',
        'device': str(lead.get('키워드', '') or '').strip(),
        'visit_address': str(lead.get('방문 주소', '') or '').strip(),
        'name': str(lead.get('고객명', '') or '').strip(),
        'inquiry_time': str(lead.get('상담 시간', '') or '').strip(),
        'location': '',
    }
    if new_visit_date:
        payload['new_visit_date'] = new_visit_date

    try:
        data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        req = urllib.request.Request(
            url, data=data,
            headers={'Content-Type': 'application/json; charset=utf-8'},
        )
        with urllib.request.urlopen(req, timeout=5) as r:
            r.read()
        logger.info(f"[SLACK/방문 list] {env_key} 호출 완료 ({lead_no})")
    except Exception as exc:
        logger.warning(f"[SLACK/방문 list] {env_key} 호출 실패 ({lead_no}): {exc}")

    # 방문 캔버스 rebuild — 정보 수정 (MODIFY) / 완료 (COMPLETE) / 취소 (CANCEL)
    # 어느 경로든 시트 상태·방문일이 바뀔 수 있어 캔버스 반영 필요 (2026-07-16).
    try:
        from dashboard.services.visit_canvas_sync import rebuild_canvas_async
        rebuild_canvas_async()
    except Exception as _vc_exc:
        logger.debug(f"[SLACK/방문 list] 캔버스 rebuild trigger 실패: {_vc_exc}")


def _process_visit_date_modify(client, body, view) -> None:
    """[📅 방문일 수정] 모달 제출 처리 — 시트 update + 메시지 chat.update.

    시작일 + (선택) 종료일 지원 — 범위 양식 (2026-07-01~03 등) 자동 조립.
    """
    metadata = json.loads(view.get("private_metadata") or "{}")
    lead_no = metadata.get("lead_no", "")
    channel = metadata.get("channel", "")
    message_ts = metadata.get("message_ts", "")
    state = view["state"]["values"]
    new_start = _v(state, "visit_date") or ''
    new_end = _v(state, "visit_date_end") or ''
    if not lead_no or not new_start:
        return
    # 범위 양식으로 조립 — end 없거나 start와 같으면 단일
    new_date_display = _format_visit_date_range(new_start, new_end)

    # 방문일 변경 감지용 — old 값 캡처 (2026-07-19)
    old_lead = _find_lead_by_no(lead_no) or {}
    old_visit_date = str(old_lead.get('방문 예정일') or '').strip().lstrip("'")

    # 1) 시트 update — escape prefix로 시리얼 변환 차단.
    # ETC- 는 Redis metadata 만 갱신 (시트에 없음).
    try:
        sheet_value = new_date_display  # E열 셀 서식 '@ 텍스트' 라 escape 불필요
        _update_lead_dispatch(lead_no, {'방문 예정일': sheet_value})
    except Exception as exc:
        logger.error(f"[SLACK/방문수정] 시트 update 실패 ({lead_no}): {exc}", exc_info=True)
        return

    # 1-2) 슬랙 List 동기화 워크플로우 — 날짜 셀 갱신 (범위/단일 통합)
    _trigger_visit_list_webhook(
        'SLACK_VISIT_MODIFY_WEBHOOK_URL', lead_no, channel, message_ts,
        new_visit_date=new_date_display,
    )

    # 2) 메시지 chat.update — 시트 lead 정보로 카드 재구성
    # conversations.history는 visit bot에 권한 없으므로 사용 X
    try:
        lead = _find_lead_by_no(lead_no) or {}
        platform = str(lead.get('플랫폼', '')).strip()
        # 거래처/기타는 그 자체 표시, 그 외(전화/홈페이지/당근/카카오톡 등)는 '온라인 (플랫폼)'
        if platform in ('거래처', '기타'):
            category = platform
            category_display = category
        else:
            category = '온라인'
            category_display = f"{category}({platform})" if platform else category

        user_id = body["user"]["id"]
        initial = _slack_user_to_initial(client, user_id) or '-'
        body_text, blocks = _build_visit_notice_blocks(
            lead_no=lead_no, category_display=category_display, initial=initial,
            visit_date=new_date_display,
            name=str(lead.get('고객명', '') or '').strip(),
            contact=str(lead.get('고객 연락처', '') or '').strip(),
            visit_address=str(lead.get('방문 주소', '') or '').strip(),
            consultation=str(lead.get('상담 내용', '') or '').strip(),
        )
        client.chat_update(
            channel=channel, ts=message_ts, text=body_text, blocks=blocks,
        )
    except Exception as exc:
        logger.error(f"[SLACK/방문수정] 메시지 update 실패 ({lead_no}): {exc}",
                     exc_info=True)

    # 방문일 변경 → dm_sent flag 있는 lead 만 담당자에게 알림 (2026-07-19)
    if old_visit_date and new_date_display and old_visit_date != new_date_display:
        try:
            from dashboard.services.visit_assignment_sync import send_visit_change_notification
            threading.Thread(
                target=send_visit_change_notification,
                args=(lead_no, old_visit_date, new_date_display, ''),
                daemon=True,
            ).start()
        except Exception as exc:
            logger.warning(f"[SLACK/방문수정] 변경 알림 예약 실패 ({lead_no}): {exc}")


# ─────────────────────────────────────────────────────────────
# [✏️ 정보 수정] 확장 모달 (2026-07-15) — 유형/일정/이름/연락처/주소/상담
# ─────────────────────────────────────────────────────────────
# ETC- pseudo-lead 는 Redis metadata, 정규 리드는 시트 셀 update.
# 유형 변경 시 ETC↔정규 전환은 별도 커밋(3~4)에서 추가. 이 커밋은
# "유형 동일" case 만 처리 (필드만 update).
_VISIT_PLATFORM_OPTIONS = ['거래처', '소개', '기타']
# 온라인 리드 플랫폼 — 편집 modal 에서 유형 dropdown 을 렌더링하지 않음
# (platform 은 원본 유입 소스로 고정. 매니저가 임의로 바꾸면 안 됨)
_ONLINE_LEAD_PLATFORMS = ('당근', '홈페이지', '카카오톡', '전화')


def _open_visit_edit_modal(client, lead_no: str, channel: str,
                            message_ts: str, trigger_id: str) -> None:
    """정보 수정 모달 open — 기존 값 pre-fill.

    온라인 리드(당근/홈페이지/카카오톡/전화) 는 유형 dropdown 을 숨긴다.
    - 원본 platform 은 metadata 에 그대로 저장 → submit 시 유지
    - dropdown 은 거래처/소개/기타 슬래시 진입 case 에서만 표시
    """
    lead = _find_lead_by_no(lead_no) or {}
    raw_platform = str(lead.get('플랫폼', '') or '').strip()
    is_online_lead = raw_platform in _ONLINE_LEAD_PLATFORMS
    # 온라인 리드는 dropdown 렌더 skip → cur_platform 은 metadata 저장용 원본
    # 거래처/소개/기타가 아니면서 온라인 리드도 아닌 예외 case 만 '거래처' fallback
    if is_online_lead:
        cur_platform = raw_platform
    elif raw_platform in _VISIT_PLATFORM_OPTIONS:
        cur_platform = raw_platform
    else:
        cur_platform = '거래처'

    # 방문 예정일 시작/종료 분리
    cur_visit_date = str(lead.get('방문 예정일', '') or '').strip()
    if cur_visit_date.startswith("'"):
        cur_visit_date = cur_visit_date[1:]
    cur_start, cur_end = _split_visit_date_range(cur_visit_date)

    cur_name = str(lead.get('고객명', '') or '').strip()
    cur_phone = str(lead.get('고객 연락처', '') or '').strip()
    cur_address = str(lead.get('방문 주소', '') or '').strip()
    cur_consultation = (
        str(lead.get('상담 내용', '') or '').strip() or
        str(lead.get('문의 내용', '') or '').strip()
    )
    if cur_consultation == '-':
        cur_consultation = ''

    metadata = json.dumps({
        'lead_no': lead_no,
        'channel': channel,
        'message_ts': message_ts,
        'original_platform': cur_platform,  # 유형 변경 감지용
    }, ensure_ascii=False)

    # 유형 셀렉트
    platform_opts = [
        {"text": {"type": "plain_text", "text": p}, "value": p}
        for p in _VISIT_PLATFORM_OPTIONS
    ]
    platform_initial = {
        "text": {"type": "plain_text", "text": cur_platform},
        "value": cur_platform,
    }

    # 날짜 초기값
    dp_start = {"type": "datepicker", "action_id": "value"}
    if cur_start:
        dp_start["initial_date"] = cur_start
    dp_end = {"type": "datepicker", "action_id": "value"}
    if cur_end:
        dp_end["initial_date"] = cur_end

    blocks = [
        {"type": "section", "text": {"type": "mrkdwn",
            "text": f"*{lead_no}* 정보 수정"}},
    ]
    # 온라인 리드는 유형 dropdown 숨김 — 원본 platform 표시만
    if is_online_lead:
        blocks.append({
            "type": "context",
            "elements": [{
                "type": "mrkdwn",
                "text": f"방문 유형 : `{cur_platform}` (온라인 리드 — 변경 불가)",
            }],
        })
    else:
        blocks.append({
            "type": "input", "block_id": "platform",
            "label": {"type": "plain_text", "text": "방문 유형"},
            "element": {
                "type": "static_select",
                "action_id": "value",
                "options": platform_opts,
                "initial_option": platform_initial,
            },
        })
    blocks.extend([
        {
            "type": "input", "block_id": "visit_date",
            "label": {"type": "plain_text", "text": "방문 예정일 (시작)"},
            "element": dp_start,
        },
        {
            "type": "input", "block_id": "visit_date_end", "optional": True,
            "label": {"type": "plain_text", "text": "방문 예정일 (종료)"},
            "hint": {"type": "plain_text",
                     "text": "범위 방문일 때만 (예: 7/1~7/3)"},
            "element": dp_end,
        },
        {
            "type": "input", "block_id": "name",
            "label": {"type": "plain_text", "text": "이름 / 상호"},
            "element": {
                "type": "plain_text_input", "action_id": "value",
                "initial_value": cur_name,
            },
        },
        {
            "type": "input", "block_id": "phone", "optional": True,
            "label": {"type": "plain_text", "text": "연락처"},
            "hint": {"type": "plain_text",
                     "text": "거래처/소개는 필수, 기타는 선택"},
            "element": {
                "type": "plain_text_input", "action_id": "value",
                "initial_value": cur_phone,
            },
        },
        {
            "type": "input", "block_id": "address",
            "label": {"type": "plain_text", "text": "방문 주소"},
            "element": {
                "type": "plain_text_input", "action_id": "value",
                "initial_value": cur_address,
            },
        },
        {
            "type": "input", "block_id": "consultation", "optional": True,
            "label": {"type": "plain_text", "text": "상담 내용"},
            "element": {
                "type": "plain_text_input", "action_id": "value",
                "multiline": True,
                "initial_value": cur_consultation,
            },
        },
    ])

    client.views_open(trigger_id=trigger_id, view={
        "type": "modal",
        "callback_id": "submit_visit_edit",
        "title": {"type": "plain_text", "text": "정보 수정"},
        "submit": {"type": "plain_text", "text": "저장"},
        "close": {"type": "plain_text", "text": "취소"},
        "private_metadata": metadata,
        "blocks": blocks,
    })


def _build_visit_edit_confirm_view(metadata: dict, state: dict) -> dict:
    """유형 변경 감지 시 확인 modal (submit → view update 로 교체)."""
    lead_no = metadata.get('lead_no', '')
    original_platform = metadata.get('original_platform', '')
    new_platform = _v(state, 'platform') or ''

    # 편집 값 metadata 에 stash (확인 후 재사용)
    new_meta = dict(metadata)
    new_meta['pending_edit'] = {
        'platform': new_platform,
        'visit_start': _v(state, 'visit_date') or '',
        'visit_end': _v(state, 'visit_date_end') or '',
        'name': (_v(state, 'name') or '').strip(),
        'phone': (_v(state, 'phone') or '').strip(),
        'address': (_v(state, 'address') or '').strip(),
        'consultation': (_v(state, 'consultation') or '').strip(),
    }

    body_text = f":warning: *{lead_no}* 방문 유형을 변경합니다.\n\n"
    body_text += f"• 기존: `{original_platform}`\n"
    body_text += f"• 변경: `{new_platform}`\n\n"

    if original_platform == '기타' and new_platform in ('거래처', '소개'):
        body_text += (
            "임시 번호(`ETC-...`) → 정식 리드 번호(`L-XXXXX`) 로 *승격* 됩니다.\n"
            "온라인 리드 시트에 등록되고 대시보드에서 조회 가능."
        )
    elif original_platform in ('거래처', '소개') and new_platform == '기타':
        body_text += (
            "정식 리드 번호(`L-...`) → 임시 번호(`ETC-XXXXX`) 로 *강등* 됩니다.\n"
            "시트에서 삭제되고 대시보드 조회 대상에서 빠집니다."
        )
    else:
        body_text += "방문 유형만 update 됩니다 (리드 번호 형식 불변)."

    return {
        "type": "modal",
        "callback_id": "submit_visit_edit_confirm",
        "title": {"type": "plain_text", "text": "변경 확인"},
        "submit": {"type": "plain_text", "text": "변경 확정"},
        "close": {"type": "plain_text", "text": "취소"},
        "private_metadata": json.dumps(new_meta, ensure_ascii=False),
        "blocks": [
            {"type": "section", "text": {"type": "mrkdwn", "text": body_text}},
        ],
    }


def _process_visit_edit_confirmed(client, body, view) -> None:
    """유형 변경 확인 후 실제 전환 처리. Redis 락으로 동시 실행 방지."""
    metadata = json.loads(view.get('private_metadata') or '{}')
    lead_no = metadata.get('lead_no', '')
    channel = metadata.get('channel', '')
    original_platform = metadata.get('original_platform', '')
    pending = metadata.get('pending_edit', {})
    new_platform = pending.get('platform', '')
    user_id = body.get('user', {}).get('id', '')

    edit_rc = _acquire_edit_lock(lead_no, ttl=30)
    if edit_rc is None:
        _notify_edit_failure(
            client, channel, user_id, lead_no,
            '다른 편집이 진행 중입니다. 30초 후 다시 시도해 주세요.',
        )
        logger.warning(
            f"[VISIT/EDIT] 락 충돌 skip: {lead_no} (editor={user_id})"
        )
        return
    try:
        logger.info(
            f"[VISIT/EDIT] 유형 변경 확정: {lead_no} "
            f"{original_platform}→{new_platform} (editor={user_id})"
        )

        if _is_etc_lead(lead_no) and new_platform in ('거래처', '소개'):
            _convert_etc_to_regular(client, body, lead_no, channel, metadata, pending)
            return

        if not _is_etc_lead(lead_no) and new_platform == '기타':
            _convert_regular_to_etc(client, body, lead_no, channel, metadata, pending)
            return

        logger.warning(
            f"[VISIT/EDIT] 미지원 전환 조합: {lead_no} "
            f"({original_platform}→{new_platform})"
        )
    finally:
        _release_edit_lock(edit_rc, lead_no)


def _convert_etc_to_regular(client, body, lead_no, channel, metadata, pending) -> None:
    """ETC-xxx → L-XXXXX 승격 (시나리오 D).

    시트 A열 (리드 No) + C열 (플랫폼) + 편집 필드만 update. 새 행 add 없음.
    """
    message_ts = metadata.get('message_ts', '')
    user_id = body.get('user', {}).get('id', '')
    new_platform = pending.get('platform', '')
    new_visit_start = pending.get('visit_start', '')
    new_visit_end = pending.get('visit_end', '')
    new_name = pending.get('name', '')
    new_phone = pending.get('phone', '')
    new_address = pending.get('address', '')
    new_consultation = pending.get('consultation', '')

    new_visit_display = _format_visit_date_range(new_visit_start, new_visit_end)
    sheet_visit_value = new_visit_display  # E열 텍스트 서식이라 escape 불필요

    # 시트 update — A열(리드 No) + C(플랫폼) + 편집 필드 + max L- 발번
    try:
        from dashboard.services.lead_service import (
            load_leads_data, _get_sheet_config, get_sheets_manager,
            invalidate_leads_cache,
        )
        import pandas as _pd
        df = load_leads_data(force_refresh=True)
        if df is None or df.empty:
            raise RuntimeError('시트 데이터 로드 실패')
        matches = df[df['리드 No'].astype(str).str.strip() == lead_no]
        if matches.empty:
            raise RuntimeError(f'lead_no {lead_no} 시트에서 못 찾음')
        sheet_row = int(matches.index[0]) + 2

        # 새 L- 발번 (max L- + 1)
        existing_nos = df['리드 No'].astype(str).str.extract(r'L-(\d+)')[0]
        existing_nos = _pd.to_numeric(existing_nos, errors='coerce').dropna()
        next_no_int = int(existing_nos.max()) + 1 if len(existing_nos) > 0 else 1
        new_lead_no = f"L-{next_no_int:05d}"

        cfg = _get_sheet_config()
        manager = get_sheets_manager()
        updates = [
            (f"A{sheet_row}", new_lead_no),        # 리드 No
            (f"C{sheet_row}", new_platform),       # 플랫폼: 기타 → 거래처/소개
            (f"E{sheet_row}", sheet_visit_value),  # 방문 예정일
            (f"F{sheet_row}", new_phone or '-'),   # 연락처
            (f"H{sheet_row}", new_name or '-'),    # 고객명
            (f"I{sheet_row}", new_address or '-'), # 방문 주소
            (f"K{sheet_row}", new_consultation),   # 상담 내용
        ]
        batch = {
            'valueInputOption': 'USER_ENTERED',
            'data': [
                {'range': f"'{cfg['sheet_name']}'!{r}", 'values': [[v]]}
                for r, v in updates
            ],
        }
        manager.service.spreadsheets().values().batchUpdate(
            spreadsheetId=cfg['sheet_id'], body=batch,
        ).execute()
        invalidate_leads_cache()
        logger.info(
            f"[VISIT/EDIT/PROMOTE] 시트 update: {lead_no} → {new_lead_no} "
            f"(row {sheet_row})"
        )
    except Exception as exc:
        logger.error(
            f"[VISIT/EDIT/PROMOTE] 시트 update 실패 ({lead_no}): {exc}",
            exc_info=True,
        )
        _notify_edit_failure(client, channel, user_id, lead_no,
                             f'시트 update 실패: {exc}')
        return

    # 카드 chat_update (헤더 lead_no 갱신)
    try:
        initial = _slack_user_to_initial(client, user_id) or '-'
        body_text, blocks = _build_visit_notice_blocks(
            lead_no=new_lead_no, category_display=new_platform, initial=initial,
            visit_date=new_visit_display,
            name=new_name, contact=new_phone,
            visit_address=new_address, consultation=new_consultation,
        )
        client.chat_update(
            channel=channel, ts=message_ts, text=body_text, blocks=blocks,
        )
    except Exception as exc:
        logger.error(
            f"[VISIT/EDIT/PROMOTE] 카드 update 실패 ({lead_no}→{new_lead_no}): {exc}",
            exc_info=True,
        )

    # List webhook — 옛 ETC 아이템 삭제 + 새 정규 아이템 add
    try:
        _trigger_visit_list_webhook(
            'SLACK_VISIT_CANCEL_WEBHOOK_URL', lead_no, channel, message_ts,
        )
    except Exception as exc:
        logger.warning(f"[VISIT/EDIT/PROMOTE] List 옛 아이템 삭제 실패: {exc}")
    try:
        _post_to_slack_list(
            client, {
                '리드 No': new_lead_no,
                '고객명': new_name, '고객 연락처': new_phone,
                '상담 시간': datetime.now().strftime('%Y.%m.%d. %H:%M'),
                '방문 주소': new_address, '문의 내용': '-',
                '플랫폼': new_platform,
            },
            modal_fields={
                'visit_date': new_visit_display, 'visit_address': new_address,
                'consultation': new_consultation, 'estimate': '',
            },
            channel=channel, message_ts=message_ts, action='visit',
        )
    except Exception as exc:
        logger.warning(f"[VISIT/EDIT/PROMOTE] List 새 아이템 add 실패: {exc}")

    # 감사 로그 답글
    try:
        editor_initial = _slack_user_to_initial(client, user_id) or '-'
        client.chat_postMessage(
            channel=channel, thread_ts=message_ts,
            text=(
                f":arrows_counterclockwise: *리드 번호 승격*: "
                f"`{lead_no}` → `{new_lead_no}` "
                f"(기타 → {new_platform}, 편집자: {editor_initial})"
            ),
            unfurl_links=False,
        )
    except Exception:
        pass

    logger.info(
        f"[VISIT/EDIT/PROMOTE] 완료: {lead_no} → {new_lead_no} "
        f"(기타 → {new_platform}, editor={user_id})"
    )


def _notify_edit_failure(client, channel, user_id, lead_no, reason) -> None:
    """편집 실패 시 매니저에게 ephemeral 알림."""
    if not (user_id and channel):
        return
    try:
        client.chat_postEphemeral(
            channel=channel, user=user_id,
            text=(
                f":x: `{lead_no}` 정보 수정 실패: {reason}\n"
                f"관리자에게 문의하거나 잠시 후 다시 시도해 주세요."
            ),
        )
    except Exception:
        pass


def _acquire_edit_lock(lead_no: str, ttl: int = 30):
    """정보 수정 동시 실행 방지 락. 성공 시 Redis client 반환, 실패 시 None.
    같은 카드에 두 매니저가 동시에 저장 시도할 때 순차 처리 강제.
    """
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        got = rc.set(f'visit_edit_lock:{lead_no}', '1', nx=True, ex=ttl)
        return rc if got else None
    except Exception as exc:
        logger.warning(f'[VISIT/EDIT] 락 획득 실패 — lock 없이 진행: {exc}')
        return None


def _release_edit_lock(rc, lead_no: str) -> None:
    if rc is None:
        return
    try:
        rc.delete(f'visit_edit_lock:{lead_no}')
    except Exception:
        pass


def _convert_regular_to_etc(client, body, lead_no, channel, metadata, pending) -> None:
    """L-XXXXX → ETC-xxx 강등 (시나리오 D).

    시트 A열 (리드 No) + C열 (플랫폼=기타) + 편집 필드만 update. 행 삭제 없음.
    """
    message_ts = metadata.get('message_ts', '')
    user_id = body.get('user', {}).get('id', '')
    new_visit_start = pending.get('visit_start', '')
    new_visit_end = pending.get('visit_end', '')
    new_name = pending.get('name', '')
    new_phone = pending.get('phone', '')
    new_address = pending.get('address', '')
    new_consultation = pending.get('consultation', '')

    new_visit_display = _format_visit_date_range(new_visit_start, new_visit_end)
    sheet_visit_value = new_visit_display  # E열 텍스트 서식이라 escape 불필요

    new_etc_lead_no = _etc_new_lead_no()

    # 시트 update — A열(리드 No=ETC-xxx) + C(플랫폼=기타) + 편집 필드
    try:
        from dashboard.services.lead_service import (
            load_leads_data, _get_sheet_config, get_sheets_manager,
            invalidate_leads_cache,
        )
        df = load_leads_data(force_refresh=True)
        if df is None or df.empty:
            raise RuntimeError('시트 데이터 로드 실패')
        matches = df[df['리드 No'].astype(str).str.strip() == lead_no]
        if matches.empty:
            raise RuntimeError(f'lead_no {lead_no} 시트에서 못 찾음')
        sheet_row = int(matches.index[0]) + 2

        cfg = _get_sheet_config()
        manager = get_sheets_manager()
        updates = [
            (f"A{sheet_row}", new_etc_lead_no),    # 리드 No: L- → ETC-xxx
            (f"C{sheet_row}", '기타'),              # 플랫폼: 정규 → 기타
            (f"E{sheet_row}", sheet_visit_value),  # 방문 예정일
            (f"F{sheet_row}", new_phone or '-'),   # 연락처
            (f"H{sheet_row}", new_name or '-'),    # 고객명
            (f"I{sheet_row}", new_address or '-'), # 방문 주소
            (f"K{sheet_row}", new_consultation),   # 상담 내용
        ]
        batch = {
            'valueInputOption': 'USER_ENTERED',
            'data': [
                {'range': f"'{cfg['sheet_name']}'!{r}", 'values': [[v]]}
                for r, v in updates
            ],
        }
        manager.service.spreadsheets().values().batchUpdate(
            spreadsheetId=cfg['sheet_id'], body=batch,
        ).execute()
        invalidate_leads_cache()
        logger.info(
            f"[VISIT/EDIT/DEMOTE] 시트 update: {lead_no} → {new_etc_lead_no} "
            f"(row {sheet_row})"
        )
    except Exception as exc:
        logger.error(
            f"[VISIT/EDIT/DEMOTE] 시트 update 실패 ({lead_no}): {exc}",
            exc_info=True,
        )
        _notify_edit_failure(client, channel, user_id, lead_no,
                             f'시트 update 실패: {exc}')
        return

    # 3) 카드 chat_update (헤더 lead_no → ETC-xxx)
    try:
        initial = _slack_user_to_initial(client, user_id) or '-'
        body_text, blocks = _build_visit_notice_blocks(
            lead_no=new_etc_lead_no, category_display='기타', initial=initial,
            visit_date=new_visit_display,
            name=new_name, contact=new_phone,
            visit_address=new_address, consultation=new_consultation,
        )
        client.chat_update(
            channel=channel, ts=message_ts, text=body_text, blocks=blocks,
        )
    except Exception as exc:
        logger.error(
            f"[VISIT/EDIT/DEMOTE] 카드 update 실패 ({lead_no}→{new_etc_lead_no}): {exc}",
            exc_info=True,
        )

    # 4) List webhook — 옛 정규 아이템 삭제 + 새 ETC 아이템 add
    try:
        _trigger_visit_list_webhook(
            'SLACK_VISIT_CANCEL_WEBHOOK_URL', lead_no, channel, message_ts,
        )
    except Exception as exc:
        logger.warning(f"[VISIT/EDIT/DEMOTE] List 옛 아이템 삭제 실패: {exc}")

    try:
        new_lead_data = {
            '리드 No': new_etc_lead_no,
            '고객명': new_name,
            '고객 연락처': new_phone,
            '이메일': str(old_lead.get('이메일', '') or '').strip(),
            '상담 시간': datetime.now().strftime('%Y.%m.%d. %H:%M'),
            '방문 주소': new_address,
            '문의 내용': new_consultation,
            '키워드': str(old_lead.get('키워드', '') or '').strip(),
            '플랫폼': '기타',
        }
        _post_to_slack_list(
            client, new_lead_data,
            modal_fields={
                'visit_date': new_visit_display,
                'visit_address': new_address,
                'consultation': new_consultation,
                'estimate': '',
            },
            channel=channel, message_ts=message_ts,
            action='visit',
        )
    except Exception as exc:
        logger.warning(f"[VISIT/EDIT/DEMOTE] List 새 아이템 add 실패: {exc}")

    # 5) 감사 로그 답글
    try:
        editor_initial = _slack_user_to_initial(client, user_id) or '-'
        client.chat_postMessage(
            channel=channel, thread_ts=message_ts,
            text=(
                f":arrows_counterclockwise: *리드 번호 강등*: "
                f"`{lead_no}` → `{new_etc_lead_no}` "
                f"(정규 → 기타, 편집자: {editor_initial})"
            ),
            unfurl_links=False,
        )
    except Exception:
        pass

    logger.info(
        f"[VISIT/EDIT/DEMOTE] 완료: {lead_no} → {new_etc_lead_no} "
        f"(정규 → 기타, editor={user_id})"
    )


def _process_visit_edit(client, body, view) -> None:
    """정보 수정 모달 submit 처리 — 유형 동일 case 필드만 update.
    유형 변경 case 는 handler 에서 confirm view 로 라우팅됨.
    Redis 락으로 동시 편집 방지.
    """
    metadata = json.loads(view.get('private_metadata') or '{}')
    lead_no = metadata.get('lead_no', '')
    channel = metadata.get('channel', '')
    message_ts = metadata.get('message_ts', '')
    original_platform = metadata.get('original_platform', '')
    user_id = body.get('user', {}).get('id', '')

    edit_rc = _acquire_edit_lock(lead_no, ttl=30)
    if edit_rc is None:
        _notify_edit_failure(
            client, channel, user_id, lead_no,
            '다른 편집이 진행 중입니다. 30초 후 다시 시도해 주세요.',
        )
        logger.warning(f"[VISIT/EDIT] 락 충돌 skip: {lead_no} (editor={user_id})")
        return
    try:
        state = view['state']['values']
        new_platform = _v(state, 'platform') or original_platform
        new_visit_start = _v(state, 'visit_date') or ''
        new_visit_end = _v(state, 'visit_date_end') or ''
        new_name = (_v(state, 'name') or '').strip()
        new_phone = (_v(state, 'phone') or '').strip()
        new_address = (_v(state, 'address') or '').strip()
        new_consultation = (_v(state, 'consultation') or '').strip()

        logger.info(
            f"[VISIT/EDIT] {lead_no} 필드 update: 이름={new_name!r}, "
            f"주소={new_address[:30]!r} (editor={user_id})"
        )
        _process_visit_edit_same_platform(
            client, body, lead_no, channel, message_ts,
            new_visit_start, new_visit_end, new_name, new_phone,
            new_address, new_consultation,
        )
    finally:
        _release_edit_lock(edit_rc, lead_no)


def _process_visit_edit_same_platform(client, body, lead_no, channel, message_ts,
                                       new_visit_start, new_visit_end, new_name,
                                       new_phone, new_address, new_consultation) -> None:
    """유형 동일 case 실제 처리 (락 안에서 호출됨)."""
    # 유형 동일 — 필드만 update
    # 방문 예정일 범위 조립
    new_visit_display = _format_visit_date_range(new_visit_start, new_visit_end)

    # E열 셀 서식 '@ 텍스트' 라 escape 불필요
    sheet_visit_value = new_visit_display

    # 방문일 변경 감지용 — old 값 캡처 (2026-07-19)
    old_lead = _find_lead_by_no(lead_no) or {}
    old_visit_date = str(old_lead.get('방문 예정일') or '').strip().lstrip("'")

    updates = {
        '방문 예정일': sheet_visit_value,
        '고객명': new_name,
        '고객 연락처': new_phone,
        '방문 주소': new_address,
        '상담 내용': new_consultation,
    }
    try:
        _update_lead_dispatch(lead_no, updates)
    except Exception as exc:
        logger.error(f"[SLACK/방문수정] update 실패 ({lead_no}): {exc}",
                     exc_info=True)
        return

    # List update webhook
    try:
        _trigger_visit_list_webhook(
            'SLACK_VISIT_MODIFY_WEBHOOK_URL', lead_no, channel, message_ts,
            new_visit_date=new_visit_display,
        )
    except Exception as exc:
        logger.warning(f"[SLACK/방문수정] List update 실패 ({lead_no}): {exc}")

    # 카드 chat_update — 새 값으로 재구성
    try:
        lead = _find_lead_by_no(lead_no) or {}
        platform = str(lead.get('플랫폼', '')).strip()
        if platform in ('거래처', '기타', '소개'):
            category_display = platform
        else:
            category_display = f"온라인({platform})" if platform else '온라인'
        user_id = body['user']['id']
        initial = _slack_user_to_initial(client, user_id) or '-'
        body_text, blocks = _build_visit_notice_blocks(
            lead_no=lead_no, category_display=category_display, initial=initial,
            visit_date=new_visit_display,
            name=new_name, contact=new_phone,
            visit_address=new_address, consultation=new_consultation,
        )
        client.chat_update(
            channel=channel, ts=message_ts, text=body_text, blocks=blocks,
        )
    except Exception as exc:
        logger.error(f"[SLACK/방문수정] 카드 update 실패 ({lead_no}): {exc}",
                     exc_info=True)

    # 방문일 변경 → dm_sent flag 있는 lead 만 담당자에게 알림 (2026-07-19)
    if old_visit_date and new_visit_display and old_visit_date != new_visit_display:
        try:
            from dashboard.services.visit_assignment_sync import send_visit_change_notification
            threading.Thread(
                target=send_visit_change_notification,
                args=(lead_no, old_visit_date, new_visit_display, new_consultation),
                daemon=True,
            ).start()
        except Exception as exc:
            logger.warning(f"[SLACK/방문수정] 변경 알림 예약 실패 ({lead_no}): {exc}")


def _process_visit_complete(client, body) -> None:
    """[✅ 방문 완료] 클릭 처리 — 슬랙 리스트 삭제 + 카드 회색 박스 변환.

    시트 상태는 변경하지 않음. 나중에 프로젝트 등록 시 자동으로 '공사 확정'으로 이동.
    상담 완료 카드와 동일한 패턴 (chat.update로 헤더 갱신 + 원본 코드 블록화).
    """
    lead_no = body["actions"][0].get("value") or ''
    channel = body["channel"]["id"]
    message_ts = body["message"]["ts"]
    user_id = body["user"]["id"]
    if not lead_no:
        return

    # 1) 슬랙 리스트에서 행 삭제 (webhook)
    _trigger_visit_list_webhook(
        'SLACK_VISIT_COMPLETE_WEBHOOK_URL', lead_no, channel, message_ts,
    )

    # 2) 원본 카드 회색 박스 변환 (상담 완료와 동일 양식)
    try:
        initial = _slack_user_to_initial(client, user_id) or '-'
        complete_time = datetime.now().strftime('%m.%d %H:%M')

        # 원본 message 의 section text 추출
        original_text = ''
        for blk in body["message"].get("blocks", []):
            if blk.get("type") == "section":
                bt = blk.get("text", {}).get("text", "")
                if bt:
                    original_text = bt
                    break
        if not original_text:
            original_text = body["message"].get("text", "")

        # `>` blockquote 마커·마크다운 강조·앞뒤 공백 제거
        cleaned_lines = [ln.lstrip('>').lstrip() for ln in original_text.split('\n')]
        cleaned_lines = [ln.replace('*', '') for ln in cleaned_lines]
        clean_text = '\n'.join(cleaned_lines)
        clean_text = re.sub(r'^[\s⠀]+|[\s⠀]+$', '', clean_text)

        header_lines = [
            "⠀",
            f":white_check_mark: *방문 완료*  `{lead_no}`",
            f"처리자 : {initial}",
            f"완료 시간 : {complete_time}",
        ]
        new_text = '\n'.join(header_lines) + f"\n\n```\n{clean_text}\n```"
        new_blocks = [
            {"type": "section", "text": {"type": "mrkdwn", "text": new_text}},
        ]
        client.chat_update(
            channel=channel, ts=message_ts, text=new_text, blocks=new_blocks,
        )
    except Exception as exc:
        logger.warning(f"[SLACK/방문완료] chat.update 실패 ({lead_no}): {exc}")

    # 3) 방문 완료 flag set + dm_sent flag 삭제 + 캔버스 rebuild trigger (2026-07-16).
    #    자동 완료 (사진 첨부) 와 flag 명 통일. 캔버스 필터가 이 flag 로 제외.
    #    dm_sent 는 완료 후 변경 알림 오작동 방지 (2026-07-19).
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        rc.setex(f'visit_auto_completed:{lead_no}', 60 * 60 * 24 * 30, '1')  # 30일
        rc.delete(f'dm_sent:{lead_no}')
    except Exception as exc:
        logger.warning(f"[SLACK/방문완료] flag set 실패 ({lead_no}): {exc}")
    try:
        from dashboard.services.visit_canvas_sync import rebuild_canvas_async
        rebuild_canvas_async()
    except Exception as exc:
        logger.debug(f"[SLACK/방문완료] 캔버스 rebuild trigger 실패 ({lead_no}): {exc}")

    logger.info(f"[SLACK/방문완료] 처리 완료: {lead_no} by {user_id}")


def _open_visit_cancel_reason_modal(client, lead_no: str, channel: str,
                                     message_ts: str, trigger_id: str) -> None:
    """방문 취소 사유 입력 모달 (2026-07-19)."""
    metadata = json.dumps({
        "lead_no": lead_no,
        "channel": channel,
        "message_ts": message_ts,
    })
    view = {
        "type": "modal",
        "callback_id": "submit_visit_cancel_reason",
        "private_metadata": metadata,
        "title": {"type": "plain_text", "text": f"방문 취소"},
        "submit": {"type": "plain_text", "text": "취소 확정"},
        "close": {"type": "plain_text", "text": "닫기"},
        "blocks": [
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": (
                        f":warning: *{lead_no}* 방문을 취소합니다.\n"
                        "_(일주일 이내 미루기는 [정보 수정] 사용)_"
                    ),
                },
            },
            {
                "type": "input",
                "block_id": "reason",
                "label": {"type": "plain_text", "text": "취소 사유"},
                "hint": {
                    "type": "plain_text",
                    "text": "예: 타업체 공사 진행\n예: 다음달로 방문 연기 요청",
                },
                "element": {
                    "type": "plain_text_input",
                    # action_id 는 반드시 "value" — _v(state, block_id) 헬퍼가
                    # state[block_id]["value"] 를 찾도록 통일된 규약 (2026-07-21 정정).
                    "action_id": "value",
                    "multiline": True,
                    "min_length": 2,
                    "max_length": 500,
                },
            },
        ],
    }
    client.views_open(trigger_id=trigger_id, view=view)


def _process_visit_cancel_confirmed(client, body, view) -> None:
    """방문 취소 사유 모달 제출 → 시트 update + 상담 내용 append + 카드 update + v21 알림.

    2026-07-19 신설. 상태 = '방문 취소' (기존 '공사 취소' 대신).
    """
    metadata = json.loads(view.get("private_metadata") or "{}")
    lead_no = metadata.get("lead_no", "")
    channel = metadata.get("channel", "")
    message_ts = metadata.get("message_ts", "")
    user_id = body["user"]["id"]
    if not lead_no:
        return
    if not _try_acquire_action_lock(lead_no, 'cancel'):
        logger.info(f'[SLACK/방문봇] visit_cancel 중복 처리 skip ({lead_no})')
        return

    state = view["state"]["values"]
    reason = (_v(state, "reason") or '').strip()
    if not reason:
        reason = '(사유 미입력)'

    # old 상담 내용 캡처 → append
    old_lead = _find_lead_by_no(lead_no) or {}
    old_note = str(old_lead.get('상담 내용') or '').strip()
    initial = _slack_user_to_initial(client, user_id) or '-'
    cancel_date = datetime.now().strftime('%Y-%m-%d')
    appended_note = (
        f"{old_note}\n─────────\n"
        f"[방문 취소 {cancel_date} {initial}]\n{reason}"
    ).strip()

    # 1) 시트 상태='방문 취소' + 상담 내용 append (ETC- 는 Redis metadata 만 갱신)
    try:
        _update_lead_dispatch(lead_no, {
            '상태': '방문 취소',
            '상담 내용': appended_note,
        })
    except Exception as exc:
        logger.error(f"[SLACK/방문취소] 시트 update 실패 ({lead_no}): {exc}",
                     exc_info=True)

    # 2) 슬랙 List 동기화 — 행 삭제
    _trigger_visit_list_webhook(
        'SLACK_VISIT_CANCEL_WEBHOOK_URL', lead_no, channel, message_ts,
    )

    # 3) 카드 회색 박스 chat.update — 취소 사유도 헤더 표시
    try:
        cancel_time = datetime.now().strftime('%Y.%m.%d. %H:%M')
        # 원본 카드 텍스트 조회 — 완료 카드와 동일한 회색 박스 스타일 위해
        # (2026-07-21: 이전엔 축약 정보만 담아 완료 카드와 비대칭이었음).
        original_text = ''
        try:
            rep = client.conversations_replies(
                channel=channel, ts=message_ts, limit=1, inclusive=True,
            )
            root = (rep.get('messages') or [{}])[0]
            for blk in root.get('blocks', []) or []:
                if blk.get('type') == 'section':
                    bt = (blk.get('text') or {}).get('text', '')
                    if bt:
                        original_text = bt
                        break
            if not original_text:
                original_text = root.get('text', '') or ''
        except Exception as _exc_repl:
            logger.debug(
                f"[SLACK/방문취소] 원본 카드 조회 실패, 재구성 fallback ({lead_no}): {_exc_repl}"
            )
        # 원본 없으면 시트 lead 로 축약 재구성 (fallback)
        if not original_text:
            lead = _find_lead_by_no(lead_no) or {}
            _visit_date = str(lead.get('방문 예정일', '')).strip().lstrip("'") or '-'
            _name = str(lead.get('고객명', '') or '').strip() or '-'
            _contact = str(lead.get('고객 연락처', '') or '').strip() or '-'
            _addr = str(lead.get('방문 주소', '') or '').strip() or '-'
            original_text = (
                f":bell: 방문 일정 취소 — `{lead_no}`\n"
                f"방문일 : {_visit_date}\n"
                f"이름 / 상호 : {_name}\n"
                f"연락처 : {_contact}\n"
                f"방문 주소 : {_addr}"
            )
        # `>` blockquote 마커·마크다운 강조·앞뒤 공백 정리 (완료 flow 와 동일)
        cleaned_lines = [ln.lstrip('>').lstrip() for ln in original_text.split('\n')]
        cleaned_lines = [ln.replace('*', '') for ln in cleaned_lines]
        clean_body = '\n'.join(cleaned_lines)
        clean_body = re.sub(r'^[\s⠀]+|[\s⠀]+$', '', clean_body)
        # shortcode → unicode 변환 (코드 블록 안에서 :bell: 이 이모지로 렌더)
        try:
            from dashboard.blueprints.slack_helpers import _normalize_shortcodes_to_unicode
            clean_body = _normalize_shortcodes_to_unicode(clean_body)
        except Exception:
            pass
        new_text = (
            f"🚫 *방문 취소*  `{lead_no}`\n"
            f"취소한 사람 : {initial}\n"
            f"취소 시간 : {cancel_time}\n"
            f"취소 사유 : {reason}\n"
            f"\n"
            f"```\n{clean_body}\n```"
        )
        new_blocks = [
            {"type": "section", "text": {"type": "mrkdwn", "text": new_text}},
            {
                "type": "actions",
                "elements": [
                    {
                        "type": "button",
                        "text": {"type": "plain_text",
                                 "text": "↩️ 취소 되돌리기", "emoji": True},
                        "style": "primary",
                        "value": lead_no,
                        "action_id": "visit_uncancel",
                        "confirm": {
                            "title": {"type": "plain_text",
                                      "text": "취소 되돌리기"},
                            "text": {"type": "plain_text",
                                     "text": "이 방문 취소를 되돌리시겠습니까?"},
                            "confirm": {"type": "plain_text",
                                        "text": "되돌리기"},
                            "deny": {"type": "plain_text", "text": "닫기"},
                        },
                    },
                ],
            },
        ]
        client.chat_update(
            channel=channel, ts=message_ts, text=new_text, blocks=new_blocks,
        )
    except Exception as exc:
        logger.error(f"[SLACK/방문취소] 카드 update 실패 ({lead_no}): {exc}",
                     exc_info=True)

    # 4) 동행 매니저에게 v21 취소 알림 (취소자 제외)
    try:
        from dashboard.services.visit_assignment_sync import send_visit_cancel_notification
        threading.Thread(
            target=send_visit_cancel_notification,
            args=(lead_no, initial, reason),
            daemon=True,
        ).start()
    except Exception as exc:
        logger.warning(f"[SLACK/방문취소] 알림 예약 실패 ({lead_no}): {exc}")

    # 5) dm_sent flag 삭제 (변경 알림 오작동 방지)
    try:
        from dashboard.utils.redis_client import get_redis_client
        get_redis_client().redis.delete(f'dm_sent:{lead_no}')
    except Exception:
        pass


def _process_visit_cancel(client, body) -> None:
    """(deprecated 2026-07-19) — 사유 입력 없이 즉시 취소하던 옛 flow.
    호환용으로 남김. 신규 clicks 는 _process_visit_cancel_confirmed 로 라우팅.
    """
    lead_no = body["actions"][0].get("value") or ''
    channel = body["channel"]["id"]
    message_ts = body["message"]["ts"]
    user_id = body["user"]["id"]
    if not lead_no:
        return

    # 1) 시트 상태='방문 취소' (2026-07-19 값 변경)
    try:
        _update_lead_dispatch(lead_no, {'상태': '방문 취소'})
    except Exception as exc:
        logger.error(f"[SLACK/방문취소] 시트 update 실패 ({lead_no}): {exc}",
                     exc_info=True)

    # 1-2) 슬랙 List 동기화 워크플로우 — 행 삭제
    _trigger_visit_list_webhook(
        'SLACK_VISIT_CANCEL_WEBHOOK_URL', lead_no, channel, message_ts,
    )

    # 2) 메시지 chat.update — 취소 양식
    try:
        initial = _slack_user_to_initial(client, user_id) or '-'
        cancel_time = datetime.now().strftime('%Y.%m.%d. %H:%M')

        # 원본 메시지의 blocks에서 section text 추출 (text 필드는 줄바꿈 깨질 위험)
        original_text = ''
        for blk in body["message"].get("blocks", []):
            if blk.get("type") == "section":
                bt = blk.get("text", {}).get("text", "")
                if bt:
                    original_text = bt
                    break
        if not original_text:
            original_text = body["message"].get("text", "")

        # 원본의 `>` blockquote 마커 제거 후 코드 블록으로 감싸기 — 흑백 회색 박스 표시
        cleaned_lines = [ln.lstrip('>').lstrip() for ln in original_text.split('\n')]
        # 마크다운 강조(*) 제거 — 코드 블록 안에서는 raw로 보이는 게 깔끔
        cleaned_lines = [ln.replace('*', '') for ln in cleaned_lines]
        clean_text = '\n'.join(cleaned_lines).strip()
        # Slack 저장 정규화로 :bell: 등 shortcode가 남아 코드블록 안에서 텍스트로 보이는 것 방지.
        from dashboard.blueprints.slack_helpers import _normalize_shortcodes_to_unicode
        clean_text = _normalize_shortcodes_to_unicode(clean_text)

        new_text = (
            f"🚫 *방문 취소*  `{lead_no}`\n"
            f"취소한 사람 : {initial}\n"
            f"취소 시간 : {cancel_time}\n"
            f"\n"
            f"```\n{clean_text}\n```"
        )
        new_blocks = [
            {"type": "section", "text": {"type": "mrkdwn", "text": new_text}},
            {
                "type": "actions",
                "elements": [
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "↩️ 취소 되돌리기", "emoji": True},
                        "style": "primary",
                        "value": lead_no,
                        "action_id": "visit_uncancel",
                        "confirm": {
                            "title": {"type": "plain_text", "text": "취소 되돌리기"},
                            "text": {"type": "plain_text",
                                     "text": "이 방문 취소를 되돌리시겠습니까?"},
                            "confirm": {"type": "plain_text", "text": "되돌리기"},
                            "deny": {"type": "plain_text", "text": "닫기"},
                        },
                    },
                ],
            },
        ]
        client.chat_update(
            channel=channel, ts=message_ts, text=new_text, blocks=new_blocks,
        )
    except Exception as exc:
        logger.error(f"[SLACK/방문취소] 메시지 update 실패 ({lead_no}): {exc}",
                     exc_info=True)

    # dm_sent flag 삭제 — 취소 후 변경 알림 오작동 방지 (2026-07-19)
    try:
        from dashboard.utils.redis_client import get_redis_client
        get_redis_client().redis.delete(f'dm_sent:{lead_no}')
    except Exception:
        pass


def _process_visit_thread_files(client, event) -> None:
    """#방문_일정 카드 thread에 첨부된 파일을 구글 드라이브로 업로드.

    폴더명: "{lead_no}_{고객명}_{방문일}" — parent는 GOOGLE_DRIVE_VISIT_FOLDER_ID
    1. thread root 메시지에서 lead_no 추출
    2. lead 정보로 폴더명 생성 (있으면 재사용)
    3. 슬랙 file URL에서 다운로드 → 드라이브 업로드
    4. thread에 답글: "사진 N장 드라이브에 저장 + 폴더 링크"
    """
    channel = event.get("channel", "")
    thread_ts = event.get("thread_ts", "")
    files = event.get("files") or []
    if not channel or not thread_ts or not files:
        return

    # 1) thread root에서 lead_no 추출
    try:
        resp = client.conversations_replies(
            channel=channel, ts=thread_ts, limit=1, inclusive=True,
        )
        msgs = resp.get("messages") or []
        if not msgs:
            return
        root_text = msgs[0].get("text", "")
        for blk in msgs[0].get("blocks", []):
            if blk.get("type") == "section" and \
                    blk.get("text", {}).get("type") == "mrkdwn":
                root_text = blk["text"].get("text", "") + "\n" + root_text
                break
    except Exception as exc:
        logger.warning(f"[SLACK/방문 사진] thread root 조회 실패: {exc}")
        return

    # L-XXXXX (정상 리드) 또는 ETC-xxxxxx (기타 방문 pseudo lead)
    m = re.search(r'L-\d{5}|ETC-[a-f0-9]{6}', root_text)
    if not m:
        logger.info("[SLACK/방문 사진] thread root에 lead_no 없음 — 스킵")
        return
    lead_no = m.group(0)

    # 기타 방문 (ETC-xxx) 은 사후관리/A/S 임시 방문 — 폴더 생성/Drive 저장 skip.
    # 안내 답글도 없음 (사진마다 답글 뜨는 게 노이즈, 카드 헤더의 ETC- 로
    # 매니저는 이미 기타 방문임을 인식).
    if _is_etc_lead(lead_no):
        logger.info(f"[SLACK/방문 사진] {lead_no} 기타 방문 — 폴더 저장 skip (조용히)")
        return

    # Race condition 방어 (2026-07-10)
    # 두 매니저가 동시에 같은 lead thread 에 사진 첨부 시 각 daemon 스레드가 동시에
    # find_or_create_folder 를 호출 → Google Drive eventual consistency 로 폴더 중복 생성.
    # TTL 은 파일 개수에 비례 (2026-07-14) — 대형 현장 50~60장 이슈 대응.
    # 파일당 다운로드 20초 timeout + Drive upload = 최악 ~10초/장, 여유롭게 4초/장.
    from dashboard.utils.redis_client import get_redis_client as _get_rc
    _lock_key = f'visit_photo_lock:{lead_no}'
    _lock_ttl = min(600, 30 + len(files) * 4)  # 1장=34s, 60장=270s, 상한 10분
    # 락 대기 (spin-wait) — 처리 중이면 skip 대신 대기했다가 획득. 매니저가 사진
    # 배치를 연속으로 올리는 UX 지원 (2026-07-15). 최대 5분 대기.
    _rc_lock = None
    _got_lock = False
    try:
        _rc_lock = _get_rc().redis
        _wait_start = time.time()
        _max_wait = 300  # 5분
        while True:
            _got_lock = _rc_lock.set(_lock_key, '1', nx=True, ex=_lock_ttl)
            if _got_lock:
                break
            if time.time() - _wait_start > _max_wait:
                logger.warning(
                    f'[SLACK/방문 사진] {lead_no} 락 대기 {_max_wait}s 초과 — skip'
                )
                return
            time.sleep(2)
        _wait_elapsed = time.time() - _wait_start
        if _wait_elapsed > 1:
            logger.info(
                f'[SLACK/방문 사진] {lead_no} 락 대기 {_wait_elapsed:.0f}s 후 획득'
            )
    except Exception as exc:
        logger.warning(f'[SLACK/방문 사진] 락 획득 실패 — 계속 진행: {exc}')
        _rc_lock = None

    # 2) 폴더명 생성 — "({이니셜}) {방문 주소} {YY.MM.DD}"
    lead = _find_lead_by_no(lead_no) or {}
    # 이니셜 — 플랫폼별 규칙 (2026-07-20 개선):
    #   거래처/기타/소개: 등록자(M열=온라인 상담자) 우선 — 대표님이 등록하고
    #     매니저 대신 보낸 경우도 폴더는 등록자 이니셜 (관리 주체 개념).
    #     본인 방문 필수는 M→N 자동 복사돼 동일 이니셜이므로 무관.
    #   온라인(그 외): List 배정 담당자(N열=영업 담당자) 우선.
    #     미배정 즉시 방문 케이스 (JK 상담 후 근처 MJ 즉시 방문) 는 N열 없으므로
    #     → 사진 업로더(event.user) 를 실제 방문자로 취급 → 업로더 이니셜.
    #     그것도 실패 시 M열(온라인 상담자) fallback.
    #   최종 fallback: 카드 "등록자 :" 정규식 → '미상'
    def _clean(v):
        s = str(v or '').strip()
        return '' if s in ('', '-', '미정') else s

    lead_platform = str(lead.get('플랫폼', '')).strip()
    _is_partner = lead_platform in ('거래처', '기타', '소개')
    if _is_partner:
        source_name = _clean(lead.get('온라인 상담자'))
    else:
        source_name = _clean(lead.get('영업 담당자'))
    initial = _to_initial(source_name) if source_name else ''

    # 업로더 fallback (event.user) — 온라인 미배정 or 거래처 M열 미기재 케이스.
    # 온라인 즉시 방문 (당일 근처 매니저가 상담 없이 다녀오는) 시나리오가 주 대상.
    if not initial:
        uploader_id = (event.get('user') or '').strip() if isinstance(event, dict) else ''
        if uploader_id:
            try:
                uploader_ini = _slack_user_to_initial(client, uploader_id)
                if uploader_ini and uploader_ini != '-':
                    initial = uploader_ini
            except Exception as exc:
                logger.debug(f'[SLACK/방문 사진] 업로더 이니셜 조회 실패: {exc}')

    # 온라인 케이스에서 N열도 업로더도 실패한 경우에만 M열(상담자) fallback
    if not initial and not _is_partner:
        m_source = _clean(lead.get('온라인 상담자'))
        if m_source:
            initial = _to_initial(m_source)

    # 최종 fallback: 카드 텍스트 파싱
    if not initial:
        m_ini = re.search(r'등록자\s*:\s*([A-Za-z가-힣]+)', root_text)
        if m_ini:
            initial = _to_initial(m_ini.group(1).strip())
    initial = initial or '미상'

    visit_address = str(lead.get('방문 주소', '') or '').strip()
    if not visit_address or visit_address == '-':
        visit_address = '주소 미상'

    today_str = datetime.now().strftime('%y.%m.%d')
    # 플랫폼 prefix — 홈페이지/전화/카카오톡/채널톡은 디폴트(없음, 광고 플랫폼 X)
    # 그 외(당근/거래처/숨고/기타)는 prefix 추가
    platform = str(lead.get('플랫폼', '') or '').strip()
    _DEFAULT_PLATFORMS = {'홈페이지', '전화', '카카오톡', '채널톡'}
    prefix = f"{platform} " if (platform and platform not in _DEFAULT_PLATFORMS) else ''

    # 사진 caption(메시지 text) = 위치(서브폴더) 용도. 상호명은 별도 답글로만.
    # 예: "1층", "2층 휴게실"
    caption = (event.get('text') or '').strip()
    location = ''
    if caption and '\n' not in caption and len(caption) <= 30:
        location = re.sub(r'[\\/:*?"<>|]', '', caption).strip()

    folder_name = f"{prefix}({initial}) {visit_address} {today_str}"
    # 폴더명에 사용 불가한 문자 정리
    folder_name = re.sub(r'[\\/:*?"<>|]', '', folder_name).strip()

    # 3) 루트 폴더 안에 lead 폴더 + 그 안에 '현장사진' 서브폴더
    parent_id = os.getenv('GOOGLE_DRIVE_VISIT_FOLDER_ID', '').strip()
    if not parent_id:
        logger.warning("[SLACK/방문 사진] GOOGLE_DRIVE_VISIT_FOLDER_ID 미설정")
        return

    from dashboard.utils.google_drive import find_or_create_folder, upload_file
    lead_folder = find_or_create_folder(folder_name, parent_id)
    if not lead_folder:
        logger.error(f"[SLACK/방문 사진] lead 폴더 생성/조회 실패: {folder_name}")
        return
    photo_folder = find_or_create_folder('현장사진', lead_folder['id'])
    if not photo_folder:
        logger.error(f"[SLACK/방문 사진] '현장사진' 서브폴더 생성/조회 실패")
        return

    # caption(위치) 있으면 현장사진/{위치}/ 서브폴더 생성, 없으면 현장사진/ 직접
    if location:
        location_folder = find_or_create_folder(location, photo_folder['id'])
        if location_folder:
            folder_id = location_folder['id']
        else:
            folder_id = photo_folder['id']
    else:
        folder_id = photo_folder['id']
    folder_link = lead_folder.get('webViewLink', '')

    bot_token = os.getenv('SLACK_VISIT_BOT_TOKEN', '').strip()
    if not bot_token:
        bot_token = os.getenv('SLACK_BOT_TOKEN', '').strip()

    # 방문 사진 최대 100MB (동영상까지 고려). 대용량 파일 다운로드는 메모리·타임아웃 위험.
    _MAX_PHOTO_BYTES = 100 * 1024 * 1024

    # 진행 답글 (2026-07-14) — 대형 현장 50~60장 시 몇 분간 침묵 → 매니저 재업로드 사고 방지.
    # 4장 이상만 켜서 소량 배치 노이즈 억제. 완료 시 이 메시지를 최종 요약으로 update.
    _total = len(files)
    _progress_ts = None
    if _total >= 4:
        try:
            _progress_resp = client.chat_postMessage(
                channel=channel, thread_ts=thread_ts,
                text=f":hourglass_flowing_sand: 사진 저장 중... (0/{_total})",
                unfurl_links=False,
            )
            _progress_ts = _progress_resp.get('ts')
        except Exception as exc:
            logger.warning(f"[SLACK/방문 사진] 진행 답글 post 실패: {exc}")

    uploaded = 0
    skipped_oversize = 0
    failed = 0
    for _idx, f in enumerate(files, 1):
        download_url = f.get('url_private_download') or f.get('url_private')
        if not download_url:
            failed += 1
            continue
        filename = f.get('name') or f.get('title') or f'photo_{f.get("id","unknown")}.jpg'
        mimetype = f.get('mimetype') or 'application/octet-stream'

        # 사전 크기 차단: Slack file object 의 size 필드 (bytes)
        size_hint = f.get('size') or 0
        if size_hint and size_hint > _MAX_PHOTO_BYTES:
            logger.warning(
                f"[SLACK/방문 사진] 파일 크기 초과 skip: "
                f"{filename} = {size_hint / 1024 / 1024:.1f}MB > 100MB"
            )
            skipped_oversize += 1
        else:
            try:
                req = urllib.request.Request(
                    download_url,
                    headers={'Authorization': f'Bearer {bot_token}'},
                )
                with urllib.request.urlopen(req, timeout=20) as r:
                    # Content-Length 이차 차단
                    try:
                        length = int(r.headers.get('Content-Length', '0'))
                    except (TypeError, ValueError):
                        length = 0
                    _too_big = False
                    if length and length > _MAX_PHOTO_BYTES:
                        logger.warning(
                            f"[SLACK/방문 사진] Content-Length 초과 skip: "
                            f"{filename} = {length / 1024 / 1024:.1f}MB"
                        )
                        skipped_oversize += 1
                        _too_big = True
                        content = None
                    else:
                        # 스트림 상한
                        content = r.read(_MAX_PHOTO_BYTES + 1)
                        if len(content) > _MAX_PHOTO_BYTES:
                            logger.warning(
                                f"[SLACK/방문 사진] 스트림 크기 초과 skip: "
                                f"{filename} > 100MB"
                            )
                            skipped_oversize += 1
                            _too_big = True
                if not _too_big and content is not None:
                    if upload_file(folder_id, filename, content, mimetype=mimetype):
                        uploaded += 1
                    else:
                        failed += 1
            except Exception as exc:
                failed += 1
                logger.error(
                    f"[SLACK/방문 사진] 다운로드/업로드 실패 ({filename}): {exc}",
                    exc_info=True,
                )

        # 진행 업데이트 — 5장마다 or 마지막 파일
        if _progress_ts and (_idx % 5 == 0 or _idx == _total):
            try:
                client.chat_update(
                    channel=channel, ts=_progress_ts,
                    text=(
                        f":hourglass_flowing_sand: 사진 저장 중... ({_idx}/{_total})"
                    ),
                )
            except Exception:
                pass  # 진행 표시 실패는 무시

    if uploaded == 0:
        # 최종 실패 안내 — 진행 답글이 있으면 그걸로 update, 없으면 새 답글.
        _reason_bits = []
        if failed:
            _reason_bits.append(f"실패 {failed}장")
        if skipped_oversize:
            _reason_bits.append(f"크기 초과 스킵 {skipped_oversize}장")
        _reason = ' / '.join(_reason_bits) or '알 수 없는 원인'
        _err_text = (
            f":x: 사진 저장 실패 — 모두 처리되지 않았습니다 ({_reason}).\n"
            f"잠시 후 다시 시도해 주세요."
        )
        try:
            if _progress_ts:
                client.chat_update(channel=channel, ts=_progress_ts, text=_err_text)
            else:
                client.chat_postMessage(
                    channel=channel, thread_ts=thread_ts, text=_err_text,
                    unfurl_links=False,
                )
        except Exception as exc:
            logger.warning(f"[SLACK/방문 사진] 실패 답글 전송 실패: {exc}")
        return

    # 4) thread → folder 매핑 저장 (상호명 답글로 폴더명 갱신용, TTL 30일)
    #    + lead_no → folder_id 역인덱스 (프로젝트 등록 모달 '리드 불러오기' 자동 채움용)
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        key = f"visit_thread:{channel}:{thread_ts}"
        rc.hset(key, mapping={
            'lead_folder_id': lead_folder['id'],
            'photo_folder_id': photo_folder['id'],
            'prefix': prefix,
            'initial': initial,
            'address': visit_address,
            'date': today_str,
            'shop_name': '',
        })
        rc.expire(key, 60 * 60 * 24 * 30)
        # 역인덱스 — lead_no 로 폴더 조회 (프로젝트 등록 시 자동 채움)
        # 프로젝트 등록까지 여유롭게 180일 TTL (몇 달 뒤 확정 케이스 대응)
        rc.set(f"visit_folder:{lead_no}", lead_folder['id'], ex=60 * 60 * 24 * 180)
    except Exception as exc:
        logger.warning(f"[SLACK/방문 사진] Redis 매핑 저장 실패: {exc}")

    # 4-2) 리드 시트 P열 (폴더 ID) 영구 저장 — Redis TTL 만료 대비 + source of truth.
    # 정상 리드는 시트 P열, ETC- 는 Redis metadata 에 저장 (기타 방문은 후속 조회 편의만).
    try:
        _update_lead_dispatch(lead_no, {'폴더 ID': lead_folder['id']})
        logger.info(f"[SLACK/방문 사진] 폴더 ID 저장: {lead_no} → {lead_folder['id']}")
    except Exception as exc:
        logger.warning(f"[SLACK/방문 사진] 폴더 ID 저장 실패 ({lead_no}): {exc}")

    # 5) thread 답글 (debounce)
    # 매니저가 배치 여러 번 올릴 때 마지막 배치 이후 15초 조용하면 그때 발송.
    # 그 사이 새 배치 도착 시 tick 갱신 → 이번 답글 skip → 마지막 배치가 발송.
    # 사진 upload 완료가 봇 event 도착보다 느려서 답글이 사진들 사이에 끼는
    # UX 방지 (2026-07-15).
    _DEBOUNCE_SEC = 15
    try:
        from dashboard.utils.redis_client import get_redis_client as _get_rc_deb
        _rc_deb = _get_rc_deb().redis
        _debounce_key = f'photo_reply_debounce:{lead_no}'
        _my_tick = f"{time.time():.6f}"
        _rc_deb.set(_debounce_key, _my_tick, ex=60)
        time.sleep(_DEBOUNCE_SEC)
        _current = _rc_deb.get(_debounce_key)
        _current_str = _current.decode() if isinstance(_current, bytes) else (_current or '')
        if _current_str != _my_tick:
            logger.info(
                f"[SLACK/방문 사진] {lead_no} 다른 배치가 이후 도착 — 이번 답글 skip"
            )
            # 진행 답글 (⏳ K/N) 이 남아있으면 삭제 (다음 배치가 최종 답글 발송)
            if _progress_ts:
                try:
                    client.chat_delete(channel=channel, ts=_progress_ts)
                except Exception:
                    pass
            # 자동 완료 처리도 skip (마지막 배치가 담당)
            try:
                if _rc_lock:
                    _rc_lock.delete(_lock_key)
            except Exception:
                pass
            return
    except Exception as exc:
        logger.warning(f"[SLACK/방문 사진] debounce 실패, 즉시 답글 진행: {exc}")

    try:
        location_suffix = f" → 현장사진/{location}" if location else ''
        # 윈도우 탐색기 경로 안내 (구글 드라이브 데스크톱 앱 동기화 경로)
        # 트리플 백틱 코드 블록 → 슬랙 데스크톱 앱 호버 시 복사 버튼 자동 표시
        win_base = os.getenv('GOOGLE_DRIVE_WINDOWS_BASE_PATH', '').strip()
        win_path_line = ''
        if win_base:
            win_path_line = (
                f"\n💻 *탐색기 경로* :\n"
                f"```{win_base}\\{folder_name}```"
            )
        # 부분 실패 표기 (일부 성공, 일부 실패/스킵)
        _partial_bits = []
        if failed:
            _partial_bits.append(f":x: 실패 {failed}장")
        if skipped_oversize:
            _partial_bits.append(f":warning: 스킵 {skipped_oversize}장 (100MB 초과)")
        _partial_line = ('\n' + ' · '.join(_partial_bits)) if _partial_bits else ''

        # 누적 카운트 — 같은 스레드 이전 배치 답글이 있으면 그 답글 update
        # (매니저가 10장씩 나눠 여러 배치 올려도 답글은 하나만 유지, 카운트 누적)
        _cumul_key = f'visit_photo_reply:{channel}:{thread_ts}'
        _prev_ts = ''
        _cumul_uploaded = uploaded
        try:
            from dashboard.utils.redis_client import get_redis_client
            _rc_reply = get_redis_client().redis
            _prev_raw = _rc_reply.hgetall(_cumul_key) or {}
            _prev = {
                (k.decode() if isinstance(k, bytes) else k):
                (v.decode() if isinstance(v, bytes) else v)
                for k, v in _prev_raw.items()
            }
            _prev_ts = _prev.get('ts', '')
            if _prev_ts:
                try:
                    _cumul_uploaded = int(_prev.get('count', '0')) + uploaded
                except ValueError:
                    _cumul_uploaded = uploaded
        except Exception as exc:
            logger.warning(f"[SLACK/방문 사진] 누적 카운트 조회 실패: {exc}")
            _rc_reply = None

        reply_text = (
            f":file_folder: 사진 {_cumul_uploaded}장을 드라이브에 저장했습니다{location_suffix}.\n"
            f"📁 {folder_name}"
            f"{_partial_line}"
            f"{win_path_line}\n"
            f":id: *폴더 ID* (새 프로젝트 등록용) :\n"
            f"```{lead_folder['id']}```\n"
            f">*상호명 추가* : 답글에 \"상호 OOO\" 입력\n"
            f">*위치 분류* : 사진 첨부시 댓글에 \"1층\" 등 함께 입력\n"
            f">*추가 사진* : 슬랙 UI 상 10장씩 나눠 올려야 하며, 같은 폴더에 이어서 저장됩니다"
        )

        # 우선순위: 이전 배치 답글 > 이번 진행 답글 > 새 답글
        _final_reply_ts = ''
        _sent = False
        if _prev_ts:
            try:
                client.chat_update(channel=channel, ts=_prev_ts, text=reply_text)
                _final_reply_ts = _prev_ts
                _sent = True
                # 진행 답글이 있으면 삭제 (누적 답글로 통합)
                if _progress_ts and _progress_ts != _prev_ts:
                    try:
                        client.chat_delete(channel=channel, ts=_progress_ts)
                    except Exception:
                        pass
            except Exception as exc:
                logger.warning(
                    f"[SLACK/방문 사진] 이전 답글 update 실패, 새 답글로 fallback: {exc}"
                )
        if not _sent and _progress_ts:
            try:
                client.chat_update(channel=channel, ts=_progress_ts, text=reply_text)
                _final_reply_ts = _progress_ts
                _sent = True
            except Exception as exc:
                logger.warning(
                    f"[SLACK/방문 사진] 진행 답글 update 실패, 새 답글로 fallback: {exc}"
                )
        if not _sent:
            try:
                resp = client.chat_postMessage(
                    channel=channel, thread_ts=thread_ts, text=reply_text,
                    unfurl_links=False,
                )
                if resp and resp.get('ts'):
                    _final_reply_ts = resp['ts']
            except Exception as exc:
                logger.warning(
                    f"[SLACK/방문 사진] 새 답글 발송 실패: {exc}"
                )

        # 누적 상태 Redis 저장 (다음 배치에서 이 답글 update)
        if _final_reply_ts and _rc_reply is not None:
            try:
                _rc_reply.hset(_cumul_key, mapping={
                    'ts': _final_reply_ts,
                    'count': str(_cumul_uploaded),
                })
                _rc_reply.expire(_cumul_key, 60 * 60 * 24 * 30)
            except Exception as exc:
                logger.warning(f"[SLACK/방문 사진] 누적 카운트 저장 실패: {exc}")
    except Exception as exc:
        logger.warning(f"[SLACK/방문 사진] thread 답글 실패: {exc}")

    # 6) 자동 방문 완료 처리 — 사진 첨부 = 방문 다녀옴 = 완료 (매니저 UX).
    # Redis flag 로 중복 방지 (여러 배치 첨부 시 첫 배치에만 완료 트리거).
    try:
        _auto_complete_visit_from_photo(
            client, channel=channel, thread_ts=thread_ts, lead_no=lead_no,
            event_user_id=event.get('user', ''),
        )
    except Exception as exc:
        logger.warning(f"[SLACK/방문 사진] 자동 방문 완료 처리 실패 ({lead_no}): {exc}")

    # 7) 락 해제 — 이후 첨부는 즉시 처리 가능하도록 (자연 만료 60초 대신 즉시)
    try:
        if _rc_lock:
            _rc_lock.delete(_lock_key)
    except Exception:
        pass


def _auto_complete_visit_from_photo(client, channel, thread_ts, lead_no,
                                      event_user_id) -> None:
    """사진 첨부 → 폴더 생성 완료 후 카드를 자동으로 [방문 완료] 처리.

    - Redis flag `visit_auto_completed:{lead_no}` (TTL 30일) 로 중복 방지
      (첫 배치 완료 시만 완료 처리, 이후 배치는 사진만 저장)
    - List 삭제 웹훅 호출 (SLACK_VISIT_COMPLETE_WEBHOOK_URL)
    - 원본 카드 회색 처리 (chat.update)
    """
    # 중복 방지 flag
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        flag_key = f'visit_auto_completed:{lead_no}'
        if rc.get(flag_key):
            logger.debug(f"[SLACK/자동완료] {lead_no} 이미 완료 처리됨 — skip")
            return
    except Exception as exc:
        logger.warning(f"[SLACK/자동완료] flag 조회 실패 — 계속 진행: {exc}")
        rc = None

    # 1) List 삭제 웹훅
    try:
        _trigger_visit_list_webhook(
            'SLACK_VISIT_COMPLETE_WEBHOOK_URL', lead_no, channel, thread_ts,
        )
    except Exception as exc:
        logger.warning(f"[SLACK/자동완료] List 웹훅 실패 ({lead_no}): {exc}")

    # 2) 원본 카드 회색 처리 — conversations.replies 로 root 메시지 blocks 재조회
    try:
        rep = client.conversations_replies(
            channel=channel, ts=thread_ts, limit=1, inclusive=True,
        )
        root = (rep.get('messages') or [{}])[0]
        original_text = ''
        for blk in root.get('blocks', []) or []:
            if blk.get('type') == 'section':
                bt = (blk.get('text') or {}).get('text', '')
                if bt:
                    original_text = bt
                    break
        if not original_text:
            original_text = root.get('text', '')

        cleaned_lines = [ln.lstrip('>').lstrip() for ln in original_text.split('\n')]
        cleaned_lines = [ln.replace('*', '') for ln in cleaned_lines]
        clean_text = '\n'.join(cleaned_lines)
        clean_text = re.sub(r'^[\s⠀]+|[\s⠀]+$', '', clean_text)

        initial = _slack_user_to_initial(client, event_user_id) or '-'
        complete_time = datetime.now().strftime('%m.%d %H:%M')
        header_lines = [
            "⠀",
            f":white_check_mark: *방문 완료 (사진 첨부 자동)*  `{lead_no}`",
            f"처리자 : {initial}",
            f"완료 시간 : {complete_time}",
        ]
        new_text = '\n'.join(header_lines) + f"\n\n```\n{clean_text}\n```"
        new_blocks = [
            {"type": "section", "text": {"type": "mrkdwn", "text": new_text}},
        ]
        client.chat_update(
            channel=channel, ts=thread_ts, text=new_text, blocks=new_blocks,
        )
        logger.info(f"[SLACK/자동완료] {lead_no} 카드 완료 처리 (처리자={initial})")
    except Exception as exc:
        logger.warning(f"[SLACK/자동완료] 카드 update 실패 ({lead_no}): {exc}")

    # 3) flag set (TTL 30일) + dm_sent flag 삭제 (2026-07-19)
    if rc is not None:
        try:
            rc.set(flag_key, '1', ex=60 * 60 * 24 * 30)
            rc.delete(f'dm_sent:{lead_no}')
        except Exception:
            pass


def _process_visit_shop_name_update(client, event) -> None:
    """thread 답글의 `상호 XXX` / `상호명 XXX` 패턴 감지 → 드라이브 폴더명 갱신."""
    channel = event.get("channel", "")
    thread_ts = event.get("thread_ts", "")
    text = (event.get("text") or '').strip()
    if not channel or not thread_ts or not text:
        return

    m = re.match(r'^상호명?\s+(.+)$', text)
    if not m:
        return
    shop_name = m.group(1).strip()
    if not shop_name:
        return
    # 파일명 사용 불가 문자 정리
    shop_name = re.sub(r'[\\/:*?"<>|]', '', shop_name).strip()
    if not shop_name:
        return

    # Redis에서 thread 매핑 조회
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        key = f"visit_thread:{channel}:{thread_ts}"
        info = rc.hgetall(key)
        logger.info(
            f"[SLACK/방문 사진] 상호명 갱신 시도 key={key} found={bool(info)} shop={shop_name!r}"
        )
    except Exception as exc:
        logger.warning(f"[SLACK/방문 사진] Redis 매핑 조회 실패: {exc}")
        return
    if not info:
        # 사진 업로드 없이 상호명만 답글 — 무시
        logger.info(f"[SLACK/방문 사진] Redis 매핑 없음 — 사진 업로드 먼저 필요")
        return
    # Redis bytes → str (decode_responses=True 면 이미 str)
    info = {(k.decode() if isinstance(k, bytes) else k):
            (v.decode() if isinstance(v, bytes) else v) for k, v in info.items()}

    lead_folder_id = info.get('lead_folder_id', '')
    if not lead_folder_id:
        return

    prefix = info.get('prefix', '')
    initial = info.get('initial', '')
    address = info.get('address', '')
    date_str = info.get('date', '')
    new_folder_name = f"{prefix}({initial}) {address} {shop_name} {date_str}"
    new_folder_name = re.sub(r'[\\/:*?"<>|]', '', new_folder_name).strip()

    from dashboard.utils.google_drive import rename_folder
    if not rename_folder(lead_folder_id, new_folder_name):
        return

    # Redis 갱신
    try:
        rc.hset(key, 'shop_name', shop_name)
    except Exception:
        pass

    # thread 답글
    try:
        client.chat_postMessage(
            channel=channel, thread_ts=thread_ts,
            text=(
                f":pencil2: 폴더명에 상호명이 추가되었습니다.\n"
                f":file_folder: {new_folder_name}"
            ),
            unfurl_links=False,
        )
    except Exception as exc:
        logger.warning(f"[SLACK/방문 사진] 상호명 답글 실패: {exc}")


def _process_visit_uncancel(client, body) -> None:
    """[↩️ 취소 되돌리기] 클릭 처리 — 시트 상태 복원 + 카드 원본 양식 복원 + list 복구 webhook."""
    lead_no = body["actions"][0].get("value") or ''
    channel = body["channel"]["id"]
    message_ts = body["message"]["ts"]
    user_id = body["user"]["id"]
    if not lead_no:
        return

    # 1) 시트 상태 → '방문 예약' 복원 (ETC- 는 Redis metadata 만 갱신)
    try:
        _update_lead_dispatch(lead_no, {'상태': '방문 예약'})
    except Exception as exc:
        logger.error(f"[SLACK/방문복원] 시트 update 실패 ({lead_no}): {exc}",
                     exc_info=True)

    # 2) list 복구 webhook 호출
    _trigger_visit_list_webhook(
        'SLACK_VISIT_RESTORE_WEBHOOK_URL', lead_no, channel, message_ts,
    )

    # 3) 메시지 chat.update — 시트 lead 정보로 원본 양식 재구성
    try:
        lead = _find_lead_by_no(lead_no) or {}
        platform = str(lead.get('플랫폼', '')).strip()
        # 카테고리는 인입 lead면 "온라인" (lead_no가 정상 카카오톡/홈페이지/당근/전화이면)
        if platform in ('전화', '거래처', '기타'):
            category = platform
            category_display = category
        else:
            category = '온라인'
            category_display = f"{category}({platform})" if platform else category

        initial = _slack_user_to_initial(client, user_id) or '-'
        visit_date_raw = str(lead.get('방문 예정일', '') or '').strip()
        if visit_date_raw.startswith("'"):
            visit_date_raw = visit_date_raw[1:]
        body_text, blocks = _build_visit_notice_blocks(
            lead_no=lead_no, category_display=category_display, initial=initial,
            visit_date=visit_date_raw,
            name=str(lead.get('고객명', '') or '').strip(),
            contact=str(lead.get('고객 연락처', '') or '').strip(),
            visit_address=str(lead.get('방문 주소', '') or '').strip(),
            consultation=str(lead.get('상담 내용', '') or '').strip(),
        )
        client.chat_update(
            channel=channel, ts=message_ts, text=body_text, blocks=blocks,
        )
    except Exception as exc:
        logger.error(f"[SLACK/방문복원] 메시지 복원 실패 ({lead_no}): {exc}",
                     exc_info=True)


# (헬퍼 함수 정의: slack_helpers.py로 이동 — _format_date_for_sheet/_v/_v_multi/
#  _to_initial/_slack_user_to_korean_name/_slack_user_to_initial/SALES_INITIALS)


# ─────────────────────────────────────────────────────────────
# 전화 문의 — 슬랙 모달 입력으로 시트 등록 + 조건부 슬랙 알림
# ─────────────────────────────────────────────────────────────
_PHONE_DEVICE_OPTIONS = [
    "천장형", "스탠드", "매립덕트", "벽걸이", "FCU", "전열교환기", "세척",
    "가정용",  # 드랍 사유 추적용 — 가정용은 취급 X
]
_PHONE_STATUS_OPTIONS = [
    ("유선 상담", "유선 상담 (시트 등록)"),
    ("문의 드랍", "문의 드랍 (시트 등록)"),
    ("방문 예약", "방문 예약 (시트 등록 + 슬랙 알림)"),
    ("견적 제출", "견적 제출 (시트 등록 + 슬랙 알림)"),
]



def _post_to_slack_list(client, lead: dict, modal_fields: dict, channel: str,
                        message_ts: str, action: str) -> bool:
    """슬랙 List 워크플로우 webhook 호출 — 모달 제출 시 자동 등록.

    Args:
        lead: 시트 행 dict (고객명/연락처/이메일/방문주소/상담시간 등)
        modal_fields: 모달 입력 dict (visit_date, visit_address, consultation, estimate)
        channel: 슬랙 채널 ID
        message_ts: 원본 메시지 ts (영구 링크용)
        action: 'visit' or 'price'
    """
    add_url = os.getenv("SLACK_LIST_WEBHOOK_URL", "").strip()
    update_url = os.getenv("SLACK_LIST_UPDATE_WEBHOOK_URL", "").strip()
    if not add_url:
        logger.debug("[SLACK/LIST] SLACK_LIST_WEBHOOK_URL 미설정 - 등록 스킵")
        return False

    # 같은 lead 첫 호출 / 재호출 판정 — Redis dedup
    # 첫 호출 → add 워크플로우 (list에 행 추가)
    # 재호출 → update 워크플로우 (같은 lead 행 갱신, contact 매칭)
    lead_no = str(lead.get('리드 No', '') or '').strip()
    is_first = True
    if lead_no:
        try:
            from dashboard.utils.redis_client import get_redis_client
            rc = get_redis_client().redis
            dedup_key = f"slack_list_posted:{lead_no}"
            # setnx — 처음만 True, 재호출은 False
            is_first = bool(rc.set(dedup_key, '1', ex=60 * 60 * 24 * 90, nx=True))
        except Exception as exc:
            logger.warning(f"[SLACK/LIST] dedup 체크 실패 ({lead_no}): {exc}")

    if is_first:
        webhook_url = add_url
        op = 'add'
    else:
        if not update_url:
            logger.info(
                f"[SLACK/LIST] 재제출이지만 SLACK_LIST_UPDATE_WEBHOOK_URL 미설정 — skip ({lead_no})"
            )
            return False
        webhook_url = update_url
        op = 'update'

    # 메시지 영구 링크 (channel/ts 둘 다 있을 때만 — 워크플로 form 흐름은 빈 상태)
    message_link = ''
    if channel and message_ts:
        try:
            permalink = client.chat_getPermalink(
                channel=channel, message_ts=message_ts,
            )
            message_link = permalink.get("permalink", "")
        except Exception:
            pass

    # 상담 내용 파싱 (장소/기기/문의)
    parts = _split_lead_content(str(lead.get('문의 내용', '') or lead.get('상담 내용', '')))

    # visit_type — Slack List 방문 유형 컬럼용 3 카테고리 (온라인/거래처/기타)
    # 거래처·소개 → 거래처, 기타 → 기타, 나머지(전화·홈페이지·카카오톡·당근·채널톡·숨고·큐플레이스·메일 등) → 온라인
    lead_platform = str(lead.get('플랫폼') or '').strip()
    if lead_platform in ('거래처', '소개'):
        visit_type_category = '거래처'
    elif lead_platform == '기타':
        visit_type_category = '기타'
    else:
        visit_type_category = '온라인'
    # 방문 예정일 — 범위 양식이면 (시작, 종료) ISO로 분리 + 합쳐진 표시 양식도 함께 전달
    visit_date_raw = str(lead.get('방문 예정일') or '').strip()
    vd_start_iso, vd_end_iso = _split_visit_date_range(visit_date_raw)
    payload = {
        "lead_no": lead_no or '-',
        "platform": lead_platform or '-',
        "visit_type": visit_type_category,
        "name": str(lead.get('고객명') or '').strip() or '-',
        "contact": str(lead.get('고객 연락처') or '').strip() or '-',
        "email": str(lead.get('이메일') or '').strip() or '-',
        "inquiry_time": str(lead.get('상담 시간') or '').strip() or '-',
        "location": parts.get('place') or '-',
        "device": parts.get('device') or str(lead.get('키워드') or '').strip() or '-',
        "visit_address": modal_fields.get('visit_address') or str(lead.get('방문 주소') or '').strip() or '-',
        "consultation": modal_fields.get('consultation') or '-',
        "details": parts.get('inquiry') or str(lead.get('문의 내용') or lead.get('상담 내용') or '').strip() or '-',
        "visit_date": modal_fields.get('visit_date') or '-',
        "visit_date_start": vd_start_iso or '-',  # 분리 변수 — Slack List datepicker 컬럼용
        "visit_date_end": vd_end_iso or '-',      # 종료일 (단일이면 '-')
        "estimate_request": modal_fields.get('estimate') or '-',
        "message_link": message_link or '-',
        "payload": f"lead_no={lead.get('리드 No')} action={action}",
    }

    try:
        req = urllib.request.Request(
            webhook_url,
            data=json.dumps(payload, ensure_ascii=False).encode('utf-8'),
            headers={'Content-Type': 'application/json; charset=utf-8'},
            method='POST',
        )
        with urllib.request.urlopen(req, timeout=5) as r:
            r.read()
        logger.info(
            f"[SLACK/LIST] webhook {op} 완료 (lead={lead.get('리드 No')} action={action})"
        )
        # 방문 캔버스 자동 sync (2026-07-15) — 시트 상태 변경 → 캔버스 rebuild
        try:
            from dashboard.services.visit_canvas_sync import rebuild_canvas_async
            rebuild_canvas_async()
        except Exception as _vc_exc:
            logger.debug(f"[SLACK/LIST] 방문 캔버스 rebuild trigger 실패: {_vc_exc}")
        return True
    except Exception as exc:
        logger.warning(f"[SLACK/LIST] webhook 호출 실패: {exc}")
        return False


def _process_visit_submission(client, body, view):
    """방문 요청 모달 제출 → 메인 시트 업데이트 + 원본 메시지에 답글 + 슬랙 List 등록"""
    metadata = json.loads(view["private_metadata"])
    lead_no = metadata["lead_no"]
    channel = metadata["channel"]
    message_ts = metadata["message_ts"]
    user_id = body["user"]["id"]

    state = view["state"]["values"]
    visit_date_raw = (_v(state, "visit_date") or '').strip()  # ISO "2026-06-25" (슬랙 표시용)
    visit_date_for_sheet = _format_date_for_sheet(visit_date_raw) if visit_date_raw else ''
    visit_address = _v(state, "visit_address")
    consultation = _v(state, "consultation")

    # 메인 시트 업데이트
    try:
        from dashboard.services.lead_service import update_lead
        update_data = {
            '상태': '방문 예약',
            '방문 예정일': visit_date_for_sheet,  # 시트 escape prefix
        }
        if visit_address:
            update_data['방문 주소'] = visit_address
        if consultation:
            update_data['상담 내용'] = consultation
        update_lead(lead_no, update_data)
    except Exception as exc:
        logger.error(f"[SLACK] 시트 업데이트 실패 ({lead_no}): {exc}", exc_info=True)

    # 슬랙 List webhook 등록 (raw 날짜 — 슬랙 List에 깔끔 표시)
    lead = _find_lead_by_no(lead_no) or {}
    _post_to_slack_list(
        client, lead,
        modal_fields={
            'visit_date': visit_date_raw,
            'visit_address': visit_address,
            'consultation': consultation,
        },
        channel=channel, message_ts=message_ts, action='visit',
    )

    # 원본 메시지에 답글 (raw 날짜 — 슬랙 표시 깔끔)
    reply_text = (
        f":white_check_mark: *방문 요청 등록* — `{lead_no}` by <@{user_id}>\n"
        f">*방문일* : {visit_date_raw or '-'}\n"
        f">*방문 주소* : {visit_address or '-'}\n"
        f">*내용 / 특이사항* : {consultation or '-'}"
    )
    try:
        client.chat_postMessage(
            channel=channel, thread_ts=message_ts,
            text=reply_text,
        )
    except Exception as exc:
        logger.error(f"[SLACK] 방문 요청 답글 실패 ({lead_no}): {exc}", exc_info=True)

    # 원본 인입 카드에 체크 reaction 추가 — 처리 완료 표시 (다른 사람이 한눈에 확인)
    try:
        client.reactions_add(
            channel=channel, timestamp=message_ts, name="white_check_mark",
        )
    except Exception as exc:
        # 이미 reaction 있거나(already_reacted) 권한 문제면 무시
        logger.debug(f"[SLACK] 방문 요청 reaction 추가 스킵 ({lead_no}): {exc}")


def _process_price_submission(client, body, view):
    """가격 문의 모달 제출 → 메인 시트 업데이트 + 원본 메시지에 답글"""
    metadata = json.loads(view["private_metadata"])
    lead_no = metadata["lead_no"]
    channel = metadata["channel"]
    message_ts = metadata["message_ts"]
    user_id = body["user"]["id"]

    state = view["state"]["values"]
    estimate = _v(state, "estimate")  # 'yes' or 'no'
    consultation = _v(state, "consultation")

    estimate_label = '요청 보냄' if estimate == 'yes' else '요청 안 보냄'
    new_status = '견적 제출' if estimate == 'yes' else '유선 상담'

    # 메인 시트 업데이트
    try:
        from dashboard.services.lead_service import update_lead
        update_data = {
            '상태': new_status,
        }
        if consultation:
            update_data['상담 내용'] = consultation
        update_lead(lead_no, update_data)
    except Exception as exc:
        logger.error(f"[SLACK] 시트 업데이트 실패 ({lead_no}): {exc}", exc_info=True)

    # 슬랙 List webhook 등록
    lead = _find_lead_by_no(lead_no) or {}
    _post_to_slack_list(
        client, lead,
        modal_fields={
            'consultation': consultation,
            'estimate': estimate_label,
        },
        channel=channel, message_ts=message_ts, action='price',
    )

    # 원본 메시지에 답글
    reply_text = (
        f":moneybag: *가격 문의 처리* — `{lead_no}` by <@{user_id}>\n"
        f">*가견적 요청* : {estimate_label}\n"
        f">*상담 내용* : {consultation or '-'}"
    )
    try:
        client.chat_postMessage(
            channel=channel, thread_ts=message_ts,
            text=reply_text,
        )
    except Exception as exc:
        logger.error(f"[SLACK] 가격 문의 답글 실패 ({lead_no}): {exc}", exc_info=True)

    # 원본 인입 카드에 체크 reaction 추가 — 처리 완료 표시
    try:
        client.reactions_add(
            channel=channel, timestamp=message_ts, name="white_check_mark",
        )
    except Exception as exc:
        logger.debug(f"[SLACK] 가격 문의 reaction 추가 스킵 ({lead_no}): {exc}")


# ─────────────────────────────────────────────────────────────
# 계산서 발행 요청 흐름 (공사 확정 카드 [💰 계산서 요청] 클릭 → 모달 → #영업_관리)
# ─────────────────────────────────────────────────────────────
def _money_kr(digits: str) -> str:
    """숫자 문자열 → '5,200,000원'. 빈 값은 '-'."""
    d = ''.join(ch for ch in (digits or '') if ch.isdigit())
    if not d:
        return '-'
    return f"{int(d):,}원"


# ─────────────────────────────────────────────────────────────
# 공사 확정 카드 편집 / 취소 헬퍼 (2026-07-09)
# ─────────────────────────────────────────────────────────────
_CONTRACT_TYPE_OPTIONS = ['외주', '내부', '일당', '기타']


def _load_active_constructors_flat() -> list:
    """활성 시공자 목록 (카테고리 flat, 이름만). 모달 multi_static_select 옵션 용."""
    try:
        from dashboard.utils.user_database import get_constructor_repository
        grouped = get_constructor_repository().get_grouped(active_only=True)
        names = []
        for cat_items in grouped.values():
            for c in cat_items:
                nm = (c.get('name') or '').strip()
                if nm and nm not in names:
                    names.append(nm)
        return names
    except Exception as exc:
        logger.warning(f'[SLACK/공사수정] 시공자 목록 로드 실패: {exc}')
        return []


def _multiselect_options(values: list) -> list:
    return [{'text': {'type': 'plain_text', 'text': v}, 'value': v} for v in values]


def _open_project_edit_modal(client, body) -> None:
    """[✏️ 내용 수정] 클릭 → 편집 가능 필드 7개 pre-fill 모달."""
    from dashboard.services.project_service import get_project_records

    trigger_id = body["trigger_id"]
    code = (body["actions"][0].get("value") or '').strip()
    channel = body.get("channel", {}).get("id", "") or body.get("container", {}).get("channel_id", "")
    message_ts = body.get("message", {}).get("ts", "") or body.get("container", {}).get("message_ts", "")

    if not code:
        return

    records = get_project_records() or []
    project = next((r for r in records if (r.get('프로젝트 코드') or '').strip() == code), None)
    if not project:
        try:
            client.chat_postEphemeral(
                channel=channel, user=body["user"]["id"],
                text=f':warning: `{code}` 프로젝트를 찾을 수 없습니다. (시트에서 삭제/이동됐을 수 있음)',
            )
        except Exception:
            pass
        return

    metadata = json.dumps({'code': code, 'channel': channel, 'message_ts': message_ts}, ensure_ascii=False)

    # pre-fill 값 준비
    def _val(field):
        v = project.get(field, '')
        return '' if v in (None, '-', 'None') else str(v).strip()

    content = _val('공사 내용')
    contractor_raw = _val('시공자')
    contract_type_raw = _val('도급 구분')
    amount_raw = project.get('총액 1', '')
    amt_str = ''
    if amount_raw not in (None, '', '-'):
        try:
            amt_str = f"{int(float(str(amount_raw).replace(',', '').strip())):,}"
        except (ValueError, TypeError):
            amt_str = str(amount_raw)
    vat_raw = project.get('부가세')
    vat_sep = (
        vat_raw is True
        or (isinstance(vat_raw, str) and vat_raw.strip().upper() in ('TRUE', 'Y', 'YES', '1'))
        or vat_raw == 1
    )
    start_raw = _val('공사 시작')
    end_raw = _val('공사 종료')
    start_date = start_raw[:10] if len(start_raw) >= 10 else ''
    end_date = end_raw[:10] if len(end_raw) >= 10 else ''

    # multi-select initial (값이 옵션에 있는 것만)
    contract_type_current = [t.strip() for t in re.split(r'[,/]', contract_type_raw) if t.strip()]
    contract_type_initial = _multiselect_options([t for t in contract_type_current if t in _CONTRACT_TYPE_OPTIONS])

    constructor_names = _load_active_constructors_flat()
    # 기존 값 중 리스트에 없으면 추가 (비활성 시공자 유지)
    contractor_current = [n.strip() for n in re.split(r'[,/]', contractor_raw) if n.strip()]
    for n in contractor_current:
        if n not in constructor_names:
            constructor_names.append(n)
    contractor_initial = _multiselect_options([n for n in contractor_current if n in constructor_names])

    vat_option = {'text': {'type': 'plain_text', 'text': 'VAT 별도'}, 'value': 'sep'}

    blocks = [
        {"type": "section", "text": {"type": "mrkdwn", "text": f"프로젝트 `{code}` 수정"}},
        {
            "type": "input", "block_id": "content", "optional": True,
            "label": {"type": "plain_text", "text": "공사 내용"},
            "element": {
                "type": "plain_text_input", "action_id": "value",
                **({"initial_value": content} if content else {}),
            },
        },
        {
            "type": "input", "block_id": "contract_type", "optional": True,
            "label": {"type": "plain_text", "text": "도급 구분 (최대 2개)"},
            "element": {
                "type": "multi_static_select", "action_id": "value",
                "max_selected_items": 2,
                "options": _multiselect_options(_CONTRACT_TYPE_OPTIONS),
                **({"initial_options": contract_type_initial} if contract_type_initial else {}),
            },
        },
        {
            "type": "input", "block_id": "contractor", "optional": True,
            "label": {"type": "plain_text", "text": "시공자"},
            "element": {
                "type": "multi_static_select", "action_id": "value",
                "options": _multiselect_options(constructor_names),
                **({"initial_options": contractor_initial} if contractor_initial else {}),
            },
        },
        {
            "type": "input", "block_id": "amount", "optional": True,
            "label": {"type": "plain_text", "text": "공사 금액"},
            "element": {
                "type": "plain_text_input", "action_id": "value",
                "placeholder": {"type": "plain_text", "text": "예: 15,000,000 (콤마·공백 무시됨)"},
                **({"initial_value": amt_str} if amt_str else {}),
            },
        },
        {
            "type": "input", "block_id": "vat", "optional": True,
            "label": {"type": "plain_text", "text": "부가세"},
            "element": {
                "type": "checkboxes", "action_id": "value",
                "options": [vat_option],
                **({"initial_options": [vat_option]} if vat_sep else {}),
            },
        },
        {
            "type": "input", "block_id": "start_date", "optional": True,
            "label": {"type": "plain_text", "text": "공사 시작"},
            "element": {
                "type": "datepicker", "action_id": "value",
                **({"initial_date": start_date} if start_date else {}),
            },
        },
        {
            "type": "input", "block_id": "end_date", "optional": True,
            "label": {"type": "plain_text", "text": "공사 종료"},
            "element": {
                "type": "datepicker", "action_id": "value",
                **({"initial_date": end_date} if end_date else {}),
            },
        },
        {
            "type": "input", "block_id": "reason",
            "label": {"type": "plain_text", "text": "수정 사유 (필수)"},
            "element": {
                "type": "plain_text_input", "action_id": "value", "multiline": True,
                "placeholder": {"type": "plain_text",
                                "text": "예: 공사 내용 변경으로 공사 금액 상향 or 하향"},
            },
        },
    ]

    view = {
        "type": "modal",
        "callback_id": "submit_project_edit",
        "private_metadata": metadata,
        "title": {"type": "plain_text", "text": "공사 내용 수정"},
        "submit": {"type": "plain_text", "text": "저장"},
        "close": {"type": "plain_text", "text": "닫기"},
        "blocks": blocks,
    }
    client.views_open(trigger_id=trigger_id, view=view)


def _process_project_edit_submission(client, body, view) -> None:
    """편집 모달 제출 → project_slack_actions.perform_edit 호출."""
    from dashboard.services.project_slack_actions import perform_edit

    metadata = json.loads(view.get("private_metadata") or "{}")
    code = metadata.get("code", "")
    if not code:
        return

    channel = metadata.get("channel", "")
    user_id = body.get("user", {}).get("id", "")

    values = view["state"]["values"]

    def _text(bid):
        return (values.get(bid, {}).get('value', {}) or {}).get('value', '') or ''

    def _multi(bid):
        opts = (values.get(bid, {}).get('value', {}) or {}).get('selected_options', []) or []
        return [o.get('value', '') for o in opts if o.get('value')]

    def _date(bid):
        return (values.get(bid, {}).get('value', {}) or {}).get('selected_date', '') or ''

    def _checked(bid):
        opts = (values.get(bid, {}).get('value', {}) or {}).get('selected_options', []) or []
        return bool(opts)

    content = _text('content').strip()
    contract_types = _multi('contract_type')
    contractors = _multi('contractor')
    amount_raw = _text('amount').strip()
    vat_sep = _checked('vat')
    start_date = _date('start_date')
    end_date = _date('end_date')
    reason = _text('reason').strip()

    # 편집할 필드만 dict 구성 — 원본 값과 다를 때만 (관리 사이트도 diff 기반)
    from dashboard.services.project_service import get_project_records
    records = get_project_records() or []
    project = next((r for r in records if (r.get('프로젝트 코드') or '').strip() == code), None)
    if not project:
        return

    updates = {}
    if content and content != (project.get('공사 내용') or '').strip():
        updates['공사 내용'] = content
    new_contract = ', '.join(contract_types)
    if new_contract != (project.get('도급 구분') or '').strip():
        updates['도급 구분'] = new_contract
    new_contractor = ', '.join(contractors)
    if new_contractor != (project.get('시공자') or '').strip():
        updates['시공자'] = new_contractor
    if amount_raw:
        digits = ''.join(ch for ch in amount_raw if ch.isdigit())
        if digits:
            new_amt = int(digits)
            try:
                cur_amt = int(float(str(project.get('총액 1', 0) or 0).replace(',', '').strip() or 0))
            except (ValueError, TypeError):
                cur_amt = 0
            if new_amt != cur_amt:
                updates['총액 1'] = new_amt
    # VAT는 체크박스 → bool. 원본과 다르면 반영.
    cur_vat_raw = project.get('부가세')
    cur_vat = (
        cur_vat_raw is True
        or (isinstance(cur_vat_raw, str) and cur_vat_raw.strip().upper() in ('TRUE', 'Y', 'YES', '1'))
        or cur_vat_raw == 1
    )
    if vat_sep != cur_vat:
        updates['부가세'] = vat_sep
    if start_date and start_date != (project.get('공사 시작') or '')[:10]:
        updates['공사 시작'] = start_date
    if end_date and end_date != (project.get('공사 종료') or '')[:10]:
        updates['공사 종료'] = end_date

    if not updates:
        try:
            client.chat_postEphemeral(
                channel=channel, user=user_id,
                text=':information_source: 변경된 필드가 없어 저장을 skip 했습니다.',
            )
        except Exception:
            pass
        return

    initial = _slack_user_to_initial(client, user_id) or '-'
    result = perform_edit(code, updates, reason, initial)

    if result.get('ok'):
        try:
            summary = ', '.join(updates.keys())
            client.chat_postEphemeral(
                channel=channel, user=user_id,
                text=f':white_check_mark: `{code}` 수정 완료 — {summary}',
            )
        except Exception:
            pass
        # #영업_관리 채널에 수정 알림 카드 발송 (양식은 계산서 요청과 유사)
        try:
            _post_project_edit_notice_card(
                client, code, project, updates, reason, initial,
            )
        except Exception as exc:
            logger.warning(f'[SLACK/공사수정] 영업_관리 알림 실패 ({code}): {exc}')
    else:
        try:
            client.chat_postEphemeral(
                channel=channel, user=user_id,
                text=f':x: `{code}` 수정 실패: {result.get("reason", "unknown")}',
            )
        except Exception:
            pass


def _fmt_edit_field_change(field: str, old_value, new_value, current_vat_after: bool) -> str:
    """수정 알림 카드 한 줄 렌더링. 부가세/총액은 사람이 읽기 쉬운 포맷으로."""
    def _money(v):
        try:
            d = int(float(str(v).replace(',', '').strip() or 0))
            return f'{d:,}원' if d else '-'
        except (ValueError, TypeError):
            return str(v) if v not in (None, '') else '-'

    def _vat_label(v):
        if v is True or (isinstance(v, str) and v.strip().upper() in ('TRUE', 'Y', 'YES', '1')) or v == 1:
            return 'VAT 별도'
        return 'VAT 없음'

    def _txt(v):
        s = str(v).strip() if v is not None else ''
        return s if s else '-'

    if field == '총액 1':
        old_disp = _money(old_value)
        new_disp = _money(new_value)
        vat_suffix = f' ({_vat_label(current_vat_after)})' if new_disp != '-' else ''
        return f'  • 공사 금액 : {old_disp} → {new_disp}{vat_suffix}'
    if field == '부가세':
        return f'  • 부가세 : {_vat_label(old_value)} → {_vat_label(new_value)}'
    return f'  • {field} : {_txt(old_value)} → {_txt(new_value)}'


def _post_project_edit_notice_card(
    client, code: str, before_project: dict, updates: dict, reason: str, initial: str,
) -> None:
    """수정 완료 후 #영업_관리 채널에 알림 카드 발송."""
    channel_id = os.getenv('SLACK_INVOICE_CHANNEL_ID', '').strip()
    if not channel_id:
        logger.debug('[SLACK/공사수정] SLACK_INVOICE_CHANNEL_ID 미설정 — 알림 skip')
        return

    biz = (before_project.get('사업자명') or '-').strip() or '-'
    addr = (before_project.get('현장 주소') or '-').strip() or '-'
    now_str = datetime.now().strftime('%m.%d %H:%M')

    # 부가세 반영된 최종 상태 (총액 라벨용)
    vat_after_raw = updates.get('부가세', before_project.get('부가세'))
    vat_after = (
        vat_after_raw is True
        or (isinstance(vat_after_raw, str) and vat_after_raw.strip().upper() in ('TRUE', 'Y', 'YES', '1'))
        or vat_after_raw == 1
    )

    change_lines = []
    # 표시 순서 고정 (모달 순서와 동일)
    field_order = ['공사 내용', '도급 구분', '시공자', '총액 1', '부가세', '공사 시작', '공사 종료']
    for f in field_order:
        if f not in updates:
            continue
        change_lines.append(_fmt_edit_field_change(f, before_project.get(f, ''), updates[f], vat_after))

    lines = [
        f'🔔 *[공사 내용 수정 알림]*  `{code}`',
        '--------------------------------------------',
        f'🏢 사업자명 : {biz}',
        f'📍 현장 주소 : {addr}',
        f'📝 수정 사유 : {reason.strip()}',
        '📋 변경 내역',
        *change_lines,
        f'👤 수정자 : {initial}  {now_str}',
        '--------------------------------------------',
    ]
    text = '⠀\n' + '\n'.join(lines)
    blocks = [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': text}},
        {'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': '⠀'}]},
    ]

    try:
        client.conversations_join(channel=channel_id)
    except Exception:
        pass
    resp = client.chat_postMessage(
        channel=channel_id, text=text, blocks=blocks, unfurl_links=False,
    )
    if resp.get('ok'):
        logger.info(
            f'[SLACK/공사수정] 영업_관리 알림 발송 완료: {code} ts={resp.get("ts")}'
        )
    else:
        logger.warning(f'[SLACK/공사수정] 영업_관리 알림 실패: {resp}')


def _process_project_cancel(client, body) -> None:
    """[❌ 공사 취소] 확인 → perform_cancel → 카드 chat.update(방문 취소 UI 스타일).

    취소자 표기는 이니셜(방문 취소·기타 알림과 통일). 감사 로그용 by_user 는
    슬랙 표시명 fallback 이니셜 사용.
    """
    from dashboard.services.project_slack_actions import perform_cancel

    code = (body["actions"][0].get("value") or '').strip()
    channel = body["channel"]["id"]
    message_ts = body["message"]["ts"]
    user_id = body["user"]["id"]
    if not code:
        return

    initial = _slack_user_to_initial(client, user_id) or '-'
    result = perform_cancel(code, initial)

    if not result.get('ok'):
        reason = result.get('reason', 'unknown')
        try:
            client.chat_postEphemeral(
                channel=channel, user=user_id,
                text=f':x: `{code}` 취소 실패: {reason}',
            )
        except Exception:
            pass
        return

    # 카드 chat.update — 방문 취소 UI 스타일 그대로 (원본 회색 처리)
    try:
        cancel_time = datetime.now().strftime('%Y.%m.%d. %H:%M')
        # 원본 텍스트는 body 대신 프로젝트 스냅샷에서 재생성.
        # (Slack이 unicode 이모지를 :bell: 같은 shortcode로 정규화 저장해
        # body에서 읽어오면 shortcode 그대로 나와 코드블록 안에서 렌더 안 됨)
        from dashboard.services.project_slack_notifier import _build_message
        from dashboard.services.business_license_handler import verify_license_exists
        snapshot = result.get('project') or {}
        try:
            license_attached = verify_license_exists(code)
        except Exception:
            license_attached = False
        original_text = _build_message(snapshot, code, license_attached=license_attached)
        cleaned = [ln.replace('*', '') for ln in original_text.split('\n')]
        clean_text = '\n'.join(cleaned).strip()

        new_text = (
            f"🚫 *고객 요청으로 공사 취소*  `{code}`\n"
            f"취소한 사람 : {initial}\n"
            f"취소 시간 : {cancel_time}\n"
            f"\n"
            f"```\n{clean_text}\n```"
        )
        new_blocks = [
            {"type": "section", "text": {"type": "mrkdwn", "text": new_text}},
            {
                "type": "actions",
                "elements": [
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "↩️ 취소 되돌리기", "emoji": True},
                        "style": "primary",
                        "value": code,
                        "action_id": "project_uncancel",
                        "confirm": {
                            "title": {"type": "plain_text", "text": "취소 되돌리기"},
                            "text": {"type": "plain_text",
                                     "text": f"{code} 공사 취소를 되돌리시겠습니까?"},
                            "confirm": {"type": "plain_text", "text": "되돌리기"},
                            "deny": {"type": "plain_text", "text": "닫기"},
                        },
                    },
                ],
            },
            {"type": "context", "elements": [{"type": "mrkdwn", "text": "⠀"}]},
        ]
        client.chat_update(channel=channel, ts=message_ts, text=new_text, blocks=new_blocks)
    except Exception as exc:
        logger.error(f"[SLACK/공사취소] chat.update 실패 ({code}): {exc}", exc_info=True)

    # #영업_관리 채널에 취소 알림 카드 발송 (계산서 요청과 유사한 양식)
    try:
        _post_project_cancel_notice_card(client, code, result.get('project') or {}, initial)
    except Exception as exc:
        logger.warning(f'[SLACK/공사취소] 영업_관리 알림 실패 ({code}): {exc}')


def _post_project_cancel_notice_card(
    client, code: str, before_project: dict, initial: str,
) -> None:
    """취소 완료 후 #영업_관리 채널에 알림 카드 발송."""
    channel_id = os.getenv('SLACK_INVOICE_CHANNEL_ID', '').strip()
    if not channel_id:
        logger.debug('[SLACK/공사취소] SLACK_INVOICE_CHANNEL_ID 미설정 — 알림 skip')
        return

    biz = (before_project.get('사업자명') or '-').strip() or '-'
    addr = (before_project.get('현장 주소') or '-').strip() or '-'

    # 금액 표시 (부가세 반영)
    amt_raw = before_project.get('총액 1', '')
    try:
        amt_int = int(float(str(amt_raw).replace(',', '').strip() or 0))
        amt_disp = f'{amt_int:,}원' if amt_int else '-'
    except (ValueError, TypeError):
        amt_disp = '-'
    vat_raw = before_project.get('부가세')
    vat_sep = (
        vat_raw is True
        or (isinstance(vat_raw, str) and vat_raw.strip().upper() in ('TRUE', 'Y', 'YES', '1'))
        or vat_raw == 1
    )
    if amt_disp != '-':
        amt_disp = f"{amt_disp} ({'VAT 별도' if vat_sep else 'VAT 없음'})"

    confirmed_raw = str(before_project.get('공사 확정', '') or '').strip()
    confirmed_disp = confirmed_raw[:10] if confirmed_raw else '-'

    now_str = datetime.now().strftime('%m.%d %H:%M')
    lines = [
        f'🔔 *[공사 취소 알림]*  `{code}`',
        '--------------------------------------------',
        f'🏢 사업자명 : {biz}',
        f'📍 현장 주소 : {addr}',
        f'💲 공사 금액 : {amt_disp}',
        f'📅 공사 확정일 : {confirmed_disp}',
        f'👤 취소자 : {initial}  {now_str}',
        '--------------------------------------------',
    ]
    text = '⠀\n' + '\n'.join(lines)
    blocks = [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': text}},
        {'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': '⠀'}]},
    ]

    try:
        client.conversations_join(channel=channel_id)
    except Exception:
        pass
    resp = client.chat_postMessage(
        channel=channel_id, text=text, blocks=blocks, unfurl_links=False,
    )
    if resp.get('ok'):
        logger.info(
            f'[SLACK/공사취소] 영업_관리 알림 발송 완료: {code} ts={resp.get("ts")}'
        )
    else:
        logger.warning(f'[SLACK/공사취소] 영업_관리 알림 실패: {resp}')


def _process_project_uncancel(client, body) -> None:
    """[↩️ 취소 되돌리기] → perform_uncancel → 카드 원본 형태로 복원."""
    from dashboard.services.project_slack_actions import perform_uncancel
    from dashboard.services.project_slack_notifier import _build_blocks
    from dashboard.services.business_license_handler import verify_license_exists
    from dashboard.services.project_service import get_project_records

    code = (body["actions"][0].get("value") or '').strip()
    channel = body["channel"]["id"]
    message_ts = body["message"]["ts"]
    user_id = body["user"]["id"]
    if not code:
        return

    display_name = _slack_user_to_korean_name(client, user_id) or user_id
    result = perform_uncancel(code, display_name)

    if not result.get('ok'):
        try:
            client.chat_postEphemeral(
                channel=channel, user=user_id,
                text=f':x: `{code}` 재개 실패: {result.get("reason", "unknown")}',
            )
        except Exception:
            pass
        return

    # 카드 원본 형태로 재렌더링
    try:
        records = get_project_records(force_refresh=True) or []
        latest = next((r for r in records if (r.get('프로젝트 코드') or '').strip() == code), None)
        if not latest:
            return
        try:
            license_attached = verify_license_exists(code)
        except Exception:
            license_attached = False
        from dashboard.services.project_slack_notifier import _thread_permalink
        permalink = _thread_permalink(channel, message_ts)
        new_blocks = _build_blocks(
            latest, code,
            license_attached=license_attached,
            thread_permalink=permalink,
        )
        biz = latest.get('사업자명') or ''
        fallback = f"[공사 확정] {code} {biz}".strip()
        client.chat_update(channel=channel, ts=message_ts, text=fallback, blocks=new_blocks)
    except Exception as exc:
        logger.error(f"[SLACK/공사재개] chat.update 실패 ({code}): {exc}", exc_info=True)


def _is_license_required(code: str) -> bool:
    """사업자등록증 첨부 검증 필수 여부.

    2026-07-21 임시 조치: '거래처' 유입만 검증 skip.
    사업자등록증 마스터 폴더 축적이 부족 (전체 거래처 3,451건 중 258건, 7.5%) 하여
    거래처 계약 시마다 재첨부 요구가 낭비. 온라인·숨고·당근·홈페이지·전화·소개·기타는
    신규 사업자 가능성 높아 검증 유지.
    거래처 마스터 파일 재사용 로직 도입 후 재검토.

    조회 실패 시 안전 default = True (검증 유지).
    """
    if not code or code == '-':
        return True
    try:
        from dashboard.services.project_service import get_project_records
        for r in get_project_records() or []:
            if (r.get('프로젝트 코드') or '').strip() == code:
                inflow = str(r.get('유입 구분') or '').strip()
                return inflow != '거래처'
        return True
    except Exception as exc:
        logger.warning(f'[SLACK/계산서] 유입 구분 조회 실패 → 검증 유지 ({code}): {exc}')
        return True


def _open_invoice_modal(client, body) -> None:
    """[💰 계산서 요청] 클릭 → 프로젝트 정보 pre-fill 모달 오픈.

    2026-07-15: 사업자등록증 검증을 모달 오픈 전으로 복귀 (사용자 요청).
      미첨부 시 ephemeral 로 안내하고 모달 자체를 열지 않음.
      Drive API 검증 ~0.5초 + views.open ~0.5초 = trigger_id 3초 안에 완료 가능.
      Drive API 예외 (지연·타임아웃) 시 검증 skip 후 모달 오픈 (submit 시점
      재검증으로 방어).
    """
    trigger_id = body["trigger_id"]
    action = body["actions"][0]
    try:
        payload = json.loads(action.get("value") or "{}")
    except Exception:
        payload = {}

    code = payload.get('code', '') or '-'
    channel_id = (body.get('channel') or {}).get('id', '')
    user_id = (body.get('user') or {}).get('id', '')
    # 클릭한 카드의 ts → 안내 모달의 스레드 permalink 링크에 사용.
    clicked_ts = ((body.get('message') or {}).get('ts') or '').strip()

    # 사업자등록증 검증 (모달 오픈 전, 2026-07-15) — 미첨부 시 안내 모달 + return.
    # 2026-07-21: '거래처' 유입만 검증 skip (_is_license_required 참조).
    # 2026-07-21: ephemeral 이 오래된 카드에선 스크롤 위로 밀려 인지 안 됨.
    #   → 안내 모달 팝업으로 변경 (화면 중앙 = 100% 인지).
    if code and code != '-' and _is_license_required(code):
        try:
            from dashboard.services.business_license_handler import verify_license_exists
            if not verify_license_exists(code):
                permalink = ''
                if clicked_ts and channel_id:
                    try:
                        pr = client.chat_getPermalink(channel=channel_id, message_ts=clicked_ts)
                        permalink = (pr.get('permalink') or '').strip()
                    except Exception as exc:
                        logger.debug(f'[SLACK/계산서] permalink 조회 실패 (무시): {exc}')
                blocks = [
                    {"type": "section", "text": {
                        "type": "mrkdwn",
                        "text": (f":warning: *`{code}` 사업자등록증이 첨부되지 않았습니다.*\n\n"
                                 f"공사 확정 카드 스레드에 사업자등록증(이미지 or PDF) 을 "
                                 f"먼저 첨부한 뒤 다시 [💰 계산서 요청] 을 눌러주세요."),
                    }},
                ]
                if permalink:
                    blocks.append({
                        "type": "actions",
                        "elements": [{
                            "type": "button",
                            "text": {"type": "plain_text", "text": "📎 공사 확정 카드로 이동", "emoji": True},
                            "url": permalink,
                            "action_id": "goto_project_thread",
                        }],
                    })
                try:
                    client.views_open(
                        trigger_id=trigger_id,
                        view={
                            "type": "modal",
                            "title": {"type": "plain_text", "text": "사업자등록증 미첨부"},
                            "close": {"type": "plain_text", "text": "닫기"},
                            "blocks": blocks,
                        },
                    )
                except Exception as exc:
                    logger.warning(f'[SLACK/계산서] 안내 모달 오픈 실패, ephemeral fallback: {exc}')
                    # trigger_id 만료 등 예외 시 ephemeral 로 fallback
                    try:
                        kwargs = {
                            'channel': channel_id, 'user': user_id,
                            'text': (f':warning: `{code}` 사업자등록증이 첨부되지 않았습니다.\n'
                                     f'공사 확정 카드 스레드에 사업자등록증(이미지 or PDF) 을 '
                                     f'첨부한 뒤 다시 [💰 계산서 요청] 을 눌러주세요.'),
                        }
                        if clicked_ts:
                            kwargs['thread_ts'] = clicked_ts
                        client.chat_postEphemeral(**kwargs)
                    except Exception as exc2:
                        logger.warning(f'[SLACK/계산서] ephemeral fallback 도 실패: {exc2}')
                logger.info(f'[SLACK/계산서] 사업자등록증 미첨부 → 안내 모달 표시 ({code})')
                return
        except Exception as exc:
            logger.warning(
                f'[SLACK/계산서] 사업자등록증 검증 실패 (모달 오픈 진행, submit 시 재검증): {exc}'
            )

    biz = payload.get('biz', '') or ''
    addr = payload.get('addr', '') or ''
    amt = payload.get('amt', '') or ''
    # pre-fill 시 콤마 자동 포맷 (사용자 가독성)
    if amt.isdigit():
        amt = f"{int(amt):,}"
    # 부가세는 계산서 발행 특성상 항상 '별도' — 필드 제거 (2026-07-13).
    email = payload.get('email', '') or ''

    # button value 는 카드 발송 시점 스냅샷 → OCR 로 갱신된 사업자명·이메일·총액이
    # 반영되지 않음. 시트 최신값이 있으면 우선 사용 (2026-07-13, 총액 2026-07-15 추가).
    if code and code != '-':
        try:
            from dashboard.services.project_service import get_project_records
            _records = get_project_records() or []
            _latest = next(
                (r for r in _records if (r.get('프로젝트 코드') or '').strip() == code),
                None,
            )
            if _latest:
                _biz_latest = (_latest.get('사업자명') or '').strip()
                if _biz_latest and _biz_latest != '-':
                    biz = _biz_latest
                _email_latest = (_latest.get('발주처 이메일') or '').strip()
                if _email_latest and _email_latest != '-':
                    email = _email_latest
                # 총액 1 재조회 (2026-07-15) — 매니저가 나중에 편집한 경우 반영
                _amt_raw = _latest.get('총액 1', '')
                try:
                    _amt_int = int(float(str(_amt_raw).replace(',', '').strip() or '0'))
                    if _amt_int > 0:
                        amt = f"{_amt_int:,}"
                except (ValueError, TypeError):
                    pass
        except Exception as exc:
            logger.warning(f'[SLACK/계산서] 최신값 조회 실패 (payload fallback): {exc}')

    metadata = json.dumps({"code": code}, ensure_ascii=False)

    # 필드별 initial_value 정책 (2026-07-13):
    #   - 사업자명·이메일: 빈값이면 initial 없이 두어 매니저가 직접 입력 (필수 항목).
    #     매니저 패턴상 '-' 접두어가 있으면 '-TEST@TEST.COM' 처럼 이어쓰는 경우가 있어
    #     아예 빈 필드로 유지 → 슬랙이 required error 로 자연스럽게 유도.
    #   - 현장 주소·금액: '-' 로 채워 매니저가 수정하거나 그대로 요청 가능.
    addr = addr or '-'
    amt = amt or '-'

    def _text_input(block_id, label, value, multiline=False, optional=False, placeholder=''):
        el = {"type": "plain_text_input", "action_id": "value"}
        if value:
            el["initial_value"] = value
        if placeholder:
            el["placeholder"] = {"type": "plain_text", "text": placeholder}
        if multiline:
            el["multiline"] = True
        blk = {
            "type": "input", "block_id": block_id,
            "label": {"type": "plain_text", "text": label},
            "element": el,
        }
        if optional:
            blk["optional"] = True
        return blk

    view = {
        "type": "modal",
        "callback_id": "submit_invoice",
        "private_metadata": metadata,
        "title": {"type": "plain_text", "text": "세금계산서 발행 요청"},
        "submit": {"type": "plain_text", "text": "요청 발송"},
        "close": {"type": "plain_text", "text": "취소"},
        "blocks": [
            {"type": "section", "text": {"type": "mrkdwn",
                "text": f"프로젝트 `{code}` 세금계산서 발행 요청"}},
            _text_input("biz", "사업자명", biz),
            _text_input("addr", "현장 주소", addr),
            # 공사 금액 (시트 원본, read-only 참고)
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": f"*공사 금액 (시트 원본)*\n{amt} 원",
                },
            },
            _text_input("amt", "계산서 발행 금액", amt),
            # VAT radio_buttons — 체크박스 여백 오클릭 사고 방지 (2026-07-16)
            {
                "type": "input", "block_id": "vat",
                "label": {"type": "plain_text", "text": "VAT (부가가치세)"},
                "element": {
                    "type": "radio_buttons",
                    "action_id": "value",
                    "initial_option": {
                        "text": {"type": "plain_text", "text": "VAT 별도"},
                        "value": "sep",
                    },
                    "options": [
                        {"text": {"type": "plain_text", "text": "VAT 별도"}, "value": "sep"},
                        {"text": {"type": "plain_text", "text": "VAT 포함"}, "value": "incl"},
                    ],
                },
            },
            # 이메일 필드 — 거래처 유입만 optional (2026-07-21).
            # 사업자등록증 검증과 동일 정책: 거래처는 마스터에 이미 있어 매번 재입력 불필요.
            _text_input(
                "email", "발행 이메일", email,
                optional=(not _is_license_required(code)),
            ),
            _text_input(
                "memo", "추가 요청사항", "",
                multiline=True, optional=True,
                placeholder='예) 청구 or 영수 발행\n예) 항목이나 비고란에 특정 내용 기재',
            ),
        ],
    }
    client.views_open(trigger_id=trigger_id, view=view)


def _notify_invoice_submit_error(client, channel_id: str, user_id: str,
                                   code: str, error_lines: list) -> None:
    """계산서 요청 검증 실패 안내 (chat.postEphemeral → DM fallback)."""
    if not user_id or not error_lines:
        return
    header = f":x: *[세금계산서 요청 반려]*  `{code or '-'}`"
    body = header + '\n' + '\n'.join(error_lines) + '\n\n_(수정 후 카드에서 다시 요청해주세요.)_'
    # 1) 채널 ephemeral (매니저가 그 채널을 보고 있으면 즉시 표시)
    if channel_id:
        try:
            client.chat_postEphemeral(channel=channel_id, user=user_id, text=body)
            return
        except Exception as exc:
            logger.warning(f'[SLACK/계산서] ephemeral 반려 안내 실패 ({code}): {exc}')
    # 2) DM fallback
    try:
        im = client.conversations_open(users=user_id)
        dm_ch = ((im.get('channel') or {}) or {}).get('id')
        if dm_ch:
            client.chat_postMessage(channel=dm_ch, text=body)
    except Exception as exc:
        logger.warning(f'[SLACK/계산서] DM 반려 안내 실패 ({code}): {exc}')


def _process_invoice_submit_bg(client, body, view) -> None:
    """계산서 요청 submit 전체 처리 (검증 + 카드 발송) — BG 스레드용.

    modal 이 이미 ack() 로 닫힌 상태에서 실행되므로, 검증 실패도 modal 오류 대신
    ephemeral/DM 안내로 처리. view.id 기반 idempotency lock 으로 중복 방어.
    """
    try:
        metadata = json.loads(view.get("private_metadata") or "{}")
        code = (metadata.get("code", "") or "").strip()
    except Exception:
        code = ''

    user_id = body.get('user', {}).get('id', '')
    channel_id = os.getenv('SLACK_INVOICE_CHANNEL_ID', '').strip()

    # 중복 submit lock — 첫 submit 만 통과. 실패 시 매니저가 다시 요청 카드에서
    # 열어 새 view_id 로 재제출 가능하므로 lock 유지해도 무방.
    _view_id = (view.get('id') or '').strip()
    if _view_id:
        try:
            from dashboard.utils.redis_client import get_redis_client
            rc = get_redis_client().redis
            if not rc.set(f'invoice_submit_lock:{_view_id}', '1', nx=True, ex=300):
                logger.info(f'[SLACK/계산서] 중복 submit skip (view_id={_view_id} code={code})')
                return
        except Exception as exc:
            logger.warning(f'[SLACK/계산서] idempotency lock 실패 (계속 진행): {exc}')

    # 검증 (병렬)
    error_lines = []
    if code and code != '-':
        from concurrent.futures import ThreadPoolExecutor

        def _check_license() -> bool:
            # 2026-07-21: '거래처' 유입은 검증 skip (마스터 재사용 로직 도입 전 임시).
            if not _is_license_required(code):
                return True
            try:
                from dashboard.services.business_license_handler import verify_license_exists
                return bool(verify_license_exists(code))
            except Exception as exc:
                logger.warning(f'[SLACK/계산서] 사업자등록증 검증 실패 (통과): {exc}')
                return True  # Drive 지연 시 통과 (관리자 후속 처리)

        def _check_vat_filled() -> bool:
            try:
                from dashboard.services.project_service import get_project_records
                records = get_project_records() or []
                for r in records:
                    if (r.get('프로젝트 코드') or '').strip() == code:
                        vat_raw = r.get('부가세')
                        if vat_raw in (None, '', ' '):
                            return False
                        if isinstance(vat_raw, str) and not vat_raw.strip():
                            return False
                        return True
                return True  # 프로젝트 못 찾으면 통과
            except Exception as exc:
                logger.warning(f'[SLACK/계산서] 부가세 필드 검증 실패 (통과): {exc}')
                return True

        with ThreadPoolExecutor(max_workers=2) as ex:
            fut_lic = ex.submit(_check_license)
            fut_vat = ex.submit(_check_vat_filled)
            lic_ok = fut_lic.result()
            vat_ok = fut_vat.result()

        if not lic_ok:
            error_lines.append(
                "• :page_facing_up: *사업자등록증 미첨부* — 공사 확정 카드 스레드에 "
                "사업자등록증(이미지·PDF)을 먼저 첨부해주세요."
            )
        if not vat_ok:
            error_lines.append(
                "• :heavy_dollar_sign: *부가세 미지정* — 관리 사이트에서 프로젝트를 "
                "편집해 부가세(포함/미포함)를 지정해주세요."
            )

    if error_lines:
        _notify_invoice_submit_error(client, channel_id, user_id, code, error_lines)
        # 반려 시 lock 해제 — 매니저가 재제출 시 같은 view.id 로 오지 않지만 안전 차원
        if _view_id:
            try:
                from dashboard.utils.redis_client import get_redis_client
                get_redis_client().redis.delete(f'invoice_submit_lock:{_view_id}')
            except Exception:
                pass
        return

    _process_invoice_submission(client, body, view)


def _process_invoice_submission(client, body, view) -> None:
    """모달 제출 → #영업_관리 채널에 계산서 요청 카드 발송."""
    channel_id = os.getenv('SLACK_INVOICE_CHANNEL_ID', '').strip()
    if not channel_id:
        logger.warning('[SLACK/계산서] SLACK_INVOICE_CHANNEL_ID 미설정 — 발송 skip')
        return

    metadata = json.loads(view.get("private_metadata") or "{}")
    code = metadata.get("code", "-") or '-'

    values = view["state"]["values"]

    def _get(block_id):
        return (values.get(block_id, {}).get('value', {}) or {}).get('value', '') or ''

    biz = _get('biz').strip() or '-'
    addr = _get('addr').strip() or '-'
    amt_raw = _get('amt').strip()
    amt_digits = ''.join(ch for ch in amt_raw if ch.isdigit())
    email = _get('email').strip() or '-'
    memo = _get('memo').strip()

    # VAT radio_buttons state — 2026-07-16 라디오 필드 재도입 (매니저 오클릭 방지)
    _vat_state = (values.get('vat', {}).get('value', {}) or {}).get('selected_option') or {}
    vat_val = _vat_state.get('value', 'sep') or 'sep'
    vat_label = 'VAT 별도' if vat_val == 'sep' else 'VAT 포함'

    amt_display = _money_kr(amt_digits)
    if amt_display != '-':
        amt_display = f"{amt_display} ({vat_label})"

    user_id = body.get("user", {}).get("id", "")
    now_str = datetime.now().strftime('%m.%d %H:%M')

    # 카드 본문
    initial = _slack_user_to_initial(client, user_id) or '-'
    lines = [
        f"🔔 *[세금계산서 발행 요청]*  `{code}`",
        "--------------------------------------------",
        f"🏢 사업자명 : {biz}",
        f"📍 현장 주소 : {addr}",
        f"💲 금액 : {amt_display}",
        f"✉️ 이메일 : {email}",
    ]
    if memo:
        lines.append(f"📝 요청사항 : {memo}")
    lines.append(f"👤 요청자 : {initial}  {now_str}")
    lines.append("--------------------------------------------")
    text = '⠀\n' + '\n'.join(lines)

    # 발행 완료 버튼 value — 완료 문구 자동 생성용.
    # 원본 카드 텍스트를 함께 저장해서 완료 처리 시 header_context 를 이걸로 재구성
    # (chat.update 된 body.message.blocks 에서 뽑아쓰면 재클릭 시 중복 표시 발생).
    complete_value = json.dumps({
        'code': code,
        'amt': amt_digits,
        'biz': biz,
        'vat': vat_val,
        'orig': text,
    }, ensure_ascii=False)

    # 카드 발송 — 발행 완료 버튼은 제거 (2026-07-13 UX 개선).
    # 매니저가 스레드에 이미지/PDF 첨부하면 handle_thread_message 가 자동으로
    # 카드 헤더·첨부 상태를 완료 표시로 update. 버튼이 매니저에게 "이미 완료?" 오해를 줌.
    blocks = [
        {"type": "section", "text": {"type": "mrkdwn", "text": text}},
        {"type": "context", "elements": [{"type": "mrkdwn", "text": "⠀"}]},
    ]

    # 카드 발송은 세금계산서 관리 알림 봇 (invoice_bot) 으로. 없으면 공사봇 fallback.
    if _invoice_slack_app is None:
        _init_invoice_slack_app()
    invoice_client = _invoice_slack_app.client if _invoice_slack_app else client

    # 봇이 채널에 없으면 자동 가입 시도 (public 채널만 성공, private면 사용자가 초대 필요)
    try:
        invoice_client.conversations_join(channel=channel_id)
    except Exception:
        pass

    resp = invoice_client.chat_postMessage(
        channel=channel_id, text=text, blocks=blocks, unfurl_links=False,
    )
    if not resp.get('ok'):
        logger.warning(f"[SLACK/계산서] 요청 카드 발송 실패: {resp}")
        return

    ts = resp.get('ts', '')
    logger.info(
        f"[SLACK/계산서] 요청 카드 발송 완료: {code} ts={ts} → {channel_id}"
    )

    # 카드 하단에 '📎 계산서 첨부 (스레드 열기)' 링크 추가 + Redis 에 metadata 저장.
    # 매니저가 스레드에 파일 첨부 시 계산서봇 handler 가 이 metadata 로
    # 카드 자동 완료 update (2026-07-13 자동 완료 전환).
    thread_url = ''
    try:
        perm_resp = invoice_client.chat_getPermalink(channel=channel_id, message_ts=ts)
        if perm_resp.get('ok'):
            base_url = perm_resp.get('permalink', '') or ''
            if base_url:
                sep = '&' if '?' in base_url else '?'
                thread_url = f"{base_url}{sep}thread_ts={ts}&cid={channel_id}"
    except Exception as perm_exc:
        logger.warning(f"[SLACK/계산서] permalink 조회 실패 (링크 생략): {perm_exc}")

    # Redis 저장 (30일 TTL) — 스레드 첨부 감지 시 auto-complete 처리용
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        rc.setex(
            f'invoice_card:{channel_id}:{ts}',
            86400 * 30,
            json.dumps({
                'code': code, 'biz': biz, 'amt': amt_digits, 'vat': vat_val,
                'email': email, 'thread_url': thread_url, 'orig_text': text,
            }, ensure_ascii=False),
        )
    except Exception as red_exc:
        logger.warning(f"[SLACK/계산서] Redis metadata 저장 실패: {red_exc}")

    # 카드 update — 첨부 안내 라인 추가
    if thread_url:
        info_block = blocks[0]
        padding_block = blocks[-1]
        attach_link_block = {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": f'📎 세금계산서 : ⬜ 미첨부 <{thread_url}|(첨부하기)>',
            },
        }
        new_blocks = [info_block, attach_link_block, padding_block]
        try:
            invoice_client.chat_update(
                channel=channel_id, ts=ts, text=text, blocks=new_blocks,
            )
        except Exception as upd_exc:
            logger.warning(f"[SLACK/계산서] 첨부 링크 추가 실패 (무시): {upd_exc}")

    # 카드 스레드에 프로젝트 사업자등록증 canonical 파일 자동 첨부.
    # (2026-07-16 UX 개선 — 매니저가 공사확정 채널 왔다갔다 안 하도록 스레드에서 즉시 열람 가능.)
    # 실패해도 카드 발송은 유지 — 조용히 warning 로그만.
    def _attach_license_to_thread():
        try:
            from dashboard.services.business_license_handler import fetch_license_canonical
            lic = fetch_license_canonical(code)
            if not lic:
                logger.info(f"[SLACK/계산서] 사업자등록증 canonical 없음 → 스레드 첨부 skip ({code})")
                return
            invoice_client.files_upload_v2(
                channel=channel_id,
                thread_ts=ts,
                file=lic['content'],
                filename=lic['file_name'],
                initial_comment=f":page_facing_up: 사업자등록증 — `{code}`",
            )
            logger.info(f"[SLACK/계산서] 사업자등록증 스레드 첨부 완료: {code} ({lic['file_name']})")
        except Exception as exc:
            logger.warning(f"[SLACK/계산서] 사업자등록증 스레드 첨부 실패 (무시, {code}): {exc}")

    threading.Thread(target=_attach_license_to_thread, daemon=True).start()


def _auto_complete_invoice_card(
    client, channel: str, message_ts: str, event: dict, meta: dict,
) -> bool:
    """스레드에 이미지/PDF 첨부 감지 → 카드 자동 완료 update.

    반환: 실제 update 됐는지 여부.
    """
    # 이미지/PDF 필터
    valid_files = []
    for f in (event.get('files') or []):
        mt = f.get('mimetype', '') or ''
        if mt.startswith('image/') or mt == 'application/pdf':
            valid_files.append(f)
    if not valid_files:
        return False

    user_id = event.get('user', '') or ''
    initial = _slack_user_to_initial(client, user_id) or '-'
    now_str = datetime.now().strftime('%m.%d %H:%M')

    orig_text = meta.get('orig_text', '') or ''
    thread_url = meta.get('thread_url', '') or ''

    # 헤더 : 🔔 요청 → ✅ 완료
    updated_text = orig_text.replace(
        '🔔 *[세금계산서 발행 요청]*',
        '✅ *[세금계산서 발행 완료]*',
        1,
    )
    # 완료 처리 라인 추가 (마지막 구분선 앞)
    _SEP = '--------------------------------------------'
    completed_line = f'✅ 처리자 : {initial}  {now_str}'
    parts = updated_text.rsplit(_SEP, 1)
    if len(parts) == 2:
        updated_text = parts[0].rstrip() + '\n' + completed_line + '\n' + _SEP + parts[1]
    else:
        updated_text += '\n' + completed_line

    # 첨부 상태 : ⬜ 미첨부 → ✅ 첨부됨 / (첨부하기) → (확인하기)
    if thread_url:
        attach_text = f'📎 세금계산서 : ✅ 첨부됨 <{thread_url}|(확인하기)>'
    else:
        attach_text = f'📎 세금계산서 : ✅ 첨부됨'
    attach_block = {'type': 'section', 'text': {'type': 'mrkdwn', 'text': attach_text}}

    new_blocks = [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': updated_text}},
        attach_block,
    ]

    # 2026-07-16: 첨부된 이미지 첫 파일을 카드에 image block 으로 embed (매니저 UX 요청).
    # PDF 는 image block 미지원 (skip). 이미지만 대상. files.sharedPublicURL 호출로
    # 파일을 public 화 → url_private + ?pub_secret=... 로 image_url 조합.
    _preview_file = next((f for f in valid_files if (f.get('mimetype') or '').startswith('image/')), None)
    if _preview_file:
        try:
            _fid = _preview_file.get('id')
            _perm_pub = _preview_file.get('permalink_public') or ''
            if not _perm_pub:
                # 아직 공개 안 됨 → sharedPublicURL 호출 (files:write scope 필요)
                _shared = client.files_sharedPublicURL(file=_fid)
                _file_info = (_shared.get('file') or {}) if _shared else {}
                _perm_pub = _file_info.get('permalink_public') or ''
                _url_private = _file_info.get('url_private') or _preview_file.get('url_private') or ''
            else:
                _url_private = _preview_file.get('url_private') or ''
            if _perm_pub and _url_private:
                _pub_secret = _perm_pub.rsplit('-', 1)[-1]
                _image_url = f'{_url_private}?pub_secret={_pub_secret}'
                new_blocks.append({
                    'type': 'image',
                    'image_url': _image_url,
                    'alt_text': '세금계산서 미리보기',
                })
        except Exception as _prev_exc:
            logger.warning(f"[SLACK/계산서] 미리보기 image block 추가 실패 (계속 진행): {_prev_exc}")

    new_blocks.append(
        {'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': '⠀'}]},
    )

    try:
        client.chat_update(
            channel=channel, ts=message_ts,
            text=(f"✅ 세금계산서 발행 완료 · {meta.get('code','')} · "
                  f"{meta.get('biz','')}"),
            blocks=new_blocks,
        )
        logger.info(
            f"[SLACK/계산서] 자동 완료 update: {meta.get('code','')} by {initial}"
        )
        return True
    except Exception as exc:
        logger.warning(f"[SLACK/계산서] 자동 완료 update 실패: {exc}")
        return False


def _process_invoice_complete(client, body) -> None:
    """[✅ 발행 완료] 클릭 처리 — 스레드 파일 첨부 검증 + 카드 회색화 + 확인 메시지."""
    channel = body["channel"]["id"]
    message_ts = body["message"]["ts"]
    user_id = body["user"]["id"]

    try:
        payload = json.loads(body["actions"][0].get("value") or "{}")
    except Exception:
        payload = {}
    code = payload.get('code', '-') or '-'
    amt_digits = payload.get('amt', '') or ''
    biz = payload.get('biz', '-') or '-'
    vat_val = payload.get('vat', 'sep') or 'sep'
    vat_label = 'VAT 별도' if vat_val == 'sep' else 'VAT 포함'
    orig_payload_text = payload.get('orig', '') or ''

    # 0) 이미 완료 처리된 카드 재클릭 방지 (2026-07-13 관측 R-TEST-KiKO 중복 표시).
    #    감지: message.text 가 완료 카드 fallback text 로 시작하거나 (신규 구조),
    #    blocks 첫 블록이 context (구 header_context 구조) 이면 skip.
    _msg = body.get('message', {}) or {}
    _msg_text = _msg.get('text', '') or ''
    _blocks = _msg.get('blocks') or []
    _already_done = (
        _msg_text.startswith('✅ 세금계산서 발행 완료')
        or (bool(_blocks) and _blocks[0].get('type') == 'context')
    )
    if _already_done:
        try:
            client.chat_postEphemeral(
                channel=channel, user=user_id,
                text=':information_source: 이미 발행 완료 처리된 카드입니다.',
            )
        except Exception:
            pass
        logger.info(f"[SLACK/계산서] 중복 클릭 skip ({code}) by {user_id}")
        return

    # 1) 스레드에 첨부 파일 있는지 검증 + 파일 정보 수집 (2026-07-10)
    has_file = False
    attached_files = []  # [{id, name, permalink, mimetype}]
    try:
        replies = client.conversations_replies(
            channel=channel, ts=message_ts, limit=200,
        )
        for m in replies.get('messages', [])[1:]:  # root 제외
            for f in (m.get('files') or []):
                if not f.get('id'):
                    continue
                has_file = True
                attached_files.append({
                    'id': f.get('id') or '',
                    'name': f.get('name') or f.get('title') or '첨부파일',
                    'permalink': f.get('permalink') or '',
                    'mimetype': f.get('mimetype') or '',
                })
    except Exception as exc:
        logger.warning(f"[SLACK/계산서] replies 조회 실패: {exc}")

    if not has_file:
        # 첨부 없음 — ephemeral로 안내 후 skip
        try:
            client.chat_postEphemeral(
                channel=channel, user=user_id,
                text=(
                    ':warning: 세금계산서 이미지/PDF를 먼저 이 스레드에 첨부한 뒤 '
                    '[✅ 발행 완료] 버튼을 눌러주세요.'
                ),
            )
        except Exception:
            pass
        logger.info(f"[SLACK/계산서] 첨부 없음 → 완료 skip ({code}) by {user_id}")
        return

    amt_display = _money_kr(amt_digits)
    biz_display = biz if biz and biz != '-' else '(사업자명 미기재)'
    initial_for_msg = _slack_user_to_initial(client, user_id) or '-'

    # 원본 요청 카드 텍스트 → context block 으로 감싸 회색 톤 + 폰트 축소.
    # payload.orig 를 우선 사용 (submission 시점 저장) — body.message.blocks 에서
    # 뽑으면 재클릭 시 이미 완료 카드 body_text 가 잡혀 중복 표시됨.
    original_text = orig_payload_text
    if not original_text:
        for b in (body.get('message', {}).get('blocks') or []):
            if b.get('type') == 'section':
                original_text = (b.get('text', {}) or {}).get('text', '') or ''
                break
    if not original_text:
        # fallback — payload 만 가지고 최소 정보 구성
        original_text = (
            f':bell: *[세금계산서 발행 요청]*  `{code}`\n'
            f':office: 사업자명 : {biz_display}\n'
            f':heavy_dollar_sign: 금액 : {amt_display} ({vat_label})'
        )

    # 스레드 permalink — 첨부는 프리뷰 대신 스레드 이동 링크로 (2026-07-13)
    thread_url = ''
    try:
        perm = client.chat_getPermalink(channel=channel, message_ts=message_ts)
        if perm.get('ok'):
            _base = perm.get('permalink', '') or ''
            if _base:
                _sep = '&' if '?' in _base else '?'
                thread_url = f'{_base}{_sep}thread_ts={message_ts}&cid={channel}'
    except Exception as exc:
        logger.warning(f'[SLACK/계산서] permalink 조회 실패 (링크 생략): {exc}')

    if thread_url:
        files_text = (
            f':paperclip: *첨부 파일* : '
            f'{len(attached_files)}개  <{thread_url}|(확인 하기)>'
        )
    else:
        files_text = f':paperclip: *첨부 파일* : {len(attached_files)}개'

    body_text = (
        f':white_check_mark: *세금계산서 발행 완료*  `{code}`\n'
        f'━━━━━━━━━━━━━━━━━━━━\n'
        f':office: *사업자명* : {biz_display}\n'
        f':moneybag: *발행 금액* : *{amt_display}*  ({vat_label})\n'
        f':bust_in_silhouette: *처리자* : {initial_for_msg}\n'
        f'{files_text}'
    )

    # 원본 요청 카드 — 코드 블록 (```...```) 으로 감싸 monospace 회색 박스로
    # 표시 (방문 취소 카드와 동일 스타일, slack_bot.py:4133 참조).
    # mrkdwn 강조(*) 제거 + shortcode → 유니코드 정규화 (코드블록 안에서는 raw 로 보이므로).
    from dashboard.blueprints.slack_helpers import _normalize_shortcodes_to_unicode
    _cleaned = [ln.replace('*', '') for ln in original_text.split('\n')]
    _cleaned = [_normalize_shortcodes_to_unicode(ln) for ln in _cleaned]
    _clean_original = '\n'.join(_cleaned).strip()
    combined_text = (
        f'{body_text}\n\n'
        f'```\n{_clean_original}\n```'
    )

    completed_blocks = [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': combined_text}},
    ]

    try:
        client.chat_update(
            channel=channel, ts=message_ts,
            text=f"✅ 세금계산서 발행 완료 · {code} · {amt_display} ({vat_label}) · {biz_display}",
            blocks=completed_blocks,
        )
    except Exception as exc:
        logger.warning(f"[SLACK/계산서] chat.update 실패 ({code}): {exc}")

    logger.info(f"[SLACK/계산서] 발행 완료: {code} by {user_id}")


# 앱 시작 시 한 번 초기화 시도
_init_slack_app()
_init_visit_slack_app()
_init_invoice_slack_app()
