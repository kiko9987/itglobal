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

_slack_app = None
_slack_handler = None
_project_slack_app = None
_project_slack_handler = None
_visit_slack_app = None
_visit_slack_handler = None
_as_slack_app = None
_as_slack_handler = None

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

    @app.action("visit_cancel")
    def handle_visit_cancel(ack, body, client):
        ack()
        def _bg():
            try:
                lead_no = body["actions"][0].get("value") or ''
                if not _try_acquire_action_lock(lead_no, 'cancel'):
                    logger.info(f'[SLACK/방문봇] visit_cancel 중복 클릭 skip ({lead_no})')
                    return
                _process_visit_cancel(client, body)
            except Exception as exc:
                logger.error(f"[SLACK/방문봇] visit_cancel 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.action("visit_complete")
    def handle_visit_complete(ack, body, client):
        ack()
        def _bg():
            try:
                lead_no = body["actions"][0].get("value") or ''
                if not _try_acquire_action_lock(lead_no, 'complete'):
                    logger.info(f'[SLACK/방문봇] visit_complete 중복 클릭 skip ({lead_no})')
                    return
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


def _register_project_handlers(app):
    """공사 현황 알림 봇 핸들러 — /공사확정 + submit_project"""

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
        # 사업자등록증 첨부 여부 검증 (모달 오픈 시점 대신 여기서 — trigger_id 만료 방지).
        # 미첨부면 modal errors 로 반려해 사용자가 스레드에 첨부 후 재제출.
        try:
            metadata = json.loads(view.get("private_metadata") or "{}")
            code = (metadata.get("code", "") or "").strip()
            if code and code != '-':
                from dashboard.services.business_license_handler import verify_license_exists
                if not verify_license_exists(code):
                    ack(response_action="errors", errors={
                        "biz": "사업자등록증이 아직 첨부되지 않았습니다. "
                               "카드 스레드에 사업자등록증(이미지 or PDF)을 첨부한 뒤 다시 요청해주세요.",
                    })
                    return
        except Exception as exc:
            # 검증 자체가 실패해도(예: Drive 지연) 발행 요청은 통과시킴 — 관리자가 후속 처리
            logger.warning(f'[SLACK/계산서] 사업자등록증 검증 실패 (통과): {exc}')

        ack()
        def _bg():
            try:
                _process_invoice_submission(client, body, view)
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
        # 스레드 파일 첨부만 처리 + 봇 자신 메시지 skip (bot_message subtype 등)
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


def _register_as_handlers(app):
    """A/S 사후 관리 봇 핸들러 — /as + 3단계 모달 흐름."""

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


def _build_as_card_text(data: dict, view_state: str = 'requested') -> str:
    """A/S 카드 본문 텍스트. view_state: requested / accepted / completed.

    공사 확정 카드와 동등한 정보량으로 렌더 — 유입 구분·발주처 담당자/연락처/이메일·
    도급 구분·시공자·공사 금액·공사 시작 추가.
    """
    # 프로젝트 상세 조회 (부족한 필드 채움 — 시트엔 없어도 카드엔 표시).
    proj = None
    code = str(data.get('프로젝트 코드', '') or '').strip()
    if code and code != '-':
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
        lines.append(f"✅ 접수자 : {data.get('접수자', '-') or '-'}  _{data.get('접수 일자', '')}_")
    if view_state == 'completed':
        lines.append("--------------------------------------------")
        lines.append(f"🎯 처리 내용 : {data.get('처리 내용', '-') or '-'}")
    lines.append("--------------------------------------------")
    return "⠀\n" + "\n".join(lines)


def _build_as_blocks(data: dict, view_state: str = 'requested') -> list:
    text = _build_as_card_text(data, view_state=view_state)
    # section 하단 구분선(-----)과 버튼 사이 여백 제거 (2026-07-09 UX).
    blocks: list = [
        {"type": "section", "text": {"type": "mrkdwn", "text": text}},
    ]
    as_no = data.get('No', '')
    if view_state == 'requested':
        blocks.append({
            "type": "actions",
            "elements": [{
                "type": "button",
                "text": {"type": "plain_text", "text": "🛠️ A/S 접수하기", "emoji": True},
                "style": "primary",
                "action_id": "as_accept_open",
                "value": as_no,
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
    """
    trigger_id = body["trigger_id"]
    as_no = (body["actions"][0].get("value") or '').strip()
    channel = body.get("channel", {}).get("id", "")
    message_ts = body.get("message", {}).get("ts", "")

    metadata = json.dumps({
        "as_no": as_no, "channel": channel, "message_ts": message_ts,
    }, ensure_ascii=False)

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
                "hint": {"type": "plain_text", "text": "서비스 기사 방문 시 작성 X"},
                "element": {
                    "type": "plain_text_input", "action_id": "value",
                    "placeholder": {"type": "plain_text", "text": "예: 김철수"},
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
    data = get_as_data(as_no) or {}
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

    data = get_as_data(as_no) or {}
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

        # 시트 반영 — 빈값/'미정'이면 '-' 로 초기화
        from dashboard.services.lead_service import update_lead
        if not assignee_name or assignee_name in ('미정', '-'):
            new_value = '-'
        else:
            new_value = assignee_name

        try:
            update_lead(lead_no, {'영업 담당자': new_value})
            logger.info(f"[SLACK/LIST] 시트 반영: {lead_no} 영업 담당자 → {new_value!r}")
            return jsonify({"ok": True, "lead_no": lead_no, "assignee": new_value})
        except Exception as exc:
            logger.error(f"[SLACK/LIST] 시트 반영 실패 ({lead_no}): {exc}", exc_info=True)
            return jsonify({"ok": False, "error": str(exc)}), 500

    except Exception as exc:
        logger.error(f"[SLACK/LIST] list-assignee 처리 오류: {exc}", exc_info=True)
        return jsonify({"ok": False, "error": str(exc)}), 500


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
def _find_lead_by_no(lead_no: str):
    """리드 No로 메인 시트 행 dict 반환"""
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

    # 상담 내용에서 장소/기기/문의 분리
    parts = _split_lead_content(str(lead.get('문의 내용', '') or lead.get('상담 내용', '')))
    name = str(lead.get('고객명') or '').strip() or '-'
    phone = str(lead.get('고객 연락처') or '').strip() or '-'
    email = str(lead.get('이메일') or '').strip() or '-'
    place = parts['place'] or '-'
    device = parts['device'] or '-'
    inquiry = parts['inquiry'] or str(lead.get('문의 내용') or lead.get('상담 내용') or '').strip() or '-'
    address = str(lead.get('방문 주소') or '').strip()
    consult_time = str(lead.get('상담 시간') or '').strip() or '-'

    # 모달 상단 - 원본 인입 정보 표시 (옛 Apps Script 패턴)
    info_blocks = [
        {"type": "section", "text": {"type": "mrkdwn",
                                      "text": f"*접수번호:* `{lead_no}`"}},
        {"type": "section", "text": {"type": "mrkdwn",
                                      "text": f"*문의시간 :* {consult_time}"}},
        {"type": "section", "text": {"type": "mrkdwn",
                                      "text": f"*이름 / 상호 :* {name}"}},
        {"type": "section", "text": {"type": "mrkdwn",
                                      "text": f"*연락처 :* {phone}"}},
        {"type": "section", "text": {"type": "mrkdwn",
                                      "text": f"*이메일 :* {email}"}},
        {"type": "section", "text": {"type": "mrkdwn",
                                      "text": f"*설치 희망 장소 :* {place}"}},
        {"type": "section", "text": {"type": "mrkdwn",
                                      "text": f"*설치 희망 기기 :* {device}"}},
        # 문의 내용은 3000자 제한 대응 — 넘치면 자동 truncate + 안내
        {"type": "section", "text": {"type": "mrkdwn",
                                      "text": slack_truncate(f"*문의 내용 :*\n{inquiry}")}},
        {"type": "divider"},
    ]

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
                "label": {"type": "plain_text", "text": "내용 / 특이사항"},
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
    # 2026-07-12 datepicker 표시 원인 확인 위한 임시 revert — 이전 placeholder +
    #   views_update 방식으로 되돌림. mobile 표시 vs datepicker 로케일 트레이드오프.
    placeholder = {
        "type": "modal",
        "callback_id": "submit_consult",
        "title": {"type": "plain_text", "text": "상담 처리"},
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
    """상담 모달 상단 인입 정보 블록 — lead 있으면 카드형 정보, 없고 lead_no만 있으면 경고."""
    if lead:
        parts = _split_lead_content(str(lead.get('문의 내용', '') or lead.get('상담 내용', '')))
        name = str(lead.get('고객명') or '').strip() or '-'
        phone = str(lead.get('고객 연락처') or '').strip() or '-'
        email = str(lead.get('이메일') or '').strip() or '-'
        consult_time = str(lead.get('상담 시간') or '').strip() or '-'
        inquiry = parts.get('inquiry') or str(lead.get('문의 내용') or lead.get('상담 내용') or '').strip() or '-'
        return [
            {"type": "section", "text": {"type": "mrkdwn", "text": (
                f"*접수번호:* `{lead_no}`\n"
                f"*문의시간:* {consult_time}\n"
                f"*이름 / 상호:* {name}\n"
                f"*연락처:* {phone}\n"
                f"*이메일:* {email}\n"
                f"*상세 문의:* {inquiry[:300]}"
            )}},
            {"type": "divider"},
        ]
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

    is_visit = (status == '방문 예약')
    is_estimate = (status == '견적 제출')

    # 두 차원 매핑 (시트 컬럼)
    category = visit_type   # 플랫폼 컬럼 = 방문 유형
    sheet_status = status   # 상태 컬럼 = 처리 유형

    # ─────────────────────────────────────────────
    # 1) 인입 리드 케이스 — 기존 lead 시트 업데이트
    # ─────────────────────────────────────────────
    if lead_no:
        try:
            from dashboard.services.lead_service import update_lead
            update_data = {'상태': sheet_status}
            if is_visit and visit_date_for_sheet:
                update_data['방문 예정일'] = visit_date_for_sheet
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
                # 옛 상담 내용은 보존 — 매니저 추가 입력은 피드백 컬럼에 저장
                update_data['상담 내용'] = consultation
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
    if is_visit:
        # 인입 lead의 플랫폼 (홈페이지/당근/카카오톡/전화) — 헤더에 부가 표시
        lead_platform = ''
        if lead_no:
            existing_lead = _find_lead_by_no(lead_no) or {}
            lead_platform = str(existing_lead.get('플랫폼', '')).strip()
        _post_visit_notice(
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
        # 1) 원본 카드 본문 회색 박스 변환 (부재중은 추후 재상담 필요 → 변환 skip)
        original_text = metadata.get("original_text", "") if isinstance(metadata, dict) else ''
        if original_text and status != '부재중':
            try:
                cancel_time = datetime.now().strftime('%m.%d %H:%M')
                initial = _slack_user_to_initial(client, user_id) or '-'
                cleaned_lines = [ln.lstrip('>').lstrip() for ln in original_text.split('\n')]
                cleaned_lines = [ln.replace('*', '') for ln in cleaned_lines]
                clean_text = '\n'.join(cleaned_lines)
                clean_text = re.sub(r'^[\s⠀]+|[\s⠀]+$', '', clean_text)
                header_lines = [
                    "⠀",
                    f":white_check_mark: *상담 완료 - {status}*",
                    f"처리자 : {initial}",
                    f"처리 시간 : {cancel_time}",
                ]
                if consultation:
                    header_lines.append(f"상담내용 : {consultation}")
                new_text = '\n'.join(header_lines) + f"\n\n```\n{clean_text}\n```"
                new_blocks = [
                    {"type": "section", "text": {"type": "mrkdwn", "text": new_text}},
                ]
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
        # chat.update 실패 시(옛 ts 삭제됐거나) chat.postMessage fallback
        reply_sent = False
        if old_reply_ts:
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
                       user_name: str = '', platform: str = '') -> tuple:
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
    if platform and platform != category:
        category_display = f"{category} ({platform})"
    else:
        category_display = category

    body_text, blocks = _build_visit_notice_blocks(
        lead_no=lead_no, category_display=category_display, initial=initial,
        visit_date=visit_date, name=name, contact=contact,
        visit_address=visit_address, consultation=consultation,
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
        return (visit_channel, ts)
    except Exception as exc:
        logger.warning(f"[SLACK/방문] #방문_일정 발송 실패: {exc}")
        return ('', '')


def _build_visit_notice_blocks(lead_no: str, category_display: str, initial: str,
                                visit_date: str, name: str, contact: str,
                                visit_address: str, consultation: str) -> tuple:
    """방문 일정 카드 양식 빌더 — (text, blocks) 반환.

    [✏️ 방문일 수정] + [🗑️ 방문 취소] 액션 버튼 포함. 카드 발송/복원 양쪽에서 재사용.
    """
    SEP = '--------------------------------------------'
    # lead_no 없으면 (거래처/기타) 헤더에 표시 안 함
    header_suffix = f"  `{lead_no}`" if lead_no else ''
    lines = [
        "⠀",
        f">:bell: *새 방문 일정* — {category_display}{header_suffix}",
        f">{SEP}",
        f">등록자 : {initial or '-'}",
        f">방문일 : {visit_date or '-'}",
        f">이름 / 상호 : {name or '-'}",
        f">연락처 : {contact or '-'}",
        f">방문 주소 : {visit_address or '-'}",
    ]
    if consultation:
        lines.append(f">상담 내용 :")
        for raw in consultation[:500].split('\n'):
            wrapped = textwrap.fill(
                raw, width=60, break_long_words=True, break_on_hyphens=False,
            ) or raw
            for ln in wrapped.split('\n'):
                lines.append(f">{ln}")
    lines.append(f">{SEP}")
    body_text = '\n'.join(lines)
    blocks = [
        {"type": "section", "text": {"type": "mrkdwn", "text": body_text}},
        {
            "type": "actions",
            "elements": [
                {
                    "type": "button",
                    "text": {"type": "plain_text", "text": "✏️ 방문일 수정", "emoji": True},
                    "value": lead_no,
                    "action_id": "visit_modify_date",
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
                    "confirm": {
                        "title": {"type": "plain_text", "text": "방문 취소"},
                        "text": {"type": "plain_text",
                                 "text": "이 방문 일정을 취소하시겠습니까?"},
                        "confirm": {"type": "plain_text", "text": "취소 확정"},
                        "deny": {"type": "plain_text", "text": "되돌리기"},
                    },
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

    # 1) 시트 update — escape prefix로 시리얼 변환 차단
    try:
        from dashboard.services.lead_service import update_lead
        # 단일이면 escape prefix, 범위는 그대로 (Google Sheets가 텍스트로 인식)
        sheet_value = f"'{new_date_display}" if '~' not in new_date_display else new_date_display
        update_lead(lead_no, {'방문 예정일': sheet_value})
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

    logger.info(f"[SLACK/방문완료] 처리 완료: {lead_no} by {user_id}")


def _process_visit_cancel(client, body) -> None:
    """[🚫 방문 취소] 클릭 처리 — 시트 상태='공사 취소' + 메시지 chat.update."""
    lead_no = body["actions"][0].get("value") or ''
    channel = body["channel"]["id"]
    message_ts = body["message"]["ts"]
    user_id = body["user"]["id"]
    if not lead_no:
        return

    # 1) 시트 상태='공사 취소'
    try:
        from dashboard.services.lead_service import update_lead
        update_lead(lead_no, {'상태': '공사 취소'})
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
            f"🚫 *고객 요청으로 방문 취소*  `{lead_no}`\n"
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

    m = re.search(r'L-\d{5}', root_text)
    if not m:
        logger.info("[SLACK/방문 사진] thread root에 lead_no 없음 — 스킵")
        return
    lead_no = m.group(0)

    # Race condition 방어 (2026-07-10)
    # 두 매니저가 동시에 같은 lead thread 에 사진 첨부 시 각 daemon 스레드가 동시에
    # find_or_create_folder 를 호출 → Google Drive eventual consistency 로 폴더 중복 생성.
    # 60초 리드별 락 (사진 업로드 자체는 보통 5~15초).
    from dashboard.utils.redis_client import get_redis_client as _get_rc
    _lock_key = f'visit_photo_lock:{lead_no}'
    try:
        _rc_lock = _get_rc().redis
        _got_lock = _rc_lock.set(_lock_key, '1', nx=True, ex=60)
        if not _got_lock:
            logger.info(
                f'[SLACK/방문 사진] {lead_no} 다른 스레드가 처리 중 — 이번 이벤트 skip'
            )
            return
    except Exception as exc:
        logger.warning(f'[SLACK/방문 사진] 락 획득 실패 — 계속 진행: {exc}')
        _rc_lock = None

    # 2) 폴더명 생성 — "({이니셜}) {방문 주소} {YY.MM.DD}"
    lead = _find_lead_by_no(lead_no) or {}
    # 이니셜 — 플랫폼별 규칙 (2026-07):
    #   거래처/기타/소개: 카드 생성자(온라인 상담자) 기준
    #   온라인(그 외): List 배정 담당자(=영업 담당자) 우선, fallback 온라인 상담자
    #   최종 fallback: 카드 "등록자 :" 정규식 → '미상'
    def _clean(v):
        s = str(v or '').strip()
        return '' if s in ('', '-', '미정') else s

    lead_platform = str(lead.get('플랫폼', '')).strip()
    if lead_platform in ('거래처', '기타', '소개'):
        source_name = _clean(lead.get('온라인 상담자'))
    else:
        source_name = _clean(lead.get('영업 담당자')) or _clean(lead.get('온라인 상담자'))
    initial = _to_initial(source_name) if source_name else ''
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

    uploaded = 0
    for f in files:
        download_url = f.get('url_private_download') or f.get('url_private')
        if not download_url:
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
            continue

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
                if length and length > _MAX_PHOTO_BYTES:
                    logger.warning(
                        f"[SLACK/방문 사진] Content-Length 초과 skip: "
                        f"{filename} = {length / 1024 / 1024:.1f}MB"
                    )
                    continue
                # 스트림 상한
                content = r.read(_MAX_PHOTO_BYTES + 1)
                if len(content) > _MAX_PHOTO_BYTES:
                    logger.warning(
                        f"[SLACK/방문 사진] 스트림 크기 초과 skip: "
                        f"{filename} > 100MB"
                    )
                    continue
            if upload_file(folder_id, filename, content, mimetype=mimetype):
                uploaded += 1
        except Exception as exc:
            logger.error(f"[SLACK/방문 사진] 다운로드/업로드 실패 ({filename}): {exc}",
                         exc_info=True)

    if uploaded == 0:
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

    # 4-2) 리드 시트 P열 (폴더 ID) 영구 저장 — Redis TTL 만료 대비 + source of truth
    try:
        from dashboard.services.lead_service import update_lead
        update_lead(lead_no, {'폴더 ID': lead_folder['id']})
        logger.info(f"[SLACK/방문 사진] 시트 P열 폴더 ID 저장: {lead_no} → {lead_folder['id']}")
    except Exception as exc:
        logger.warning(f"[SLACK/방문 사진] 시트 P열 저장 실패 ({lead_no}): {exc}")

    # 5) thread 답글
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
        reply_text = (
            f":file_folder: 사진 {uploaded}장을 드라이브에 저장했습니다{location_suffix}.\n"
            f"📁 {folder_name}"
            f"{win_path_line}\n"
            f":id: *폴더 ID* (새 프로젝트 등록용) :\n"
            f"```{lead_folder['id']}```\n"
            f">*상호명 추가* : 답글에 \"상호 OOO\" 입력\n"
            f">*위치 분류* : 사진 첨부시 댓글에 \"1층\" 등 함께 입력"
        )
        client.chat_postMessage(
            channel=channel, thread_ts=thread_ts, text=reply_text,
            unfurl_links=False,
        )
    except Exception as exc:
        logger.warning(f"[SLACK/방문 사진] thread 답글 실패: {exc}")

    # 6) thread root(방문 일정 카드)에 ✅ reaction 추가 — 사진 등록 완료 표시
    try:
        client.reactions_add(
            channel=channel, timestamp=thread_ts, name="white_check_mark",
        )
    except Exception as exc:
        # 이미 reaction 있으면 already_reacted — 무시
        logger.debug(f"[SLACK/방문 사진] reaction 추가 스킵: {exc}")

    # 7) 락 해제 — 이후 첨부는 즉시 처리 가능하도록 (자연 만료 60초 대신 즉시)
    try:
        if _rc_lock:
            _rc_lock.delete(_lock_key)
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

    # 1) 시트 상태 → '방문 예약' 복원
    try:
        from dashboard.services.lead_service import update_lead
        update_lead(lead_no, {'상태': '방문 예약'})
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
        f'👤 수정자 : {initial}  _{now_str}_',
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
        f'👤 취소자 : {initial}  _{now_str}_',
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


def _open_invoice_modal(client, body) -> None:
    """[💰 계산서 요청] 클릭 → 프로젝트 정보 pre-fill 모달 오픈.

    2026-07-09: trigger_id 만료(3초) 방지 — Drive API 검증은 모달 오픈 후
    submit 시점으로 이동. 여기선 오직 pre-fill 후 즉시 views.open 만.
    """
    trigger_id = body["trigger_id"]
    action = body["actions"][0]
    try:
        payload = json.loads(action.get("value") or "{}")
    except Exception:
        payload = {}

    code = payload.get('code', '') or '-'
    biz = payload.get('biz', '') or ''
    addr = payload.get('addr', '') or ''
    amt = payload.get('amt', '') or ''
    # pre-fill 시 콤마 자동 포맷 (사용자 가독성)
    if amt.isdigit():
        amt = f"{int(amt):,}"
    vat = payload.get('vat', 'sep')  # 'sep' | 'incl'
    email = payload.get('email', '') or ''

    metadata = json.dumps({"code": code}, ensure_ascii=False)

    def _text_input(block_id, label, value, placeholder='', multiline=False, optional=False):
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

    vat_options = [
        {"text": {"type": "plain_text", "text": "VAT 별도"}, "value": "sep"},
        {"text": {"type": "plain_text", "text": "VAT 포함"}, "value": "incl"},
    ]
    vat_initial = vat_options[0] if vat == 'sep' else vat_options[1]

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
            _text_input("biz", "사업자명", biz, "예: (주)크리스아이티"),
            _text_input("addr", "현장 주소", addr, "예: 용인 수지구 포은대로59번길 37 1001호"),
            _text_input("amt", "금액", amt, "예: 4,900,000 (콤마·공백 무시됨)"),
            {
                "type": "input", "block_id": "vat",
                "label": {"type": "plain_text", "text": "부가세 처리"},
                "element": {
                    "type": "radio_buttons", "action_id": "value",
                    "options": vat_options,
                    "initial_option": vat_initial,
                },
            },
            _text_input("email", "발행 이메일", email, "예: crissit23@crissit.com"),
            _text_input("memo", "추가 요청사항", "", "선택 사항", multiline=True, optional=True),
        ],
    }
    client.views_open(trigger_id=trigger_id, view=view)


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

    vat_selected = (
        values.get('vat', {}).get('value', {}).get('selected_option', {})
    )
    vat_val = vat_selected.get('value', 'sep') if vat_selected else 'sep'
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

    blocks = [
        {"type": "section", "text": {"type": "mrkdwn", "text": text}},
        {
            "type": "actions",
            "elements": [
                {
                    "type": "button",
                    "text": {"type": "plain_text", "text": "✅ 발행 완료", "emoji": True},
                    "action_id": "invoice_complete",
                    "value": complete_value,
                    "style": "primary",
                },
            ],
        },
        {"type": "context", "elements": [{"type": "mrkdwn", "text": "⠀"}]},
    ]

    # 봇이 채널에 없으면 자동 가입 시도 (public 채널만 성공, private면 사용자가 초대 필요)
    try:
        client.conversations_join(channel=channel_id)
    except Exception:
        pass

    resp = client.chat_postMessage(
        channel=channel_id, text=text, blocks=blocks, unfurl_links=False,
    )
    if resp.get('ok'):
        ts = resp.get('ts', '')
        logger.info(
            f"[SLACK/계산서] 요청 카드 발송 완료: {code} ts={ts} → {channel_id}"
        )
        # 카드 하단에 '📎 계산서 첨부 (스레드 열기)' 링크 추가 (2026-07-10)
        # 회계 매니저(샛별)가 링크 클릭 → 자동으로 이 카드 스레드로 이동 → 파일 첨부
        try:
            perm_resp = client.chat_getPermalink(channel=channel_id, message_ts=ts)
            if perm_resp.get('ok'):
                base_url = perm_resp.get('permalink', '')
                if base_url:
                    sep = '&' if '?' in base_url else '?'
                    thread_url = f"{base_url}{sep}thread_ts={ts}&cid={channel_id}"
                    # 블록 순서: [info] → [📎 첨부 링크] → [✅ 발행 완료 버튼] → [padding]
                    # (첨부가 완료보다 먼저 나오게 — 자연스러운 액션 순서)
                    info_block = blocks[0]      # section (사업자·주소·금액 등)
                    actions_block = blocks[1]   # actions (발행 완료 버튼)
                    padding_block = blocks[-1]  # context (⠀)
                    # 사업자등록증 첨부 안내와 동일한 폰트/패턴 (section 블록, 상태 + (첨부하기) 링크)
                    attach_link_block = {
                        "type": "section",
                        "text": {
                            "type": "mrkdwn",
                            "text": f'📎 세금계산서 : ⬜ 미첨부 <{thread_url}|(첨부하기)>',
                        },
                    }
                    # 첨부 라인과 발행 완료 버튼 사이에 여백 한 줄
                    # (사업자등록증-버튼 사이 여백 패턴 project_slack_notifier.py:218 참조)
                    spacer_block = {'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': '⠀'}]}
                    new_blocks = [info_block, attach_link_block, spacer_block, actions_block, padding_block]
                    client.chat_update(
                        channel=channel_id, ts=ts, text=text, blocks=new_blocks,
                    )
        except Exception as perm_exc:
            logger.warning(f"[SLACK/계산서] permalink 링크 추가 실패 (무시): {perm_exc}")
    else:
        logger.warning(f"[SLACK/계산서] 요청 카드 발송 실패: {resp}")


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
