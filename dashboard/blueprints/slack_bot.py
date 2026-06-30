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

_slack_app = None
_slack_handler = None
_project_slack_app = None
_project_slack_handler = None
_visit_slack_app = None
_visit_slack_handler = None

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

        logger.info("[SLACK] 봇 초기화 완료 ✅")
        return True

    except Exception as exc:
        logger.error(f"[SLACK] 봇 초기화 실패: {exc}", exc_info=True)
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
        logger.info("[SLACK/방문봇] 초기화 완료 ✅")
        return True
    except Exception as exc:
        logger.error(f"[SLACK/방문봇] 초기화 실패: {exc}", exc_info=True)
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
    """방문 일정 알림 봇 핸들러 — [✏️ 방문일 수정] / [🗑️ 방문 취소]"""

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
                try:
                    msg_text = body["message"].get("text", "")
                    m = re.search(r'방문일\s*:\s*(\d{4}-\d{2}-\d{2})', msg_text)
                    current_date = m.group(1) if m else ''
                except Exception:
                    current_date = ''
                metadata = json.dumps({
                    "lead_no": lead_no, "channel": channel, "message_ts": message_ts,
                }, ensure_ascii=False)
                dp_element = {"type": "datepicker", "action_id": "value"}
                if current_date:
                    dp_element["initial_date"] = current_date
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
                            "label": {"type": "plain_text", "text": "새 방문 예정일"},
                            "element": dp_element,
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
        ack()
        def _bg():
            try:
                _process_project_submission(client, body, view)
            except Exception as exc:
                logger.error(f"[SLACK/공사확정] submit 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

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

    logger.info(
        "[SLACK/공사봇] 핸들러 등록 완료: /공사확정, submit_project, "
        "options(company_name)"
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
            except Exception as exc:
                logger.error(f"[ChannelTalk→] thread 답글 처리 예외: {exc}", exc_info=True)

    # ④ 인입 알림 메시지의 [방문 요청] 버튼
    # ⓑ [📋 상담하기] 통합 버튼 — 인입 카드 모든 처리 흐름의 단일 진입점
    @app.action("button_consult")
    def handle_button_consult(ack, body, client):
        ack()
        try:
            _open_consult_modal(client, body, from_slash=False)
        except Exception as exc:
            logger.error(f"[SLACK] button_consult 실패: {exc}", exc_info=True)

    # 채널톡 카드 [🔗 기존 lead 연결] — 같은 사람이 다른 채널로도 인입했을 때
    @app.action("link_existing_lead")
    def handle_link_existing_lead(ack, body, client):
        ack()
        try:
            chat_id = body["actions"][0]["value"]
            channel = body["channel"]["id"]
            message_ts = body["message"]["ts"]
            _open_link_lead_modal(client, body, chat_id, channel, message_ts)
        except Exception as exc:
            logger.error(f"[SLACK] link_existing_lead 실패: {exc}", exc_info=True)

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
        try:
            metadata = json.loads(view["private_metadata"])
            chat_id = metadata.get("chat_id", "")
            channel = metadata.get("channel", "")
            message_ts = metadata.get("message_ts", "")
            state = view["state"]["values"]
            # external_select 결과 — selected_option.value = lead_no
            sel = state.get("target_lead_no", {}).get("link_lead_search", {}).get("selected_option")
            target_lead_no = (sel or {}).get("value", "").strip().upper() if sel else ""
            if not re.match(r"^L-\d{5}$", target_lead_no):
                ack(response_action="errors", errors={
                    "target_lead_no": "검색해서 lead를 선택해주세요"
                })
                return
            target_lead = _find_lead_by_no(target_lead_no)
            if not target_lead:
                ack(response_action="errors", errors={
                    "target_lead_no": f"{target_lead_no} 시트에 없는 lead 입니다"
                })
                return
            ack()
            _link_chat_to_existing_lead(client, chat_id, target_lead_no, channel, message_ts)
        except Exception as exc:
            logger.error(f"[SLACK] submit_link_lead 실패: {exc}", exc_info=True)
            try:
                ack()
            except Exception:
                pass

    # ⓓ /방문 슬래시 명령 — 거래처/기타 방문 직접 등록
    @app.command("/방문")
    def handle_visit_command(ack, command, client):
        ack()
        trigger_id = command.get("trigger_id", "")
        channel = command.get("channel_id", "")
        user_id = command.get("user_id", "")
        if not trigger_id:
            return
        # /방문은 lead_no 없이 통합 모달 호출 (거래처/기타 신규 등록용)
        fake_body = {
            "trigger_id": trigger_id,
            "user": {"id": user_id},
            "channel_id": channel,
        }
        try:
            _open_consult_modal(client, fake_body, from_slash=True)
        except Exception as exc:
            logger.error(f"[SLACK] /방문 실패: {exc}", exc_info=True)

    # ⓒ 통합 상담 모달 제출
    @app.view("submit_consult")
    def handle_submit_consult(ack, body, client, view):
        # 안전장치: 방문 예약인데 날짜 미선택 시 차단
        state = view["state"]["values"]
        status = _v(state, "status")
        visit_date = (_v(state, "visit_date") or '').strip()
        if status == '방문 예약' and not visit_date:
            ack(
                response_action="errors",
                errors={"visit_date": "방문 예약 시 방문 예정일을 선택해주세요."},
            )
            return
        ack()
        def _bg():
            try:
                _process_consult_submission(client, body, view)
            except Exception as exc:
                logger.error(f"[SLACK] submit_consult 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.action("button_visit")
    def handle_button_visit(ack, body, client):
        ack()
        try:
            _open_inquiry_modal(client, body, action='visit')
        except Exception as exc:
            logger.error(f"[SLACK] button_visit 실패: {exc}", exc_info=True)

    # ⑦ 인입 알림 메시지의 [가격 문의] 버튼
    @app.action("button_price")
    def handle_button_price(ack, body, client):
        ack()
        try:
            _open_inquiry_modal(client, body, action='price')
        except Exception as exc:
            logger.error(f"[SLACK] button_price 실패: {exc}", exc_info=True)

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

    @app.command("/전화")
    def handle_phone_command(ack, command, client):
        ack()
        text = command.get("text", "").strip().lower()
        trigger_id = command.get("trigger_id", "")
        channel = command.get("channel_id", "")
        user_id = command.get("user_id", "")

        # 인자 분기
        if text in ("안내", "setup", "help"):
            # 채널에 안내 메시지 + [+ 전화 문의 등록] 버튼 발송 (관리자가 핀 고정용)
            _post_phone_setup_message(client, channel)
            return

        # 기본: 모달 열기
        if not trigger_id:
            return
        try:
            _open_phone_modal(client, trigger_id, channel, user_id)
        except Exception as exc:
            logger.error(f"[SLACK] /전화 모달 실패: {exc}", exc_info=True)

    # ⑪ [+ 전화 문의 등록] 버튼 (채널 고정 메시지의 버튼)
    @app.action("button_phone")
    def handle_button_phone(ack, body, client):
        ack()
        try:
            _open_phone_modal(
                client,
                body["trigger_id"],
                body["channel"]["id"],
                body["user"]["id"],
            )
        except Exception as exc:
            logger.error(f"[SLACK] button_phone 실패: {exc}", exc_info=True)

    # ⑪-2 채널 책갈피용 App Shortcut — 슬랙 콘솔에서 callback_id="phone_inquiry_shortcut"
    # 으로 Global Shortcut 등록 + 채널 책갈피로 추가하면 1클릭으로 모달 호출
    @app.shortcut("phone_inquiry_shortcut")
    def handle_phone_shortcut(ack, body, client):
        ack()
        try:
            trigger_id = body.get("trigger_id", "")
            user_id = (body.get("user") or {}).get("id", "")
            channel = (body.get("channel") or {}).get("id", "") or \
                os.getenv('SLACK_LEAD_CHANNEL', '').strip()
            _open_phone_modal(client, trigger_id, channel, user_id)
        except Exception as exc:
            logger.error(f"[SLACK] 전화 shortcut 실패: {exc}", exc_info=True)

    # ⑫ 전화 문의 모달 제출
    @app.view("submit_phone")
    def handle_submit_phone(ack, body, client, view):
        ack()
        # 시트 로드(3500+행) + 등록이 3초 넘을 수 있어 백그라운드 스레드로 처리
        # → 슬랙 3초 timeout 회피, 모달 정상 닫힘
        def _bg():
            try:
                _process_phone_submission(client, body, view)
            except Exception as exc:
                logger.error(f"[SLACK] submit_phone 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

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

    logger.info(
        "[SLACK] 메인 봇 핸들러 등록 완료: /상태, /전화, /청소, app_mention, message(DM), "
        "button_visit, button_price, button_phone, submit_visit, submit_price, submit_phone, "
        "sweep_confirm, sweep_cancel"
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
                "label": {"type": "plain_text", "text": "사업자명 (고객사) — 선택, 1글자 입력시 검색"},
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
                "label": {"type": "plain_text", "text": "발주처 담당자 (선택)"},
                "element": {"type": "plain_text_input", "action_id": "value"},
            },
            {
                "type": "input", "block_id": "contact", "optional": True,
                "label": {"type": "plain_text", "text": "발주처 연락처 (선택)"},
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
@slack_bp.route("/events", methods=["POST"])
def slack_events():
    """슬랙 → 우리 서버 webhook (메인 봇: 모든 이벤트/명령/인터랙션 통합 endpoint)"""
    if _slack_handler is None:
        if not _init_slack_app():
            return jsonify({"error": "Slack bot not configured"}), 503

    return _slack_handler.handle(request)


@slack_bp.route("/project-events", methods=["POST"])
def slack_project_events():
    """슬랙 → 공사 현황 알림 봇 전용 endpoint (/공사확정 슬래시 + 모달)"""
    if _project_slack_handler is None:
        if not _init_project_slack_app():
            return jsonify({"error": "Project Slack bot not configured"}), 503

    return _project_slack_handler.handle(request)


@slack_bp.route("/visit-events", methods=["POST"])
def slack_visit_events():
    """슬랙 → 방문 일정 알림 봇 전용 endpoint (날짜 수정/취소 액션)"""
    if _visit_slack_handler is None:
        if not _init_visit_slack_app():
            return jsonify({"error": "Visit Slack bot not configured"}), 503

    return _visit_slack_handler.handle(request)


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
    parts = _split_lead_content(str(lead.get('상담 내용', '')))
    name = str(lead.get('고객명') or '').strip() or '-'
    phone = str(lead.get('고객 연락처') or '').strip() or '-'
    email = str(lead.get('이메일') or '').strip() or '-'
    place = parts['place'] or '-'
    device = parts['device'] or '-'
    inquiry = parts['inquiry'] or str(lead.get('상담 내용') or '').strip() or '-'
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
        {"type": "section", "text": {"type": "mrkdwn",
                                      "text": f"*상세 문의 내용 :*\n{inquiry}"}},
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
                                 channel: str, message_ts: str) -> None:
    """채널톡 채팅을 기존 lead에 통합.
    - 시트 lead의 피드백 컬럼에 카톡 메시지 메모 추가
    - Redis pending lead 삭제 (이 채팅은 더 이상 새 lead 등록 안 함)
    - 슬랙 thread에 통합 완료 안내
    """
    try:
        from dashboard.utils.redis_client import get_redis_client
        from dashboard.services.lead_service import update_lead, get_lead_by_no
        from datetime import datetime
        rc = get_redis_client().redis
        pending_key = f'channeltalk_pending_lead:{chat_id}'
        pending_raw = rc.get(pending_key)
        chat_memo_parts = [f"[{datetime.now().strftime('%m/%d %H:%M')} 카톡 추가 문의 통합]"]
        if pending_raw:
            pending = json.loads(
                pending_raw.decode('utf-8') if isinstance(pending_raw, bytes) else pending_raw
            )
            chat_memo_parts.append(f"닉네임: {pending.get('user_name', '-')}")
            chat_memo_parts.append(f"메시지: {pending.get('first_message', '-')}")
        chat_memo = '\n'.join(chat_memo_parts)

        # 기존 피드백에 추가 (덮어쓰지 않음)
        existing = get_lead_by_no(target_lead_no) or {}
        old_feedback = (existing.get('피드백') or '').strip()
        new_feedback = (old_feedback + '\n\n' + chat_memo).strip() if old_feedback else chat_memo
        update_lead(target_lead_no, {'피드백': new_feedback})

        # Redis 삭제 — 이 채팅은 새 lead 등록 안 함
        rc.delete(pending_key)

        # 슬랙 thread 안내
        if channel and message_ts:
            try:
                client.chat_postMessage(
                    channel=channel, thread_ts=message_ts,
                    text=f":link: 기존 lead `{target_lead_no}`에 통합 완료. "
                         f"이 채팅의 추가 메시지는 시트에 별도 등록되지 않습니다."
                )
            except Exception:
                pass

        # 연결된 기존 lead의 원본 카드에 ✅ reaction 추가 (시각적 처리 완료 표시)
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
                        # 이미 reaction 있거나 권한 문제는 무시
                        logger.debug(f"[SLACK/link] reaction 추가 skip ({target_lead_no}): {exc}")
        except Exception as exc:
            logger.warning(f"[SLACK/link] 원본 카드 reaction 실패: {exc}")

        logger.info(
            f"[SLACK/link] chat_id={chat_id} → {target_lead_no} 통합 완료"
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

    # placeholder 모달 (3초 trigger_id 제약 회피)
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
    except Exception as exc:
        logger.error(f"[SLACK/상담] placeholder 실패: {exc}", exc_info=True)
        return

    # lead_no 있으면 시트 조회 (인입 케이스 prefill)
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
            # 시트에서 매칭
            from dashboard.services.lead_service import load_leads_data
            df = load_leads_data(force_refresh=True)  # 시트 정리 직후 stale 방지
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
        'visit_address': (str(lead.get('방문 주소') or '').strip() if lead else ''),
        # 옛 상담 내용은 카드에 이미 표시 — 모달은 통화 후 추가 메모만 받음 (피드백 컬럼에 저장)
        'consultation': '',
    }
    full_view = _build_consult_view(info_blocks, metadata, prefilled)
    try:
        client.views_update(view_id=view_id, view=full_view)
    except Exception as exc:
        logger.error(f"[SLACK/상담] views_update 실패: {exc}", exc_info=True)


def _build_consult_info_blocks(lead: dict | None, lead_no: str) -> list:
    """상담 모달 상단 인입 정보 블록 — lead 있으면 카드형 정보, 없고 lead_no만 있으면 경고."""
    if lead:
        parts = _split_lead_content(str(lead.get('상담 내용', '')))
        name = str(lead.get('고객명') or '').strip() or '-'
        phone = str(lead.get('고객 연락처') or '').strip() or '-'
        consult_time = str(lead.get('상담 시간') or '').strip() or '-'
        inquiry = parts.get('inquiry') or str(lead.get('상담 내용') or '').strip() or '-'
        return [
            {"type": "section", "text": {"type": "mrkdwn", "text": (
                f"*접수번호:* `{lead_no}`\n"
                f"*문의시간:* {consult_time}\n"
                f"*이름 / 상호:* {name}\n"
                f"*연락처:* {phone}\n"
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
        "placeholder": {"type": "plain_text", "text": "처리 유형 선택"},
        "options": [
            {"text": {"type": "plain_text", "text": label}, "value": v}
            for v, label in _CONSULT_STATUS_OPTIONS
        ],
    }
    if initial_status:
        status_element["initial_option"] = initial_status

    def _text_input(block_id, label, optional=True, multiline=False, placeholder=None):
        elem = {"type": "plain_text_input", "action_id": "value"}
        if multiline:
            elem["multiline"] = True
        if placeholder:
            elem["placeholder"] = {"type": "plain_text", "text": placeholder}
        val = (prefilled.get(block_id) or '').strip()
        if val:
            elem["initial_value"] = val[:300]
        return {
            "type": "input", "block_id": block_id, "optional": optional,
            "label": {"type": "plain_text", "text": label},
            "element": elem,
        }

    vd_element = {"type": "datepicker", "action_id": "value"}
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

    input_blocks = []
    if not is_lead_card:
        input_blocks.append({
            "type": "input", "block_id": "visit_type",
            "label": {"type": "plain_text", "text": "방문 유형"},
            "element": visit_type_element,
        })
    input_blocks.extend([
        {
            "type": "input", "block_id": "status",
            "label": {"type": "plain_text", "text": "처리 유형"},
            "element": status_element,
        },
        {
            "type": "input", "block_id": "visit_date", "optional": True,
            "label": {"type": "plain_text",
                      "text": "방문 예정일 (방문 예약 시 입력)"},
            "hint": {"type": "plain_text",
                     "text": "처리 유형이 '방문 예약'이 아니면 무시됩니다."},
            "element": vd_element,
        },
        {
            "type": "input", "block_id": "visit_date_end", "optional": True,
            "label": {"type": "plain_text", "text": "방문 종료일 (범위 시 입력)"},
            "hint": {"type": "plain_text",
                     "text": "여러 날 방문 (예: 7/1~7/3) 일 때만 입력. 단일이면 비워두세요."},
            "element": {"type": "datepicker", "action_id": "value"},
        },
    ])

    input_blocks.extend([
        _text_input("name", "이름 / 상호"),
        _text_input("contact", "연락처", placeholder="010-1234-5678"),
        _text_input("visit_address", "방문 주소"),
        _text_input("consultation", "추가 상담 메모 (옵션)",
                    multiline=True,
                    placeholder="통화/방문 후 추가 정보, 특이사항 등 — 시트 피드백 컬럼에 저장"),
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
            if visit_address:
                update_data['방문 주소'] = visit_address
            if consultation:
                # 옛 상담 내용은 보존 — 매니저 추가 입력은 피드백 컬럼에 저장
                update_data['피드백'] = consultation
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
                '이메일': '-',
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
                '이메일': '-',
                '고객명': name or '-',
                '방문 주소': visit_address or '-',
                '상담 내용': consultation or '-',
                '키워드': '-',
                '온라인 상담자': counselor,
                '영업 담당자': '',
                '마지막 연락일': '',
                '피드백': consultation or '',
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
        SEP = '---------------------------------------------'
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
    SEP = '---------------------------------------------'
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
                    "style": "primary",
                    "value": lead_no,
                    "action_id": "visit_modify_date",
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

    # visit_type — 슬랙 워크플로가 시트 C열(플랫폼)에 매핑 → 시트 값 그대로 사용
    _lead_platform = str(lead.get('플랫폼', '') or '').strip()
    # 방문 예정일 분리 — start/end ISO 변수도 함께 전달
    _vd_raw = str(lead.get('방문 예정일', '') or '').strip()
    _vd_start, _vd_end = _split_visit_date_range(_vd_raw)
    payload = {
        'lead_no': lead_no or '-',
        'platform': _lead_platform or '-',
        'visit_type': _lead_platform or '온라인',
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
    """[📅 방문일 수정] 모달 제출 처리 — 시트 update + 메시지 chat.update."""
    metadata = json.loads(view.get("private_metadata") or "{}")
    lead_no = metadata.get("lead_no", "")
    channel = metadata.get("channel", "")
    message_ts = metadata.get("message_ts", "")
    state = view["state"]["values"]
    new_date = _v(state, "visit_date") or ''
    if not lead_no or not new_date:
        return

    # 1) 시트 update — escape prefix로 시리얼 변환 차단
    try:
        from dashboard.services.lead_service import update_lead
        update_lead(lead_no, {'방문 예정일': f"'{new_date}"})
    except Exception as exc:
        logger.error(f"[SLACK/방문수정] 시트 update 실패 ({lead_no}): {exc}", exc_info=True)
        return

    # 1-2) 슬랙 List 동기화 워크플로우 — 날짜 셀 갱신
    _trigger_visit_list_webhook(
        'SLACK_VISIT_MODIFY_WEBHOOK_URL', lead_no, channel, message_ts,
        new_visit_date=new_date,
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
            visit_date=new_date,
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

        new_text = (
            f":no_entry_sign: *고객 요청으로 방문 취소*\n"
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

    # 2) 폴더명 생성 — "({이니셜}) {방문 주소} {YY.MM.DD}"
    lead = _find_lead_by_no(lead_no) or {}
    # 이니셜 — 시트의 영업 담당자 우선, 없으면 카드 본문의 "등록자 :" 라인 추출
    sales_rep = str(lead.get('영업 담당자', '') or '').strip()
    initial = _to_initial(sales_rep) if sales_rep else ''
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

    uploaded = 0
    for f in files:
        download_url = f.get('url_private_download') or f.get('url_private')
        if not download_url:
            continue
        filename = f.get('name') or f.get('title') or f'photo_{f.get("id","unknown")}.jpg'
        mimetype = f.get('mimetype') or 'application/octet-stream'
        # 사진/비디오만 (기타 PDF 등은 OK이지만 의도 사진)
        try:
            req = urllib.request.Request(
                download_url,
                headers={'Authorization': f'Bearer {bot_token}'},
            )
            with urllib.request.urlopen(req, timeout=20) as r:
                content = r.read()
            if upload_file(folder_id, filename, content, mimetype=mimetype):
                uploaded += 1
        except Exception as exc:
            logger.error(f"[SLACK/방문 사진] 다운로드/업로드 실패 ({filename}): {exc}",
                         exc_info=True)

    if uploaded == 0:
        return

    # 4) thread → folder 매핑 저장 (상호명 답글로 폴더명 갱신용, TTL 30일)
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
    except Exception as exc:
        logger.warning(f"[SLACK/방문 사진] Redis 매핑 저장 실패: {exc}")

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


def _open_phone_modal(client, trigger_id: str, channel: str, user_id: str):
    """[전화 문의 등록] 모달 열기"""
    metadata = json.dumps({"channel": channel, "user_id": user_id}, ensure_ascii=False)
    modal = {
        "type": "modal",
        "callback_id": "submit_phone",
        "title": {"type": "plain_text", "text": "전화 문의 등록"},
        "submit": {"type": "plain_text", "text": "등록"},
        "close": {"type": "plain_text", "text": "취소"},
        "private_metadata": metadata,
        "blocks": [
            {
                "type": "input", "block_id": "name",
                "label": {"type": "plain_text", "text": "고객명 / 상호 (선택)"},
                "element": {"type": "plain_text_input", "action_id": "value"},
                "optional": True,
            },
            {
                "type": "input", "block_id": "phone",
                "label": {"type": "plain_text", "text": "연락처"},
                "element": {
                    "type": "plain_text_input", "action_id": "value",
                    "placeholder": {"type": "plain_text", "text": "010-1234-5678"},
                },
            },
            {
                "type": "input", "block_id": "email",
                "label": {"type": "plain_text", "text": "이메일 (선택)"},
                "element": {"type": "plain_text_input", "action_id": "value"},
                "optional": True,
            },
            {
                "type": "input", "block_id": "status",
                "label": {"type": "plain_text", "text": "상태"},
                "element": {
                    "type": "static_select", "action_id": "value",
                    "initial_option": {
                        "text": {"type": "plain_text", "text": "유선 상담 (시트 등록)"},
                        "value": "유선 상담",
                    },
                    "options": [
                        {"text": {"type": "plain_text", "text": label}, "value": v}
                        for v, label in _PHONE_STATUS_OPTIONS
                    ],
                },
            },
            {
                "type": "input", "block_id": "device",
                "label": {"type": "plain_text", "text": "설치 희망 기기 (선택, 멀티)"},
                "element": {
                    "type": "multi_static_select", "action_id": "value",
                    "placeholder": {"type": "plain_text", "text": "기기 선택"},
                    "options": [
                        {"text": {"type": "plain_text", "text": d}, "value": d}
                        for d in _PHONE_DEVICE_OPTIONS
                    ],
                },
                "optional": True,
            },
            {
                "type": "input", "block_id": "address",
                "label": {"type": "plain_text", "text": "방문 주소 (선택)"},
                "element": {
                    "type": "plain_text_input", "action_id": "value",
                    "placeholder": {"type": "plain_text", "text": "예: 강남구 테헤란로 152"},
                },
                "optional": True,
            },
            {
                "type": "input", "block_id": "visit_date",
                "label": {"type": "plain_text", "text": "방문 예정일 (방문 예약 시 입력)"},
                "element": {"type": "datepicker", "action_id": "value"},
                "optional": True,
            },
            {
                "type": "input", "block_id": "visit_date_end",
                "label": {"type": "plain_text", "text": "방문 종료일 (범위 시 입력)"},
                "hint": {"type": "plain_text",
                         "text": "여러 날 방문일 때만 입력. 단일이면 비워두세요."},
                "element": {"type": "datepicker", "action_id": "value"},
                "optional": True,
            },
            {
                "type": "input", "block_id": "inquiry",
                "label": {"type": "plain_text", "text": "상담 내용 (선택)"},
                "element": {
                    "type": "plain_text_input", "action_id": "value",
                    "multiline": True,
                    "placeholder": {"type": "plain_text",
                                    "text": "통화에서 받은 정보, 상담 내용 등"},
                },
                "optional": True,
            },
        ],
    }
    client.views_open(trigger_id=trigger_id, view=modal)


def _post_phone_setup_message(client, channel: str):
    """채널에 전화 문의 등록 안내 메시지 + [+ 등록] 버튼 발송 (관리자가 핀 고정)"""
    if not channel:
        return
    blocks = [
        {"type": "section", "text": {"type": "mrkdwn",
                                      "text": ":telephone_receiver: *전화 문의 받으셨나요?*\n"
                                              "_아래 버튼을 누르거나 `/전화` 입력하시면 등록 모달이 뜹니다._"}},
        {
            "type": "actions",
            "elements": [
                {
                    "type": "button",
                    "text": {"type": "plain_text", "text": "+ 전화 문의 등록"},
                    "style": "primary",
                    "action_id": "button_phone",
                    "value": "open",
                },
            ],
        },
    ]
    client.chat_postMessage(
        channel=channel,
        blocks=blocks,
        text="전화 문의 등록 안내",
    )


def _recover_phone_lead_from_sheet(now, phone: str, status: str) -> Optional[str]:
    """일시 에러 후 시트에서 최근 2분 이내 같은 번호 lead 찾기 + 행 색 재설정.

    Google API의 자체 retry로 시트엔 보통 등록돼 있어, 응답만 못 받은 케이스를 회복.
    Returns: lead_no (찾으면) or None
    """
    try:
        from datetime import timedelta
        from dashboard.services.lead_service import (
            load_leads_data, _get_sheet_config, get_sheets_manager,
            LEAD_COLUMN_ORDER,
        )
        from dashboard.services.lead_sync import (
            _parse_consult_dt, _reset_row_background,
        )

        # 재문의 감지 — 캐시 사용 (5분 stale 허용, sync 직후 항상 새로고침되므로 신선)
        main_df = load_leads_data()
        phone_digits = _re.sub(r'\D', '', phone or '')
        if not phone_digits or main_df is None or main_df.empty:
            return None

        norm = main_df['고객 연락처'].astype(str).str.replace(r'\D', '', regex=True)
        matches = main_df[norm == phone_digits].copy()
        if matches.empty:
            return None

        cutoff = now - timedelta(minutes=2)
        recent = []
        for _, m_row in matches.iterrows():
            dt = _parse_consult_dt(m_row.get('상담 시간'))
            if dt and dt >= cutoff:
                recent.append((dt, str(m_row.get('리드 No', ''))))
        if not recent:
            return None

        recent.sort(key=lambda x: x[0], reverse=True)
        lead_no = recent[0][1]

        # 회복된 행의 색 재설정 (위 행 색 상속 방지)
        try:
            cfg = _get_sheet_config()
            if cfg:
                idx_match = main_df.index[main_df['리드 No'].astype(str) == lead_no]
                if len(idx_match):
                    row_num = int(idx_match[0]) + 2  # 헤더 1행
                    updated_range = f"'{cfg['sheet_name']}'!A{row_num}:O{row_num}"
                    _reset_row_background(
                        get_sheets_manager(), cfg['sheet_id'],
                        cfg['sheet_name'], updated_range,
                        num_cols=len(LEAD_COLUMN_ORDER),
                        statuses=[status],
                    )
                    logger.info(
                        f"[SLACK/전화] 행 색 재설정 완료 (row {row_num}, {status})"
                    )
        except Exception as exc3:
            logger.warning(f"[SLACK/전화] 행 색 재설정 실패: {exc3}")

        return lead_no
    except Exception as exc:
        logger.warning(f"[SLACK/전화] 시트 회복 실패: {exc}")
        return None


def _process_phone_submission(client, body, view):
    """전화 문의 모달 제출 → 시트 등록 + 상태 분기 슬랙 알림"""
    import re as _re
    metadata = json.loads(view["private_metadata"])
    channel = metadata.get("channel", "")
    user_id = metadata.get("user_id") or body["user"]["id"]

    state = view["state"]["values"]
    name = _v(state, "name").strip() or '-'
    phone_raw = _v(state, "phone").strip()
    email = _v(state, "email").strip() or '-'
    status = _v(state, "status").strip() or '유선 상담'
    address_raw = _v(state, "address").strip()
    # datepicker는 ISO 형식("2026-06-25") 반환 — 단일/범위 표시 양식 처리
    visit_date_raw = _v(state, "visit_date").strip()
    visit_date_end_raw = _v(state, "visit_date_end").strip()
    visit_date_display = _format_visit_date_range(visit_date_raw, visit_date_end_raw)
    visit_date = _format_date_for_sheet(visit_date_display) if visit_date_display else '-'
    # 슬랙 카드 표시용 raw — 범위 양식 또는 단일
    visit_date_raw = visit_date_display
    inquiry = _v(state, "inquiry").strip() or '-'
    devices = _v_multi(state, "device")
    device_str = ', '.join(devices) if devices else '-'
    place = '-'  # 전화 문의 모달에서는 장소 필드 제거 (통화로 주소만 받음)

    # 슬랙 user → 시트 L열 (온라인 상담자) 한국 이름 매핑
    counselor = _slack_user_to_korean_name(client, user_id) or '-'

    # 연락처 정규화
    from dashboard.services.lead_helpers import (
        normalize_phone, extract_keywords_from_sources,
    )
    phone = normalize_phone(phone_raw) or phone_raw or '-'

    # 키워드 = device 값에서 vocab 매칭
    keyword = extract_keywords_from_sources(device_str) or '-'

    # 주소: 사용자가 모달에 직접 입력한 값 그대로 사용
    # (카카오 자동 검증을 적용하면 건물명/시설명(예: "한울요양원")이 잘림)
    address = address_raw or '-'

    # 상담 시간 = 지금
    now = datetime.now()
    consult_time = now.strftime('%Y.%m.%d. %H:%M')

    lead = {
        '리드 No': '',
        '상담 시간': consult_time,
        '플랫폼': '전화',
        '상태': status,
        '방문 예정일': visit_date,
        '고객 연락처': phone,
        '이메일': email,
        '고객명': name,
        '방문 주소': address,
        '상담 내용': inquiry,
        '키워드': keyword,
        '온라인 상담자': counselor,
        '영업 담당자': '',
        '마지막 연락일': '',
        '피드백': '',
        '_meta_place': place,
        '_meta_device': device_str,
        '_meta_inquiry': inquiry,
        '_meta_consult_dt': now,
        '_meta_address_level': '',
    }

    # 중복 등록 감지 — 같은 번호 24시간 이내 lead 있으면 신규 등록 X, 기존 update
    existing_lead_no = ''
    try:
        from dashboard.services.lead_service import load_leads_data
        from dashboard.services.lead_sync import _get_existing_phone_lookup
        from datetime import timedelta
        main_df = load_leads_data(force_refresh=False)
        phone_lookup = _get_existing_phone_lookup(main_df)
        phone_digits = _re.sub(r'\D', '', phone)
        if phone_digits and phone_digits in phone_lookup:
            prev = phone_lookup[phone_digits]
            if prev and prev[0].get('consult_dt'):
                prev_dt = prev[0].get('consult_dt')
                prev_lead_no = prev[0].get('lead_no', '')
                if prev_dt and (now - prev_dt) < timedelta(hours=24):
                    # 24시간 이내 같은 번호 → 기존 lead update (신규 등록 X)
                    existing_lead_no = prev_lead_no
                    logger.info(
                        f"[SLACK/전화] 중복 감지 → 기존 lead update "
                        f"(phone={phone}, lead_no={existing_lead_no}, by={user_id})"
                    )
            if prev and prev[0].get('consult_dt'):
                if (now - prev[0]['consult_dt']).total_seconds() > 3600:
                    lead['_meta_previous_leads'] = prev
    except Exception as exc:
        logger.warning(f"[SLACK/전화] 재문의 감지 실패: {exc}")

    # 시트 등록 / update — 같은 번호 24시간 이내 lead 있으면 update, 없으면 신규
    lead_no = None
    permanent_error = False
    if existing_lead_no:
        # 기존 lead update — 신규 발번 X, 기존 행만 갱신
        try:
            from dashboard.services.lead_service import update_lead
            update_data = {
                '상태': status,
                '방문 예정일': visit_date,
                '고객명': name,
                '이메일': email,
                '방문 주소': address,
                '상담 내용': inquiry,
                '키워드': keyword,
                '온라인 상담자': counselor,
            }
            update_lead(existing_lead_no, update_data)
            lead_no = existing_lead_no
            logger.info(f"[SLACK/전화] 기존 lead update 완료: {lead_no}")
        except Exception as exc:
            logger.error(f"[SLACK/전화] 기존 lead update 실패 ({existing_lead_no}): {exc}",
                         exc_info=True)
            permanent_error = True

    for attempt in range(2):
        if lead_no:
            break  # 기존 update 성공 또는 신규 등록 성공
        if attempt > 0:
            logger.warning(f"[SLACK/전화] 1차 등록 실패 — 5초 후 재시도")
            time.sleep(5)
            # 재시도 직전 한 번 더 검색 — 1차 응답 못 받았지만 늦게 시트에 들어왔을 가능성
            lead_no = _recover_phone_lead_from_sheet(now, phone, status)
            if lead_no:
                logger.info(f"[SLACK/전화] 재시도 전 시트 회복: {lead_no}")
                break
        try:
            from dashboard.services.lead_sync import _append_leads_to_main
            lead_nos = _append_leads_to_main([lead])
            if lead_nos:
                lead_no = lead_nos[0]
                break
        except Exception as exc:
            err_lower = str(exc).lower()
            is_transient = (
                'ssl' in err_lower or 'wrong_version' in err_lower
                or 'timeout' in err_lower or 'connection' in err_lower
            )
            if not is_transient:
                logger.error(f"[SLACK/전화] 시트 등록 실패 (non-transient): {exc}",
                             exc_info=True)
                permanent_error = True
                break
            logger.warning(f"[SLACK/전화] 시도 {attempt+1} 일시 에러: {exc}")
            time.sleep(2)
            lead_no = _recover_phone_lead_from_sheet(now, phone, status)
            if lead_no:
                break

    if not lead_no:
        # 2회 시도 모두 실패 또는 영구 에러
        msg = (
            ":x: 시트 등록에 실패했습니다 (자동 재시도 2회 후). "
            "잠시 후 시트에서 직접 확인하시고, 미등록 시 다시 입력해 주세요."
            if not permanent_error else
            ":x: 시트 등록에 실패했습니다 (영구 에러). "
            "관리자에게 알리고 모달을 다시 띄워 재등록 부탁드립니다."
        )
        try:
            client.chat_postEphemeral(
                channel=channel or user_id, user=user_id, text=msg,
            )
        except Exception:
            pass
        return

    # 방문 예약일 때 #방문_일정 채널에 발송 (통합 모달과 동일 흐름)
    is_visit = (status == '방문 예약')
    if is_visit:
        # 시트 escape prefix 제거 (슬랙 표시용)
        visit_date_display = visit_date[1:] if visit_date.startswith("'") else visit_date
        if visit_date_display in ('-', ''):
            visit_date_display = ''
        try:
            _post_visit_notice(
                client, lead_no=lead_no, category='온라인', user_id=user_id,
                visit_date=visit_date_display, name=name, contact=phone,
                visit_address=address, consultation=inquiry,
                platform='전화',
            )
        except Exception as exc:
            logger.error(f"[SLACK/전화] #방문_일정 발송 실패: {exc}", exc_info=True)

    # 확인 메시지 (ephemeral) — 방문 예약은 채널 공지가 영수증, 그 외는 짧은 한 줄
    if not is_visit:
        try:
            action_label = '재등록 (update)' if existing_lead_no else '등록'
            confirm = f":white_check_mark: *{phone}* {action_label} — {status} `{lead_no}`"
            client.chat_postEphemeral(channel=channel or user_id, user=user_id, text=confirm)
        except Exception as exc:
            logger.warning(f"[SLACK/전화] 확인 메시지 실패: {exc}")


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
    parts = _split_lead_content(str(lead.get('상담 내용', '')))

    # visit_type — 슬랙 워크플로가 시트 C열(플랫폼)에 매핑 → lead 실제 플랫폼 사용
    # 채팅(카카오톡/채널톡) / 홈페이지 / 전화 / 당근 — 시트 값 유지 보장
    lead_platform = str(lead.get('플랫폼') or '').strip()
    # 방문 예정일 — 범위 양식이면 (시작, 종료) ISO로 분리 + 합쳐진 표시 양식도 함께 전달
    visit_date_raw = str(lead.get('방문 예정일') or '').strip()
    vd_start_iso, vd_end_iso = _split_visit_date_range(visit_date_raw)
    payload = {
        "lead_no": lead_no or '-',
        "platform": lead_platform or '-',
        "visit_type": lead_platform or "온라인",
        "name": str(lead.get('고객명') or '').strip() or '-',
        "contact": str(lead.get('고객 연락처') or '').strip() or '-',
        "email": str(lead.get('이메일') or '').strip() or '-',
        "inquiry_time": str(lead.get('상담 시간') or '').strip() or '-',
        "location": parts.get('place') or '-',
        "device": parts.get('device') or str(lead.get('키워드') or '').strip() or '-',
        "visit_address": modal_fields.get('visit_address') or str(lead.get('방문 주소') or '').strip() or '-',
        "consultation": modal_fields.get('consultation') or '-',
        "details": parts.get('inquiry') or str(lead.get('상담 내용') or '').strip() or '-',
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
            update_data['피드백'] = consultation
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
            update_data['피드백'] = consultation
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


# 앱 시작 시 한 번 초기화 시도
_init_slack_app()
_init_visit_slack_app()
