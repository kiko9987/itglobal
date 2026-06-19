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
import time
import threading
import logging
import json
import urllib.request
from datetime import date, datetime
from flask import Blueprint, request, jsonify

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

slack_bp = Blueprint('slack_bot', __name__, url_prefix='/slack')


# ─────────────────────────────────────────────────────────────
# 활성화 여부 + slack_bolt App 초기화
# ─────────────────────────────────────────────────────────────
_BOT_ENABLED = os.getenv('SLACK_BOT_ENABLED', 'false').lower() == 'true'
_BOT_TOKEN = os.getenv('SLACK_BOT_TOKEN', '')
_SIGNING_SECRET = os.getenv('SLACK_SIGNING_SECRET', '')

_slack_app = None
_slack_handler = None

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
                resp = send_manager_message(chat_id, manager_id, text)
                logger.info(f"[ChannelTalk→] 메시지 발신: text={text[:40]!r}, resp_ok={resp is not None}")

                # 직원 응답했으니 미배정 알림 큐에서 제거
                from dashboard.services.channeltalk_threads import remove_pending
                remove_pending(chat_id)
                if resp:
                    # 슬랙 reaction으로 전송 성공 표시
                    try:
                        client.reactions_add(
                            channel=event["channel"],
                            timestamp=event["ts"],
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

    # ⑬ /청소 슬래시 명령 — 채널 메시지 일괄 청소 (봇이 보낸 메시지만)
    @app.command("/청소")
    def handle_sweep_command(ack, command, client, respond):
        ack()
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

    # ⑯ /공사확정 슬래시 명령 — 모바일/슬랙에서 공사 확정 등록
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

    # ⑰ 공사확정 모달 제출
    @app.view("submit_project")
    def handle_submit_project(ack, body, client, view):
        ack()
        # 시트 로드 + 등록이 3초 넘을 수 있어 백그라운드
        def _bg():
            try:
                _process_project_submission(client, body, view)
            except Exception as exc:
                logger.error(f"[SLACK/공사확정] submit 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    logger.info(
        "[SLACK] 핸들러 등록 완료: /상태, /전화, /청소, /공사확정, app_mention, message(DM), "
        "button_visit, button_price, button_phone, submit_visit, submit_price, submit_phone, "
        "submit_project, sweep_confirm, sweep_cancel"
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


def _human_duration(seconds: int) -> str:
    if seconds >= 86400:
        return f"{seconds // 86400}일"
    if seconds >= 3600:
        return f"{seconds // 3600}시간"
    if seconds >= 60:
        return f"{seconds // 60}분"
    return f"{seconds}초"


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
                "label": {"type": "plain_text", "text": "사업자명 (고객사)"},
                "element": {"type": "plain_text_input", "action_id": "value"},
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
    company_name = (_v(state, "company_name") or '').strip()
    address = (_v(state, "address") or '').strip()
    customer = (_v(state, "customer") or '').strip()
    contact = (_v(state, "contact") or '').strip()
    start_date = _v(state, "start_date")
    end_date = _v(state, "end_date")
    content = (_v(state, "content") or '').strip()
    amount_raw = (_v(state, "amount") or '').strip()
    vat_separate = bool(_v_multi(state, "vat"))

    # 영업 담당자: 슬랙 사용자 → 한국 이름
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
    """슬랙 → 우리 서버 webhook (모든 이벤트/명령/인터랙션 통합 endpoint)"""
    if _slack_handler is None:
        if not _init_slack_app():
            return jsonify({"error": "Slack bot not configured"}), 503

    return _slack_handler.handle(request)


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
                "label": {"type": "plain_text", "text": "방문 날짜"},
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


def _v(state, block_id, default=''):
    """모달 state.values에서 안전하게 값 추출 (datepicker / text / select 자동 분기)"""
    try:
        item = state[block_id]["value"]
        return (
            item.get("selected_date")
            or item.get("value")
            or (item.get("selected_option") or {}).get("value")
            or default
        )
    except Exception:
        return default


def _slack_user_to_korean_name(client, user_id: str) -> str:
    """슬랙 user_id → SALES_EMAILS 매핑 한국 이름 (fallback: display_name/real_name)"""
    if not user_id:
        return ''
    try:
        resp = client.users_info(user=user_id)
        if not resp.get("ok"):
            return ''
        profile = resp["user"]["profile"]
        email = (profile.get("email") or '').strip().lower()

        # SALES_EMAILS 역매칭
        try:
            sales_emails = json.loads(os.getenv("SALES_EMAILS", "{}"))
        except Exception:
            sales_emails = {}
        for name, mapped in sales_emails.items():
            if str(mapped).strip().lower() == email:
                return name

        # Fallback: display_name / real_name
        return (profile.get("display_name")
                or profile.get("real_name")
                or '').strip()
    except Exception as exc:
        logger.warning(f"[SLACK] users_info 실패 ({user_id}): {exc}")
        return ''


def _v_multi(state, block_id) -> list:
    """멀티 선택 체크박스/multi_static_select 값 추출"""
    try:
        item = state[block_id]["value"]
        opts = item.get("selected_options") or []
        return [o.get("value") for o in opts if o.get("value")]
    except Exception:
        return []


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
    visit_date = _v(state, "visit_date").strip() or '-'  # 방문 예약 시 datepicker
    inquiry = _v(state, "inquiry").strip() or '-'
    devices = _v_multi(state, "device")
    device_str = ', '.join(devices) if devices else '-'
    place = '-'  # 전화 문의 모달에서는 장소 필드 제거 (통화로 주소만 받음)

    # 슬랙 user → 시트 L열 (온라인 상담자) 한국 이름 매핑
    counselor = _slack_user_to_korean_name(client, user_id) or '-'

    # 연락처 정규화
    from dashboard.services.lead_helpers import (
        normalize_phone, extract_keywords_from_sources, extract_korean_address,
    )
    phone = normalize_phone(phone_raw) or phone_raw or '-'

    # 키워드 = device 값에서 vocab 매칭
    keyword = extract_keywords_from_sources(device_str) or '-'

    # 주소 카카오 검증
    address = '-'
    if address_raw:
        try:
            from dashboard.services.address_resolver import resolve_address
            r = extract_korean_address(address_raw)
            ra = r[0] if r else None
            rl = r[1] if r else ''
            verified, lv = resolve_address(address_raw, ra, rl)
            address = verified or address_raw
        except Exception as exc:
            logger.warning(f"[SLACK/전화] 주소 검증 실패: {exc}")
            address = address_raw

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

    # 재문의 감지 (같은 번호 옛 lead 1시간 이상 전)
    try:
        from dashboard.services.lead_service import load_leads_data
        from dashboard.services.lead_sync import _get_existing_phone_lookup
        main_df = load_leads_data(force_refresh=False)
        phone_lookup = _get_existing_phone_lookup(main_df)
        phone_digits = _re.sub(r'\D', '', phone)
        if phone_digits and phone_digits in phone_lookup:
            prev = phone_lookup[phone_digits]
            if prev and prev[0].get('consult_dt'):
                if (now - prev[0]['consult_dt']).total_seconds() > 3600:
                    lead['_meta_previous_leads'] = prev
    except Exception as exc:
        logger.warning(f"[SLACK/전화] 재문의 감지 실패: {exc}")

    # 시트 등록 (SSL 일시 에러 시 시트에서 lead 검색 후 정상 흐름 이어가기)
    lead_no = None
    try:
        from dashboard.services.lead_sync import _append_leads_to_main
        lead_nos = _append_leads_to_main([lead])
        lead_no = lead_nos[0] if lead_nos else None
    except Exception as exc:
        err_lower = str(exc).lower()
        is_ssl_error = 'ssl' in err_lower or 'wrong_version' in err_lower
        if is_ssl_error:
            # SSL 일시 에러 — google API 자동 retry로 시트엔 보통 등록됨
            # 같은 연락처 최근 lead를 시트에서 찾아 lead_no 회복
            logger.warning(f"[SLACK/전화] SSL 일시 에러, 시트 확인 중: {exc}")
            time.sleep(2)
            try:
                from dashboard.services.lead_service import load_leads_data
                main_df = load_leads_data(force_refresh=True)
                phone_digits = _re.sub(r'\D', '', phone)
                if phone_digits and main_df is not None and not main_df.empty:
                    norm = main_df['고객 연락처'].astype(str).str.replace(r'\D', '', regex=True)
                    matches = main_df[norm == phone_digits]
                    if not matches.empty:
                        lead_no = str(matches.iloc[-1]['리드 No'])
                        logger.info(f"[SLACK/전화] SSL 에러 후 시트에서 lead 확인: {lead_no}")
            except Exception as exc2:
                logger.warning(f"[SLACK/전화] SSL 후 시트 검증 실패: {exc2}")
        else:
            logger.error(f"[SLACK/전화] 시트 등록 실패: {exc}", exc_info=True)

    if not lead_no:
        # 진짜 등록 실패
        try:
            client.chat_postEphemeral(
                channel=channel or user_id, user=user_id,
                text=":x: 시트 등록 실패. 잠시 후 시트에서 직접 확인 부탁드립니다.",
            )
        except Exception:
            pass
        return

    # 상태 분기 — 방문 예약/견적 제출만 채널에 "전화 등록 결과 보고" 카드 발송
    # (인입 알림 카드가 아니라 결과 공유용 — 버튼 없음, 이중 컨택 방지가 목적)
    notify_slack = status in ('방문 예약', '견적 제출')
    if notify_slack:
        try:
            _post_phone_registration_notice(
                client, lead, lead_no, status, user_id,
                channel=os.getenv('SLACK_LEAD_CHANNEL', '').strip(),
            )
        except Exception as exc:
            logger.error(f"[SLACK/전화] 등록 보고 발송 실패: {exc}", exc_info=True)

    # 확인 메시지 (ephemeral) — 전화번호 메인 + 리드 No 보조
    if notify_slack:
        confirm = f":white_check_mark: *{phone}* 등록 완료 — {status} (채널 공지 완료) `{lead_no}`"
    else:
        confirm = f":white_check_mark: *{phone}* 등록 완료 — {status} (시트만 기록) `{lead_no}`"
    try:
        client.chat_postEphemeral(channel=channel or user_id, user=user_id, text=confirm)
    except Exception as exc:
        logger.warning(f"[SLACK/전화] 확인 메시지 실패: {exc}")


def _post_phone_registration_notice(client, lead: dict, lead_no: str, status: str,
                                     user_id: str, channel: str):
    """전화 모달 등록 후 채널에 결과 보고 카드 발송.

    인입 알림 카드(버튼 있음)가 아니라 *결과 공유*용 단순 카드.
    의도: 다른 영업 담당자가 같은 고객한테 이중 컨택하지 않도록 공유.

    양식:
        ✅ 전화 접수 → 방문 예약 등록  `L-02909`
        📞 010-8942-0275 (김서아)
        📍 군포 엘에스로 13 신일IT유토지식산업센터 1012호
        📅 방문 예정: 2026-06-20
        👤 등록: @박정우
    """
    if not channel:
        return

    phone = (lead.get('고객 연락처') or '').strip() or '-'
    name = (lead.get('고객명') or '').strip()
    address = (lead.get('방문 주소') or '').strip()
    visit_date = (lead.get('방문 예정일') or '').strip()

    lines = [f":white_check_mark: *전화 접수 → {status} 등록*  `{lead_no}`"]

    phone_line = f":telephone_receiver: {phone}"
    if name and name not in ('-', ''):
        phone_line += f" ({name})"
    lines.append(phone_line)

    if address and address not in ('-', ''):
        lines.append(f":round_pushpin: {address}")

    if status == '방문 예약' and visit_date and visit_date not in ('-', ''):
        lines.append(f":calendar: 방문 예정: {visit_date}")

    lines.append(f":bust_in_silhouette: 등록: <@{user_id}>")

    try:
        client.chat_postMessage(
            channel=channel, text='\n'.join(lines), unfurl_links=False,
        )
    except Exception as exc:
        logger.warning(f"[SLACK/전화] 채널 보고 발송 실패: {exc}")


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
    webhook_url = os.getenv("SLACK_LIST_WEBHOOK_URL", "").strip()
    if not webhook_url:
        logger.debug("[SLACK/LIST] SLACK_LIST_WEBHOOK_URL 미설정 - 등록 스킵")
        return False

    # 메시지 영구 링크
    message_link = ''
    try:
        permalink = client.chat_getPermalink(channel=channel, message_ts=message_ts)
        message_link = permalink.get("permalink", "")
    except Exception:
        pass

    # 상담 내용 파싱 (장소/기기/문의)
    parts = _split_lead_content(str(lead.get('상담 내용', '')))

    payload = {
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
        logger.info(f"[SLACK/LIST] webhook 등록 완료 (lead={lead.get('리드 No')} action={action})")
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
    visit_date = _v(state, "visit_date")
    visit_address = _v(state, "visit_address")
    consultation = _v(state, "consultation")

    # 메인 시트 업데이트
    try:
        from dashboard.services.lead_service import update_lead
        update_data = {
            '상태': '방문 예약',
            '방문 예정일': visit_date,
        }
        if visit_address:
            update_data['방문 주소'] = visit_address
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
            'visit_date': visit_date,
            'visit_address': visit_address,
            'consultation': consultation,
        },
        channel=channel, message_ts=message_ts, action='visit',
    )

    # 원본 메시지에 답글
    reply_text = (
        f"✅ *방문 요청 등록* — `{lead_no}` by <@{user_id}>\n"
        f">*방문 날짜* : {visit_date or '-'}\n"
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
        f"💰 *가격 문의 처리* — `{lead_no}` by <@{user_id}>\n"
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


# 앱 시작 시 한 번 초기화 시도
_init_slack_app()
