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

# file_shared 이벤트 중복 방지 (5분 TTL)
_processed_file_ids: dict = {}  # file_id → timestamp
_FILE_DEDUP_TTL_SEC = 300
_processed_file_ids_lock = threading.Lock()


def _is_file_already_processed(file_id: str) -> bool:
    """파일이 최근 처리됐는지 체크하고, 새로 처리할 거면 등록."""
    if not file_id:
        return False
    now = time.time()
    with _processed_file_ids_lock:
        # 오래된 항목 정리
        for fid in list(_processed_file_ids.keys()):
            if now - _processed_file_ids[fid] > _FILE_DEDUP_TTL_SEC * 2:
                del _processed_file_ids[fid]
        # 5분 내 동일 file_id면 중복으로 간주
        if file_id in _processed_file_ids:
            return True
        _processed_file_ids[file_id] = now
        return False


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

    # ① 슬래시 명령: /당근 (당근마켓 리드폼 처리)
    @app.command("/당근")
    def handle_karrot(ack, command, respond):
        ack()
        user_id = command.get("user_id", "")
        text = command.get("text", "").strip().lower()

        # 인자에 따라 분기
        if text in ("도움말", "help", ""):
            respond({
                "response_type": "ephemeral",
                "blocks": [
                    {
                        "type": "header",
                        "text": {"type": "plain_text", "text": "🥕 당근 리드폼 자동 처리"},
                    },
                    {
                        "type": "section",
                        "text": {
                            "type": "mrkdwn",
                            "text": (
                                "*사용법*\n"
                                "1️⃣ 당근 비즈프로필에서 리드폼 엑셀을 다운로드\n"
                                "2️⃣ 이 채널에 *드래그앤드롭* (또는 📎 첨부)\n"
                                "3️⃣ 봇이 자동으로 복호화 + 신규만 추출 + 시트 등록\n\n"
                                "*명령어*\n"
                                "• `/당근` 또는 `/당근 도움말` — 이 도움말\n"
                                "• `/당근 상태` — 마지막 처리 시점 + 시트의 당근 리드 수\n\n"
                                "_⚠️ 아직 자동 처리 기능 구현 중입니다. 곧 활성화됩니다._"
                            ),
                        },
                    },
                ],
            })
        elif text in ("상태", "status"):
            # 시트에서 당근 리드 마지막 처리 시점 조회
            try:
                from dashboard.services.lead_service import load_leads_data
                df = load_leads_data()
                if df is None or df.empty:
                    respond({"response_type": "ephemeral", "text": "📭 시트 데이터 없음"})
                    return

                karrot_rows = df[df['플랫폼'].astype(str).str.strip() == '당근']
                total = len(karrot_rows)
                last_consult = ''
                if total > 0 and '상담 시간' in karrot_rows.columns:
                    last_consult = karrot_rows['상담 시간'].dropna().astype(str).max()

                respond({
                    "response_type": "ephemeral",
                    "blocks": [
                        {
                            "type": "header",
                            "text": {"type": "plain_text", "text": "🥕 당근 리드 현황"},
                        },
                        {
                            "type": "section",
                            "fields": [
                                {"type": "mrkdwn", "text": f"*시트의 당근 리드:*\n{total}건"},
                                {"type": "mrkdwn", "text": f"*마지막 응답 일시:*\n{last_consult or '(없음)'}"},
                            ],
                        },
                    ],
                })
            except Exception as exc:
                logger.error(f"[SLACK] /당근 상태 실패: {exc}", exc_info=True)
                respond({"response_type": "ephemeral", "text": f"❌ 상태 조회 실패: {exc}"})
        else:
            respond({
                "response_type": "ephemeral",
                "text": (
                    f"❓ 알 수 없는 옵션: `{text}`\n"
                    "`/당근 도움말` 또는 `/당근 상태` 를 사용하세요."
                ),
            })

        logger.info(f"[SLACK] /당근 처리: user={user_id}, text={text!r}")

    # ② 슬래시 명령: /상태 (사이트 헬스체크)
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

    # ③ 봇 멘션 이벤트 (예: @ITG관리봇 안녕)
    @app.event("app_mention")
    def handle_mention(event, say):
        user = event.get("user", "")
        text = event.get("text", "")
        say(f"<@{user}> 부르셨나요? `/당근`, `/상태` 명령을 사용해보세요.")

    # ④ DM + 채널톡 thread 답글 통합 처리
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

        # ④-1. DM: 안내 메시지
        if channel_type == "im":
            text = event.get("text", "")
            say(f"메시지 받았습니다: _{text}_\n슬래시 명령 `/당근`, `/상태`도 사용 가능합니다.")
            return

        # ④-2. 채널 thread 답글 — 채널톡 thread면 채널톡으로 forward
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

    # ⑤ 파일 업로드 이벤트 — 당근 엑셀 자동 처리
    @app.event("file_shared")
    def handle_file_shared(event, client, request):
        file_id = event.get("file_id", "")

        # 슬랙 retry 체크 (헤더 X-Slack-Retry-Num): 이미 처리 시도된 거니까 무시
        retry_num = request.headers.get("X-Slack-Retry-Num") if request else None
        retry_reason = request.headers.get("X-Slack-Retry-Reason") if request else None
        if retry_num:
            logger.info(f"[SLACK] file_shared retry 무시: file_id={file_id} retry={retry_num} reason={retry_reason}")
            return

        # 5분 내 동일 file_id 중복 방지 (file_id별 dedup)
        if _is_file_already_processed(file_id):
            logger.info(f"[SLACK] file_shared 중복 무시: file_id={file_id}")
            return

        logger.info(f"[SLACK] file_shared 수신: file_id={file_id}")
        if not file_id:
            return

        # 파일 정보 조회 (이름·다운로드 URL·채널)
        try:
            info = client.files_info(file=file_id)
        except Exception as exc:
            logger.error(f"[SLACK] files_info 실패: {exc}", exc_info=True)
            return

        f = info.get("file", {}) or {}
        fname = (f.get("name") or "").lower()
        channels = f.get("channels") or []
        groups = f.get("groups") or []
        target_channel = (channels + groups)[0] if (channels or groups) else None

        if not target_channel:
            logger.info(f"[SLACK] file_shared 무시 (채널 없음): {fname}")
            return

        # .xlsx 가 아니면 무시 (당근 엑셀만 처리, 향후 현장 사진은 별도 이벤트로)
        if not fname.endswith(".xlsx"):
            logger.info(f"[SLACK] file_shared 무시 (xlsx 아님): {fname}")
            return

        # 당근 리드 엑셀로 추정되는지 (파일명에 '당근' 또는 '리드폼' 포함)
        is_karrot = ('당근' in fname) or ('리드폼' in fname) or ('karrot' in fname)
        if not is_karrot:
            client.chat_postMessage(
                channel=target_channel,
                text=f"📎 `{f.get('name','')}` 받았습니다. 당근 리드 파일은 파일명에 *당근* 또는 *리드폼* 이 포함되어야 자동 처리됩니다.",
            )
            return

        # 무거운 작업은 별도 스레드 → 슬랙에 빠른 ack 보장 (retry 방지)
        threading.Thread(
            target=_process_karrot_file_async,
            args=(client, target_channel, f),
            daemon=True,
            name=f"karrot-{file_id[:8]}",
        ).start()

    # ⑥ 인입 알림 메시지의 [방문 요청] 버튼
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

    logger.info(
        "[SLACK] 핸들러 등록 완료: /당근, /상태, app_mention, message(DM), file_shared, "
        "button_visit, button_price, submit_visit, submit_price"
    )


def _process_karrot_file_async(client, channel: str, file_info: dict):
    """당근 엑셀 다운로드 → 파싱 → 시트 일괄 등록 → 결과 메시지"""
    fname = file_info.get("name", "")
    download_url = file_info.get("url_private_download") or file_info.get("url_private")

    if not download_url:
        client.chat_postMessage(channel=channel, text=f"❌ `{fname}` 다운로드 URL 없음")
        return

    # 처리 시작 알림
    try:
        ack_msg = client.chat_postMessage(
            channel=channel,
            text=f"🥕 `{fname}` 처리 중... (10~30초 소요)",
        )
        ack_ts = ack_msg.get("ts")
    except Exception:
        ack_ts = None

    business_number = os.getenv("KARROT_BUSINESS_NUMBER", "")
    if not business_number:
        client.chat_postMessage(channel=channel, text="❌ `KARROT_BUSINESS_NUMBER` 환경변수 미설정")
        return

    # 다운로드 (Bot 토큰 필요)
    import requests
    try:
        resp = requests.get(
            download_url,
            headers={"Authorization": f"Bearer {_BOT_TOKEN}"},
            timeout=30,
        )
        resp.raise_for_status()
        file_bytes = resp.content
    except Exception as exc:
        logger.error(f"[SLACK] 파일 다운로드 실패: {exc}", exc_info=True)
        client.chat_postMessage(channel=channel, text=f"❌ 파일 다운로드 실패: {exc}")
        return

    # 파싱 + 시트 등록
    try:
        from dashboard.services.karrot_parser import process_karrot_excel, append_leads_to_sheet
        result = process_karrot_excel(file_bytes, business_number)

        # 신규 등록
        lead_nos = []
        if result["new"]:
            lead_nos = append_leads_to_sheet(result["new"])
    except ValueError as exc:
        logger.error(f"[SLACK] 당근 파싱 실패: {exc}")
        client.chat_postMessage(channel=channel, text=f"❌ 처리 실패: {exc}")
        return
    except Exception as exc:
        logger.error(f"[SLACK] 당근 처리 실패: {exc}", exc_info=True)
        client.chat_postMessage(channel=channel, text=f"❌ 처리 중 오류: {type(exc).__name__}: {exc}")
        return

    # ─────────────────────────────────────────────────────────
    # 결과 처리
    # ─────────────────────────────────────────────────────────
    new_count = result["new_count"]
    sheet_karrot_count = result.get("sheet_karrot_count", 0)

    # 신규 0건: 처리 중 메시지 삭제하고 짧은 알림만
    if new_count == 0:
        msg = (
            f"🥕 신규 리드 없음 _(엑셀 {result['total']}건 / "
            f"이미 등록 {result['duplicates']}건, 당근 누적 {sheet_karrot_count}건)_"
        )
        try:
            if ack_ts:
                client.chat_update(channel=channel, ts=ack_ts, text=msg)
            else:
                client.chat_postMessage(channel=channel, text=msg)
        except Exception as exc:
            logger.error(f"[SLACK] 결과 메시지 전송 실패: {exc}", exc_info=True)
        return

    # 신규 N건: 처리 중 메시지를 짧은 헤더로 교체 (또는 삭제 후 새로)
    try:
        if ack_ts:
            client.chat_update(
                channel=channel, ts=ack_ts,
                text=f"🥕 당근 신규 리드 {new_count}건 도착",
            )
    except Exception as exc:
        logger.warning(f"[SLACK] ack 메시지 업데이트 실패: {exc}")

    # 각 신규 리드를 응답 시각 오름차순으로 채널에 개별 메시지 전송
    # (가장 최신 응답이 마지막에 떠서 영업 알림에 잘 띄게)
    for lead, lead_no in zip(result["new"], lead_nos):
        try:
            text = _format_karrot_message(lead, lead_no)
            client.chat_postMessage(channel=channel, text=text, unfurl_links=False)
        except Exception as exc:
            logger.error(f"[SLACK] 리드 메시지 전송 실패 ({lead_no}): {exc}", exc_info=True)


def _format_karrot_message(lead: dict, lead_no: str) -> str:
    """
    당근 신규 리드 1건을 사용자 정의 양식으로 포맷.

    양식:
        *당근 문의 (방문)*
        2026.05.29. 09:42   이샛별   010-9025-9352   상가 / 상업시설 / 의료시설
        천장형   이천 이섭대천로 1407번길 8 2층   "기존 천장형 에어컨이 있는데..."
    """
    consult_time = (lead.get("상담 시간") or "").strip() or "-"
    name = (lead.get("고객명") or "").strip() or "-"
    phone = (lead.get("고객 연락처") or "").strip() or "-"
    place = (lead.get("_meta_place") or "").strip() or "-"
    device = (lead.get("_meta_device") or "").strip() or "-"
    address = (lead.get("방문 주소") or "").strip() or "-"
    inquiry = (lead.get("_meta_inquiry") or "").strip() or "-"

    # 사용자 양식 그대로 (공백 3칸 구분자)
    line1 = f"{consult_time}   {name}   {phone}   {place}"
    line2 = f"{device}   {address}   \"{inquiry}\""

    # 헤더에 리드 No도 살짝 포함 (운영자가 사이트와 매칭하기 쉽게)
    return f"*당근 문의 (방문)*  `{lead_no}`\n{line1}\n{line2}"

    logger.info("[SLACK] 핸들러 등록 완료: /당근, /상태, app_mention, message(DM), file_shared")


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
        # 자동 분기
        return (
            item.get("selected_date")
            or item.get("value")
            or (item.get("selected_option") or {}).get("value")
            or default
        )
    except Exception:
        return default


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
