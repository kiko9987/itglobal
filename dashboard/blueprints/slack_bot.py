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

# 수금 관리 알림 봇 (별도 토큰/secret) — 입금 카드 발송(payment_sync)은 WebClient,
# 여기선 [🗑 삭제] 버튼 인터랙션만 처리 (정정·취소 회색 카드 정리)
_PAYMENT_BOT_TOKEN = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '')
_PAYMENT_SIGNING_SECRET = os.getenv('SLACK_PAYMENT_SIGNING_SECRET', '')

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
_payment_slack_app = None
_payment_slack_handler = None

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


def _init_payment_slack_app():
    """수금 관리 알림 봇 — 입금 카드 [🗑 삭제] 버튼 인터랙션 처리용 Bolt App.

    카드 발송은 payment_sync 가 WebClient 로 함. 여기선 정정·취소 회색 카드의
    [🗑 삭제] 버튼(경영지원 한정)만 처리. Interactivity Request URL = /slack/payment-events.
    """
    global _payment_slack_app, _payment_slack_handler

    if not _BOT_ENABLED:
        return False
    if not _PAYMENT_BOT_TOKEN:
        logger.warning("[SLACK/수금봇] SLACK_PAYMENT_BOT_TOKEN 미설정 — 인터랙션 비활성화")
        return False
    if not _PAYMENT_SIGNING_SECRET:
        logger.warning("[SLACK/수금봇] SLACK_PAYMENT_SIGNING_SECRET 미설정 — 인터랙션 비활성화")
        return False

    try:
        from slack_bolt import App
        from slack_bolt.adapter.flask import SlackRequestHandler

        _payment_slack_app = App(
            token=_PAYMENT_BOT_TOKEN,
            signing_secret=_PAYMENT_SIGNING_SECRET,
            process_before_response=True,
        )
        _payment_slack_handler = SlackRequestHandler(_payment_slack_app)

        _register_payment_handlers(_payment_slack_app)
        _verify_bot_token(_payment_slack_app.client, '수금봇')
        logger.info("[SLACK/수금봇] 초기화 완료 ✅")
        return True
    except Exception as exc:
        logger.error(f"[SLACK/수금봇] 초기화 실패: {exc}", exc_info=True)
        return False


def _register_payment_handlers(app):
    """수금봇 핸들러 — 정정·취소 회색 카드 [🗑 삭제] 버튼 (경영지원 한정)."""

    @app.action("payment_card_delete")
    def handle_payment_card_delete(ack, body, client):
        ack()
        user = (body.get('user') or {}).get('id', '')
        channel = (body.get('channel') or {}).get('id', '')
        ts = (body.get('message') or {}).get('ts', '')
        # 정산 카드 삭제는 경영지원(황샛별)만 — 정산 핀 ✅·금액 게이트와 동일 사상.
        if user != _SETTLEMENT_CHECKER_ID:
            try:
                client.chat_postEphemeral(
                    channel=channel, user=user,
                    text=':lock: 정산 카드 삭제는 경영지원(황샛별)만 가능합니다.',
                )
            except Exception:
                pass
            return
        if not channel or not ts:
            return

        def _bg():
            try:
                client.chat_delete(channel=channel, ts=ts)
                logger.info(f'[SLACK/수금봇] 정정 카드 삭제: ts={ts} by {user}')
            except Exception as exc:
                logger.warning(f'[SLACK/수금봇] chat_delete 실패 (ts={ts}): {exc}')
        threading.Thread(target=_bg, daemon=True).start()

    # ─── 은행 입금 SMS 인입 → 프로젝트·수금단계 지정 (2026-08-12) ───
    # /sms/inbound webhook 이 #수금_입력 채널에 [지정] 버튼 카드 게시 →
    # 매니저가 모달에서 프로젝트(자동완성)+단계 선택 → 시트 U/V/W 셀 메모 기록 →
    # 기존 payment_sync 폴러가 감지해 #수금_관리 정식 카드 게시.
    @app.action("payment_intake_open")
    def handle_payment_intake_open(ack, body, client):
        ack()
        def _bg():
            try:
                intake_id = body["actions"][0]["value"]
                channel = body["channel"]["id"]
                message_ts = body["message"]["ts"]
                _open_payment_intake_modal(client, body, intake_id, channel, message_ts)
            except Exception as exc:
                logger.error(f"[SLACK/수금봇] payment_intake_open 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.options("payment_intake_project")
    def handle_payment_intake_options(ack, body):
        """external_select — 프로젝트 코드/사업자명 자동완성."""
        try:
            query = (body.get("value") or "").strip()
            ack(options=_build_payment_intake_options(query, limit=30))
        except Exception as exc:
            logger.error(f"[SLACK/수금봇] payment_intake options 실패: {exc}", exc_info=True)
            try:
                ack(options=[])
            except Exception:
                pass

    @app.action("payment_intake_project")
    def handle_payment_intake_select(ack, body, client):
        """프로젝트 선택 → 상세(상호·주소·담당자·공사내용) 삽입 재렌더 (반복거래처 확인용)."""
        ack()
        def _bg():
            try:
                view = body.get("view") or {}
                action = (body.get("actions") or [{}])[0]
                sel = action.get("selected_option") or {}
                code = (sel.get("value") or "").strip()
                if not code:
                    return
                meta = json.loads(view.get("private_metadata") or "{}")
                state = (view.get("state") or {}).get("values", {})
                stage_value = (_v(state, "stage") or "").strip()
                d = _load_intake(meta.get("intake_id", ""))
                from dashboard.services.as_service import get_project_details
                details = get_project_details(code) or {}
                new_view = _build_payment_intake_view(
                    meta.get("intake_id", ""), meta.get("channel", ""),
                    meta.get("message_ts", ""), d.get("text") or "",
                    selected_project=sel, project_details=details, stage_value=stage_value)
                client.views_update(view_id=view.get("id"), hash=view.get("hash"), view=new_view)
            except Exception as exc:
                logger.error(f"[SLACK/수금봇] payment_intake 프로젝트 선택 상세 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.view("submit_payment_intake")
    def handle_submit_payment_intake(ack, body, view, client):
        # 매니저는 프로젝트·수금단계만 '지정' → 시트 기록은 경영지원(황샛별) 확인 후.
        try:
            state = view["state"]["values"]
            sel = (state.get("project", {})
                        .get("payment_intake_project", {})
                        .get("selected_option"))
            project_code = (sel or {}).get("value", "").strip() if sel else ""
            stage = (_v(state, "stage") or "").strip()
            errors = {}
            if not project_code:
                errors["project"] = "프로젝트를 검색해서 선택해주세요."
            if stage not in ("계약금", "중도금", "잔금"):
                errors["stage"] = "수금 단계를 선택해주세요."
            if errors:
                ack(response_action="errors", errors=errors)
                return
            ack()
        except Exception as exc:
            logger.error(f"[SLACK/수금봇] submit_payment_intake 검증 실패: {exc}", exc_info=True)
            try:
                ack()
            except Exception:
                pass
            return

        def _bg():
            try:
                meta = json.loads(view.get("private_metadata") or "{}")
                user_id = (body.get("user") or {}).get("id", "")
                intake_id = meta.get("intake_id", "")
                channel = meta.get("channel", "")
                message_ts = meta.get("message_ts", "")
                d = _load_intake(intake_id)
                preview = d.get("preview") or {}
                amount = int(preview.get("amount") or 0)
                memo = d.get("text") or ""
                # 지정 내용만 저장 (아직 시트 기록 X — 샛별 확인 대기)
                _update_intake(intake_id, designation={
                    "project_code": project_code, "stage": stage,
                    "amount": amount, "memo": memo, "by": user_id,
                })
                if channel and message_ts:
                    client.chat_update(
                        channel=channel, ts=message_ts,
                        text=f"확인 대기: {project_code} · {stage}",
                        blocks=_build_intake_pending_blocks(
                            intake_id, project_code, stage, amount, memo, user_id),
                    )
            except Exception as exc:
                logger.error(f"[SLACK/수금봇] submit_payment_intake 처리 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.action("payment_intake_redesignate")
    def handle_payment_intake_redesignate(ack, body, client):
        """[✏️ 재지정] — 프로젝트/단계 다시 지정 (모달 재오픈). 아무 매니저 가능."""
        ack()
        def _bg():
            try:
                intake_id = body["actions"][0]["value"]
                channel = body["channel"]["id"]
                message_ts = body["message"]["ts"]
                _open_payment_intake_modal(client, body, intake_id, channel, message_ts)
            except Exception as exc:
                logger.error(f"[SLACK/수금봇] payment_intake_redesignate 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.action("payment_intake_confirm")
    def handle_payment_intake_confirm(ack, body, client):
        """[✅ 확인 후 기록] — 경영지원(황샛별)만. 이때 비로소 시트에 값+메모 기록."""
        ack()
        user = (body.get("user") or {}).get("id", "")
        channel = (body.get("channel") or {}).get("id", "")
        ts = (body.get("message") or {}).get("ts", "")
        intake_id = (body.get("actions") or [{}])[0].get("value", "")
        # 시트 기록 확인은 경영지원(황샛별)만 — 정산 ✅·금액 게이트와 동일 사상.
        if user != _SETTLEMENT_CHECKER_ID:
            _intake_ephemeral(client, channel, user,
                              ":lock: 시트 기록 확인은 경영지원만 가능합니다.")
            return

        def _bg():
            try:
                d = _load_intake(intake_id)
                des = d.get("designation") or {}
                project_code = (des.get("project_code") or "").strip()
                stage = (des.get("stage") or "").strip()
                amount = int(des.get("amount") or 0)
                memo = des.get("memo") or (d.get("text") or "")
                if not project_code or stage not in ("계약금", "중도금", "잔금"):
                    _intake_ephemeral(client, channel, user,
                                      ":warning: 지정 정보가 없습니다. [✏️ 재지정] 후 다시 확인해주세요.")
                    return
                if amount <= 0:
                    _intake_ephemeral(client, channel, user,
                                      ":warning: 금액이 자동 인식되지 않았습니다. 스레드에서 확인 후 수동 처리해주세요.")
                    return
                ok, old_num, new_num, err = _commit_intake_to_sheet(
                    project_code, stage, amount, memo, user)
                if not ok:
                    _intake_ephemeral(client, channel, user, f":warning: 기록 실패: {err}")
                    return
                if channel and ts:
                    client.chat_update(
                        channel=channel, ts=ts,
                        text=f"✅ {project_code} · {stage} {amount:,}원 확인 완료",
                        blocks=_build_intake_done_blocks(
                            project_code, stage, amount, memo, des.get("by", ""), user),
                    )
                    # 온라인 리드처럼 카드에 ✅ 리액션 = '처리 완료' 신호 (채널 스캔 가시성)
                    try:
                        _react_card_handled(client, channel, ts)
                    except Exception:
                        pass
                try:
                    from dashboard.utils.redis_client import get_redis_client
                    get_redis_client().redis.delete(f"sms_intake:{intake_id}")
                except Exception:
                    pass
            except Exception as exc:
                logger.error(f"[SLACK/수금봇] payment_intake_confirm 실패: {exc}", exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    logger.info(
        "[SLACK/수금봇] 핸들러 등록 완료: payment_card_delete, payment_intake_open, "
        "payment_intake_project, submit_payment_intake, payment_intake_confirm, "
        "payment_intake_redesignate"
    )


# ─────────────────────────────────────────────
# 수금 SMS 인입 — 모달/자동완성/시트 기록 헬퍼 (2026-08-12)
# ─────────────────────────────────────────────

def _build_payment_intake_options(query: str, limit: int = 30) -> list:
    """수금 인입 모달 프로젝트 자동완성 — 코드/사업자명/주소 매칭.

    반복 거래처(같은 상호 수십 건) 구분을 위해:
      - 라벨에 현장 주소 표시 (동일 상호 구분)
      - 주소도 검색어 대상 (예: '에스엘 역삼')
      - 정렬: 미수금 있는(수금 대기) 건 우선 → 최근(코드 번호 큰) 순
        → 옛 완납 프로젝트는 아래로 밀려 헷갈림 감소.
    """
    try:
        from dashboard.services.project_service import get_project_records
        recs = get_project_records() or []
    except Exception as exc:
        logger.warning(f"[SLACK/수금봇] 프로젝트 레코드 로드 실패: {exc}")
        return []
    q = (query or "").strip().lower()

    def _codenum(code):
        m = re.search(r'\d+', code)
        return int(m.group()) if m else 0

    def _owes(r):
        try:
            return 1 if abs(float(r.get('미수금') or 0)) >= 1 else 0
        except (TypeError, ValueError):
            return 0

    matched = []
    for r in recs:
        code = str(r.get('프로젝트 코드') or '').strip()
        if not code:
            continue
        biz = str(r.get('사업자명') or '').strip()
        addr = str(r.get('현장 주소') or '').strip()
        if q and q not in f"{code} {biz} {addr}".lower():
            continue
        # 사업자명이 길어도 주소(동일 상호 구분 핵심)가 안 잘리게 상호는 축약.
        biz_short = (biz[:13] + '…') if len(biz) > 14 else (biz or '-')
        label = f"{code} | {biz_short}"
        if addr:  # 남는 공간에 주소 채움 (Slack 라벨 75자 제한)
            remain = 74 - len(label) - 3
            if remain >= 5:
                label += " | " + addr[:remain]
        matched.append((_owes(r), _codenum(code), code, label[:75]))

    # 미수금 있는 것 우선 → 그다음 최근(코드 큰) 순
    matched.sort(key=lambda t: (t[0], t[1]), reverse=True)
    return [
        {"text": {"type": "plain_text", "text": label}, "value": code}
        for _owe, _num, code, label in matched[:limit]
    ]


def _open_payment_intake_modal(client, body, intake_id, channel, message_ts):
    """[지정] 버튼 → 프로젝트(자동완성)+수금단계만 지정하는 모달.

    금액·메모는 문자에서 자동 인식(읽기전용 표시)하고, 실제 시트 기록은 경영지원
    (황샛별)이 '확인 후 기록' 을 눌렀을 때만 반영된다. 여기선 '지정'만 한다.
    """
    trigger_id = body["trigger_id"]
    d = _load_intake(intake_id)
    text = d.get("text") or ""
    view = _build_payment_intake_view(intake_id, channel, message_ts, text)
    client.views_open(trigger_id=trigger_id, view=view)


def _build_payment_intake_view(intake_id, channel, message_ts, text,
                               selected_project=None, project_details=None, stage_value=None):
    """수금 지정 모달 view 빌더.

    프로젝트 선택 시(dispatch_action) 상세(project_details)를 삽입하고, 이미 고른
    단계(stage_value)를 보존해 재렌더한다. (A/S 요청 모달과 동일 사상 — 반복거래처
    상세 확인용)
    """
    blocks = [
        {"type": "section", "text": {"type": "mrkdwn", "text": f"```{text}```"}},
        {"type": "divider"},
    ]
    project_el = {
        "type": "external_select",
        "action_id": "payment_intake_project",
        "min_query_length": 1,
        "placeholder": {"type": "plain_text", "text": "코드·상호명·주소 검색"},
    }
    if selected_project:
        project_el["initial_option"] = selected_project
    blocks.append({
        "type": "input", "block_id": "project", "dispatch_action": True,
        "label": {"type": "plain_text", "text": "프로젝트"},
        "element": project_el,
    })
    if project_details:
        dd = project_details
        detail = "\n".join([
            f"*✅ 선택한 프로젝트*  `{dd.get('code', '')}`",
            f"• 사업자명 : {dd.get('biz', '-')}",
            f"• 현장 주소 : {dd.get('address', '-')}",
            f"• 담당자 : {dd.get('manager', '-')}",
            f"• 공사 내용 : {(dd.get('work_content', '-') or '-')[:120]}",
        ])
        blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": detail}})
    stage_el = {
        "type": "static_select",
        "action_id": "value",
        "placeholder": {"type": "plain_text", "text": "선택"},
        "options": [
            {"text": {"type": "plain_text", "text": t}, "value": t}
            for t in ("계약금", "중도금", "잔금")
        ],
    }
    if stage_value in ("계약금", "중도금", "잔금"):
        stage_el["initial_option"] = {
            "text": {"type": "plain_text", "text": stage_value}, "value": stage_value}
    blocks.append({
        "type": "input", "block_id": "stage",
        "label": {"type": "plain_text", "text": "수금 단계"},
        "element": stage_el,
    })
    blocks.append({"type": "context", "elements": [
        {"type": "mrkdwn", "text": ":lock: 실제 시트 기록은 경영지원 확인 후 반영됩니다."}
    ]})
    return {
        "type": "modal",
        "callback_id": "submit_payment_intake",
        "private_metadata": json.dumps({
            "intake_id": intake_id, "channel": channel, "message_ts": message_ts,
        }),
        "title": {"type": "plain_text", "text": "수금 지정"},
        "submit": {"type": "plain_text", "text": "지정"},
        "close": {"type": "plain_text", "text": "취소"},
        "blocks": blocks,
    }


def _intake_ephemeral(client, channel, user_id, text):
    if not channel or not user_id:
        return
    try:
        client.chat_postEphemeral(channel=channel, user=user_id, text=text)
    except Exception:
        pass


def _load_intake(intake_id):
    """Redis 인입 레코드 로드 (text/preview/designation)."""
    try:
        from dashboard.utils.redis_client import get_redis_client
        raw = get_redis_client().redis.get(f"sms_intake:{intake_id}")
        return json.loads(raw) if raw else {}
    except Exception as exc:
        logger.warning(f"[SLACK/수금봇] 인입 로드 실패: {exc}")
        return {}


def _update_intake(intake_id, **fields):
    """Redis 인입 레코드 부분 갱신 (TTL 7일 유지)."""
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        raw = rc.get(f"sms_intake:{intake_id}")
        d = json.loads(raw) if raw else {}
        d.update(fields)
        rc.set(f"sms_intake:{intake_id}", json.dumps(d), ex=60 * 60 * 24 * 7)
    except Exception as exc:
        logger.warning(f"[SLACK/수금봇] 인입 갱신 실패: {exc}")


def _commit_intake_to_sheet(project_code, stage, amount, memo_text, slack_user_id):
    """프로젝트 행 조회 → U/V/W 셀에 금액 값(기존값+합산)과 메모(append) 기록.

    카드 발송 트리거 조건이 '해당 stage 셀 값 > 0' 이므로 값과 노트를 둘 다 쓴다
    (SB 수동 흐름과 동일). 카드 갱신은 호출자(확인 핸들러)가 담당.
    Returns: (ok: bool, old_num: int, new_num: int, err: str)
    """
    from dashboard.constants import PAYMENT_FIELD_TO_COLUMN
    from dashboard.services.lead_service import get_sheets_manager

    col = PAYMENT_FIELD_TO_COLUMN.get(stage)
    if not col:
        return False, 0, 0, f"알 수 없는 단계 {stage}"
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    if not sheet_id or not sheet_name:
        return False, 0, 0, "시트 설정 오류(GOOGLE_SHEET_ID/NAME)"

    manager = get_sheets_manager()
    row = manager.find_row_by_project_code(sheet_id, project_code, f"{sheet_name}!A:A")
    if not row:
        return False, 0, 0, f"{project_code} 행을 시트에서 못 찾음"

    cell = f"{col}{row}"
    # 1) 금액 값 — 기존 값에 합산 (트리거 조건: 값 > 0)
    old_val_raw = manager.get_cell_value(sheet_id, sheet_name, cell)
    try:
        old_num = int(float(old_val_raw)) if str(old_val_raw).strip() not in ('', 'None') else 0
    except (ValueError, TypeError):
        old_num = 0
    new_num = old_num + int(amount)
    if not manager.update_cell_value(sheet_id, sheet_name, cell, new_num):
        return False, old_num, new_num, "금액 기록 실패"

    # 2) 메모(노트) — 기존 있으면 append (분납 대비). 실패해도 값은 기록됨.
    old_note = (manager.get_cell_note(sheet_id, sheet_name, cell) or '').rstrip()
    new_note = f"{old_note}\n\n{memo_text.strip()}" if old_note else memo_text.strip()
    if not manager.update_cell_note(sheet_id, sheet_name, cell, new_note):
        logger.warning(f"[SLACK/수금봇] 셀 메모 기록 실패(값은 기록됨): {cell}")

    try:
        from dashboard.utils.cache_invalidation import smart_invalidate
        smart_invalidate(f"cell_notes_{sheet_id}")
    except Exception:
        pass

    try:
        from dashboard.utils.user_database import get_audit_repository
        get_audit_repository().log_action(
            user_email=f"slack:{slack_user_id}",
            action='SMS_INTAKE_PAYMENT',
            details=f"수금 SMS 인입 확인기록 → {project_code} {stage} +{amount:,}원 (값 {old_num:,}→{new_num:,})",
            project_code=project_code,
            field_name=f"{stage}_수금",
            old_value=f"{old_num:,}",
            new_value=f"{new_num:,} (+{amount:,})",
            ip_address=None,
        )
    except Exception as exc:
        logger.warning(f"[SLACK/수금봇] 감사 로그 실패: {exc}")

    logger.info(f"[SLACK/수금봇] 수금 인입 기록: {project_code} {stage} +{amount:,}원 "
                f"값 {old_num:,}→{new_num:,} ({cell}) by {slack_user_id}")
    return True, old_num, new_num, ""


def _build_intake_pending_blocks(intake_id, project_code, stage, amount, memo, by_user):
    """지정 완료 → 경영지원 확인 대기 카드 (문자 원문 + 확인/재지정 버튼)."""
    amt = f"{amount:,}원" if amount else '—'
    warn = "" if amount else "\n:warning: 금액 자동인식 실패 — 확인 전 스레드로 금액 확인 필요"
    return [
        {"type": "section", "text": {"type": "mrkdwn", "text": (
            f"*🕓 확인 대기 — 경영지원 확인 후 기록*\n"
            f"*{project_code}*  ·  *{stage}*  ·  *{amt}*   ·   지정 <@{by_user}>{warn}")}},
        {"type": "section", "text": {"type": "mrkdwn", "text": f"```{(memo or '').strip()}```"}},
        {"type": "actions", "elements": [
            {"type": "button", "text": {"type": "plain_text", "text": "✅ 확인 후 기록"},
             "style": "primary", "action_id": "payment_intake_confirm", "value": intake_id},
            {"type": "button", "text": {"type": "plain_text", "text": "✏️ 재지정"},
             "action_id": "payment_intake_redesignate", "value": intake_id},
        ]},
        {"type": "context", "elements": [
            {"type": "mrkdwn", "text": f"intake `{intake_id}`"}
        ]},
    ]


def _build_intake_done_blocks(project_code, stage, amount, memo_text, by_user, confirmed_by):
    """온라인 리드 완료 카드와 동일 양식 — ✅ 헤더 + 메타 + 원문 code block, 버튼 없음."""
    who = (f"지정 <@{by_user}> · 확인 <@{confirmed_by}>" if by_user
           else f"확인 <@{confirmed_by}>")
    header = '\n'.join([
        '⠀',
        f":white_check_mark: *확인 완료 - {stage}*  `{project_code}`",
        f"금액 : {amount:,}원",
        f"처리 : {who}",
    ])
    text = f"{header}\n\n```\n{memo_text.strip()}\n```"
    return [
        {"type": "section", "text": {"type": "mrkdwn", "text": text}},
    ]


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
        _dups = result.get('duplicates') or []
        if _dups:
            lines.append(f":rotating_light: *중복 배정 {len(_dups)}건 — 같은 건이 2곳에 걸림 (캔버스 확인)*")
            for d in _dups[:8]:
                _k = '전화' if d.get('kind') == 'phone' else '주소'
                lines.append(f"  · [{_k}] {d.get('key')} → {' / '.join(d.get('assignees', []))}")
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
    _dups = result.get('duplicates') or []
    if _dups:
        lines.append(f"   :rotating_light: *중복 배정 {len(_dups)}건 — 같은 건이 2곳에 걸림 (커밋 전 캔버스 정리)*")
        for d in _dups[:8]:
            _k = '전화' if d.get('kind') == 'phone' else '주소'
            lines.append(f"      · [{_k}] {d.get('key')} → {' / '.join(d.get('assignees', []))}")
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
                not_license = bool(result.get('not_license'))
                is_card = bool(result.get('is_card'))

                # reaction 최종 상태 반영
                if saved and not_license:
                    _safe_react('warning')  # 저장은 됐으나 사업자등록증 아닐 가능성(오첨부)
                elif saved and not skipped:
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
                # 사업자등록증 오첨부 감지 (2026-08-10 G3879-JSH 카드 이미지) — 조용히 넘기지 말고 안내.
                if not_license and saved:
                    if is_card:
                        lines.append(
                            ":warning: 첨부가 *카드 이미지*로 보입니다 — 사업자등록증이 아닙니다. "
                            "올바른 사업자등록증인지 확인해 주세요. _(카드결제 건이면 무시하셔도 됩니다.)_"
                        )
                    else:
                        lines.append(
                            ":warning: 첨부에서 *사업자등록증 정보(사업자번호·상호)를 못 찾았습니다*. "
                            "사업자등록증이 맞는지 확인해 주세요. _(카드결제 건이면 무시하셔도 됩니다.)_"
                        )
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

            # ③-2-a. lead 카드 스레드 자동 감지 (부재중/드랍) — 2026-07-26
            # 매니저가 짧은 텍스트만 남겨도 시트·카드 자동 갱신. 매칭 X 시 forward 로 계속.
            try:
                if _try_auto_thread_status(client, event):
                    return
            except Exception as exc:
                logger.warning(f"[SLACK/자동감지] 예외 (무시): {exc}")

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
        # 재상담 신원 덮어쓰기 방어 (2026-07-31 L-03367) — 다른 고객의 완료 카드에
        # 잘못 [재상담] 눌러 기존 리드 고객명/연락처를 덮어쓰는 것 방지. 신원이
        # 바뀌면 경고 배너를 얹어 값 보존 재렌더 → 매니저가 확인 후 재제출해야 진행.
        # (같은 고객 재상담은 pre-fill 그대로라 안 걸림 → 정상 흐름 무마찰)
        try:
            _meta = json.loads(view.get("private_metadata") or "{}")
            if _meta.get("lead_no") and not _meta.get("_identity_confirmed"):
                _old = _find_lead_by_no(_meta["lead_no"]) or {}
                _idc = _consult_identity_changed(
                    _old, _v(state, "name"), _v(state, "contact"))
                if _old and _idc.get("changed"):
                    ack(response_action="update",
                        view=_build_consult_identity_confirm_view(_meta, state, _idc))
                    return
        except Exception as _gexc:
            logger.warning(f"[SLACK/상담] 신원 확인 게이트 skip (fallback 진행): {_gexc}")
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


# 정산 처리 완료 = 체크 리액션 → 자동 고정 해제 (2026-07-28)
# 사용자 요구: ✅(white_check_mark)만, 경영지원(황샛별)이 눌렀을 때만.
_SETTLE_DONE_REACTIONS = {'white_check_mark'}  # ✅ 만
# 황샛별 Slack ID (sb@itg-aircon.com). env 로 override 가능.
_SETTLEMENT_CHECKER_ID = os.getenv('SLACK_SETTLEMENT_CHECKER_ID', '').strip() or 'U0BHC2JV7U5'

# ── 공사 금액/부가세 수정 요청 (경영지원 ✅ 반영) ────────────────────────────
# 금액·부가세는 즉시 반영 대신 #영업_관리 요청 카드 발송 → 황샛별 ✅ 시 시스템이 반영.
# (PM 어드민의 '금액 변경은 어드민만' 사상. 오직 황샛별 ✅ 에서만 금액 확정.)
_AMOUNT_EDIT_FIELDS = ('총액 1', '부가세')


def _norm_edit_val(v) -> str:
    """공사 수정 diff 비교용 정규화 — '-'·빈문자·공백을 모두 빈값으로 취급.

    시트에 '-' 로 저장된 필드(도급 구분·시공자 등)가 모달 빈칸('') 제출과
    거짓 diff 나서 불필요한 '공사 내용 수정 알림' 카드 + '-'→'' 덮어쓰기를
    유발하던 것 방지 (2026-08-07 R3906-TH). 실제 값 변경/삭제는 그대로 감지.
    """
    s = str(v).strip() if v is not None else ''
    return '' if s == '-' else s


def _amt_int(v) -> int:
    """금액 문자열/숫자를 정수로 정규화 (콤마·공백·빈값 방어). 반영 여부 비교용."""
    try:
        return int(float(str(v).replace(',', '').strip() or 0))
    except (ValueError, TypeError):
        return 0


def _amount_request_applied(proj, before: dict, updates: dict):
    """금액 요청이 실제로 시트/캐시에 반영됐는지 판정.

    요청 필드(총액 1·부가세) 중 하나라도 요청 전(before) 값과 달라졌으면 반영됨(True).
    proj(레코드) 없으면 None(확인 불가 → 호출부는 차단하지 않음). 전부 그대로면 False.
    실제 반영값이 요청값과 달라도(경영지원 재량 조정) '변경됨'이면 반영으로 인정.
    """
    if proj is None:
        return None
    before = before or {}
    updates = updates or {}
    if '총액 1' in updates and _amt_int(proj.get('총액 1')) != _amt_int(before.get('총액 1')):
        return True
    if '부가세' in updates and _vat_is_sep(proj.get('부가세')) != _vat_is_sep(before.get('부가세')):
        return True
    return False


def _vat_is_sep(v) -> bool:
    """부가세 값(bool/str/int) → VAT 별도 여부."""
    return v is True or (isinstance(v, str) and v.strip().upper() in ('TRUE', 'Y', 'YES', '1')) or v == 1


def _invoice_client():
    """세금계산서 관리 알림 봇 WebClient. 미가용 시 None."""
    global _invoice_slack_app
    if _invoice_slack_app is None:
        try:
            _init_invoice_slack_app()
        except Exception:
            pass
    return _invoice_slack_app.client if _invoice_slack_app is not None else None


def _project_client():
    """공사 현황 알림 봇 WebClient (공사 정보 수정 카드 발송/수정용). 미가용 시 None.

    금액 수정 요청 카드는 '공사' 건이므로 공사봇 명의로 발송·갱신.
    ✅(reaction_added) 이벤트는 계산서봇이 수신하지만, 카드 chat_update 는
    카드를 올린 공사봇 클라이언트로 해야 함(같은 봇만 자기 메시지 수정 가능).
    """
    global _project_slack_app
    if _project_slack_app is None:
        try:
            _init_project_slack_app()
        except Exception:
            pass
    return _project_slack_app.client if _project_slack_app is not None else None


def _dm_client():
    """개인 DM 발송용 WebClient — im:write 있는 메인봇 토큰 사용.

    계산서봇·공사봇은 im:write 가 없어 conversations.open(DM 개설) 불가.
    메인봇(SLACK_BOT_TOKEN)은 DM 가능(검증됨) → 완료 DM 은 메인봇으로 발송.
    """
    from slack_sdk import WebClient
    tok = os.getenv('SLACK_BOT_TOKEN', '').strip()
    return WebClient(token=tok) if tok else None


def _maybe_auto_pin_settlement(client, channel: str, msg: dict) -> None:
    """정산 요청 메시지(입금내역·세금계산서)면 자동 고정(pin).

    #영업_관리 top-level 메시지 대상. 계좌번호(452/255/352) or 세금계산서 양식
    (MM/DD G/R 금액원 거래처) 매칭 시 pin. 처리는 담당자가 고정 해제로.
    """
    from dashboard.services.pin_remind import is_settlement_message, _msg_full_text
    text = _msg_full_text(msg)
    if not is_settlement_message(text):
        return
    ts = msg.get('ts')
    if not ts:
        return
    # 메시지당 1회만 자동 고정 — 담당자가 처리 후 해제한 걸 편집·언펄(message_changed)로
    # 재고정하지 않도록 Redis nx 가드 (90일).
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        if not rc.set(f'settlement_pin_seen:{channel}:{ts}', '1', nx=True, ex=60 * 60 * 24 * 90):
            return
    except Exception:
        pass
    try:
        client.pins_add(channel=channel, timestamp=ts)
        logger.info(f'[SLACK/정산핀] 자동 고정: ts={ts} | {text[:40]!r}')
    except Exception as exc:
        if 'already_pinned' in str(exc):
            return
        logger.warning(f'[SLACK/정산핀] pins.add 실패 (ts={ts}): {exc}')


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

        # 정산 요청 자동 고정 (2026-07-28) — #영업_관리 top-level 입금/세금계산서 메시지.
        # 첨부·언펄 편차 대응: 원본(빈 subtype/file_share) + message_changed 둘 다 검사.
        # 봇 메시지·스레드 답글 제외. pins.add 는 already_pinned 처리 (중복 무해).
        if channel and not event.get('bot_id'):
            _pin_msg = None
            if subtype in ('', 'file_share', None) and not thread_ts:
                _pin_msg = event
            elif subtype == 'message_changed':
                _cand = event.get('message') or {}
                _cts = _cand.get('thread_ts')
                if not (_cts and _cts != _cand.get('ts')):  # 스레드 답글 edit 제외
                    _pin_msg = _cand
            if _pin_msg and not _pin_msg.get('bot_id'):
                _pm, _pch = dict(_pin_msg), channel

                def _bg_pin():
                    try:
                        _maybe_auto_pin_settlement(client, _pch, _pm)
                    except Exception as exc:
                        logger.debug(f'[SLACK/정산핀] 자동 고정 실패 (무시): {exc}')
                threading.Thread(target=_bg_pin, daemon=True).start()

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

    @app.event("reaction_added")
    def handle_invoice_reaction_added(event, client):
        """#영업_관리 정산 메시지에 체크(✅) 리액션 → 자동 고정 해제 (2026-07-28).

        경영지원이 처리 완료 시 핀 수동 해제 대신 체크만 하면 목록에서 빠짐.
        """
        reaction = event.get('reaction', '')
        user = event.get('user', '')
        item = event.get('item', {}) or {}
        channel = item.get('channel', '')
        ts = item.get('ts', '')
        inv_ch = os.getenv('SLACK_INVOICE_CHANNEL_ID', '').strip()
        # 진단 로그 (DEBUG) — 이벤트 도착·필터 사유 확인용. 평소 미출력.
        logger.debug(
            f'[SLACK/정산핀] reaction_added 수신: reaction={reaction} user={user} '
            f'ch={channel} inv_ch={inv_ch} type={item.get("type")}'
        )
        if reaction not in _SETTLE_DONE_REACTIONS:
            return
        if item.get('type') != 'message':
            return
        if not ts or not channel or channel != inv_ch:
            logger.debug(f'[SLACK/정산핀] skip — 채널 불일치 (ch={channel} != inv_ch={inv_ch})')
            return
        # 경영지원(황샛별)이 누른 체크만 인정 — 아무나 ✅ 눌러 해제되면 안 됨.
        if user != _SETTLEMENT_CHECKER_ID:
            logger.debug(f'[SLACK/정산핀] skip — 체커 아님 (user={user} != checker={_SETTLEMENT_CHECKER_ID})')
            return

        def _bg():
            try:
                # 1) 공사 금액 수정 요청 카드면 → perform_edit 반영 + 요청자 DM (핀 로직 skip)
                if _maybe_apply_amount_request(client, channel, ts, user):
                    return
                # 2) 정산 핀 자동 해제
                client.pins_remove(channel=channel, timestamp=ts)
                logger.info(f'[SLACK/정산핀] 체크 리액션 → 자동 고정 해제: ts={ts}')
            except Exception as exc:
                if 'no_pin' in str(exc):
                    return  # 고정 안 돼있던 메시지 — 무해
                logger.warning(f'[SLACK/정산핀] pins.remove 실패 (ts={ts}): {exc}')
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
                    code = p["code"]
                    biz = (p.get('biz') or '').strip() or '사업자 비어 있음'
                    addr = (p.get('address') or '').strip()
                    # 사업자명 축약 → 주소(반복거래처 구분 핵심)가 안 잘리게 (수금 인입과 동일)
                    biz_short = (biz[:13] + '…') if len(biz) > 14 else biz
                    label = f'{code} | {biz_short}'
                    if addr:
                        remain = 74 - len(label) - 3
                        if remain >= 5:
                            label += ' | ' + addr[:remain]
                    options.append({
                        "text": {"type": "plain_text", "text": label[:75]},
                        "value": code[:75],
                    })
                ack(options=options)
            except Exception as exc:
                logger.warning(f"[SLACK/AS/options] 실패: {exc}", exc_info=True)
                ack(options=[])
        else:
            ack(options=[])

    @app.action("value")
    def handle_as_block_action(ack, body, client):
        """모달 내 external_select 선택 → pre-fill / 코드없음 체크박스 → 모드 스왑."""
        ack()
        if not body.get("view"):
            return
        action = (body.get("actions") or [{}])[0]
        bid = action.get("block_id")
        if bid == "as_project_code":
            def _bg():
                try:
                    _update_as_modal_with_project(client, body, action)
                except Exception as exc:
                    logger.error(f"[SLACK/AS] 모달 갱신 실패: {exc}", exc_info=True)
            threading.Thread(target=_bg, daemon=True).start()
        elif bid == "as_manual_toggle":
            def _bg2():
                try:
                    _toggle_as_manual_mode(client, body, action)
                except Exception as exc:
                    logger.error(f"[SLACK/AS] 수동모드 토글 실패: {exc}", exc_info=True)
            threading.Thread(target=_bg2, daemon=True).start()

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

    # 코드 없는 수동 등록 A/S (코드 이전 공사) → 프로젝트 파생 필드 생략, 간소 렌더
    if not (code and code != '-'):
        lines.append("🏷️ 구분 : 코드 이전 공사 (수동 등록)")
        lines.append(f"📍 현장 주소 : {_pick('현장주소', 'address')}")
        _mwork = _pick('공사내용', 'work_content')
        if _mwork and _mwork != '-':
            lines.append(f"📋 공사 내용 : {_mwork}")
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


# A/S 담당자 드롭다운 제외 대상 — 경영지원팀 (영업 아님). email localpart 소문자.
_AS_MANAGER_EXCLUDE = {'kiko', 'sb'}


def _active_manager_options() -> list:
    """A/S 수동 등록용 담당자(영업) static_select 옵션 — 활성 영업 인원만.

    users.db 에서 퇴사자(resigned*)·테스트(@example.com)·경영지원팀(_AS_MANAGER_EXCLUDE) 제외.
    value=email. 2026-07-27 코드 이전 공사 A/S 등록 시 담당자 직접 지정용.
    """
    opts: list = []
    try:
        from dashboard.utils.user_database import UserDatabase
        db = UserDatabase()
        rows = []
        for u in db.get_all_users():
            email = (u.get('email') or '').strip()
            name = (u.get('name') or '').strip()
            if not email or not name:
                continue
            local = email.split('@')[0].strip()
            if local.lower().startswith('resigned') or email.lower().endswith('@example.com'):
                continue
            if local.lower() in _AS_MANAGER_EXCLUDE:
                continue
            ini = local.upper()
            if ini == 'KIKO':
                ini = 'KiKO'
            rows.append((name, ini, email))
        rows.sort(key=lambda r: r[0])
        for name, ini, email in rows:
            label = f'{name} ({ini})'
            opts.append({
                "text": {"type": "plain_text", "text": label[:75]},
                "value": email[:75],
            })
    except Exception as exc:
        logger.warning(f'[SLACK/AS] 담당자 옵션 로드 실패: {exc}')
    return opts


def _as_request_view_blocks(
    initial_project_option: Optional[dict] = None,
    project_details: Optional[dict] = None,
    initial_request_content: str = '',
    manual_mode: bool = False,
) -> list:
    """요청 모달 blocks — 프로젝트 선택 전/후 + 코드 없는 수동 입력 모드 공용.

    manual_mode=True (코드 이전 공사) → 코드 external_select 대신 현장 주소·공사 내용·
    담당자(영업) 직접 입력. 최상단 체크박스로 두 모드 전환 (views.update, 2026-07-27).
    """
    # 코드 없음 토글 (actions 블록 — 즉시 block_actions 발동, dispatch_action 불필요)
    _toggle_opt = {
        "text": {"type": "plain_text", "text": "프로젝트 코드 없음 (코드 이전 공사)"},
        "value": "manual",
    }
    toggle_checkbox = {
        "type": "checkboxes", "action_id": "value", "options": [_toggle_opt],
    }
    if manual_mode:
        toggle_checkbox["initial_options"] = [_toggle_opt]
    blocks: list = [{
        "type": "actions", "block_id": "as_manual_toggle",
        "elements": [toggle_checkbox],
    }]

    if manual_mode:
        # 코드 없는 공사 — 현장 정보·담당자 직접 입력
        addr_el = {
            "type": "plain_text_input", "action_id": "value",
            "placeholder": {"type": "plain_text",
                            "text": "예: 반포자이 아파트, 서울 서초구 신반포로 …"},
        }
        blocks.append({
            "type": "input", "block_id": "as_manual_address",
            "label": {"type": "plain_text", "text": "현장명 / 현장 주소"},
            "element": addr_el,
        })
        work_el = {
            "type": "plain_text_input", "action_id": "value", "multiline": True,
            "placeholder": {"type": "plain_text", "text": "예: 천장형 4way 3대 설치 (2023년)"},
        }
        blocks.append({
            "type": "input", "block_id": "as_manual_work", "optional": True,
            "label": {"type": "plain_text", "text": "공사 내용 (선택)"},
            "element": work_el,
        })
        mgr_options = _active_manager_options()
        if mgr_options:
            mgr_el = {
                "type": "static_select", "action_id": "value",
                "placeholder": {"type": "plain_text", "text": "담당자 선택 (영업)"},
                "options": mgr_options,
            }
        else:
            # users.db 로드 실패 fallback — 이름 직접 입력
            mgr_el = {
                "type": "plain_text_input", "action_id": "value",
                "placeholder": {"type": "plain_text", "text": "담당자 이름 (예: 고광일)"},
            }
        blocks.append({
            "type": "input", "block_id": "as_manual_manager",
            "label": {"type": "plain_text", "text": "담당자 (영업) — 알림 DM 대상"},
            "element": mgr_el,
        })
    else:
        # 기존 — 확정 프로젝트 코드 검색·선택
        project_element = {
            "type": "external_select", "action_id": "value",
            "min_query_length": 1,
            "placeholder": {"type": "plain_text", "text": "예: G3745 / R3845 (1글자부터 검색)"},
        }
        if initial_project_option:
            project_element["initial_option"] = initial_project_option
        blocks.append({
            "type": "input", "block_id": "as_project_code",
            "label": {"type": "plain_text", "text": "프로젝트 코드 (검색해서 선택)"},
            "element": project_element,
            "dispatch_action": True,  # 선택 즉시 block_actions 발동해 상세 pre-fill
        })
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

    # 공통 — A/S 요청 내용
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


def _toggle_as_manual_mode(client, body, action) -> None:
    """'프로젝트 코드 없음' 체크박스 토글 → 수동/일반 모드 스왑 후 views.update."""
    checked = bool(action.get("selected_options"))
    view = body["view"]
    view_id = view.get("id", '')
    view_hash = view.get("hash", '')
    metadata = view.get("private_metadata", '') or json.dumps({}, ensure_ascii=False)
    # 이미 입력한 A/S 요청 내용 보존
    current_content = ''
    try:
        current_content = (
            (view.get("state", {}) or {}).get("values", {})
            .get("request_content", {}).get("value", {})
            .get("value", '') or ''
        )
    except Exception:
        current_content = ''
    new_view = {
        "type": "modal",
        "callback_id": "submit_as_request",
        "private_metadata": metadata,
        "title": {"type": "plain_text", "text": "A/S 요청"},
        "submit": {"type": "plain_text", "text": "제출"},
        "close": {"type": "plain_text", "text": "취소"},
        "blocks": _as_request_view_blocks(
            manual_mode=checked,
            initial_request_content=current_content,
        ),
    }
    try:
        client.views_update(view_id=view_id, hash=view_hash, view=new_view)
    except Exception as exc:
        logger.warning(f"[SLACK/AS] 수동모드 토글 views_update 실패: {exc}")


def _send_as_manager_dm(client, as_no: str, project_code: str, card_data: dict,
                        proj: Optional[dict] = None,
                        card_channel: str = '', card_ts: str = '',
                        override: Optional[dict] = None) -> bool:
    """A/S 요청 등록 시 프로젝트 담당자(영업)에게 알림 DM.

    버튼 없는 안내 전용 — "담당하신 현장에 A/S 접수" + 채널 카드 permalink.
    접수는 다른 사람이 할 수 있으므로 담당자에겐 인지만 시킴 (2026-07-27 사용자 요구).

    override={'name','email'} 전달 시 코드 기반 조회 대신 그 담당자로 발송
    (코드 없는 수동 등록 A/S — 담당자 직접 지정).

    Returns: 발송 성공 여부. 담당자 미식별·이미 발송 시 False.
    """
    from dashboard.services.as_service import resolve_project_manager

    if override and (override.get('email') or '').strip():
        mgr_name = (override.get('name') or '').strip()
        email = (override.get('email') or '').strip()
    else:
        mgr_name, email = resolve_project_manager(project_code, proj=proj)
    if not email:
        logger.info(f'[SLACK/AS] 담당자 미식별 ({project_code or "수동"}) — DM skip')
        return False

    # 재발송 방지 (view_submission 재전송 대응)
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        if not rc.set(f'as_manager_dm:{as_no}', email, nx=True, ex=60 * 60 * 24 * 30):
            logger.info(f'[SLACK/AS] 담당자 DM 이미 발송 ({as_no}) — skip')
            return False
    except Exception:
        pass

    # 채널 카드 permalink (있으면 상세 보기 링크)
    permalink = ''
    if card_channel and card_ts:
        try:
            _pl = client.chat_getPermalink(channel=card_channel, message_ts=card_ts)
            permalink = (_pl or {}).get('permalink', '') or ''
        except Exception as exc:
            logger.debug(f'[SLACK/AS] permalink 조회 실패 ({as_no}): {exc}')

    biz = (proj or {}).get('biz', '-') or '-'
    address = card_data.get('현장주소') or (proj or {}).get('address', '-') or '-'
    work = card_data.get('공사내용') or (proj or {}).get('work_content', '-') or '-'
    req_content = card_data.get('요청 내용', '-') or '-'
    requester = card_data.get('요청자', '-') or '-'

    lines = [
        '⠀',
        f'🔔 *담당하신 현장에 A/S 요청이 접수되었습니다*  `{as_no}`',
        '--------------------------------------------',
    ]
    if project_code:
        lines.append(f'🔗 프로젝트 코드 : `{project_code}`')
    else:
        lines.append('🏷️ 코드 이전 공사 (수동 등록)')
    if biz and biz != '-':
        lines.append(f'🏢 사업자명 : {biz}')
    lines.append(f'📍 현장 주소 : {address}')
    if work and work != '-':
        lines.append(f'📋 공사 내용 : {work}')
    lines.append('--------------------------------------------')
    lines.append(f'📝 A/S 요청 내용 : {req_content}')
    lines.append(f'👤 요청자 : {requester}')
    lines.append('--------------------------------------------')
    if permalink:
        lines.append(f'🔗 <{permalink}|사후 관리 채널에서 상세 보기>')
    lines.append('_접수는 사후 관리 채널에서 진행됩니다._')
    lines.append('⠀')
    dm_text = '\n'.join(lines)

    try:
        u = client.users_lookupByEmail(email=email)
        uid = u['user']['id']
        r = client.chat_postMessage(
            channel=uid, text=dm_text,
            unfurl_links=False, unfurl_media=False,
        )
        if r.get('ok'):
            logger.info(f'[SLACK/AS] 담당자 DM 발송 완료: {as_no} → {mgr_name}({email})')
            return True
        logger.warning(f'[SLACK/AS] 담당자 DM 응답 not ok ({as_no}): {r.get("error")}')
    except Exception as exc:
        logger.warning(f'[SLACK/AS] 담당자 DM 발송 예외 ({as_no}, {email}): {exc}')
    return False


def _process_as_request_submission(client, body, view) -> None:
    """요청 제출 → (일반) 프로젝트 조회 또는 (수동) 직접 입력 → 시트 append → 카드 발송.

    수동 모드 = '프로젝트 코드 없음' 체크 (코드 이전 공사). as_manual_address 블록 존재로 판별.
    """
    from dashboard.services.as_service import get_project_details, create_as_row

    values = view["state"]["values"]
    user_id = body.get("user", {}).get("id", "")
    requester_initial = _slack_user_to_initial(client, user_id) or '-'

    def _pt(block_id: str) -> str:
        try:
            return ((values.get(block_id, {}).get("value", {}) or {}).get("value", '') or '').strip()
        except Exception:
            return ''

    request_content = _pt("request_content")
    is_manual = "as_manual_address" in values

    if is_manual:
        # 코드 없는 공사 — 현장 정보·담당자 직접 입력
        address = _pt("as_manual_address")
        work_content = _pt("as_manual_work")
        if not address:
            logger.warning('[SLACK/AS] 수동 등록 현장 주소 누락')
            return
        # 담당자: static_select(value=email) 또는 fallback plain_text(이름)
        mgr_name, mgr_email = '', ''
        mgr_state = values.get("as_manual_manager", {}).get("value", {}) or {}
        opt = mgr_state.get("selected_option")
        if opt:
            mgr_email = (opt.get("value") or '').strip()
            mgr_name = re.sub(
                r'\s*\([^)]*\)\s*$', '',
                (opt.get("text", {}) or {}).get("text", '') or '',
            ).strip()
        else:
            mgr_name = (mgr_state.get("value") or '').strip()
            if mgr_name:
                from dashboard.services.as_service import _email_from_manager_name
                mgr_email = _email_from_manager_name(mgr_name)
        project_code = ''
        details: dict = {}
        as_no, row_num = create_as_row(
            project_code='', address=address, work_content=work_content,
            work_end='', request_content=request_content, requester=requester_initial,
        )
        dm_override = {'name': mgr_name, 'email': mgr_email}
    else:
        # 기존 — 확정 프로젝트 코드 선택
        project_code = ''
        try:
            opt = values.get("as_project_code", {}).get("value", {}).get("selected_option", {})
            project_code = (opt or {}).get("value", "") or ''
        except Exception:
            pass
        project_code = project_code.strip()
        if not project_code:
            logger.warning('[SLACK/AS] 프로젝트 코드 누락')
            return
        details = get_project_details(project_code) or {}
        address = details.get('address', '')
        work_content = details.get('work_content', '')
        as_no, row_num = create_as_row(
            project_code=project_code,
            address=address,
            work_content=work_content,
            work_end=details.get('work_end', ''),
            request_content=request_content,
            requester=requester_initial,
        )
        dm_override = None

    channel = os.getenv('SLACK_AS_CHANNEL', '').strip()
    if not channel:
        logger.warning('[SLACK/AS] SLACK_AS_CHANNEL 미설정 — 카드 발송 skip')
        return

    card_data = {
        'No': as_no,
        '프로젝트 코드': project_code,
        '현장주소': address,
        '공사내용': work_content,
        '공사 종료일': '' if is_manual else details.get('work_end', ''),
        '요청 내용': request_content,
        '요청자': requester_initial,
    }
    text = f"[A/S 요청] {as_no} {project_code or '(코드 없음)'}"
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
        logger.info(f'[SLACK/AS] 요청 카드 발송 완료: {as_no} ts={ts} (manual={is_manual})')
        # 담당자(영업)에게 알림 DM — "담당 현장에 A/S 접수" (버튼 없는 안내)
        try:
            _send_as_manager_dm(client, as_no, project_code, card_data,
                                proj=(details or None), card_channel=channel, card_ts=ts,
                                override=dm_override)
        except Exception as exc:
            logger.warning(f'[SLACK/AS] 담당자 DM 예외 (무시): {exc}')


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

    # 진단 로그 (DEBUG) — 필요 시 이 앱 도착 이벤트 타입 확인용. 평소 미출력.
    try:
        _b = request.get_json(silent=True) or {}
        _ev = (_b.get('event') or {})
        logger.debug(
            f"[SLACK/계산서봇] 이벤트 수신: outer={_b.get('type')} "
            f"event={_ev.get('type')} reaction={_ev.get('reaction')} user={_ev.get('user')}"
        )
    except Exception:
        pass

    if _is_slack_retry_duplicate():
        return jsonify({"ok": True, "dedup": True}), 200

    return _invoice_slack_handler.handle(request)


@slack_bp.route("/payment-events", methods=["POST"])
def slack_payment_events():
    """슬랙 → 수금 관리 알림 봇 전용 endpoint (정정 카드 [🗑 삭제] 버튼 인터랙션)"""
    if _payment_slack_handler is None:
        if not _init_payment_slack_app():
            return jsonify({"error": "Payment Slack bot not configured"}), 503

    if _is_slack_retry_duplicate():
        return jsonify({"ok": True, "dedup": True}), 200

    return _payment_slack_handler.handle(request)


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
                category_display = f"온라인 ({platform})" if platform else '온라인'

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


def _replace_last_consult_content(old_text: str, new_content: str) -> str:
    """옛 K열 상담 내용의 마지막 회차 content 만 new_content 로 교체.

    2026-07-24 도입 — [정보 수정] 저장 시 이전 회차 이력 유지 + 최신 회차만 편집.
    옛 형식 (헤더 없음) 이거나 회차 파싱 실패 시 new_content 로 통째 교체.

    예:
      old = '[07.20 · 유선] 첫 회차 ─── [07.24 · 방문 예약] 최신 회차'
      new_content = '최신 회차 수정본'
      → '[07.20 · 유선] 첫 회차 ─── [07.24 · 방문 예약] 최신 회차 수정본'
    """
    if not old_text:
        return new_content or ''
    entries = _parse_consultation_entries(old_text)
    if not entries:
        return new_content or ''
    def _rebuild_entry(e, body):
        if e.get('time') and e.get('ini') and e.get('status'):
            return f"[{e['time']} {e['ini']} · {e['status']}] {body}"
        return body
    last = entries[-1]
    last_str = _rebuild_entry(last, (new_content or '').strip())
    prev_strs = [_rebuild_entry(e, e['content']) for e in entries[:-1]]
    if prev_strs:
        divider = f'\n{_CONSULT_DIVIDER}\n'
        return divider.join(prev_strs + [last_str])
    return last_str


def _parse_consultation_entries(text: str) -> list:
    """저장된 상담 내용 → 회차별 dict 리스트.

    각 회차: {'time':..., 'ini':..., 'status':..., 'content':...}
    옛 형식(헤더 없음) 은 ini/status 빈 값으로 content 만 채워 반환.
    """
    entries = []
    if not text:
        return entries
    # 두 divider 형태 대응: 표준 '\n─────────\n'(9칸 라인) + 방문완료 등이 붙이는
    #   인라인 ' ─── '(3칸). ─ 3개 이상 연속을 구분자로 분할.
    #   (2026-07-30 L-03371 등: 인라인 divider 미분할 → 뒤 회차 헤더가 content 에 남던 이슈)
    for chunk in re.split(r'\s*─{3,}\s*', text):
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


# ─────────────────────────────────────────────────────────────
# 스레드 텍스트 자동 감지 → 부재중/드랍 자동 처리 (2026-07-26)
#
# 매니저가 lead 카드 스레드에 짧은 텍스트만 남기고 [상담하기] 모달 안 눌러도
# 시트·카드 자동 갱신. 자연어 오탐 방지 위해 정확 매치·짧은 텍스트만 인식.
# ─────────────────────────────────────────────────────────────
_AUTO_STATUS_ABSENT_RE = re.compile(
    r'^(부재중|부재|전화\s*안\s*받음|안\s*받음|노쇼|no\s*show)\s*[.!?~,]*\s*$',
    re.IGNORECASE,
)
_AUTO_STATUS_DROP_RE = re.compile(
    r'^(문의\s*)?드랍(\s*처리)?\s*[.!?~,]*\s*$',
    re.IGNORECASE,
)


def _detect_auto_thread_status(text: str) -> Optional[str]:
    """스레드 reply 텍스트 → 자동 처리 상태 매핑. None 이면 자동 처리 대상 아님.

    자연어 오탐 방지 — 20자 이하 단일 라인 텍스트만 인식.
    """
    if not text:
        return None
    stripped = text.strip()
    if len(stripped) > 20 or '\n' in stripped:
        return None
    if _AUTO_STATUS_ABSENT_RE.match(stripped):
        return '부재중'
    if _AUTO_STATUS_DROP_RE.match(stripped):
        return '문의 드랍'
    return None


def _extract_lead_no_from_root(root_msg: dict) -> str:
    """Root 메시지 (lead 카드) 에서 lead_no 추출. `L-XXXXX` 또는 `ETC-xxxxxx` 패턴."""
    text_pool = [root_msg.get('text') or '']
    for blk in root_msg.get('blocks') or []:
        _bt = blk.get('type')
        if _bt == 'section':
            _t = (blk.get('text') or {}).get('text', '')
            if _t:
                text_pool.append(_t)
        elif _bt == 'context':
            for el in blk.get('elements') or []:
                if isinstance(el, dict):
                    _t = el.get('text', '')
                    if _t:
                        text_pool.append(_t)
    joined = '\n'.join(text_pool)
    m = re.search(r'\b(L-\d{4,6}|ETC-[a-f0-9]{4,10})\b', joined)
    return m.group(1) if m else ''


def _is_absent_badge_block(b: dict) -> bool:
    """카드 blocks 중 부재중 배지 블록인지 판정 (모달·자동감지 공통, 2026-08-06).

    배지 양식: section text 에 '*부재중*' + '처리 시간' (구 context 양식도 감지).
    새 배지 prepend 전 기존 부재중 배지를 **전부 제거**해 스택(배지 여러 개 쌓임)을
    방지·자가치유하는 데 사용. (모달 경로가 '마지막 시도' 로 오판정해 매번 삽입 →
    배지 중복되던 L-03527 사고.)
    """
    if not isinstance(b, dict):
        return False
    if b.get('type') == 'section':
        t = ((b.get('text') or {}).get('text') or '')
        return '*부재중*' in t and '처리 시간' in t
    if b.get('type') == 'context':
        return any(
            '부재중' in ((el.get('text') or '') if isinstance(el, dict) else '')
            for el in (b.get('elements') or [])
        )
    return False


def _apply_auto_absent_badge(client, channel: str, thread_ts: str, root: dict,
                              lead_no: str, initial: str, hdr_time: str):
    """부재중 배지 삽입 — 원본 body 유지 + 상단 section 배지 (task #32 정책).

    이미 배지 있으면 갱신 (재시도 케이스 → 회차·시각만 업데이트).
    """
    from dashboard.utils.redis_client import get_redis_client
    rc = get_redis_client().redis
    _count = int(rc.incr(f'consult_missed_count:{lead_no}') or 1)
    rc.expire(f'consult_missed_count:{lead_no}', 60 * 60 * 24 * 90)

    badge_text = '\n'.join([
        '⠀',
        f':arrows_counterclockwise: *부재중* (총 *{_count}회*)',
        f'처리자 : {initial}',
        f'처리 시간 : {hdr_time}',
        '상담 내용 : 부재중',
    ])
    badge_block = {'type': 'section', 'text': {'type': 'mrkdwn', 'text': badge_text}}
    # 기존 부재중 배지 전부 제거 후 최신 1개 prepend (스택 방지·자가치유)
    existing = [b for b in (root.get('blocks') or []) if not _is_absent_badge_block(b)]
    existing.insert(0, badge_block)
    client.chat_update(
        channel=channel, ts=thread_ts,
        text=root.get('text', '') or '',
        blocks=existing,
    )


def _apply_auto_dropped_card(client, channel: str, thread_ts: str, root: dict,
                              lead_no: str, initial: str, hdr_time: str,
                              reason_text: str):
    """문의 드랍 회색 완료 카드 — 원본 code block + 회색 헤더 + [재상담] 버튼."""
    original_text = ''
    for blk in root.get('blocks') or []:
        if blk.get('type') == 'section':
            original_text = (blk.get('text') or {}).get('text', '')
            break
    if '```' in original_text and '상담 완료' in original_text:
        return  # 이미 회색 처리됨

    cleaned = [ln.lstrip('>').lstrip().replace('*', '') for ln in original_text.split('\n')]
    clean_text = re.sub(r'^[\s⠀]+|[\s⠀]+$', '', '\n'.join(cleaned))
    try:
        from dashboard.blueprints.slack_helpers import _normalize_shortcodes_to_unicode
        clean_text = _normalize_shortcodes_to_unicode(clean_text)
    except Exception:
        pass

    header_lines = [
        '⠀',
        f':white_check_mark: *상담 완료 - 문의 드랍*  `{lead_no}`',
        f'처리자 : {initial}',
        f'처리 시간 : {hdr_time}',
        f'상담 내용 : {reason_text}',
    ]
    new_text = '\n'.join(header_lines) + f'\n\n```\n{clean_text}\n```'
    new_blocks = [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': new_text}},
        {'type': 'actions', 'elements': [{
            'type': 'button',
            'text': {'type': 'plain_text', 'text': '✏️ 재상담', 'emoji': True},
            'value': lead_no,
            'action_id': 'button_consult',
        }]},
    ]
    client.chat_update(
        channel=channel, ts=thread_ts, text=new_text, blocks=new_blocks,
    )


def _react_card_handled(client, channel: str, ts: str) -> bool:
    """리드 카드 본체에 ✅(white_check_mark) 리액션 = '리드 처리 완료' 신호.

    모달 완료·기존 lead 연결·자동 스레드 감지(부재중/드랍) 등 모든 처리 경로가
    이 헬퍼로 카드 root 에 ✅ 를 달아 "카드 ✅ = 리드 처리 완료" 신호를 통일
    (2026-08-04). 기존은 자동감지만 카드가 아닌 매니저 답글에 ✅ 를 달아
    채널 스캔 시 처리 여부가 안 보였음.

    already_reacted(재처리) 등 실패는 debug 로그만 — 카드 회색화로 이미 표시되므로
    치명적 아님. 조용한 except:pass 대신 로그로 원인 추적 가능하게.

    Returns: 성공 True / 실패·skip False.
    """
    if not (channel and ts):
        return False
    try:
        client.reactions_add(channel=channel, timestamp=ts, name='white_check_mark')
        return True
    except Exception as exc:
        logger.debug(f'[SLACK] 카드 ✅ 리액션 skip ({channel}/{ts}): {exc}')
        return False


def _try_auto_thread_status(client, event: dict) -> bool:
    """스레드 텍스트 자동 감지 → lead 상태 처리.

    Returns:
      True — 자동 처리 대상 스레드 (성공/이미 완료된 skip 모두 포함).
             호출자는 이후 채널톡 forward 등을 건너뛰어야 함.
      False — 자동 처리 대상 아님 (텍스트 unmatch, lead 카드 아님, 채널 mismatch).
             호출자는 기존 flow 계속 진행.
    """
    text = (event.get('text') or '').strip()
    new_status = _detect_auto_thread_status(text)
    if not new_status:
        return False

    channel = event.get('channel', '')
    thread_ts = event.get('thread_ts', '')
    user_id = event.get('user', '')
    event_ts = event.get('ts', '')
    if not (channel and thread_ts and user_id and event_ts):
        return False

    # 채널 필터 — 온라인_문의 채널 only (오탐 예방)
    try:
        lead_channel_setting = os.getenv('SLACK_LEAD_CHANNEL', '').strip()
        if not lead_channel_setting:
            return False
        from dashboard.services.lead_sync import _resolve_channel_id
        if channel != _resolve_channel_id(client, lead_channel_setting):
            return False
    except Exception:
        return False

    # Redis dedup — 동일 이벤트 재처리 방지 (bolt 재전송 대응)
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        if not rc.set(f'auto_thread_status:{event_ts}', '1', nx=True, ex=600):
            return True
    except Exception:
        pass

    # Root fetch → lead_no 추출
    try:
        rp = client.conversations_replies(
            channel=channel, ts=thread_ts, limit=1, inclusive=True,
        )
    except Exception as exc:
        logger.warning(f'[SLACK/자동감지] replies fetch 실패: {exc}')
        return False
    root = ((rp.get('messages') or [{}])[0]) if rp else {}
    lead_no = _extract_lead_no_from_root(root)
    if not lead_no:
        return False  # lead 카드 아님 — 기존 flow 계속

    try:
        from dashboard.services.lead_service import get_lead_by_no, update_lead
        lead = get_lead_by_no(lead_no)
    except Exception as exc:
        logger.warning(f'[SLACK/자동감지] lead 조회 실패 ({lead_no}): {exc}')
        return True  # lead_no 는 감지됐으니 forward 는 방지
    if not lead:
        return True
    cur_status = str(lead.get('상태') or '').strip()
    # 재처리 허용: 부재중 재시도만. 이미 완료된 lead 는 안내 후 skip.
    _allow_retry = {'상담 대기', '부재중'} if new_status == '부재중' else {'상담 대기'}
    if cur_status not in _allow_retry:
        try:
            client.chat_postEphemeral(
                channel=channel, user=user_id, thread_ts=thread_ts,
                text=(
                    f':information_source: `{lead_no}` 은(는) 이미 `{cur_status}` 상태입니다.\n'
                    '재상담이 필요하면 [✏️ 재상담] 버튼을 눌러주세요.'
                ),
            )
        except Exception:
            pass
        logger.info(f'[SLACK/자동감지] {lead_no} 이미 {cur_status} → skip')
        return True

    initial = _slack_user_to_initial(client, user_id) or '-'
    manager_name = _slack_user_to_korean_name(client, user_id) or ''
    try:
        _dt = datetime.fromtimestamp(float(event_ts))
    except Exception:
        _dt = datetime.now()
    hdr_time = _dt.strftime('%m.%d %H:%M')

    old_k = str(lead.get('상담 내용') or '').strip()
    if new_status == '부재중':
        marker = f'[{hdr_time} {initial} · 부재중] 부재중'
    else:
        marker = f'[{hdr_time} {initial} · 문의 드랍] {text}'
    new_k = f'{old_k} ─── {marker}' if old_k and old_k != '-' else marker

    try:
        _update = {'상태': new_status, '상담 내용': new_k}
        if manager_name:
            _update['온라인 상담자'] = manager_name
        update_lead(lead_no, _update)
    except Exception as exc:
        logger.error(f'[SLACK/자동감지] 시트 update 실패 ({lead_no}): {exc}', exc_info=True)
        return True

    try:
        if new_status == '부재중':
            _apply_auto_absent_badge(client, channel, thread_ts, root, lead_no, initial, hdr_time)
        else:
            _apply_auto_dropped_card(client, channel, thread_ts, root, lead_no, initial, hdr_time, text)
    except Exception as exc:
        logger.warning(f'[SLACK/자동감지] 카드 회색화 실패 ({lead_no}): {exc}')

    # 매니저 답글에 ✅ — "봇이 인지·처리함" 피드백 (채널톡 forward 답글 ✅와 동일 의미)
    try:
        client.reactions_add(
            channel=channel, timestamp=event_ts, name='white_check_mark',
        )
    except Exception:
        pass
    # 카드 본체에도 ✅ — "리드 처리 완료" 채널 스캔 신호 (모달·연결 경로와 통일)
    _react_card_handled(client, channel, thread_ts)

    logger.info(
        f'[SLACK/자동감지] {lead_no} → {new_status} '
        f'(매니저 {initial}, text={text[:20]!r})'
    )
    return True


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


def _grey_out_merged_chat_card(client, channel: str, message_ts: str,
                               chat_lead_no: str, target_lead_no: str,
                               initial: str = '-') -> bool:
    """기존 lead 통합 시 채팅 카드를 회색 '문의 드랍 (통합)' 처리.

    원본 문의 정보는 code block 으로 보존, `채팅 열기` 링크는 clickable 유지,
    [상담하기]·[기존 lead 연결] 버튼은 제거 (재처리 방지). 2026-07-27.
    """
    try:
        _rp = client.conversations_replies(
            channel=channel, ts=message_ts, limit=1, inclusive=True,
        )
        _root = ((_rp.get('messages') or [{}])[0]) if _rp else {}
        _blocks = _root.get('blocks') or []
        # 첫 section 원본 텍스트 추출
        orig_text = ''
        for b in _blocks:
            if b.get('type') == 'section':
                orig_text = (b.get('text') or {}).get('text', '')
                break
        if not orig_text:
            orig_text = _root.get('text', '') or ''

        # 이미 회색 처리됐으면 skip (재통합 재시도 방어)
        if '상담 완료 - 문의 드랍' in orig_text and '통합' in orig_text:
            return False

        # `채팅 열기` 링크 라인 분리 (clickable 유지), 나머지는 code block 으로
        chat_link_line = ''
        body_lines = []
        for ln in orig_text.split('\n'):
            if '채팅 열기' in ln and '<http' in ln:
                chat_link_line = ln.strip()
                continue
            body_lines.append(ln)
        # code block 용 정리 — blockquote/마크업/여백 제거
        cleaned = [l.lstrip('>').lstrip().replace('*', '') for l in body_lines]
        clean_text = re.sub(r'^[\s⠀]+|[\s⠀]+$', '', '\n'.join(cleaned))

        _now = datetime.now().strftime('%m.%d %H:%M')
        header = '\n'.join([
            '⠀',
            f':white_check_mark: *상담 완료 - 문의 드랍*  `{chat_lead_no}`',
            f'처리자 : {initial}',
            f'처리 시간 : {_now}',
            f'상담 내용 : → `{target_lead_no}` 로 통합',
            '--------------------------------------------',
        ])
        new_text = header + f'\n```\n{clean_text}\n```'
        if chat_link_line:
            new_text += f'\n{chat_link_line}'
        new_text += '\n⠀'

        client.chat_update(
            channel=channel, ts=message_ts,
            text=f'상담 완료 - 문의 드랍 {chat_lead_no} → {target_lead_no} 통합',
            blocks=[{'type': 'section', 'text': {'type': 'mrkdwn', 'text': new_text}}],
        )
        logger.info(f'[SLACK/link] 채팅 카드 회색 통합 처리: {chat_lead_no} → {target_lead_no}')
        return True
    except Exception as exc:
        logger.warning(f'[SLACK/link] 채팅 카드 회색 처리 실패 ({chat_lead_no}): {exc}')
        return False


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

        # === 채팅 카드 회색 '문의 드랍 (통합)' 처리 (2026-07-27) ===
        # 기존엔 카드 본문 변화 없이 thread 안내만 → 활성 카드로 오인. 회색 처리 +
        # 버튼 제거로 재처리 방지. 채팅 열기 링크는 유지.
        if channel and message_ts and chat_lead_no:
            try:
                _linker_ini = _slack_user_to_initial(client, slack_user_id) or '-'
            except Exception:
                _linker_ini = '-'
            _grey_out_merged_chat_card(
                client, channel, message_ts, chat_lead_no, target_lead_no,
                initial=_linker_ini,
            )

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

        # === 원본(target) lead 카드에 ✅ reaction (리드 처리 완료, 공통 헬퍼) ===
        try:
            card_info = rc.get(f'lead_card_msg:{target_lead_no}')
            if card_info:
                card_info_s = card_info.decode('utf-8') if isinstance(card_info, bytes) else card_info
                if '|' in card_info_s:
                    target_channel, target_ts = card_info_s.split('|', 1)
                    _react_card_handled(client, target_channel, target_ts)
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


def _consult_identity_changed(old_lead: dict, new_name: str, new_contact: str) -> dict:
    """재상담이 기존 리드의 신원(고객명/연락처)을 다른 값으로 바꾸는지 판정.

    2026-07-31 L-03367: 매니저가 다른 고객의 완료 카드에서 [재상담] 을 눌러 전혀
    다른 고객 정보를 입력 → 기존 리드 고객명·주소가 조용히 덮어써져 원 리드가
    소실됐다. 같은 고객 재상담이면 고객명·연락처가 (pre-fill 그대로라) 동일해 무해.
    둘 중 하나가 다른 값으로 바뀌면 '다른 카드에 잘못 누른' 신호로 본다.
    공백·하이픈만 다른 경우는 동일 취급, 옛 값이 비면 비교 불가 → 변경 아님.

    Returns: {'changed': bool, 'name': (old,new)|None, 'contact': (old,new)|None}
    """
    def _n(s):
        return re.sub(r'\s+', '', str(s or '')).strip()

    def _np(s):
        return re.sub(r'\D', '', str(s or ''))

    old_name, old_contact = _n(old_lead.get('고객명')), _np(old_lead.get('고객 연락처'))
    nn, nc = _n(new_name), _np(new_contact)
    name_diff = bool(nn and old_name and nn != old_name)
    contact_diff = bool(nc and old_contact and nc != old_contact)
    return {
        'changed': name_diff or contact_diff,
        'name': (str(old_lead.get('고객명') or ''), str(new_name or '')) if name_diff else None,
        'contact': (str(old_lead.get('고객 연락처') or ''), str(new_contact or '')) if contact_diff else None,
    }


def _build_consult_identity_confirm_view(metadata: dict, state: dict, idc: dict) -> dict:
    """재상담 신원 변경 확인 view — 상담 모달을 값 보존 재렌더 + 경고 배너 prepend.

    resubmit 시 metadata['_identity_confirmed']=True 라 게이트를 통과한다
    (callback_id 는 그대로 submit_consult → 별도 view 핸들러 불필요).
    """
    def _cur(bid):
        return (_v(state, bid) or '').strip()

    prefilled = {
        'visit_type': _cur("visit_type") or '온라인',
        'status': _cur("status") or '유선 상담',
        'visit_date': _cur("visit_date"),
        'visit_date_end': _cur("visit_date_end"),
        'name': _cur("name"),
        'contact': _cur("contact"),
        'email': _cur("email"),
        'visit_address': _cur("visit_address"),
        'consultation': _cur("consultation"),
    }
    new_meta = dict(metadata)
    new_meta['_identity_confirmed'] = True
    lead_no = metadata.get('lead_no', '')
    diff_lines = []
    if idc.get('name'):
        diff_lines.append(f"• 고객명: `{idc['name'][0] or '-'}` → `{idc['name'][1] or '-'}`")
    if idc.get('contact'):
        diff_lines.append(f"• 연락처: `{idc['contact'][0] or '-'}` → `{idc['contact'][1] or '-'}`")
    warn = (
        f":warning: *이 재상담이 기존 리드 `{lead_no}` 의 고객 정보를 덮어씁니다.*\n"
        + "\n".join(diff_lines) + "\n\n"
        ":point_right: *같은 고객* 이면 그대로 아래 [등록] 을 다시 눌러 진행하세요.\n"
        ":point_right: *다른 고객* 이면 이 창을 닫고 [전화 문의 등록하기] 로 새로 "
        "등록하세요. (그대로 진행하면 기존 리드 정보가 사라집니다)"
    )
    info_blocks = [
        {"type": "section", "text": {"type": "mrkdwn", "text": warn}},
        {"type": "divider"},
    ]
    return _build_consult_view(
        info_blocks, json.dumps(new_meta, ensure_ascii=False), prefilled)


def _backup_consult_overwrite(lead_no: str, old_lead: dict, idc: dict, user_id: str) -> None:
    """재상담이 신원을 덮어쓰기 직전 이전 행 스냅샷을 Redis 90일 보관 + WARNING.

    확인 게이트를 통과해 진행하더라도(오확인 포함) 원 리드를 복구할 수 있게 한다.
    key: consult_overwrite_backup:{lead_no}
    """
    _fields = ('리드 No', '고객명', '고객 연락처', '이메일', '방문 주소', '방문 예정일',
               '상태', '상담 내용', '플랫폼', '상담 시간', '온라인 상담자', '영업 담당자')
    snap = {k: str(old_lead.get(k, '') or '') for k in _fields}
    snap['_overwritten_by'] = user_id
    try:
        from dashboard.utils.redis_client import get_redis_client
        get_redis_client().redis.set(
            f'consult_overwrite_backup:{lead_no}',
            json.dumps(snap, ensure_ascii=False), ex=60 * 60 * 24 * 90,
        )
    except Exception as exc:
        logger.warning(f"[SLACK/상담] 백업 저장 실패 ({lead_no}): {exc}")
    logger.warning(
        f"[SLACK/상담] ⚠ 재상담 신원 덮어쓰기 ({lead_no}) by {user_id}: "
        f"name={idc.get('name')} contact={idc.get('contact')} "
        f"— 이전 스냅샷 백업(consult_overwrite_backup:{lead_no}, 90일)"
    )


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

    # 방문 모달 주소 정규화 + 미검증 배지 (2026-07-30 / 2026-08-01) — verified 만 정정,
    # 미verified(도로명·번지 오타)면 raw 유지 + '확인 필요' addr_note → 방문 카드 배지.
    _visit_addr_note = None
    if visit_address:
        visit_address, _visit_addr_note = _normalize_visit_address_if_verified(visit_address)

    # 본인 방문 필수 라디오 (2026-07-17) — JW 담당자 배정 참고용
    _assign_state = (state.get('assign_self', {}).get('value', {}) or {}).get('selected_option') or {}
    assign_self_yes = (_assign_state.get('value') == 'yes')

    is_visit = (status == '방문 예약')
    is_estimate = (status == '견적 제출')

    # 방문 예약 + 본인 방문 필수 → 상담 내용 앞에 태그 프리픽스
    # (시트 저장 → 방문 카드/캔버스/List 자동 반영, 매니저는 이니셜로 자신임을 확인)
    if is_visit and assign_self_yes:
        _register_initial = _slack_user_to_initial(client, user_id) or '-'
        _tag = f':man-raising-hand: 본인 방문 필수({_register_initial})'
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
                update_data['방문 주소'] = visit_address  # 상단에서 이미 정규화됨
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
            # 재상담 신원 덮어쓰기 백업 (2026-07-31 L-03367) — 게이트를 통과해
            # 진행하더라도(오확인 포함) 원 리드를 복구할 수 있게 이전 스냅샷 보관.
            try:
                _old_for_backup = _find_lead_by_no(lead_no) or {}
                _idc_bk = _consult_identity_changed(_old_for_backup, name, contact)
                if _idc_bk.get('changed'):
                    _backup_consult_overwrite(lead_no, _old_for_backup, _idc_bk, user_id)
            except Exception as _bkexc:
                logger.warning(f"[SLACK/상담] 덮어쓰기 백업 실패 ({lead_no}): {_bkexc}")
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
            platform=lead_platform, addr_note=_visit_addr_note,
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
            # 부재중 — 원본 카드 body 유지 (재시도 시 문의 내용·주소 참고 필수) +
            # 상단에 section 크기 배지. 재클릭 시 배지만 갈아끼우기 (시각·회차·사유 갱신).
            # (2026-07-23) context → section 승격, 사유 라인 포함.
            try:
                from dashboard.utils.redis_client import get_redis_client
                _rc = get_redis_client().redis
                _count_key = f'consult_missed_count:{lead_no}'
                _count = int(_rc.incr(_count_key) or 1)
                _rc.expire(_count_key, 60 * 60 * 24 * 90)
            except Exception:
                _count = 1

            # 부재중 사유 — 회차별(1차/2차) 표기 (2026-08-06 사용자 요청). K열 재상담
            # 이력에서 각 회차 content. 2회 이상이면 회차별, 1회면 단일 라인.
            _entries = _parse_consultation_entries(full_consultation) if full_consultation else []
            _badge_lines = [
                '⠀',  # 봇 헤더와 배지 사이 여백 (다른 완료 카드와 동일)
                # 부재중은 원본 body 그대로 노출 (lead_no 이미 원본에 있음) → 배지에 lead_no 중복 X
                f':arrows_counterclockwise: *부재중* (총 *{_count}회*)',
                f'처리자 : {_initial_for_card}',
                f'처리 시간 : {_now_for_card}',
            ]
            if _entries and len(_entries) >= 2:
                for _i, _e in enumerate(_entries):
                    _c = (_e.get('content') or '').strip() or '-'
                    _badge_lines.append(f'상담 내용 ({_i + 1}차) : {_c[:200]}')
            else:
                _reason = ''
                if _entries:
                    _reason = (_entries[-1].get('content') or '').strip()
                elif consultation:
                    _reason = consultation.strip()
                if _reason:
                    _badge_lines.append(f'상담 내용 : {_reason[:200]}')
            _badge_text = '\n'.join(_badge_lines)
            try:
                # 기존 카드 blocks fetch → 부재중 배지 전부 제거 후 최신 1개만 prepend.
                # (교체 판정을 '마지막 시도'로 하던 버그로 매번 삽입돼 배지가 쌓이던 것
                #  해소 + 이미 쌓인 카드 자가치유. L-03527 사고, 2026-08-06.)
                _rp = client.conversations_replies(channel=channel, ts=message_ts, limit=1, inclusive=True)
                _root = ((_rp.get('messages') or [{}])[0]) if _rp else {}
                _existing_blocks = [
                    b for b in (_root.get('blocks') or [])
                    if not _is_absent_badge_block(b)
                ]
                _badge_block = {
                    'type': 'section',
                    'text': {'type': 'mrkdwn', 'text': _badge_text},
                }
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
                _hdr_lno = f"  `{lead_no}`" if lead_no else ""
                header_lines = [
                    "⠀",
                    f":white_check_mark: *상담 완료 - {status}*{_hdr_lno}",
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

        # 2) 원본 카드 ✅ reaction — "리드 처리 완료" (공통 헬퍼)
        _react_card_handled(client, channel, message_ts)

        # 3) thread reply 발송 (slack UI가 reply count 표시 갱신하도록 마지막에)
        # 2026-07-23 정책 개편: 방문 예약만 방문_일정 카드 permalink 링크 reply.
        # 유선 상담/견적/드랍/부재중은 원본 카드 회색 처리로 상태 확인·검색 가능하므로
        # thread reply skip (같은 lead 가 검색 결과에 중복 노출되는 노이즈 제거).
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

            reply_sent = False
            if old_reply_ts:
                # 방문 카드 unfurl 재생성 위해 delete + repost.
                try:
                    client.chat_delete(channel=channel, ts=old_reply_ts)
                except Exception as exc:
                    logger.warning(
                        f"[SLACK/상담] 옛 reply 삭제 실패 — 새 reply 발송: {exc}"
                    )
            if not reply_sent:
                try:
                    resp = client.chat_postMessage(
                        channel=channel, thread_ts=message_ts, text=reply_text,
                    )
                    if resp and resp.get('ok') and resp.get('ts'):
                        try:
                            rc.set(reply_key, resp['ts'], ex=60 * 60 * 24 * 90)
                        except Exception:
                            pass
                except Exception as exc:
                    logger.error(f"[SLACK/상담] thread reply 실패: {exc}", exc_info=True)
        else:
            # 방문 예약이 아니면 thread reply skip. 옛 reply 가 남아있으면 정리.
            if old_reply_ts:
                try:
                    client.chat_delete(channel=channel, ts=old_reply_ts)
                    try:
                        rc.delete(reply_key)
                    except Exception:
                        pass
                except Exception as exc:
                    logger.debug(
                        f"[SLACK/상담] 옛 reply 정리 실패 ({lead_no}): {exc}"
                    )
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
    if _kind not in ('normalized', 'failed', 'note_only'):
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
    _moved = (addr_note.get('moved_notes') or '').strip()
    if _kind == 'normalized':
        _region_warn = bool(addr_note.get('region_warn'))
        _head = (
            f":rotating_light: 방금 등록한 `{lead_no}` 방문 주소가 "
            f"카카오 API 로 정정되며 *시/구가 바뀌었어요*. 오방문 위험이 있어 "
            f"확인 부탁드립니다.\n\n"
            if _region_warn else
            f":mag: 방금 등록한 `{lead_no}` 방문 주소가 "
            f"카카오 API 로 자동 정정됐어요.\n\n"
        )
        text = (
            f"{_head}"
            f"  원본: {_orig}\n"
            f"  정정: {_norm}\n\n"
            "정정된 주소가 맞는지 위 카드에서 확인 부탁드립니다.\n"
            "잘못 매핑됐다면 [✏️ 정보 수정] 으로 다시 입력해주세요."
        )
        # 주소에 특이사항이 함께 있어 상담으로 옮긴 경우 안내 + 습관 유도
        #   (2026-08-06: 정정+노트 동시 케이스도 note_only 와 동일 안내, 사용자 요청)
        if _moved:
            text += (
                f"\n\n:memo: 주소에 함께 있던 특이사항은 상담 내용으로 옮겼습니다.\n"
                f"  옮긴 내용: {_moved}\n"
                "*다음부터는 방문 주소 필드에는 주소만* 넣어주세요 — "
                "특이사항은 상담 내용 필드에 넣어주시면 정확합니다."
            )
    elif _kind == 'failed':
        text = (
            f":warning: 방금 등록한 `{lead_no}` 방문 주소를 "
            f"카카오 API 가 인식하지 못했습니다.\n"
            "위 카드에서 [✏️ 정보 수정] 으로 정확한 주소를 다시 입력해주세요.\n\n"
            f"  입력값: {_orig}"
        )
    else:  # note_only
        text = (
            f":memo: 방금 등록한 `{lead_no}` 방문 주소에 특이사항이 함께 있어 "
            f"자동으로 상담 내용으로 옮겼습니다.\n\n"
            f"  옮긴 내용: {_moved}\n\n"
            "*다음부터는 방문 주소 필드에는 주소만*, "
            "특이사항(방문 전 연락 요망 등) 은 상담 내용 필드에 넣어주세요.\n"
            "카카오 주소 정규화가 어긋날 수 있어 원문에 넣어주시는 게 정확합니다."
        )

    # 정규화·실패 알림 발송 시 특이사항 이동도 있으면 뒷부분에 안내 append
    if _kind in ('normalized', 'failed') and _moved:
        text += (
            f"\n\n:memo: 참고 — 방문 주소에 함께 있던 특이사항 "
            f"`{_moved}` 은 상담 내용으로 옮겨두었습니다.\n"
            "*다음부터는 특이사항은 상담 내용 필드에 넣어주세요.*"
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

    - 청크 안쪽 leading/trailing 공백은 마크 밖으로 (`*text *` 형태 회피 — Slack이 리터럴로 렌더).
    - 한글/영문/숫자가 marker 에 딱 붙으면 Slack이 리터럴로 렌더하므로 Word Joiner 삽입.
    - 순수 공백 청크는 raw 반환.
    """
    lead = ''
    while chunk and chunk[0] in ' \t':
        lead += chunk[0]
        chunk = chunk[1:]
    trail = ''
    while chunk and chunk[-1] in ' \t':
        trail = chunk[-1] + trail
        chunk = chunk[:-1]
    if not chunk:
        return lead + trail
    left = _WORD_JOINER if prev_ch and prev_ch.isalnum() else ''
    right = _WORD_JOINER if next_ch and next_ch.isalnum() else ''
    return f'{lead}{left}{marker}{chunk}{marker}{right}{trail}'


def _highlight_addr_diff(original: str, converted: str) -> tuple:
    """원본↔변환 diff 청크를 길이별 스타일로 감싼 (orig, conv) 튜플 반환.

    - 모든 diff 청크(1자 포함)                 : 볼드(*chunk*)

    2026-08-06 통일: 기존엔 1자 차이(벨↔밸)만 홑따옴표('chunk'), 2자↑만 볼드로
    분기했으나 표기 통일성 위해 1자도 볼드로 일원화 (사용자 요청). 청크가
    한글/영문/숫자 사이에 낀 경우 Word Joiner 로 mrkdwn word boundary 확보.

    주의: blockquote(>) 컨텍스트 전용. 회색 코드블록(```) 안에선 mrkdwn 리터럴이라 사용 X.
    """
    if not original or not converted or original == converted:
        return original, converted
    from difflib import SequenceMatcher
    sm = SequenceMatcher(None, original, converted, autojunk=False)
    opcodes = sm.get_opcodes()
    # 공백/문장부호만 있는 diff 청크는 equal 로 취급 (하이라이트 스킵).
    # insert('' vs ' ') / delete(' ' vs '') 도 순수 공백이면 skip (2026-07-23 ETC-678632)
    def _is_noise(chunk: str) -> bool:
        # 빈 문자열도 noise (insert/delete 짝의 반대편 처리)
        return not re.search(r'[가-힣A-Za-z0-9]', chunk)
    meaningful_chunks = [(t, i1, i2, j1, j2) for t, i1, i2, j1, j2 in opcodes
                        if t != 'equal' and not (_is_noise(original[i1:i2]) and _is_noise(converted[j1:j2]))]
    # 청크가 너무 많으면 (오탈자 여러 곳 + 공백 정정) 하이라이트 자체가 노이즈 →
    # 두 줄 대조만으로도 원본↔변환 쉽게 대조 가능하므로 하이라이트 스킵.
    if len(meaningful_chunks) > 4:
        return original, converted

    orig_parts = []
    conv_parts = []
    for tag, i1, i2, j1, j2 in opcodes:
        o = original[i1:i2]
        n = converted[j1:j2]
        if tag == 'equal' or (_is_noise(o) and _is_noise(n)):
            orig_parts.append(o)
            conv_parts.append(n)
            continue
        marker = '*'
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
        _register_line += ' (:man-raising-hand: *본인 방문 필수*)'
    # 방문 주소 렌더 — 온라인 lead 카드와 동일 스타일로 통일 (2026-07-22)
    # addr_note 로 정규화 결과 판별:
    #   normalized + 원본 != 변환 → 원본/변환 두 줄
    #   failed → 방문 주소 + [주소 확인 필요] 배지
    #   그 외 (배지 없음) → 방문 주소 한 줄
    from dashboard.services.lead_helpers import is_blank_address, ADDRESS_MISSING_LABEL
    _addr_lines = []
    if is_blank_address(visit_address):
        # 주소 미입력이면 배지/원본·변환 비교 없이 단일 안내 (2026-07-27)
        _addr_lines.append(f">방문 주소 : {ADDRESS_MISSING_LABEL}")
    elif addr_note and isinstance(addr_note, dict):
        _kind = addr_note.get('kind', '')
        _orig = (addr_note.get('original') or '').strip()
        if (_kind in ('normalized', 'note_only')
                and _orig and _orig != (visit_address or '').strip()):
            # note_only 포함 (2026-08-04): 정규화 없이 노트만 이동돼도 원본(노트 포함)/변환
            # 두 줄로 히스토리 보존 (원본 최대 유지 요청). 정규화 케이스와 동일 렌더.
            # 개행 → 공백 정규화 (2026-07-24 L-03361): 원본/변환 안 개행이 남으면
            # blockquote 다음 줄이 인용 밖으로 삐져나와 잘못 렌더됨.
            _orig_norm = re.sub(r'\s+', ' ', _orig).strip()
            _conv_norm = re.sub(r'\s+', ' ', (visit_address or '').strip())
            _orig_hl, _conv_hl = _highlight_addr_diff(_orig_norm, _conv_norm)
            _addr_lines.append(f">*원본 주소* : {_orig_hl}")
            _addr_lines.append(f">*변환 주소* : {_conv_hl}")
            # 시/구 변경 경고 (2026-08-06 L-03659): 여러 도시 공유 도로명 오매칭 방어
            if addr_note.get('region_warn'):
                _addr_lines.append(
                    ">:rotating_light: *[시/구가 바뀜 — 오방문 주의, 주소 확인 요망]*"
                )
        elif _kind == 'failed':
            _va_norm = re.sub(r'\s+', ' ', visit_address or '-').strip()
            _addr_lines.append(
                f">방문 주소 : {_va_norm}  :warning: *[주소 확인 필요]*"
            )
    if not _addr_lines:
        _va_norm = re.sub(r'\s+', ' ', visit_address or '-').strip()
        _addr_lines.append(f">방문 주소 : {_va_norm}")
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

    # 날짜 유실 방어 (2026-07-30): MODIFY/RESTORE 워크플로는 new_visit_date 로 List
    #   방문 예정일 셀을 세팅한다. 호출자가 빈 값으로 부르면 그 셀이 지워지는 사고
    #   (소급 스크립트가 new_visit_date 없이 MODIFY 호출 → 날짜 4건 유실). 안 넘겼으면
    #   시트의 현재 방문 예정일로 채워 유실 방지. CANCEL/COMPLETE 는 행 삭제라 무영향.
    if not new_visit_date:
        _cur_vd = str(lead.get('방문 예정일', '') or '').strip().lstrip("'")
        if _cur_vd:
            new_visit_date = _cur_vd

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
        # 2026-07-30 L-03421: K열 재상담 헤더(`[MM.DD HH:MM 이니셜 · 상태]`)가 List
        #   상담 컬럼에 그대로 노출되던 이슈. _post_to_slack_list(8507)는 이미 stripper
        #   경유하나 이 방문 수정/취소/완료 경로는 raw 전송 → 최신 회차 content 만 전달.
        'consultation': _extract_latest_consult_content(str(lead.get('상담 내용', '') or '').strip()),
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
            category_display = f"{category} ({platform})" if platform else category

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
    # 2026-07-24: 재상담 헤더 (`[MM.DD HH:MM 이니셜 · 상태]`) 포함된 값이 모달
    # initial_value 로 노출되면 매니저 UX 나쁨. 최신 회차 content 만 pre-fill.
    # 매니저가 편집 후 저장 시 헤더는 유실되지만 방문 카드 회색 처리에 이력 남음.
    if cur_consultation:
        cur_consultation = _extract_latest_consult_content(cur_consultation) or cur_consultation

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

    # List·Redis in-place 마이그레이션 (2026-08-06: webhook delete+add 불안정 대체)
    if not _migrate_visit_list_row(
        lead_no, new_lead_no, address=new_address,
        visit_date=new_visit_display, consultation=new_consultation,
        platform=new_platform,
    ):
        try:  # 옛 행 못 찾음 → 신규 add fallback
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
            logger.warning(f"[VISIT/EDIT/PROMOTE] List add fallback 실패: {exc}")
    _migrate_lead_redis_keys(lead_no, new_lead_no)

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

    # 4) List·Redis in-place 마이그레이션 (2026-08-06 L-03491→ETC 사고: webhook
    #    delete+add 불안정 + 방문유형·플랫폼·Redis 키 미이관). 직접 API 로 한 행 갱신.
    if not _migrate_visit_list_row(
        lead_no, new_etc_lead_no, address=new_address,
        visit_date=new_visit_display, consultation=new_consultation,
        platform='기타',
    ):
        try:  # 옛 행 못 찾음 → 신규 add fallback
            _post_to_slack_list(
                client, {
                    '리드 No': new_etc_lead_no, '고객명': new_name,
                    '고객 연락처': new_phone,
                    '이메일': str(old_lead.get('이메일', '') or '').strip(),
                    '상담 시간': datetime.now().strftime('%Y.%m.%d. %H:%M'),
                    '방문 주소': new_address, '문의 내용': new_consultation,
                    '키워드': str(old_lead.get('키워드', '') or '').strip(),
                    '플랫폼': '기타',
                },
                modal_fields={
                    'visit_date': new_visit_display, 'visit_address': new_address,
                    'consultation': new_consultation, 'estimate': '',
                },
                channel=channel, message_ts=message_ts, action='visit',
            )
        except Exception as exc:
            logger.warning(f"[VISIT/EDIT/DEMOTE] List add fallback 실패: {exc}")
    _migrate_lead_redis_keys(lead_no, new_etc_lead_no)

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

    # 2026-07-24: 상담 내용은 이전 회차 이력 유지 + 마지막 회차만 편집값으로 교체.
    # (매니저는 pre-fill 로 최신 회차 content 만 봤으므로 편집값도 최신 회차)
    _old_k = str(old_lead.get('상담 내용', '') or '').strip()
    new_consultation_merged = _replace_last_consult_content(_old_k, new_consultation)

    # 방문 주소 정규화 편입 (2026-08-13 L-03655): [정보 수정] 편집 경로가 정규화를
    #   통째 건너뛰어(raw 저장) 카카오 verify·시/도 축약·건물명 부착이 하나도 안 됐음
    #   ('경기도 남양주시 …' 그대로). L-03399(수동 편집은 원본/변환 배지 안 붙임)는
    #   '배지 미표시' 의도였는데 '정규화 미실행' 부작용이 됨. → 정규화는 실행하되,
    #   성공(시/구 동일)은 배지 없이 단일 라인 유지(L-03399 존중), 미검증(failed)·
    #   시/구 변경(region_warn)은 오방문 방어 배지 유지.
    _edit_addr_note = None
    if new_address:
        try:
            _norm_addr, _addr_note = _normalize_visit_address_if_verified(new_address)
            new_address = _norm_addr
            if _addr_note and (_addr_note.get('kind') == 'failed'
                               or _addr_note.get('region_warn')):
                _edit_addr_note = _addr_note
        except Exception as _exc:
            logger.warning(f"[SLACK/방문수정] 주소 정규화 실패 (raw 유지) ({lead_no}): {_exc}")

    updates = {
        '방문 예정일': sheet_visit_value,
        '고객명': new_name,
        '고객 연락처': new_phone,
        '방문 주소': new_address,
        '상담 내용': new_consultation_merged,
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
    # 2026-07-27: 수동 [정보 수정] 은 매니저가 직접 입력한 값이므로 '원본/변환' 두 줄
    #   (자동 변환 뉘앙스) 대신 최종 주소 한 줄만 표시. addr_note 안 만듦.
    #   (L-03399 계기 — H타워 수동 입력이 '변환 주소' 로 표기돼 자동 변환 오해)
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
            addr_note=_edit_addr_note,  # 성공=None(단일 라인), 미검증·시/구변경만 배지
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


def _mark_visit_complete_on_sheet(lead_no: str, initial: str, dt_str: str) -> None:
    """시트 K열 (상담 내용) 에 `[MM.DD HH:MM 이니셜 · 방문 완료]` 마커 append.

    2026-07-24 이중 필터 도입 — Redis flag (`visit_auto_completed:*`) 대량 손실 사고 대비.
    캔버스 sync 가 이 마커로도 완료 판정 → Redis 만 의존하지 않음.

    - ETC 리드는 시트에 없어서 skip
    - 이미 마커 있으면 중복 append 방지
    - dt_str 형식: 'MM.DD HH:MM' (호출자에서 이미 datetime.now().strftime)
    """
    if _is_etc_lead(lead_no):
        return
    try:
        lead = _find_lead_by_no(lead_no) or {}
        cur = str(lead.get('상담 내용', '') or '').strip()
        marker = f'[{dt_str} {initial} · 방문 완료]'
        if re.search(r'·\s*방문 완료\]', cur):
            return  # 이미 마커 있음
        # 재상담 형식과 통일 — 기존 값 뒤에 ─── 구분자 + 마커 append
        if cur and cur != '-':
            new_content = f'{cur} ─── {marker}'
        else:
            new_content = marker
        _update_lead_dispatch(lead_no, {'상담 내용': new_content})
        logger.info(f'[SLACK/방문완료] 시트 마커 append: {lead_no} → {marker}')
    except Exception as exc:
        logger.warning(f'[SLACK/방문완료] 시트 마커 append 실패 ({lead_no}): {exc}')


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

    # 3-2) 시트 K열 (상담 내용) 에 방문 완료 마커 append (2026-07-24 이중 필터)
    #      Redis flag 손실 사고 (07-24 대량 삭제) 대비. 시트가 source of truth.
    #      캔버스 sync 가 K열 마커도 확인 → flag/시트 둘 중 하나만 있어도 필터 성공.
    #      ETC 리드는 시트에 없어서 skip.
    _mark_visit_complete_on_sheet(lead_no, initial=initial, dt_str=complete_time)

    try:
        from dashboard.services.visit_canvas_sync import rebuild_canvas_async
        rebuild_canvas_async()
    except Exception as exc:
        logger.debug(f"[SLACK/방문완료] 캔버스 rebuild trigger 실패 ({lead_no}): {exc}")

    # 4) 폴더 사전 생성 (2026-07-24 투 트랙): 외부 LTE 환경에서 사진 대량 첨부 실패 시
    #    사무실 복귀 후 첨부해도 이 폴더 안에 저장되도록 완료 버튼 시점에 미리 생성.
    #    ETC 리드는 helper 안에서 skip. 이미 폴더 있으면 (사진 먼저 첨부된 경우 등) skip.
    try:
        _folder = _ensure_visit_folder(
            client, lead_no,
            channel=channel, thread_ts=message_ts,
            uploader_id=user_id,
        )
        if _folder and _folder.get('created_now'):
            # 신규 생성 시 스레드에 폴더 링크 안내 (매니저가 이름 정정·복귀 후 사진 첨부 가이드)
            _link = _folder.get('lead_folder_link', '')
            _name = _folder.get('folder_name', '')
            _reply_text = (
                f":file_folder: 방문 폴더 생성됨: <{_link}|{_name}>\n"
                f"_사무실 복귀 후 이 스레드에 사진 첨부 시 위 폴더 안 `현장사진/` 에 저장됩니다._\n"
                f"_(폴더명 정정이 필요하면 Drive 에서 이름 바꿔주세요. 폴더 ID 로 연결돼있어 무관합니다.)_"
            )
            try:
                client.chat_postMessage(
                    channel=channel, thread_ts=message_ts,
                    text=_reply_text, unfurl_links=False,
                )
            except Exception as exc:
                logger.warning(f"[SLACK/방문완료] 폴더 안내 reply 실패 ({lead_no}): {exc}")
    except Exception as exc:
        logger.warning(f"[SLACK/방문완료] 폴더 사전 생성 실패 ({lead_no}): {exc}")

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
            from dashboard.services.lead_helpers import is_blank_address, ADDRESS_MISSING_LABEL
            _addr = str(lead.get('방문 주소', '') or '').strip()
            _addr = ADDRESS_MISSING_LABEL if is_blank_address(_addr) else _addr
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
            f"취소자 : {initial}\n"
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
            f"취소자 : {initial}\n"
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


def _ensure_visit_folder(
    client, lead_no: str,
    channel: str = '', thread_ts: str = '',
    uploader_id: str = '', root_text: str = '',
    caption: str = '',
) -> Optional[dict]:
    """방문 lead 폴더 확보 (조회 or 생성) — 완료 버튼 & 사진 첨부 공용.

    우선순위 (2026-07-24 투 트랙 도입):
      1. Redis `visit_folder:{lead_no}` → folder_id 사용 (이름 재계산 X)
      2. 시트 P열 폴더 ID → folder_id 사용 (Redis 재저장)
      3. 둘 다 없음 → 이름 계산 (기존 규칙) + find_or_create_folder + Redis + P열 write

    folder_id 로 접근하므로 매니저가 Drive 에서 폴더명 수정해도 그대로 사용.
    folder_id 유효성 (삭제·이동) 은 get_folder_name 으로 검증 → 실패 시 신규 생성 fallback.

    ETC- 리드는 폴더 skip (기타 방문은 후속 관리 안 함).

    Returns:
        {
            'lead_folder_id': str,
            'lead_folder_link': str,     # webViewLink (created_now=True 일 때만 신뢰)
            'folder_name': str,          # Drive 실제 이름 or 재계산 이름
            'prefix': str, 'initial': str, 'address': str, 'date': str,
            'created_now': bool,         # 이번 호출로 만들어졌는지
            'source': str,               # 'redis' / 'sheet' / 'new'
        }
        or None (ETC 리드 or 설정 미비)
    """
    if _is_etc_lead(lead_no):
        return None

    parent_id = os.getenv('GOOGLE_DRIVE_VISIT_FOLDER_ID', '').strip()
    if not parent_id:
        logger.warning("[SLACK/방문 폴더] GOOGLE_DRIVE_VISIT_FOLDER_ID 미설정")
        return None

    from dashboard.utils.google_drive import find_or_create_folder, get_folder_name
    from dashboard.utils.redis_client import get_redis_client as _get_rc

    lead = _find_lead_by_no(lead_no) or {}

    # 1) Redis 매핑 → 유효성 검증 후 사용
    rc = None
    try:
        rc = _get_rc().redis
        cached_id = rc.get(f'visit_folder:{lead_no}')
        if isinstance(cached_id, bytes):
            cached_id = cached_id.decode()
        if cached_id:
            name = get_folder_name(cached_id)
            if name:
                return {
                    'lead_folder_id': cached_id,
                    'lead_folder_link': f'https://drive.google.com/drive/folders/{cached_id}',
                    'folder_name': name,
                    'prefix': '', 'initial': '', 'address': '', 'date': '',
                    'created_now': False,
                    'source': 'redis',
                }
            logger.info(
                f'[SLACK/방문 폴더] {lead_no} Redis folder_id 접근 실패 → 시트/신규 fallback'
            )
    except Exception as exc:
        logger.warning(f'[SLACK/방문 폴더] Redis 조회 실패 ({lead_no}): {exc}')

    # 2) 시트 P열 → 유효성 검증 후 사용 + Redis 재저장
    sheet_folder_id = str(lead.get('폴더 ID', '') or '').strip()
    if sheet_folder_id:
        name = get_folder_name(sheet_folder_id)
        if name:
            try:
                if rc is not None:
                    rc.set(
                        f'visit_folder:{lead_no}', sheet_folder_id,
                        ex=60 * 60 * 24 * 180,
                    )
            except Exception:
                pass
            return {
                'lead_folder_id': sheet_folder_id,
                'lead_folder_link': f'https://drive.google.com/drive/folders/{sheet_folder_id}',
                'folder_name': name,
                'prefix': '', 'initial': '', 'address': '', 'date': '',
                'created_now': False,
                'source': 'sheet',
            }
        logger.info(
            f'[SLACK/방문 폴더] {lead_no} 시트 P열 folder_id 접근 실패 → 신규 생성'
        )

    # 3) 신규 생성 — 이름 계산 (기존 로직 재사용)
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

    if not initial and uploader_id:
        try:
            uploader_ini = _slack_user_to_initial(client, uploader_id)
            if uploader_ini and uploader_ini != '-':
                initial = uploader_ini
        except Exception as exc:
            logger.debug(f'[SLACK/방문 폴더] 업로더 이니셜 조회 실패: {exc}')

    if not initial and not _is_partner:
        m_source = _clean(lead.get('온라인 상담자'))
        if m_source:
            initial = _to_initial(m_source)

    if not initial and root_text:
        m_ini = re.search(r'등록자\s*:\s*([A-Za-z가-힣]+)', root_text)
        if m_ini:
            initial = _to_initial(m_ini.group(1).strip())
    initial = initial or '미상'

    visit_address = str(lead.get('방문 주소', '') or '').strip()
    if not visit_address or visit_address == '-':
        visit_address = '주소 미상'

    today_str = datetime.now().strftime('%y.%m.%d')
    _DEFAULT_PLATFORMS = {'홈페이지', '전화', '카카오톡', '채널톡'}
    prefix = f"{lead_platform} " if (lead_platform and lead_platform not in _DEFAULT_PLATFORMS) else ''

    folder_name = f"{prefix}({initial}) {visit_address} {today_str}"
    folder_name = re.sub(r'[\\/:*?"<>|]', '', folder_name).strip()

    lead_folder = find_or_create_folder(folder_name, parent_id)
    if not lead_folder:
        logger.error(f"[SLACK/방문 폴더] lead 폴더 생성/조회 실패: {folder_name}")
        return None

    # Redis + P열 write
    try:
        if rc is not None:
            rc.set(f'visit_folder:{lead_no}', lead_folder['id'], ex=60 * 60 * 24 * 180)
    except Exception as exc:
        logger.warning(f'[SLACK/방문 폴더] Redis 저장 실패 ({lead_no}): {exc}')
    try:
        _update_lead_dispatch(lead_no, {'폴더 ID': lead_folder['id']})
        logger.info(
            f"[SLACK/방문 폴더] 신규 생성: {lead_no} → {folder_name} "
            f"({lead_folder['id']})"
        )
    except Exception as exc:
        logger.warning(f'[SLACK/방문 폴더] P열 저장 실패 ({lead_no}): {exc}')

    return {
        'lead_folder_id': lead_folder['id'],
        'lead_folder_link': lead_folder.get('webViewLink', ''),
        'folder_name': folder_name,
        'prefix': prefix, 'initial': initial,
        'address': visit_address, 'date': today_str,
        'created_now': True,
        'source': 'new',
    }


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

    # 2) lead 폴더 확보 — helper 로 위임 (Redis→P열→신규 순위, 이름 무관 folder_id 재사용)
    uploader_id = (event.get('user') or '').strip() if isinstance(event, dict) else ''
    caption = (event.get('text') or '').strip()
    location = ''
    if caption and '\n' not in caption and len(caption) <= 30:
        location = re.sub(r'[\\/:*?"<>|]', '', caption).strip()

    _folder_info = _ensure_visit_folder(
        client, lead_no,
        channel=channel, thread_ts=thread_ts,
        uploader_id=uploader_id, root_text=root_text,
        caption=caption,
    )
    if not _folder_info:
        logger.error(f"[SLACK/방문 사진] {lead_no} lead 폴더 확보 실패 — skip")
        return

    from dashboard.utils.google_drive import find_or_create_folder, upload_file, list_folder_filenames
    lead_folder = {
        'id': _folder_info['lead_folder_id'],
        'webViewLink': _folder_info['lead_folder_link'],
    }
    folder_name = _folder_info['folder_name']
    # helper 는 신규 생성 시에만 prefix/initial/address/date 채움. 사후 첨부는 빈값.
    prefix = _folder_info.get('prefix', '')
    initial = _folder_info.get('initial', '')
    visit_address = _folder_info.get('address', '')
    today_str = _folder_info.get('date', '')

    # 3) lead 폴더 안에 '현장사진' 서브폴더 (사진 첨부 시점에만 생성)
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
    _progress_owned = False  # 이번 배치가 새로 post 한 답글인지 (재사용이면 False)
    # 스레드당 답글 하나만 유지 — 이미 답글(이전 배치의 진행/최종)이 있으면 재사용해
    # in-place update. 배치마다 새 답글을 올렸다 지우면 팔로워에게 '누가 보냈는지 모를
    # 빈 스레드 알림'(삭제된 유령 메시지)이 쌓임. 재사용하면 지울 일이 없어 유령 0. (2026-08-12)
    _cumul_key_e = f'visit_photo_reply:{channel}:{thread_ts}'
    try:
        from dashboard.utils.redis_client import get_redis_client as _get_rc_e
        _ex_ts = _get_rc_e().redis.hget(_cumul_key_e, 'ts')
        _ex_ts = _ex_ts.decode() if isinstance(_ex_ts, bytes) else (_ex_ts or '')
    except Exception:
        _ex_ts = ''
    if _ex_ts:
        # 기존 답글 재사용 — 새 답글 post 안 함 (삭제 불필요 → 유령 알림 방지)
        _progress_ts = _ex_ts
        if _total >= 4:
            try:
                client.chat_update(
                    channel=channel, ts=_progress_ts,
                    text=f":hourglass_flowing_sand: 사진 저장 중... (0/{_total})")
            except Exception as exc:
                logger.warning(f"[SLACK/방문 사진] 진행 답글 재사용 실패: {exc}")
                _progress_ts = None
    if not _progress_ts and _total >= 4:
        try:
            _progress_resp = client.chat_postMessage(
                channel=channel, thread_ts=thread_ts,
                text=f":hourglass_flowing_sand: 사진 저장 중... (0/{_total})",
                unfurl_links=False,
            )
            _progress_ts = _progress_resp.get('ts')
            _progress_owned = bool(_progress_ts)
            # 새 답글 ts 즉시 기록 → 동시/후속 배치가 이 답글을 재사용(새 답글 X)
            if _progress_ts:
                try:
                    from dashboard.utils.redis_client import get_redis_client as _get_rc_s
                    _rc_s = _get_rc_s().redis
                    _rc_s.hset(_cumul_key_e, 'ts', _progress_ts)
                    _rc_s.expire(_cumul_key_e, 60 * 60 * 24 * 7)
                except Exception:
                    pass
        except Exception as exc:
            logger.warning(f"[SLACK/방문 사진] 진행 답글 post 실패: {exc}")

    # 배치 상태 Redis 기록 — 업로드 hang·재시작으로 루프가 끊겨도 남은 파일을 복구할 수
    # 있게 {folder_id, 파일목록(name·url·mime)} 저장. 정상 완료 시 삭제. (2026-08-12)
    # (기존 복구는 conversations.history 에서 '저장 중' 스레드 답글을 찾았으나 history 는
    #  스레드 답글을 안 돌려줘 한 번도 작동 못 함 → Redis 기반으로 대체.)
    _batch_key = f"photo_batch:{channel}:{event.get('ts') or thread_ts}"
    try:
        from dashboard.utils.redis_client import get_redis_client as _get_rc_b
        _get_rc_b().redis.set(_batch_key, json.dumps({
            'channel': channel, 'thread_ts': thread_ts, 'folder_id': folder_id,
            'reply_ts': _progress_ts or '',
            'files': [{'name': f.get('name') or f.get('title') or '',
                       'url': f.get('url_private_download') or f.get('url_private') or '',
                       'mime': f.get('mimetype') or 'image/jpeg'} for f in files],
            'created': time.time(),
        }), ex=2 * 86400)
    except Exception as _bexc:
        logger.warning(f"[SLACK/방문 사진] 배치 상태 기록 실패: {_bexc}")

    # 멱등 dedup — 이미 폴더에 있는 파일명은 재업로드 skip (중복 file_shared 이벤트·
    # 재처리·재시작 복구 시 중복 방지, 2026-07-30)
    _existing_names = list_folder_filenames(folder_id)
    uploaded = 0
    skipped_oversize = 0
    failed = 0
    dup_skipped = 0
    for _idx, f in enumerate(files, 1):
        download_url = f.get('url_private_download') or f.get('url_private')
        if not download_url:
            failed += 1
            continue
        filename = f.get('name') or f.get('title') or f'photo_{f.get("id","unknown")}.jpg'
        mimetype = f.get('mimetype') or 'application/octet-stream'

        # 이미 폴더에 있으면 재업로드 skip (멱등) — 저장된 것으로 카운트
        if filename in _existing_names:
            uploaded += 1
            dup_skipped += 1
            if _progress_ts and (_idx % 5 == 0 or _idx == _total):
                try:
                    client.chat_update(
                        channel=channel, ts=_progress_ts,
                        text=f":hourglass_flowing_sand: 사진 저장 중... ({_idx}/{_total})")
                except Exception:
                    pass
            continue

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

    # 루프 정상 완료(실패 0) → 복구 대상에서 제거. 실패/hang 이면 키 유지 → 복구가 재시도.
    if failed == 0:
        try:
            from dashboard.utils.redis_client import get_redis_client as _get_rc_d
            _get_rc_d().redis.delete(_batch_key)
        except Exception:
            pass

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
            # 재사용한 답글(이전 배치의 성공 메시지)은 에러로 덮어쓰지 않음 — 새 답글로 안내
            if _progress_ts and _progress_owned:
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
    # visit_folder:{lead_no} 역인덱스 + P열 write 는 _ensure_visit_folder 안에서 이미 완료.
    # 여기서는 thread 컨텍스트 (photo_folder_id 포함) 만 저장.
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
            # 진행 답글은 삭제하지 않음 — 스레드당 답글 하나를 재사용하므로 마지막
            # 배치가 이 답글을 최종 요약으로 update 함. (삭제 시 유령 알림 발생, 2026-08-12)
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


def _visit_not_yet_reached(visit_date) -> bool:
    """방문 예정일(범위면 시작일)이 오늘보다 미래면 True = '아직 안 다녀옴'.

    사진 첨부 자동완료 게이트용 (2026-08-05, L-03575 도면 오완료 사고). 방문 전
    첨부는 참고자료(도면 등)라 완료로 보면 안 됨. 파싱 실패·빈값이면 False —
    기존 동작(자동완료 허용) 유지해 현장 사진 UX 안 깨뜨림.
    """
    try:
        from dashboard.services.visit_assignment_sync import _parse_visit_date_start
        from datetime import date as _date
        start = _parse_visit_date_start(str(visit_date or '').strip())
        return bool(start and start > _date.today())
    except Exception:
        return False


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

    # 방문일 미래면 = 아직 안 다녀옴 = 첨부 사진은 참고자료(도면 등) → 자동완료 skip.
    #   사진은 이미 드라이브에 저장됨(완료만 막음). flag 도 set 안 함 → 실제 방문일에
    #   현장사진 첨부하면 그때 정상 자동완료. (L-03575 사고: 08-06 방문 건에 도면 참고
    #   첨부가 08-05 자동완료를 유발. 거래처 방문요청은 등록 시 도면 첨부가 흔함.)
    _lead_for_date = _find_lead_by_no(lead_no) or {}
    if _visit_not_yet_reached(_lead_for_date.get('방문 예정일')):
        logger.info(
            f"[SLACK/자동완료] {lead_no} 방문일({_lead_for_date.get('방문 예정일')!r}) "
            f"미래 — 참고사진으로 보고 자동완료 skip (사진은 저장됨)"
        )
        return

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

    # 3-2) 시트 K열 (상담 내용) 에 방문 완료 마커 append (2026-07-24 이중 필터)
    _mark_visit_complete_on_sheet(lead_no, initial=initial, dt_str=complete_time)


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
            category_display = f"{category} ({platform})" if platform else category

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



def _extract_latest_consult_content(text: str) -> str:
    """K열 (상담 내용) 값에서 최신 회차 content 만 추출.

    저장 형식: `[MM.DD HH:MM 이니셜 · 상태] content ─── [다음 헤더] content ...`
    → 마지막 회차 content 만 반환. 헤더·이전 회차 다 버림.
    파싱 실패 시 원문 그대로 반환.

    2026-07-24 L-03343 관측: 슬랙 List `상담 내용` 컬럼에 헤더 그대로 노출되던 이슈 대응.
    방문 카드·캔버스 렌더는 이미 최신 회차만 표시 (task #19), List sync 도 통일.
    """
    if not text:
        return ''
    text = str(text).strip()
    if not text:
        return ''
    try:
        entries = _parse_consultation_entries(text)
        # 최신 회차부터 역순으로 내용 있는 회차를 찾음. 방문 완료·부재중 등
        #   빈 content 마커 회차는 건너뜀 (2026-07-30 L-03371: 최신이 빈 방문완료라
        #   entries[-1].content 가 '' → 원문 전체 반환하며 헤더 노출되던 이슈).
        for e in reversed(entries):
            content = (e.get('content') or '').strip()
            if content:
                return content
    except Exception:
        pass
    return text


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
    # 2026-07-30: 문의 내용 비고 상담 내용(K열)에 재상담 헤더가 있으면 details 노출 +
    #   장소/기기 파싱이 헤더 prefix 때문에 깨짐 → 헤더 제거(최신 회차) 후 파싱.
    parts = _split_lead_content(
        _extract_latest_consult_content(str(lead.get('문의 내용', '') or lead.get('상담 내용', ''))))

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
        # 2026-07-24 L-03374 fix: 시트값 우선 (정규화 후 최종본). modal_fields 는
        #   매니저 raw 입력 (예: '성현로 135번안길') 이라 방문 카드·캔버스와 불일치 발생.
        #   시트는 address_resolver 로 정규화된 값 (`성현로135번안길`) → 이걸 truth 로.
        "visit_address": str(lead.get('방문 주소') or '').strip() or modal_fields.get('visit_address') or '-',
        # 2026-07-24: K열 헤더 (`[MM.DD HH:MM 이니셜 · 상태]`) 포함된 값이 넘어오는 케이스
        #   (재편집 promote·update 경로) → 최신 회차 content 만 파싱해 payload 전달.
        #   방문 카드·캔버스는 이미 최신 회차만 표시 (task #19). List sync 도 통일.
        # 2026-08-04: modal_fields 누락 시 시트 lead 로 backfill — 부분 modal_fields 로
        #   호출(소급 정정 등)하면 빈 값이 '-' 로 기존 List 컬럼(상담 내용·방문 요청일)을
        #   덮어써 지우는 사고 방지 (L-03530 계기, eb2b462 와 동일 사상). 정상 모달 flow 는
        #   modal_fields 에 값이 다 있어 no-op.
        "consultation": _extract_latest_consult_content(
            modal_fields.get('consultation') or str(lead.get('상담 내용') or ''),
        ) or '-',
        "details": parts.get('inquiry') or str(lead.get('문의 내용') or lead.get('상담 내용') or '').strip() or '-',
        "visit_date": modal_fields.get('visit_date') or visit_date_raw or '-',
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


# --- 리드 번호 강등/승격 시 List·Redis in-place 마이그레이션 (2026-08-06) ---
# webhook delete+add 방식은 불안정 (L-03491→ETC-429d99 사고: 옛 행 잔존 + 방문유형·
#   플랫폼 미변경 + Redis 키 미이관). 직접 API 로 한 행을 in-place 갱신.
_VLIST_COL = {
    'lead': 'Col087VA2RG3G', 'addr': 'Col0BDHF203UL', 'date': 'Col088QNV75NU',
    'consult': 'Col08C6LTR681', 'vtype': 'Col0BCYB7SPHV', 'platform': 'Col089R4CR59P',
}
_VLIST_VTYPE_OPT = {'기타': 'OptEAZ7ROVZ', '거래처': 'OptVWBRAB5M', '온라인': 'Opt8SJW8M1Y'}
_VLIST_PLATFORM_OPT = {
    '-': 'Opt4L327EU3', '전화': 'OptVGEOC5ZE', '당근': 'Opt7424AJFY',
    '홈페이지': 'OptEWNT3KYR', '카카오톡': 'OptRKJ0FD6O', '채널톡': 'OptA1YP2AZZ',
    '숨고': 'OptRF68T6X1', '큐플레이스': 'OptYHURABE1', '메일': 'Opt378Z95JK',
    '거래처': 'OptIDNGR1EJ', '소개': 'OptFUV84AJ9', '기타': 'OptKF07W16X',
}


def _vlist_rt(text: str) -> list:
    """List rich_text 셀 값."""
    return [{'type': 'rich_text', 'elements': [
        {'type': 'rich_text_section',
         'elements': [{'type': 'text', 'text': text}]}]}]


def _migrate_visit_list_row(old_no: str, new_no: str, address: str = '',
                            visit_date: str = '', consultation: str = '',
                            platform: str = '') -> bool:
    """강등/승격 시 List 행을 직접 API 로 in-place 마이그레이션.

    old_no(또는 이미 바뀐 new_no) 행을 찾아 lead_no + 방문유형(select) + 플랫폼(select)
    + 주소/날짜/상담 을 한 번에 갱신. 행 못 찾으면 False (호출부가 신규 add fallback).
    """
    import urllib.request as _u
    token = (os.getenv('SLACK_VISIT_BOT_TOKEN', '').strip()
             or os.getenv('SLACK_BOT_TOKEN', '').strip())
    lid = os.getenv('SLACK_VISIT_LIST_ID', '').strip()
    if not token or not lid:
        return False
    try:
        from slack_sdk import WebClient
        c = WebClient(token=token)
        resp = c.api_call('slackLists.items.list', http_verb='GET',
                          params={'list_id': lid})
        items = resp.data.get('items', []) if hasattr(resp, 'data') else resp.get('items', [])
        row_id = None
        for it in items:
            for f in it.get('fields', []):
                if f.get('column_id') == _VLIST_COL['lead'] and \
                        f.get('text') in (old_no, new_no):
                    row_id = it['id']
                    break
            if row_id:
                break
        if not row_id:
            return False
        vtype = ('거래처' if platform in ('거래처', '소개')
                 else '기타' if platform == '기타' else '온라인')
        cells = [
            {'row_id': row_id, 'column_id': _VLIST_COL['lead'], 'rich_text': _vlist_rt(new_no)},
            {'row_id': row_id, 'column_id': _VLIST_COL['vtype'], 'select': [_VLIST_VTYPE_OPT[vtype]]},
        ]
        if platform in _VLIST_PLATFORM_OPT:
            cells.append({'row_id': row_id, 'column_id': _VLIST_COL['platform'],
                          'select': [_VLIST_PLATFORM_OPT[platform]]})
        if address:
            cells.append({'row_id': row_id, 'column_id': _VLIST_COL['addr'], 'rich_text': _vlist_rt(address)})
        if visit_date:
            cells.append({'row_id': row_id, 'column_id': _VLIST_COL['date'], 'rich_text': _vlist_rt(visit_date)})
        if consultation:
            cells.append({'row_id': row_id, 'column_id': _VLIST_COL['consult'],
                          'rich_text': _vlist_rt(_extract_latest_consult_content(consultation))})
        body = {'list_id': lid, 'cells': cells}
        req = _u.Request('https://slack.com/api/slackLists.items.update',
                         data=json.dumps(body).encode('utf-8'),
                         headers={'Content-Type': 'application/json; charset=utf-8',
                                  'Authorization': f'Bearer {token}'}, method='POST')
        with _u.urlopen(req, timeout=8) as r:
            ok = json.loads(r.read()).get('ok')
        logger.info(f"[VISIT/MIGRATE] List row {old_no}→{new_no} 갱신 ok={ok}")
        return bool(ok)
    except Exception as exc:
        logger.warning(f"[VISIT/MIGRATE] List 마이그레이션 실패 ({old_no}→{new_no}): {exc}")
        return False


def _migrate_lead_redis_keys(old_no: str, new_no: str) -> None:
    """강등/승격 시 lead_no 로 키잉된 Redis 추적 키 이관 (카드ts·List dedup·완료·폴더 등)."""
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
    except Exception:
        return
    for prefix, ttl in [
        ('visit_notice_msg', 60 * 60 * 24 * 180),
        ('slack_list_posted', 60 * 60 * 24 * 90),
        ('visit_auto_completed', None), ('visit_folder', None),
        ('consult_reply', None), ('lead_card_msg', None),
    ]:
        try:
            v = rc.get(f'{prefix}:{old_no}')
            if v is None:
                continue
            v = v.decode() if isinstance(v, bytes) else v
            _ttl = ttl
            if _ttl is None:
                _t = rc.ttl(f'{prefix}:{old_no}')
                _ttl = _t if (_t and _t > 0) else None
            rc.set(f'{prefix}:{new_no}', v, ex=_ttl)
            rc.delete(f'{prefix}:{old_no}')
        except Exception:
            pass
    logger.info(f"[VISIT/MIGRATE] Redis 키 이관 {old_no}→{new_no}")


_REGION_SIDO_RE = re.compile(
    r'\s*(서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|'
    r'전북|전남|경북|경남|제주)'
)
_REGION_GU_RE = re.compile(r'([가-힣]{2,}(?:구|군))(?:\s|$)')


def _addr_region_sig(addr: str) -> tuple:
    """주소에서 (시/도, 첫 구/군) 시그니처 추출 — 지역 교차확인용."""
    a = addr or ''
    m1 = _REGION_SIDO_RE.match(a)
    m2 = _REGION_GU_RE.search(a)
    return (m1.group(1) if m1 else '', m2.group(1) if m2 else '')


def _region_changed(raw: str, norm: str) -> bool:
    """입력(raw)과 정규화(norm)의 시/도·구가 '명시적으로' 다른지 (오매칭 감지, L-03659).

    '경인로 789'처럼 여러 도시 공유 도로명에서 카카오가 매니저가 명시한 시/구를
    다른 시/구로 바꾸는 위험한 오매칭 → 시/구 바뀌면 True → 카드 ⚠️ 플래그.
    한쪽이 시/구 정보 없으면(비교 불가) False(오탐 방지).
    """
    rs, rg = _addr_region_sig(raw)
    ns, ng = _addr_region_sig(norm)
    if rg and ng and rg != ng:      # 구/군 양쪽 존재 + 다름 (부평구↔영등포구)
        return True
    if rs and ns and rs != ns:      # 시/도 양쪽 존재 + 다름 (인천↔서울)
        return True
    return False


def _normalize_visit_address_if_verified(raw_addr: str) -> tuple:
    """방문 모달 입력 주소를 카카오 verified 면 정규화값, 아니면 raw 유지 + 미검증 배지.

    Returns: (주소, addr_note)
      - verified & 정정됨: (정규화, {'kind':'normalized',...}) — 방문 카드 원본/변환 2줄 + ephemeral
      - verified & 원본동일: (raw, None)                       — 배지 없음
      - 미verified:        (raw, {'kind':'failed',...})       — '주소 확인 필요' 배지

    2026-07-30: 당근/온라인 intake·전화 sync(lead_sync)는 이미 resolve_address 로
    정규화하는데 방문 모달(상담하기·방문요청 submit)만 raw 저장돼 규격 불일치였음
    (특히 채널톡·카톡은 모달에서 주소 첫 입력). verified 확신 있을 때만 정정.
    2026-08-01: 미verified(도로명·번지가 카카오 미확인 = 오타 가능) 시 addr_note 반환
      → 방문 카드에 '확인 필요' 배지. 호수 오타는 건물서 소통되지만 **도로명·번지를
      고객이 부른 것과 다르게 받아적으면 엉뚱한 건물 방문 = 돌이킬 수 없음** → 최소
      방어. (당근/홈페이지 intake·전화 sync 경로와 동일 사상 — 매니저 모달만 무방비였음)
    """
    raw = (raw_addr or '').strip()
    if not raw:
        return raw, None
    try:
        from dashboard.services import address_resolver as _ar
        from dashboard.services import lead_helpers as _lh
        _rx = _lh.extract_korean_address(raw)
        _norm, _lv = _ar.resolve_address(
            raw, _rx[0] if _rx else None, _rx[1] if _rx else '',
        )
        if _lv == 'verified' and _norm:
            if _norm != raw:
                # verified & 정정됨 → 원본/변환 2줄 + 등록자 ephemeral (2026-08-01):
                #   워크플로(전화/거래처) 경로와 동일하게 매니저가 "내 입력 → 카카오
                #   정정" 을 카드에서 확인해 오정규화를 잡도록 통일. (모달만 무표시였음)
                logger.info(f"[SLACK/방문주소] 모달 주소 정규화: '{raw}' → '{_norm}'")
                _note = {'kind': 'normalized', 'original': raw, 'normalized': _norm}
                # 시/구 교차확인 (2026-08-06 L-03659): 입력에 명시한 시/구와 변환
                #   결과의 시/구가 다르면 = 여러 도시 공유 도로명 오매칭 위험 →
                #   ⚠️ 플래그(silent override 금지, 오방문 방어). 매니저가 카드에서 확인.
                if _region_changed(raw, _norm):
                    _note['region_warn'] = True
                    logger.warning(
                        f"[SLACK/방문주소] ⚠️ 시/구 변경 감지 — 오매칭 의심: "
                        f"'{raw}' → '{_norm}'"
                    )
                return _norm, _note
            # verified & 원본 동일 → 배지 없음 (조용히 통과)
            return _norm, None
        # 미verified — 도로명+번지가 카카오에 확인 안 됨 → raw 유지 + 확인 필요 배지
        logger.info(
            f"[SLACK/방문주소] 모달 주소 미검증(도로·번지 확인 실패) — raw + 배지: '{raw}'"
        )
        return raw, {'kind': 'failed', 'original': raw, 'normalized': ''}
    except Exception as exc:
        logger.warning(f"[SLACK/방문주소] 모달 정규화 실패 (raw 유지): {exc}")
    return raw, None


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
    # 방문 모달 주소 정규화 + 미검증 배지 (2026-07-30 / 2026-08-01) — verified 만 정정,
    #   미verified(도로·번지 오타)면 raw 유지 + 답글에 '확인 필요' 경고.
    _visit_addr_note = None
    if visit_address:
        visit_address, _visit_addr_note = _normalize_visit_address_if_verified(visit_address)
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
    _addr_warn = (
        ">:warning: *주소 확인 필요* — 도로명·번지가 확인되지 않았습니다. 재확인 요망\n"
        if _visit_addr_note else ""
    )
    reply_text = (
        f":white_check_mark: *방문 요청 등록* — `{lead_no}` by <@{user_id}>\n"
        f">*방문일* : {visit_date_raw or '-'}\n"
        f">*방문 주소* : {visit_address or '-'}\n"
        + _addr_warn
        + f">*내용 / 특이사항* : {consultation or '-'}"
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

    vat_option = {'text': {'type': 'plain_text', 'text': 'VAT 별도'},
                  'description': {'type': 'plain_text', 'text': '체크 해제 시 VAT 없음'},
                  'value': 'sep'}

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
            "type": "context",
            "elements": [{
                "type": "mrkdwn",
                "text": ("ℹ️ 금액 · VAT는 *이 칸에 입력*하면 경영지원 확인 후 반영(요청) — "
                         "사유란에만 적으면 반영 안 됩니다. 그 외 항목은 즉시 반영"),
            }],
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
    if content and _norm_edit_val(content) != _norm_edit_val(project.get('공사 내용')):
        updates['공사 내용'] = content
    new_contract = ', '.join(contract_types)
    if _norm_edit_val(new_contract) != _norm_edit_val(project.get('도급 구분')):
        updates['도급 구분'] = new_contract
    new_contractor = ', '.join(contractors)
    if _norm_edit_val(new_contractor) != _norm_edit_val(project.get('시공자')):
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

    # 금액·부가세는 경영지원 ✅ 반영 요청 / 나머지는 즉시 반영 (한 모달 분기 제출)
    amount_updates = {k: v for k, v in updates.items() if k in _AMOUNT_EDIT_FIELDS}
    direct_updates = {k: v for k, v in updates.items() if k not in _AMOUNT_EDIT_FIELDS}

    applied_fields: list = []
    direct_failed_reason = ''
    # 1) 즉시 반영 (금액·부가세 외)
    if direct_updates:
        result = perform_edit(code, direct_updates, reason, initial)
        if result.get('ok'):
            applied_fields = list(direct_updates.keys())
            # 원본 공사 확정 카드도 최신 내용으로 재렌더 (2026-08-06). 기존엔 시트·PM 만
            # 반영되고 카드는 옛 내용으로 stale. latest_data 에 수정값 merge 로 넘겨
            # write-behind(시트 반영 지연) 레이스 회피.
            try:
                from dashboard.services.project_slack_notifier import refresh_project_card_license
                refresh_project_card_license(code, latest_data={**project, **direct_updates})
            except Exception as exc:
                logger.warning(f'[SLACK/공사수정] 공사 확정 카드 재렌더 실패 ({code}): {exc}')
            try:
                _post_project_edit_notice_card(client, code, project, direct_updates, reason, initial)
            except Exception as exc:
                logger.warning(f'[SLACK/공사수정] 영업_관리 알림 실패 ({code}): {exc}')
        else:
            direct_failed_reason = result.get('reason', 'unknown')

    # 2) 금액·부가세 → 경영지원 반영 요청 카드 (계산서봇 발송, ✅ 시 반영)
    request_sent = False
    if amount_updates:
        try:
            req_ts = _post_amount_edit_request_card(project, amount_updates, reason, user_id, initial)
            request_sent = bool(req_ts)
        except Exception as exc:
            logger.error(f'[SLACK/공사금액] 요청 카드 발송 예외 ({code}): {exc}', exc_info=True)

    # 3) 요청자 안내
    _notify_project_edit_result(
        client, channel, user_id, code,
        applied_fields, direct_failed_reason, amount_updates, request_sent,
    )


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


def _notify_project_edit_result(
    client, channel: str, user_id: str, code: str,
    applied_fields: list, direct_failed_reason: str,
    amount_updates: dict, request_sent: bool,
) -> None:
    """공사 정보 수정 제출 결과를 요청자에게 ephemeral 로 안내."""
    _label = {'총액 1': '공사 금액', '부가세': '부가세'}
    parts: list = []
    if applied_fields:
        names = ', '.join(_label.get(f, f) for f in applied_fields)
        parts.append(f':white_check_mark: `{code}` 수정 완료 — {names}')
    if direct_failed_reason:
        parts.append(f':x: `{code}` 일부 항목 수정 실패: {direct_failed_reason}')
    if amount_updates:
        amt_names = ', '.join(_label.get(f, f) for f in amount_updates)
        if request_sent:
            parts.append(
                f':hourglass_flowing_sand: *{amt_names}* 은(는) 경영지원에 *수정 요청*으로 전달됐습니다. '
                f'경영지원이 직접 반영 후 DM 으로 알려드립니다.'
            )
        else:
            parts.append(
                f':warning: *{amt_names}* 수정 요청 전달에 실패했습니다. 경영지원에 직접 문의해주세요.'
            )
    if not parts:
        parts.append(':information_source: 변경된 필드가 없어 저장을 skip 했습니다.')
    try:
        client.chat_postEphemeral(channel=channel, user=user_id, text='\n'.join(parts))
    except Exception:
        pass


def _post_amount_edit_request_card(
    project: dict, amount_updates: dict, reason: str,
    requester_id: str, requester_initial: str,
):
    """공사 금액/부가세 수정 요청 카드 → #영업_관리 (공사봇 발송).

    금액은 경영지원(황샛별)이 PM 에서 **직접 반영** 후 ✅ 하면 카드 완료 + 요청자 DM.
    (시스템은 시트를 쓰지 않음.) 공사 건이라 '공사 현황 알림 봇' 명의. ✅ 이벤트는
    계산서봇이 받지만 카드 갱신은 공사봇 클라이언트로. Returns ts / None. Redis pending.
    """
    proj = _project_client()
    if proj is None:
        logger.warning('[SLACK/공사금액] 공사봇 미가용 — 금액 수정 요청 카드 발송 불가')
        return None
    channel_id = os.getenv('SLACK_INVOICE_CHANNEL_ID', '').strip()
    if not channel_id:
        logger.warning('[SLACK/공사금액] SLACK_INVOICE_CHANNEL_ID 미설정')
        return None

    code = (project.get('프로젝트 코드') or '').strip()
    biz = (project.get('사업자명') or '-').strip() or '-'
    addr = (project.get('현장 주소') or '-').strip() or '-'
    now_str = datetime.now().strftime('%m.%d %H:%M')
    vat_after = _vat_is_sep(amount_updates.get('부가세', project.get('부가세')))

    change_lines = [
        _fmt_edit_field_change(f, project.get(f, ''), amount_updates[f], vat_after)
        for f in _AMOUNT_EDIT_FIELDS if f in amount_updates
    ]
    lines = [
        f'🔔 *[공사 금액 수정 요청]*  `{code}`',
        '--------------------------------------------',
        f'🏢 사업자명 : {biz}',
        f'📍 현장 주소 : {addr}',
        f'📝 수정 사유 : {reason.strip()}',
        '📋 요청 내역',
        *change_lines,
        f'👤 요청자 : {requester_initial}  {now_str}',
        '--------------------------------------------',
    ]
    text = '⠀\n' + '\n'.join(lines)
    blocks = [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': text}},
        {'type': 'context', 'elements': [{'type': 'mrkdwn',
            'text': '경영지원이 *직접 반영*한 뒤 ✅ 하면 요청자에게 완료 DM 이 전송됩니다.'}]},
    ]
    try:
        proj.conversations_join(channel=channel_id)
    except Exception:
        pass
    resp = proj.chat_postMessage(channel=channel_id, text=text, blocks=blocks, unfurl_links=False)
    if not resp.get('ok'):
        logger.warning(f'[SLACK/공사금액] 요청 카드 발송 실패 ({code}): {resp}')
        return None
    ts = resp.get('ts')
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        payload = {
            'code': code,
            'updates': amount_updates,
            'reason': reason.strip(),
            'requester_id': requester_id,
            'requester_initial': requester_initial,
            'before': {f: project.get(f, '') for f in _AMOUNT_EDIT_FIELDS},
            'biz': biz, 'addr': addr, 'requested_at': now_str,
            'card_text': text,
        }
        rc.set(f'project_amount_req:{channel_id}:{ts}',
               json.dumps(payload, ensure_ascii=False, default=str),
               ex=60 * 60 * 24 * 60)
    except Exception as exc:
        logger.warning(f'[SLACK/공사금액] pending 저장 실패 ({code}): {exc}')
    logger.info(f'[SLACK/공사금액] 수정 요청 카드 발송: {code} ts={ts}')
    return ts


def _maybe_apply_amount_request(client, channel: str, ts: str, checker_user_id: str) -> bool:
    """✅(황샛별) on 금액 수정 요청 카드 → 카드 '반영 완료' 갱신 + 요청자 완료 DM.

    금액·부가세는 **경영지원(황샛별)이 PM 에서 직접 수정**하므로 시스템은 시트를 쓰지 않음.
    ✅ 는 '직접 반영했다'는 완료 신호 → 카드 갱신 + 요청자 통보만. 요청 카드 아니면 False.
    """
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
    except Exception:
        return False
    key = f'project_amount_req:{channel}:{ts}'
    raw = rc.get(key)
    if not raw:
        return False
    # 중복 처리 방지 (동시/재전송) — 먼저 잡은 이벤트만
    if not rc.set(f'{key}:proc', '1', nx=True, ex=120):
        return True
    try:
        data = json.loads(raw.decode() if isinstance(raw, bytes) else raw)
    except Exception:
        rc.delete(f'{key}:proc')
        return True

    code = data.get('code', '')
    requester_id = data.get('requester_id', '')
    requester_ini = data.get('requester_initial', '-')
    checker_ini = _slack_user_to_initial(client, checker_user_id) or 'SB'
    updates = data.get('updates', {}) or {}
    before = data.get('before', {}) or {}

    # ── PM 미반영 상태에서 ✅ 감지 (2026-08-07) ──
    # 금액은 시스템이 안 쓰고 경영지원이 PM 에서 직접 반영하는 구조라, PM 반영 없이
    # ✅ 만 누르면 요청자에게 '완료' DM 이 잘못 나가고 카드도 완료로 바뀜. 실제 시트에
    # 요청 필드가 하나도 안 바뀌었으면 완료 처리하지 않고 경고 + pending 유지(재클릭 허용).
    from dashboard.services.project_service import get_project_records
    from dashboard.services.project_slack_notifier import refresh_project_card_license

    def _find_proj(force):
        try:
            recs = get_project_records(force_refresh=force) or []
            return next((r for r in recs if (r.get('프로젝트 코드') or '').strip() == code), None)
        except Exception as exc:
            logger.warning(f'[SLACK/공사금액] 프로젝트 조회 실패 ({code}, force={force}): {exc}')
            return None

    # 시트(권위) 우선 → 미반영이면 캐시(PM 편집 시 즉시 갱신) 재확인 → write-behind 지연 오탐 방지
    proj_sheet = _find_proj(True)
    applied = _amount_request_applied(proj_sheet, before, updates)
    render_proj = proj_sheet
    if applied is False:
        proj_cache = _find_proj(False)
        if _amount_request_applied(proj_cache, before, updates) is True:
            applied = True
            render_proj = proj_cache

    if applied is False:
        # 시트·캐시 모두 반영 없음 → PM 미반영 의심. 완료 처리 skip, pending 유지.
        logger.warning(
            f'[SLACK/공사금액] ⚠️ PM 미반영 상태 ✅ 감지 ({code}) by {checker_ini} — 완료 처리 skip'
        )
        _warn_amount_not_applied(client, channel, ts, code, checker_user_id, before, updates, render_proj)
        rc.delete(f'{key}:proc')  # 락만 해제 → PM 반영 후 ✅ 재클릭(뗐다 다시) 시 재처리 가능
        return True

    # ── 정상 완료 처리 (반영 확인됨 또는 조회 불가로 검증 skip) ──
    rc.delete(key)  # pending 소비
    logger.info(f'[SLACK/공사금액] ✅ 완료 처리 {code} by {checker_ini} (요청 {requester_ini}) — 경영지원 직접 반영')
    _mark_amount_request_done(channel, ts, data, checker_ini)
    _dm_amount_request_done(requester_id, code, data, checker_ini)
    # 원본 공사 확정 카드도 최신 금액으로 재렌더 (2026-08-06). 이미 조회한 레코드 재사용.
    try:
        if render_proj:
            refresh_project_card_license(code, latest_data=render_proj)
    except Exception as exc:
        logger.warning(f'[SLACK/공사금액] 공사 확정 카드 재렌더 실패 ({code}): {exc}')
    return True


def _warn_amount_not_applied(
    client, channel: str, ts: str, code: str, checker_user_id: str,
    before: dict, updates: dict, proj,
) -> None:
    """✅ 눌렀지만 시트·캐시에 금액 반영이 없을 때 경고 (완료 처리 안 함).

    요청 카드 thread 에 경고(공사봇) + 누른 사람에게 ephemeral. 카드는 '요청'
    상태 유지 → PM 반영 후 ✅ 를 뗐다 다시 눌러야 완료됨. 2026-08-07.
    """
    vat_after = _vat_is_sep(updates.get('부가세', before.get('부가세')))
    req_lines = [
        _fmt_edit_field_change(f, before.get(f, ''), updates[f], vat_after)
        for f in _AMOUNT_EDIT_FIELDS if f in updates
    ]
    cur_amt = _amt_int(proj.get('총액 1')) if proj else None
    cur_line = f'\n• 현재 시트 금액 : {cur_amt:,}원' if cur_amt is not None else ''
    warn = (
        f'⚠️ *`{code}` 변경 요청한 금액이 아직 시트에 반영되지 않았습니다.*\n'
        f'PM 사이트나 구글 시트에 요청 금액을 반영한 뒤 ✅ 를 눌러주세요.\n'
        + ('\n'.join(req_lines) + '\n' if req_lines else '')
        + cur_line.lstrip('\n')
        + '\n\n_반영 후, 체크(✅)를 한 번 뗐다가 다시 눌러주세요._'
    )
    # 요청 카드 thread 경고 (카드를 올린 공사봇으로)
    try:
        proj_client = _project_client()
        if proj_client:
            proj_client.chat_postMessage(channel=channel, thread_ts=ts, text=warn)
    except Exception as exc:
        logger.warning(f'[SLACK/공사금액] 미반영 경고 thread 실패 ({code}): {exc}')
    # 누른 사람에게 ephemeral (즉시 인지)
    try:
        client.chat_postEphemeral(channel=channel, user=checker_user_id, text=warn)
    except Exception as exc:
        logger.warning(f'[SLACK/공사금액] 미반영 경고 ephemeral 실패 ({code}): {exc}')


def _mark_amount_request_done(channel: str, ts: str, data: dict, checker_ini: str) -> None:
    """반영 완료 후 요청 카드 갱신 (헤더 → 완료, 반영자·시각 추가).

    카드는 공사봇이 올렸으므로 공사봇 클라이언트로 chat_update (같은 봇만 수정 가능).
    """
    proj = _project_client()
    if proj is None:
        logger.warning('[SLACK/공사금액] 공사봇 미가용 — 카드 갱신 skip')
        return
    now_str = datetime.now().strftime('%m.%d %H:%M')
    orig = data.get('card_text', '') or ''
    new_text = orig.replace('🔔 *[공사 금액 수정 요청]*', '✅ *[공사 금액 수정 완료]*')
    new_text += f'\n☑️ 반영 : {checker_ini}  {now_str}'
    blocks = [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': new_text}},
        {'type': 'context', 'elements': [{'type': 'mrkdwn',
            'text': f'✅ {checker_ini} 확인·반영 완료 · 요청자에게 DM 전송됨'}]},
    ]
    try:
        proj.chat_update(channel=channel, ts=ts, text=new_text, blocks=blocks)
    except Exception as exc:
        logger.warning(f'[SLACK/공사금액] 카드 갱신 실패 (ts={ts}): {exc}')


def _dm_amount_request_done(requester_id: str, code: str, data: dict, checker_ini: str) -> None:
    """요청자에게 금액 반영 완료 DM (im:write 있는 메인봇으로 발송)."""
    if not requester_id:
        return
    updates = data.get('updates', {}) or {}
    before = data.get('before', {}) or {}
    vat_after = _vat_is_sep(updates.get('부가세', before.get('부가세')))
    lines = [f':white_check_mark: *공사 금액 수정 요청 반영 완료*  `{code}`']
    lines.append('_(요청 내역)_')
    for f in _AMOUNT_EDIT_FIELDS:
        if f in updates:
            lines.append(_fmt_edit_field_change(f, before.get(f, ''), updates[f], vat_after))
    lines.append(f'\n경영지원({checker_ini})이 직접 반영 처리했습니다.')
    text = '\n'.join(lines)
    # DM 은 im:write 있는 메인봇으로 (계산서봇은 DM 개설 불가)
    dm = _dm_client()
    if dm is None:
        logger.warning('[SLACK/공사금액] 메인봇 토큰 미설정 — 완료 DM skip')
        return
    try:
        im = dm.conversations_open(users=requester_id)
        dm_ch = ((im.get('channel') or {}) or {}).get('id')
        if dm_ch:
            dm.chat_postMessage(channel=dm_ch, text=text)
    except Exception as exc:
        logger.warning(f'[SLACK/공사금액] 완료 DM 실패 (requester={requester_id}): {exc}')


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
            f"취소자 : {initial}\n"
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


def _partner_status_warn(biz: str) -> str:
    """상호명으로 거래처 탭 폐업/휴업 조회 → 경고 문구 (없으면 ''). 카드·모달 공용."""
    if not biz or biz == '-':
        return ''
    try:
        from dashboard.services.partner_status_sync import lookup_partner_status_by_name
        st = lookup_partner_status_by_name(biz)
    except Exception:
        return ''
    if not st:
        return ''
    f0 = st['flagged'][0]
    stt = f0.get('status') or f"{st['label']}자"
    bno = f0.get('bno') or '-'
    if st.get('ambiguous'):
        return (f"⚠️ *국세청 조회 — 동일 상호 중 {st['label']} 이력 있음* "
                f"(`{bno}` {stt}). 발행 전 사업자번호 확인 필요.")
    return (f"⚠️ *국세청 조회 — {st['label']} 거래처* (`{bno}` {stt}). "
            f"발행 전 사업자번호·최신 등록증 확인 필요.")


def _build_invoice_modal_view(code, biz, addr, amt, email, metadata, partner_warn='') -> dict:
    """세금계산서 요청 모달 view dict. open / 백그라운드 update 공용 (2026-07-28).

    partner_warn: 폐업/휴업 경고 문구. 있으면 헤더 바로 아래 section 으로 표시.
    """
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
            "label": {"type": "plain_text", "text": label}, "element": el,
        }
        if optional:
            blk["optional"] = True
        return blk

    blocks = [
        {"type": "section", "text": {"type": "mrkdwn",
            "text": f"프로젝트 `{code}` 세금계산서 발행 요청"}},
    ]
    if partner_warn:
        # 헤더 바로 아래 폐업/휴업 경고 (요청 전 인지)
        blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": partner_warn}})
    blocks += [
        _text_input("biz", "사업자명", biz),
        _text_input("addr", "현장 주소", addr),
        {"type": "section", "text": {"type": "mrkdwn",
            "text": f"*공사 금액 (시트 원본)*\n{amt} 원"}},
        _text_input("amt", "계산서 발행 금액", amt),
        {
            "type": "input", "block_id": "vat",
            "label": {"type": "plain_text", "text": "VAT (부가가치세)"},
            "element": {
                "type": "radio_buttons", "action_id": "value",
                "initial_option": {"text": {"type": "plain_text", "text": "VAT 별도"}, "value": "sep"},
                "options": [
                    {"text": {"type": "plain_text", "text": "VAT 별도"}, "value": "sep"},
                    {"text": {"type": "plain_text", "text": "VAT 포함"}, "value": "incl"},
                ],
            },
        },
        _text_input("email", "발행 이메일", email),
        _text_input(
            "memo", "추가 요청사항", "",
            multiline=True, optional=True,
            placeholder='예) 청구 or 영수 발행\n예) 항목이나 비고란에 특정 내용 기재',
        ),
    ]
    return {
        "type": "modal", "callback_id": "submit_invoice",
        "private_metadata": metadata,
        "title": {"type": "plain_text", "text": "세금계산서 발행 요청"},
        "submit": {"type": "plain_text", "text": "요청 발송"},
        "close": {"type": "plain_text", "text": "취소"},
        "blocks": blocks,
    }


def _refresh_invoice_modal_from_sheet(client, view_id, view_hash, code,
                                       opened_biz, addr, amt, opened_email) -> None:
    """계산서 모달 open 직후, 시트 최신값(사업자명·주소·이메일)으로 갱신.

    [💰 계산서 요청] 버튼 payload 는 공사확정 발송 시점 스냅샷이라, 이후 웹/시트에서
    채워진 사업자명 등을 못 따라감(예: G3901-SJ 발송 시 사업자명 빈값 → 이후 '(주)미덕원'
    채워졌으나 버튼은 옛 값). open 은 payload 로 즉시(trigger 안전), 직후 views.update 로
    최신값 반영 (2026-07-28).
    """
    import re
    from dashboard.services.as_service import get_project_details

    def _n(s):
        return re.sub(r'\s+', '', str(s or '')).strip()

    d = get_project_details(code) or {}
    fresh_biz = (d.get('biz') or '').strip()
    fresh_addr = (d.get('address') or '').strip()
    new_biz, new_addr, new_email = opened_biz, addr, opened_email
    changed = False
    if fresh_biz and fresh_biz != '-' and _n(fresh_biz) != _n(opened_biz):
        new_biz = fresh_biz
        changed = True
        # 사업자명 최신화 → 거래처 탭 이메일(계산서용) 재계산 우선 반영
        try:
            from dashboard.services.partner_status_sync import get_cached_partner_email
            _ce = get_cached_partner_email(fresh_biz)
            if _ce:
                new_email = _ce
        except Exception:
            pass
    if fresh_addr and fresh_addr != '-' and _n(fresh_addr) != _n(addr or ''):
        new_addr = fresh_addr
        changed = True

    # 금액 최신화 (2026-08-03): [계산서 요청] payload 는 공사확정 발송 시점 스냅샷이라,
    # 이후 공사 금액을 수정(웹/PM/경영지원 반영)하면 모달에 옛 금액이 뜸. 시트 최신
    # 총액(amount_raw)으로 갱신. 모달 입력칸 형식(콤마 숫자, 원·VAT 텍스트 없음)에 맞춤.
    new_amt = amt
    _fresh_amt = (d.get('amount_raw') or '').strip()
    if _fresh_amt.isdigit() and re.sub(r'[^\d]', '', str(amt or '')) != _fresh_amt:
        new_amt = f'{int(_fresh_amt):,}'
        changed = True

    # 폐업/휴업 경고 — 최신 사업자명 기준으로 거래처 탭 상태 조회해 모달 헤더 하단 표시
    partner_warn = _partner_status_warn(new_biz or opened_biz)

    # 사업자명·주소가 바뀌었거나 폐업 경고가 있으면 모달 갱신
    if not (changed or partner_warn):
        return
    metadata = json.dumps({"code": code}, ensure_ascii=False)
    view = _build_invoice_modal_view(code, new_biz, new_addr, new_amt, new_email, metadata,
                                     partner_warn=partner_warn)
    try:
        client.views_update(view_id=view_id, hash=view_hash, view=view)
        logger.info(
            f'[SLACK/계산서] 모달 최신값 반영 ({code}): 사업자명={new_biz!r}'
            + (' | ⚠️폐업/휴업 경고' if partner_warn else '')
        )
    except Exception as exc:
        # hash 불일치(매니저가 이미 입력 중) 등은 조용히 skip — 입력값 보존
        logger.debug(f'[SLACK/계산서] 모달 views_update skip ({code}): {exc}')


def _open_invoice_modal(client, body) -> None:
    """[💰 계산서 요청] 클릭 → 프로젝트 정보 pre-fill 모달 오픈.

    2026-07-24 (Redis 손실 사고 후속): trigger_id 3초 제약 대응.
      이전에는 모달 오픈 전 사업자등록증 검증 (Drive API) + 시트 최신값
      재조회 (get_project_records, 3910행 로드) 를 수행했으나, Redis 캐시
      손실 후 시트 재로드가 5~10초 걸려 trigger_id 만료 (expired_trigger_id)
      로 모달 자체가 안 뜨는 사고. → payload snapshot 만으로 즉시 modal 오픈.
      사업자등록증 검증·이메일 필수 판정은 submit 시 (_check_license) 로 위임.
      submit 시 반려 → 매니저 재입력 UX.
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
    # 발행 이메일 자동채움 (2026-07-28, 우선순위: 거래처 탭 > 발주처).
    # 세금계산서 모달은 '계산서용' 이메일이 필요 — 거래처 탭 이메일(홈택스 발행
    # 이력 기반)이 계산서용이라 **최우선**. 발주처 이메일(온라인 리드=견적용)은
    # fallback (법인은 견적≠계산서 이메일이 흔함, 개인은 대개 동일). 매니저 수정 가능.
    # trigger_id 3초 안전 (Redis HGET O(1)).
    email = ''
    if biz and biz != '-':
        try:
            from dashboard.services.partner_status_sync import get_cached_partner_email
            email = get_cached_partner_email(biz) or ''
        except Exception as _eexc:
            logger.warning(f'[SLACK/계산서] 거래처 이메일 pre-fill 실패 (무시): {_eexc}')
    if not email:
        email = payload.get('email', '') or ''  # 발주처(견적) 이메일 fallback

    metadata = json.dumps({"code": code}, ensure_ascii=False)

    # payload 스냅샷으로 즉시 오픈 (trigger_id 3초 안전).
    view = _build_invoice_modal_view(code, biz, addr, amt, email, metadata)
    try:
        resp = client.views_open(trigger_id=trigger_id, view=view)
    except Exception as exc:
        logger.warning(f'[SLACK/계산서] 모달 오픈 실패: {exc}')
        return

    # 버튼 payload 는 공사확정 발송 시점 값 → 이후 채워진 사업자명·주소·이메일을
    # 백그라운드로 재조회해 views.update 로 반영 (2026-07-28, G3901-SJ 계기).
    _view = (resp or {}).get('view') or {}
    view_id, view_hash = _view.get('id', ''), _view.get('hash', '')
    if view_id:
        def _bg_refresh():
            try:
                _refresh_invoice_modal_from_sheet(
                    client, view_id, view_hash, code, biz, addr, amt, email,
                )
            except Exception as exc:
                logger.warning(f'[SLACK/계산서] 모달 최신값 반영 실패 (무시): {exc}')
        threading.Thread(target=_bg_refresh, daemon=True).start()


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
    # 2026-07-24: 세금계산서 요청은 #계산서_관리 채널로 분리. 공사수정·취소는 #영업_관리 유지.
    channel_id = os.getenv('SLACK_INVOICE_REQUEST_CHANNEL_ID', '').strip() \
        or os.getenv('SLACK_INVOICE_CHANNEL_ID', '').strip()

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
    """모달 제출 → #계산서_관리 채널에 계산서 요청 카드 발송.

    (2026-07-24) 세금계산서 요청 카드는 #계산서_관리 로 분리 발송.
    공사수정·취소 알림 (SLACK_INVOICE_CHANNEL_ID) 과는 채널 분리.
    """
    channel_id = os.getenv('SLACK_INVOICE_REQUEST_CHANNEL_ID', '').strip() \
        or os.getenv('SLACK_INVOICE_CHANNEL_ID', '').strip()
    if not channel_id:
        logger.warning('[SLACK/계산서] SLACK_INVOICE_REQUEST_CHANNEL_ID 미설정 — 발송 skip')
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

    # 국세청 조회 기반 폐업/휴업 경고 (거래처 탭 상호 매칭, 2026-07-28).
    # 폐업 번호로 세금계산서 발행하는 사고 방지 — 카드에 경고 라인 삽입 (모달과 공용 helper).
    partner_warn = _partner_status_warn(biz)

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
    ]
    if partner_warn:
        lines.append(partner_warn)
    lines += [
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
    # (2026-07-24 retry 3회 + 실패 시 안내 reply — SSL 일시 실패로 누락되지 않도록.)
    def _attach_license_to_thread():
        from dashboard.services.business_license_handler import fetch_license_canonical
        backoffs = [0, 2, 5]  # 즉시 → 2초 → 5초 (총 3회)
        last_exc = None
        for attempt, delay in enumerate(backoffs, start=1):
            if delay:
                time.sleep(delay)
            try:
                lic = fetch_license_canonical(code)
                if not lic:
                    logger.info(f"[SLACK/계산서] 사업자등록증 canonical 없음 → 스레드 첨부 skip ({code})")
                    return  # canonical 자체가 없는 건 재시도 무의미
                invoice_client.files_upload_v2(
                    channel=channel_id,
                    thread_ts=ts,
                    file=lic['content'],
                    filename=lic['file_name'],
                    initial_comment=f":page_facing_up: 사업자등록증 — `{code}`",
                )
                logger.info(f"[SLACK/계산서] 사업자등록증 스레드 첨부 완료 (attempt={attempt}): {code} ({lic['file_name']})")
                return
            except Exception as exc:
                last_exc = exc
                logger.warning(f"[SLACK/계산서] 사업자등록증 첨부 시도 {attempt}/3 실패 ({code}): {exc}")

        # 3회 모두 실패 — 스레드에 매니저 안내 reply (수동 첨부 유도)
        logger.error(f"[SLACK/계산서] 사업자등록증 스레드 첨부 최종 실패 ({code}): {last_exc}")
        try:
            invoice_client.chat_postMessage(
                channel=channel_id,
                thread_ts=ts,
                text=(
                    f":warning: 사업자등록증 자동 첨부 실패 (`{code}`) — "
                    f"이 스레드에 파일을 수동으로 첨부해주세요."
                ),
                unfurl_links=False, unfurl_media=False,
            )
        except Exception as notify_exc:
            logger.warning(f"[SLACK/계산서] 실패 안내 reply 발송 실패 ({code}): {notify_exc}")

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


def _recover_stuck_visit_photo_uploads():
    """서버 시작 시, 재시작 등으로 중단된 '사진 저장 중' 배치를 자동 복구 (2026-07-30).

    #방문_일정 최근 히스토리에서 '사진 저장 중… (N/M)' 로 멈춘 봇 메시지를 찾아,
    같은 스레드의 배치 파일 메시지를 멱등(dedup) 재업로드 → 누락분만 올리고 메시지 완료 갱신.
    배치 처리 스레드가 restart 로 죽으면 in-flight 파일이 유실되던 취약점 대응.
    """
    import time as _t
    import urllib.request as _ur
    channel = os.getenv('SLACK_VISIT_CHANNEL', '').strip()
    token = (os.getenv('SLACK_VISIT_BOT_TOKEN', '').strip()
             or os.getenv('SLACK_BOT_TOKEN', '').strip())
    if not channel or not token:
        return
    try:
        from slack_sdk import WebClient
        from dashboard.utils.google_drive import (
            list_folder_filenames, upload_file, find_or_create_folder)
        from dashboard.utils.redis_client import get_redis_client
        client = WebClient(token=token)
        rc = get_redis_client().redis
    except Exception as exc:
        logger.warning(f'[SLACK/방문사진복구] 초기화 실패: {exc}')
        return

    now = _t.time()
    cutoff = now - 2 * 86400          # 최근 2일만
    min_age = now - 180              # 3분 이내(진행 중일 수 있음)는 제외
    stuck = []
    cur = None
    try:
        for _ in range(6):
            resp = client.conversations_history(
                channel=channel, limit=200, **({'cursor': cur} if cur else {}))
            msgs = resp.get('messages', [])
            if not msgs:
                break
            stop = False
            for m in msgs:
                ts = float(m.get('ts', 0) or 0)
                if ts < cutoff:
                    stop = True
                    break
                if ('사진 저장 중' in (m.get('text', '') or '')
                        and m.get('thread_ts') and ts < min_age):
                    stuck.append(m)
            cur = resp.get('response_metadata', {}).get('next_cursor')
            if stop or not cur:
                break
    except Exception as exc:
        logger.warning(f'[SLACK/방문사진복구] history 조회 실패: {exc}')
        return

    if not stuck:
        logger.info('[SLACK/방문사진복구] 중단된 배치 없음')
        return
    logger.info(f'[SLACK/방문사진복구] 중단 의심 {len(stuck)}건 점검')

    for sm in stuck:
        thread_ts, stuck_ts = sm.get('thread_ts'), sm.get('ts')
        try:
            reps = client.conversations_replies(
                channel=channel, ts=thread_ts, limit=100).get('messages', [])
        except Exception:
            continue
        # 배치 = stuck_ts 직전, files 있는 메시지 중 가장 최근
        batch = None
        for rm in reps:
            if rm.get('files') and float(rm['ts']) < float(stuck_ts):
                if batch is None or float(rm['ts']) > float(batch['ts']):
                    batch = rm
        if not batch:
            continue
        files = batch.get('files') or []
        _raw = rc.hgetall(f'visit_thread:{channel}:{thread_ts}') or {}
        info = {(k.decode() if isinstance(k, bytes) else k):
                (v.decode() if isinstance(v, bytes) else v) for k, v in _raw.items()}
        photo_fid = info.get('photo_folder_id', '')
        if not photo_fid:
            logger.warning(f'[SLACK/방문사진복구] {thread_ts} 폴더 매핑 없음 — skip(수동)')
            continue
        target = photo_fid
        caption = (batch.get('text') or '').strip()
        if caption and '\n' not in caption and len(caption) <= 30:
            loc = re.sub(r'[\\/:*?"<>|]', '', caption).strip()
            if loc:
                lf = find_or_create_folder(loc, photo_fid)
                if lf:
                    target = lf['id']
        existing = list_folder_filenames(target)
        recovered = []
        for f in files:
            name = f.get('name') or f.get('title')
            if not name or name in existing:
                continue
            url = f.get('url_private_download') or f.get('url_private')
            if not url:
                continue
            try:
                req = _ur.Request(url, headers={'Authorization': f'Bearer {token}'})
                content = _ur.urlopen(req, timeout=30).read()
                if upload_file(target, name, content, mimetype=f.get('mimetype', 'image/jpeg')):
                    recovered.append(name)
            except Exception as exc:
                logger.warning(f'[SLACK/방문사진복구] 업로드 실패 ({name}): {exc}')
        try:
            _sfx = (f' (서버 재시작으로 중단됐던 {len(recovered)}장 복구)'
                    if recovered else '')
            client.chat_update(
                channel=channel, ts=stuck_ts,
                text=f':white_check_mark: 사진 {len(files)}장을 드라이브에 저장했습니다.{_sfx}')
        except Exception:
            pass
        if recovered:
            logger.info(f'[SLACK/방문사진복구] {thread_ts} {len(recovered)}장 복구: {recovered}')


def _recover_photo_batches():
    """Redis 배치 상태(photo_batch:*) 기반으로 중단된 방문 사진 업로드 복구 (2026-08-12).

    _process_visit_thread_files 가 루프 직전 photo_batch:{ch}:{ts} 에
    {folder_id, files[{name,url,mime}], reply_ts, created} 기록 → 정상 완료 시 삭제.
    hang·재시작으로 루프가 끊기면 키가 남음 → 폴더에 없는 파일만 멱등 재업로드.
    (구 _recover_stuck_visit_photo_uploads 는 conversations.history 로 스레드 답글을
     찾아 한 번도 작동 못 했음 — 이 함수가 대체.)
    """
    import urllib.request as _ur
    token = (os.getenv('SLACK_VISIT_BOT_TOKEN', '').strip()
             or os.getenv('SLACK_BOT_TOKEN', '').strip())
    try:
        from dashboard.utils.google_drive import list_folder_filenames, upload_file
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
    except Exception as exc:
        logger.warning(f'[SLACK/방문사진복구] 초기화 실패: {exc}')
        return
    try:
        keys = [k.decode() if isinstance(k, bytes) else k
                for k in rc.scan_iter('photo_batch:*', count=200)]
    except Exception as exc:
        logger.warning(f'[SLACK/방문사진복구] scan 실패: {exc}')
        return
    if not keys:
        logger.info('[SLACK/방문사진복구] 대기 배치 없음')
        return
    now = time.time()
    checked = rec_total = 0
    client = None
    for k in keys:
        raw = rc.get(k)
        if not raw:
            continue
        try:
            b = json.loads(raw.decode() if isinstance(raw, bytes) else raw)
        except Exception:
            rc.delete(k)
            continue
        # 3분 이내 배치는 진행 중일 수 있어 skip
        if now - float(b.get('created', 0) or 0) < 180:
            continue
        fid = b.get('folder_id')
        bfiles = b.get('files') or []
        if not fid or not bfiles:
            rc.delete(k)
            continue
        checked += 1
        try:
            existing = list_folder_filenames(fid)
        except Exception as exc:
            logger.warning(f'[SLACK/방문사진복구] 폴더 조회 실패 ({fid}): {exc}')
            continue
        missing = [f for f in bfiles if f.get('name') and f['name'] not in existing]
        if not missing:
            rc.delete(k)
            continue
        rec = []
        for f in missing:
            name, url = f.get('name'), f.get('url')
            if not name or not url:
                continue
            try:
                req = _ur.Request(url, headers={'Authorization': f'Bearer {token}'})
                content = _ur.urlopen(req, timeout=30).read()
                if upload_file(fid, name, content, mimetype=f.get('mime', 'image/jpeg')):
                    rec.append(name)
            except Exception as exc:
                logger.warning(f'[SLACK/방문사진복구] 업로드 실패 ({name}): {exc}')
        rec_total += len(rec)
        if rec:
            logger.info(f'[SLACK/방문사진복구] {b.get("thread_ts")} {len(rec)}장 복구: {rec}')
        # 재확인 후 다 채워졌으면 키 삭제 + (reply_ts 있으면) 스레드 답글 최종 갱신
        try:
            still = [f for f in bfiles
                     if f.get('name') and f['name'] not in list_folder_filenames(fid)]
        except Exception:
            still = missing
        if not still:
            rc.delete(k)
            reply_ts = b.get('reply_ts')
            if reply_ts and rec and token:
                try:
                    if client is None:
                        from slack_sdk import WebClient
                        client = WebClient(token=token)
                    client.chat_update(
                        channel=b.get('channel'), ts=reply_ts,
                        text=(f':white_check_mark: 사진 {len(bfiles)}장을 드라이브에 '
                              f'저장했습니다. (중단됐던 {len(rec)}장 자동 복구)'))
                except Exception:
                    pass
    logger.info(f'[SLACK/방문사진복구] 점검 {checked}배치 · 복구 {rec_total}장')


# 앱 시작 시 한 번 초기화 시도
_init_slack_app()
_init_visit_slack_app()
_init_invoice_slack_app()
_init_payment_slack_app()

# 재시작으로 중단된 방문 사진 업로드 배치 자동 복구 (백그라운드 — 앱 init 여유 후, 2026-07-30)
try:
    import threading as _recov_thr

    def _delayed_visit_photo_recovery():
        import time as _rt
        _rt.sleep(25)
        try:
            _recover_photo_batches()
        except Exception as _exc:
            logger.warning(f'[SLACK/방문사진복구] 실행 예외: {_exc}')

    _recov_thr.Thread(target=_delayed_visit_photo_recovery, daemon=True).start()
except Exception:
    pass
