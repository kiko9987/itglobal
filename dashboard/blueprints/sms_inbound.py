# -*- coding: utf-8 -*-
"""은행 입금 SMS 인입 webhook — 폰 SMS 포워딩 → 슬랙 #수금_입력 버튼 카드.

흐름:
1. 폰(안드 2 + 아이폰 1) SMS 포워딩 앱 → POST /sms/inbound {token, sender, text}
2. 기기 토큰 인증(SMS_INBOUND_TOKENS) → 은행 발신번호 allowlist(선택) → 입금 문자 판별
3. 잔액 라인 제거(sms_intake.strip_balance) → 통장 잔고 전 직원 노출 차단
4. Redis 중복제거 + 원문(잔액 제거본) 보관(sms_intake:{id})
5. #수금_입력 채널에 [프로젝트·수금단계 지정] 버튼 카드 게시(수금봇 토큰)
   → 매니저가 모달에서 프로젝트+단계 선택 → 시트 U/V/W 셀 메모 기록(slack_bot.py)
   → 기존 payment_sync 폴러가 감지 → #수금_관리 정식 카드 (여기부턴 기존 파이프라인)

보안: 로그인 세션 없는 기계 webhook. security_middleware 가 /sms/ 는 CSRF 우회
(기기 토큰으로 자체 인증). 실제 값은 URL/쿼리에 싣지 않고 JSON body 로만 받는다.
"""

import json
import os
import time

from flask import Blueprint, jsonify, request

from dashboard.services.sms_intake import (
    active_display, dedup_hash, has_business_account, looks_like_payment,
    parse_preview, strip_balance,
)
from dashboard.utils.logging_config import get_logger
from dashboard.utils.redis_client import get_redis_client

logger = get_logger(__name__)

sms_bp = Blueprint('sms_inbound', __name__, url_prefix='/sms')

_INTAKE_TTL = 60 * 60 * 24 * 7   # 원문 보관 7일 (모달 제출까지 여유)
_DEDUP_TTL = 60 * 60 * 24        # 중복 무시 24시간


def _load_device_tokens() -> dict:
    """SMS_INBOUND_TOKENS='jw:secretA,yg:secretB,ip:secretC' → {secret: device}.

    폰별 비밀키 → 한 대 유출돼도 개별 폐기 가능. 값(secret)은 로그에 남기지 않는다.
    """
    raw = os.getenv('SMS_INBOUND_TOKENS', '').strip()
    tokens = {}
    for pair in raw.split(','):
        pair = pair.strip()
        if not pair or ':' not in pair:
            continue
        device, secret = pair.split(':', 1)
        secret = secret.strip()
        if secret:
            tokens[secret] = device.strip() or 'unknown'
    return tokens


def _bank_senders() -> list:
    """SMS_BANK_SENDERS='1544xxxx,15881111' → 허용 발신번호 substring 목록 (선택)."""
    raw = os.getenv('SMS_BANK_SENDERS', '').strip()
    return [s.strip() for s in raw.split(',') if s.strip()]


@sms_bp.route('/inbound', methods=['POST'])
def sms_inbound():
    """폰 SMS 포워딩 앱 → 은행 입금 문자 인입.

    응답: 정상/무시/중복은 200 (폰 앱 재전송 폭주 방지), 토큰 불일치만 401.
    """
    tokens = _load_device_tokens()
    if not tokens:
        logger.error('[SMS_INBOUND] SMS_INBOUND_TOKENS 미설정 — 인입 비활성화')
        return jsonify({'error': 'not configured'}), 503

    data = request.get_json(silent=True) or {}
    token = (data.get('token') or request.headers.get('X-SMS-Token') or '').strip()
    device = tokens.get(token)
    if not device:
        logger.warning(f'[SMS_INBOUND] 토큰 불일치 — 거부 (ip={request.remote_addr})')
        return jsonify({'error': 'unauthorized'}), 401

    sender = (data.get('sender') or '').strip()
    text = (data.get('text') or '').strip()
    if not text:
        return jsonify({'status': 'empty'}), 200

    # 발신번호 allowlist (설정된 경우에만 적용)
    senders = _bank_senders()
    if senders and not any(s in sender for s in senders):
        logger.info(f'[SMS_INBOUND] 발신번호 미허용 skip: {sender!r} (device={device})')
        return jsonify({'status': 'ignored', 'reason': 'sender'}), 200

    # 입금 문자 판별 (광고·인증·택배 등 배제)
    if not looks_like_payment(text):
        logger.info(f'[SMS_INBOUND] 입금 문자 아님 skip (device={device})')
        return jsonify({'status': 'ignored', 'reason': 'not_payment'}), 200

    # 사업자 통장(452/255/352) 입금만 통과 — 개인 계좌 입금(같은 은행이라도) 배제
    if not has_business_account(text):
        logger.info(f'[SMS_INBOUND] 사업자 통장 계좌 아님 skip — 개인 입금 배제 (device={device})')
        return jsonify({'status': 'ignored', 'reason': 'not_business_account'}), 200

    # Redis (중복제거·원문 보관)
    rc = None
    try:
        rc = get_redis_client().redis
    except Exception as exc:
        logger.warning(f'[SMS_INBOUND] Redis 접근 실패 (dedup/보관 skip): {exc}')

    intake_id = dedup_hash(sender, text)
    if rc is not None:
        if not rc.set(f'sms_intake:seen:{intake_id}', '1', nx=True, ex=_DEDUP_TTL):
            logger.info(f'[SMS_INBOUND] 중복 문자 skip: id={intake_id} (device={device})')
            return jsonify({'status': 'duplicate', 'id': intake_id}), 200

    # 잔액 제거 (통장 잔고 노출 차단)
    clean = strip_balance(text)
    if not clean:
        return jsonify({'status': 'ignored', 'reason': 'empty_after_strip'}), 200

    preview = parse_preview(clean)

    # 원문(잔액 제거본) 보관 — 모달 제출 시 시트에 기록할 내용
    if rc is not None:
        try:
            rc.set(f'sms_intake:{intake_id}', json.dumps({
                'text': clean, 'sender': sender, 'device': device,
                'preview': preview, 'ts': int(time.time()),
            }), ex=_INTAKE_TTL)
        except Exception as exc:
            logger.warning(f'[SMS_INBOUND] 원문 보관 실패: {exc}')

    # 슬랙 인입 카드 게시
    if not _post_intake_card(intake_id, clean, preview):
        return jsonify({'status': 'card_failed', 'id': intake_id}), 200

    logger.info(f"[SMS_INBOUND] 인입 카드 게시: id={intake_id} device={device} "
                f"partner={(preview or {}).get('partner')!r}")
    return jsonify({'status': 'ok', 'id': intake_id}), 200


def _post_intake_card(intake_id: str, clean_text: str, preview: dict) -> bool:
    """#수금_입력 채널에 [프로젝트·수금단계 지정] 버튼 카드 게시 (수금봇 토큰)."""
    channel = os.getenv('SLACK_PAYMENT_INTAKE_CHANNEL', '').strip()
    bot_token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    if not channel or not bot_token:
        logger.error('[SMS_INBOUND] SLACK_PAYMENT_INTAKE_CHANNEL/BOT_TOKEN 미설정 — 카드 미게시')
        return False
    try:
        from slack_sdk import WebClient
        slack = WebClient(token=bot_token)
        slack.chat_postMessage(
            channel=channel,
            text='입금 문자 도착 — 프로젝트 지정 필요',
            blocks=_build_intake_blocks(intake_id, clean_text, preview),
        )
        return True
    except Exception as exc:
        logger.error(f'[SMS_INBOUND] 카드 게시 실패: {exc}', exc_info=True)
        return False


def _build_intake_blocks(intake_id: str, clean_text: str, preview: dict) -> list:
    # 문자 원문(여러 줄) 그대로 노출 — 매니저가 원문 보고 프로젝트 식별. 잔액만 제거된 상태.
    return [
        {"type": "context", "elements": [{"type": "mrkdwn", "text": "⠀"}]},   # 상단 여백
        {"type": "section", "text": {"type": "mrkdwn",
            "text": "*💰 입금 문자 도착 — 프로젝트 지정 필요*"}},
        {"type": "section", "text": {"type": "mrkdwn", "text": active_display(clean_text)}},
        {"type": "actions", "elements": [{
            "type": "button",
            "text": {"type": "plain_text", "text": "📌 프로젝트·수금단계 지정"},
            "style": "primary",
            "action_id": "payment_intake_open",
            "value": intake_id,
        }]},
        {"type": "context", "elements": [{"type": "mrkdwn", "text": "⠀"}]},   # 하단 여백
    ]
