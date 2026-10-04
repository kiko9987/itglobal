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
    active_display, dedup_hash, has_business_account, is_bank_interest,
    looks_like_cash, looks_like_payment, looks_like_withdrawal, normalize_cash_layout,
    normalize_deposit_layout, parse_preview, strip_balance, strip_bold_markers,
)
from dashboard.utils.logging_config import get_logger
from dashboard.utils.redis_client import get_redis_client

logger = get_logger(__name__)

sms_bp = Blueprint('sms_inbound', __name__, url_prefix='/sms')

_INTAKE_TTL = 60 * 60 * 24 * 7   # 원문 보관 7일 (모달 제출까지 여유)
_DEDUP_TTL = 60 * 60 * 24        # 중복 무시 24시간
_OUTFLOW_TTL = 60 * 60 * 24 * 14  # 출금 문자 서버 보관 14일 (슬랙 미노출)
_OUTFLOW_INDEX = 'sms_outflow:index'   # zset(score=수신 시각) — 기간 조회용


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

    # 입금 문자 판별 → 사업자통장 → dedup → 잔액제거 → Redis → 🔗카드 (공통 파이프라인)
    result = ingest_deposit(text, source=f'sms:{device}')
    if result.get('status') == 'ok':
        logger.info(f"[SMS_INBOUND] 인입 카드 게시: id={result.get('id')} device={device} "
                    f"partner={(result.get('preview') or {}).get('partner')!r}")
    else:
        logger.info(f"[SMS_INBOUND] skip: {result.get('status')}/{result.get('reason','')} (device={device})")
    result.pop('preview', None)
    return jsonify(result), 200


def ingest_deposit(text: str, source: str = 'sms', cash_receiver: str = '') -> dict:
    """은행 입금 문자 1건 처리 — 폰 포워딩과 채널 붙여넣기 공통 코어.

    판별(looks_like_payment) → 사업자통장(has_business_account) → Redis dedup →
    잔액 제거 → 미리보기 파싱 → Redis 보관 → #수금_입력 🔗 인입 카드 게시.
    source: 추적용 라벨('sms:{device}' 또는 'channel:{user_id}'). dedup 키엔 미반영
    (같은 입금이 폰·수동 양쪽으로 와도 본문 동일 → 첫 1건만 카드).
    Returns: {'status': 'ok'|'ignored'|'duplicate'|'card_failed', 'id'?, 'reason'?, 'preview'?}
    """
    # 렌더된 슬랙 메시지 복사·붙여넣기로 각 줄이 '*줄*' 볼드 마크다운으로 딸려온 경우 제거
    # (2026-09-29 아이티플레이 건 — 표시 '∗줄∗'·거래처 오파싱·시트 '*' 오염 동시 유발).
    # 폰/수동/forward 공통 코어라 여기 한 곳이면 전 경로 정상화. 계좌 마스킹 내부 별표는 보존.
    text = strip_bold_markers(text)
    # 현금 수령 판별 — '현금' + 금액, 사업자계좌 없음 (은행 SMS 아닌 매니저 자유문장).
    # 계좌 있으면 은행 입금이므로 현금으로 오인 안 함. (2026-08 현금 인입 도입)
    is_cash = looks_like_cash(text) and not has_business_account(text)
    if not is_cash:
        if not looks_like_payment(text):
            # 회사 계좌 출금 문자 — 슬랙 카드 없이 서버에만 보관 (법인 간 이체 짝짓기·진단용)
            if looks_like_withdrawal(text) and has_business_account(text):
                return record_outflow(text, source)
            return {'status': 'ignored', 'reason': 'not_payment'}
        # 사업자 통장(452/255/352) 입금만 통과 — 개인 계좌 입금(같은 은행이라도) 배제
        if not has_business_account(text):
            return {'status': 'ignored', 'reason': 'not_business_account'}
        # 은행 예금 이자·결산 입금 제외 — 프로젝트 입금 아님 (적요 '2026년결산' 등)
        if is_bank_interest(text):
            return {'status': 'ignored', 'reason': 'bank_interest'}

    # Redis (중복제거·원문 보관)
    rc = None
    try:
        rc = get_redis_client().redis
    except Exception as exc:
        logger.warning(f'[SMS_INBOUND] Redis 접근 실패 (dedup/보관 skip): {exc}')

    intake_id = dedup_hash(source, text)
    if rc is not None:
        if not rc.set(f'sms_intake:seen:{intake_id}', '1', nx=True, ex=_DEDUP_TTL):
            return {'status': 'duplicate', 'id': intake_id}

    if is_cash:
        # 현금 자유문장 → 표준 메모 ('MM/DD 입금 X원 / 현금 수령 (수령자)'). 잔액·계좌 불필요.
        # 수령자 미기재 시 cash_receiver(올린 사람 이니셜)로 채움 (채널 인입 경로에서 전달).
        clean = normalize_cash_layout(text, default_receiver=cash_receiver)
        converted = False
    else:
        # 잔액 제거 (통장 잔고 노출 차단) → 은행별 압축 양식 필드 줄바꿈 재구성(농협 등)
        clean = strip_balance(text)
        if not clean:
            return {'status': 'ignored', 'reason': 'empty_after_strip'}
        clean_conv = normalize_deposit_layout(clean)
        converted = clean_conv != clean   # 원본 양식 자동 변환됐는지 (농협 등)
        clean = clean_conv

    preview = parse_preview(clean)
    if converted:
        preview['converted'] = True   # 카드에 '농협 자동 변환' 배지 표시용
    if is_cash:
        preview['cash'] = True         # 카드에 '현금 수령 변환' 배지 표시용

    # 원문(잔액 제거본) 보관 — 모달 제출 시 시트에 기록할 내용
    if rc is not None:
        try:
            rc.set(f'sms_intake:{intake_id}', json.dumps({
                'text': clean, 'sender': source, 'device': source,
                'preview': preview, 'ts': int(time.time()),
            }), ex=_INTAKE_TTL)
        except Exception as exc:
            logger.warning(f'[SMS_INBOUND] 원문 보관 실패: {exc}')

    if not _post_intake_card(intake_id, clean, preview):
        return {'status': 'card_failed', 'id': intake_id}
    return {'status': 'ok', 'id': intake_id, 'preview': preview}


def record_outflow(text: str, source: str = 'sms') -> dict:
    """회사 계좌 출금 문자 → 서버(Redis)에만 보관. 슬랙 카드·알림 없음 (2026-10-04 사용자 결정).

    용도: ①법인 간 이체 짝짓기 증거(예 하나 출금 3,014,000원 → 1분 뒤 기업 입금 같은 금액·
    같은 이름, G4125-YM) ②'무시된 문자가 뭐였나' 진단 — 예전엔 출금을 받자마자 버려 원문이
    남지 않았다. 잔액 제거본만 저장(통장 잔고 비노출 원칙 동일). 폰 3대 동시 포워딩 → 본문 dedup.
    """
    try:
        rc = get_redis_client().redis
    except Exception as exc:
        logger.warning(f'[SMS_INBOUND] Redis 접근 실패 (출금 보관 skip): {exc}')
        return {'status': 'ignored', 'reason': 'outflow_unstored'}
    oid = dedup_hash(source, text)
    if not rc.set(f'sms_outflow:seen:{oid}', '1', nx=True, ex=_DEDUP_TTL):
        return {'status': 'duplicate', 'reason': 'outflow', 'id': oid}
    clean = strip_balance(text)
    pv = parse_preview(clean) or {}
    now = int(time.time())
    rec = {
        'text': clean, 'source': source, 'ts': now,
        'amount': abs(int(pv.get('amount') or 0)),
        'partner': (pv.get('partner') or '').strip(),
        'bank': pv.get('bank') or '', 'acct_code': pv.get('acct_code') or '',
        'date_md': pv.get('date_md') or '',
    }
    try:
        rc.set(f'sms_outflow:{oid}', json.dumps(rec, ensure_ascii=False), ex=_OUTFLOW_TTL)
        rc.zadd(_OUTFLOW_INDEX, {oid: now})
        rc.zremrangebyscore(_OUTFLOW_INDEX, 0, now - _OUTFLOW_TTL)
    except Exception as exc:
        logger.warning(f'[SMS_INBOUND] 출금 보관 실패: {exc}')
        return {'status': 'ignored', 'reason': 'outflow_unstored'}
    logger.info(f"[SMS_INBOUND] 출금 문자 서버 보관(슬랙 미노출): {rec['bank']}({rec['acct_code']}) "
                f"{rec['amount']:,}원 {rec['partner']!r} id={oid}")
    return {'status': 'ignored', 'reason': 'outflow_recorded', 'id': oid}


def recent_outflows(since_ts: int = 0) -> list:
    """보관된 출금 문자 목록(수신 시각 오름차순). since_ts 이후만. 짝짓기·진단 조회용."""
    try:
        rc = get_redis_client().redis
        ids = rc.zrangebyscore(_OUTFLOW_INDEX, since_ts, '+inf')
        out = []
        for i in ids:
            i = i.decode() if isinstance(i, bytes) else i
            raw = rc.get(f'sms_outflow:{i}')
            if raw:
                out.append(dict(json.loads(raw), id=i))
        return out
    except Exception as exc:
        logger.warning(f'[SMS_INBOUND] 출금 목록 조회 실패: {exc}')
        return []


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
        resp = slack.chat_postMessage(
            channel=channel,
            text='입금 문자 도착 — 프로젝트 지정 필요',
            blocks=_build_intake_blocks(intake_id, clean_text, preview),
        )
        # 미처리 박제 — 프로젝트 지정·기록 전까지 카드 고정(pin). 경영지원 [✅ 확인 후
        # 기록] 시 자동 해제. 고정 목록 = 아직 처리 안 된 입금. pins:write scope 없으면
        # 실패해도 카드 게시는 정상(무해).
        ts = (resp or {}).get('ts')
        if ts:
            try:
                slack.pins_add(channel=channel, timestamp=ts)
            except Exception as exc:
                logger.warning(f'[SMS_INBOUND] 인입 카드 고정 실패(무해, pins:write 확인 필요): {exc}')
            if not (preview or {}).get('cash'):
                _suggest_transfer_async(slack, channel, ts, intake_id, clean_text, preview)
        return True
    except Exception as exc:
        logger.error(f'[SMS_INBOUND] 카드 게시 실패: {exc}', exc_info=True)
        return False


def _suggest_transfer_async(slack, channel: str, ts: str, intake_id: str,
                            clean_text: str, preview: dict) -> None:
    """게시한 입금 카드가 법인 간 이체로 들어온 돈인지 백그라운드로 점검 → 카드에 안내·버튼 추가.

    대기 중인 자금 이동 요청(시트 메모 원장, fund_transfer)과 도착 사업자·금액·이름 앞부분이
    맞으면 '[🔁 법인 간 이체로 처리]' + 일치 안내, 맞는 게 없어도 같은 사업자로 가는 대기 요청이
    있으면 '[🔁 법인 간 이체]'(직접 고르기). 대기 요청 자체가 없으면 카드 그대로(평소 입금 깔끔).
    """
    import threading

    def _bg():
        try:
            from dashboard.services import fund_transfer as ft
            to = ft.deposit_entity(clean_text, preview)
            if to not in ('G', 'R'):
                return
            reqs = ft.pending_requests()
            if not any(r.get('to') == to for r in reqs):
                return
            cands = ft.match_deposit(clean_text, preview, reqs)
            evidence = bool(cands) and bool(
                ft.outflow_evidence(cands[0], int(time.time()) - 3 * 86400))
            # 그 사이 매니저가 먼저 지정했으면 카드를 되돌리지 않는다
            raw = get_redis_client().redis.get(f'sms_intake:{intake_id}')
            if not raw or (json.loads(raw).get('designation')):
                return
            slack.chat_update(
                channel=channel, ts=ts, text='입금 문자 도착 — 프로젝트 지정 필요',
                blocks=_build_intake_blocks(intake_id, clean_text, preview, transfer={
                    'cands': cands, 'evidence': evidence}))
            if cands:
                logger.info(f"[FUND_MOVE] 이체 입금 후보 안내: {intake_id} → "
                            f"{[(c['code'], c['stage']) for c in cands]} 출금증거={evidence}")
        except Exception as exc:
            logger.warning(f'[FUND_MOVE] 이체 후보 점검 실패(무해): {exc}')

    threading.Thread(target=_bg, daemon=True).start()


def _build_intake_blocks(intake_id: str, clean_text: str, preview: dict,
                         transfer: dict = None) -> list:
    """transfer: {'cands': [대기 요청…], 'evidence': bool} — 법인 간 이체 안내·버튼 (2026-10-04)."""
    # 온라인/방문 카드와 동일 구조 — 헤더·구분선·본문·구분선을 한 섹션에 전부 '>' 인용으로
    # 넣어 섹션 간 여백 제거(2026-08-14). 문자 원문 그대로 노출(잔액만 제거).
    from dashboard.services.sms_intake import INTAKE_SEP, quoted_body
    # 헤더 은행 라벨 = 계좌로 판정(사업자까지 정확) → 미매칭 시 은행명 기준 기존 라벨.
    # 기업은행이라도 하도급지킴이 계좌는 '(글로벌그룹)' (2026-09-30).
    from dashboard.services.itg_accounts import match_account
    _acct = match_account(clean_text)
    bank = (preview or {}).get('bank') or ''
    bank_label = _acct.label if _acct else {
        '기업': '기업은행 (글로벌)',
        '하나': '하나은행 (글로벌그룹)',
        '농협': '농협은행 (N통장)',
    }.get(bank, '')
    is_cash = bool((preview or {}).get('cash'))
    if is_cash:
        bank_label = '현금'   # 은행명 자리에 '현금' — 은행 카드 헤더 틀 유지
    header = '새 입금 내역 알림' + (f' - {bank_label}' if bank_label else '')
    lines = ["⠀", f">🔔 *{header}*", f">{INTAKE_SEP}",
             *quoted_body(clean_text), f">{INTAKE_SEP}"]
    elements = [
        {
            "type": "button",
            "text": {"type": "plain_text", "text": "🔗 프로젝트 지정"},
            "action_id": "payment_intake_open",
            "value": intake_id,
        },
        {
            # 통합 입금(한 문자에 여러 프로젝트 금액) → 프로젝트별 분할 지정
            "type": "button",
            "text": {"type": "plain_text", "text": "⚖️ 프로젝트 분할 지정"},
            "action_id": "payment_intake_split_open",
            "value": intake_id,
        },
    ]
    blocks = [{"type": "section", "text": {"type": "mrkdwn", "text": '\n'.join(lines)}}]
    if transfer is not None:
        from dashboard.services.fund_transfer import ENTITY
        cands = (transfer.get('cands') or [])[:3]
        if cands:
            descs = [f"`{c['code']}` {c['stage']} ({c.get('bank', '')} → {ENTITY.get(c['to'], c['to'])} · "
                     f"{int(c['amount']):,}원 · {c.get('partner', '')})" for c in cands]
            ev = ' · 출발 계좌 출금 문자 확인됨' if transfer.get('evidence') else ''
            blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": (
                f"🔁 *법인 간 이체로 보입니다* — 자금 이동 요청 {' / '.join(descs)}와 일치{ev}\n"
                f"_경영지원이 이체한 돈이면 아래 버튼으로 처리하세요. 고객이 보낸 돈이면 평소처럼 지정._")}})
            tbtns = [{
                "type": "button", "style": "primary",
                "text": {"type": "plain_text",
                         "text": "🔁 법인 간 이체로 처리" if len(cands) == 1 else f"🔁 이체: {c['code']}"},
                "action_id": f"payment_intake_transfer_{k}",
                "value": json.dumps({'iid': intake_id, 'code': c['code'], 'stage': c['stage']},
                                    ensure_ascii=False),
            } for k, c in enumerate(cands)]
            elements = tbtns + elements
        else:
            elements.append({
                "type": "button",
                "text": {"type": "plain_text", "text": "🔁 법인 간 이체"},
                "action_id": "payment_intake_transfer_pick",
                "value": intake_id,
            })
    blocks.append({"type": "actions", "elements": elements})
    # 원본 압축 문자(농협 등)가 표준 양식으로 자동 변환된 경우 배지 표시
    if (preview or {}).get('converted'):
        blocks.append({"type": "context", "elements": [
            {"type": "mrkdwn", "text": "🔄 _원본 농협 문자를 표준 양식으로 자동 변환한 카드입니다._"}]})
    if is_cash:
        blocks.append({"type": "context", "elements": [
            {"type": "mrkdwn", "text": "💵 _현금 수령 메시지를 표준 양식으로 자동 변환한 카드입니다._"}]})
    return blocks
