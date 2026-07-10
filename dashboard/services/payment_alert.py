"""수금 관리 매니저 실수 감지 + 슬랙 알림 (2026-07-10).

감지 카테고리 (6개):
    1. note_missing       — 메모 없이 금액만 입력 (10분 지연)
    2. amount_typo        — 총액 대비 10배 차이 금액 (즉시)
    3. required_missing   — 메모에 파트너/은행/날짜 중 하나 빠짐 (10분 지연)
    4. complete_untick    — 미수금=0 인데 수금완료(AA) 미체크 (10분 지연)
    6. unpaid_invalid     — X 미수금이 음수 or 총액 초과 (즉시)
    7. unknown_initial    — 매니저 이니셜이 등록 목록에 없음 (즉시)

발송 방식:
    Ephemeral (매니저만 보이는 채널 메시지) + 매니저 개인 DM 병행.
    매니저가 채널 안 봐도 DM 은 놓치지 않음.

중복 방지:
    Redis 키 `payment_alert:notified:{code}:{category}` (24h TTL).

지연 알림 (10분):
    Redis 키 `payment_alert:pending:{code}:{category}` 에 최초 감지 시각 저장.
    10분 지나야 알림 발송. 그 사이 매니저가 정정하면 감지 자체 skip.

환경변수:
    SLACK_MANAGER_SB_USER_ID  — 황샛별 매니저 슬랙 user ID (없으면 알림 skip)
    SLACK_PAYMENT_BOT_TOKEN   — 발송용 봇 토큰
    SLACK_PAYMENT_CHANNEL     — 수금 관리 채널 ID
"""
from __future__ import annotations

import os
import re
import time
from typing import Dict, List, Optional

from dashboard.utils.logging_config import get_logger
from dashboard.utils.redis_client import get_redis_client

logger = get_logger(__name__)

# 지연 알림 임계값 (초). 10분 = 600.
DELAYED_THRESHOLD = 600
# 알림 중복 방지 TTL (초). 24시간.
NOTIFIED_TTL = 60 * 60 * 24

CATEGORY_META = {
    'note_missing':      {'severity': 'delayed',   'title': '📝 메모를 채워주세요'},
    'amount_typo':       {'severity': 'immediate', 'title': '⚠️ 금액 자릿수 확인'},
    'required_missing':  {'severity': 'delayed',   'title': '📝 메모 정보 부족'},
    'complete_untick':   {'severity': 'delayed',   'title': '✅ 수금완료 체크'},
    'unpaid_invalid':    {'severity': 'immediate', 'title': '⚠️ 미수금 값 확인'},
    'unknown_initial':   {'severity': 'immediate', 'title': '⚠️ 이니셜 확인'},
    # 2026-07-10 추가 (daily 스캔 전용)
    'reference_mismatch':      {'severity': 'immediate', 'title': '🔗 참조 프로젝트 대응 기록 없음'},
    'sheet_note_mismatch':     {'severity': 'immediate', 'title': '⚠️ 메모-시트 금액 불일치'},
    # 2026-07-11 추가 — 통합 입금 케이스 (informational, 검토 불필요).
    #   같은 메모가 여러 프로젝트에 반복되고 메모합=시트 개별값 합이면
    #   통합 입금으로 자동 인식. 매니저는 그룹 요약만 확인.
    'unified_payment':         {'severity': 'info',      'title': ':white_check_mark: 통합 입금 확인'},
}


# ─────────────────────────────────────────────
# 감지 함수 6개
# ─────────────────────────────────────────────

def detect_note_missing(row: Dict, payments: List[Dict]) -> Optional[Dict]:
    """카테고리 1: U/V/W 값 있는데 메모 파싱 결과가 fallback (partner=`-` date=`-`).

    AA(수금완료) 체크된 옛 프로젝트는 skip — 이미 완료된 건은 매니저가 손댈 필요 없음.
    """
    if row.get('aa', False):
        return None
    stages_with_val = [(s, row.get(s, 0)) for s in ('계약금', '중도금', '잔금') if row.get(s, 0) > 0]
    if not stages_with_val:
        return None
    incomplete_stages = []
    for stage_name, _val in stages_with_val:
        sp = [p for p in payments if p.get('stage') == stage_name]
        if sp and sp[-1].get('date_md') == '-' and sp[-1].get('partner') == '-':
            incomplete_stages.append(stage_name)
    if not incomplete_stages:
        return None
    return {
        'stages': incomplete_stages,
        'body': (
            f"{', '.join(incomplete_stages)} 금액이 입력됐지만 메모가 비어있어요. "
            f"시트 메모에 입금 정보(일시/입금액/계좌/적요) 추가해주세요."
        ),
    }


def detect_amount_typo(row: Dict, payments: List[Dict]) -> Optional[Dict]:
    """카테고리 2: payment 금액이 총액의 10배 이상 = 자릿수 오타 (예: 26,800,0000원).

    아랫쪽 (10% 미만) 은 정상 계약금 비율이라 skip.
    AA(수금완료) 옛 프로젝트도 skip.
    """
    if row.get('aa', False):
        return None
    total_t = row.get('total_t', 0)
    if total_t <= 0 or not payments:
        return None
    suspects = []
    for p in payments:
        amt = p.get('amount', 0)
        if amt <= 0:
            continue
        if amt >= total_t * 10:  # 자릿수 하나 이상 많음 → 오타 확실
            suspects.append(f"{p.get('stage', '-')} {amt:,}원")
    if not suspects:
        return None
    return {
        'body': f"금액이 총액({total_t:,}원)의 10배 이상이에요: {', '.join(suspects)}. 자릿수 확인해주세요.",
    }


def detect_required_missing(row: Dict, payments: List[Dict]) -> Optional[Dict]:
    """카테고리 3: 메모는 있는데 파트너/은행/날짜 중 하나 빠짐.

    AA(수금완료) 옛 프로젝트는 skip.
    현금 케이스 완화 (2026-07-10 매니저 확인): partner='현금' 이면 은행 정보 요구 안 함.
    """
    if row.get('aa', False):
        return None
    issues = []
    for p in payments:
        stage = p.get('stage', '-')
        partner = p.get('partner') or ''
        bank = p.get('bank') or ''
        miss = []
        if not partner or partner == '-':
            miss.append('입금자')
        # 현금이면 은행 정보 요구 안 함 (자연스러운 무-은행 상태)
        if not bank and partner != '현금':
            miss.append('은행')
        if not p.get('date_md') or p['date_md'] == '-':
            miss.append('일시')
        if miss:
            issues.append(f"{stage} — {'/'.join(miss)} 없음")
    if not issues:
        return None
    return {
        'body': f"메모에 필수 정보가 부족해요: {'; '.join(issues)}.",
    }


def detect_complete_untick(row: Dict, payments: List[Dict]) -> Optional[Dict]:
    """카테고리 4: 미수금=0 인데 AA 체크 안 됨."""
    total_t = row.get('total_t', 0)
    unpaid = row.get('unpaid', 0)
    aa_chk = row.get('aa', False)
    if total_t <= 0 or unpaid != 0 or aa_chk:
        return None
    return {
        'body': f"미수금이 0인데 수금완료 체크가 안 되어 있어요. AA 컬럼 체크해주세요.",
    }


def detect_unpaid_invalid(row: Dict) -> Optional[Dict]:
    """카테고리 6: 미수금 이상 케이스.

    시트 X 열 정의: X = (U+V+W) - T
      - X < 0 : 정상 미납 (총액 대비 입금 부족) — 감지 대상 아님
      - X == 0 : 완납
      - X > 0 : 초과 입금 — 매니저 실수 or 정정 필요

    감지 케이스:
      1) AA=True 인데 unpaid != 0 : 수금완료 체크했는데 미납/초과 있음
      2) AA=False 인데 unpaid > 0 : 초과 입금 상태로 방치됨 (완료 처리 필요)
    """
    unpaid = row.get('unpaid', 0)
    if row.get('aa', False):
        if unpaid != 0:
            # 반올림 오차 (|unpaid| < 100원) 는 무시 — 매니저 확인 반영 (2026-07-10)
            #   1원, 몇십원 단위는 카드 수수료 계산 등에서 자연 발생
            if abs(unpaid) < 100:
                return None
            direction = '더 받아야' if unpaid < 0 else '초과 입금'
            return {
                'body': f"수금완료 체크됐는데 {abs(unpaid):,}원 {direction} 상태예요. 확인해주세요.",
            }
        return None
    # AA=False: 초과 입금 방치 케이스만
    if unpaid > 0:
        return {'body': f"입금액이 총액보다 {unpaid:,}원 초과했어요. 확인해주세요."}
    return None


def detect_unknown_initial(row: Dict, known_initials: set) -> Optional[Dict]:
    """카테고리 7: 매니저 이니셜이 등록 목록에 없음.

    AA(수금완료) 옛 프로젝트는 skip — 옛 이니셜 (퇴사 매니저 등) 노이즈 방지.
    """
    if row.get('aa', False):
        return None
    code = row.get('code', '')
    m = re.match(r'^[GRN]\d{4}-([A-Z]+)$', code)
    if not m:
        return None
    ini = m.group(1)
    if ini in known_initials:
        return None
    return {
        'body': f"프로젝트 코드의 이니셜 '{ini}' 이 등록 목록에 없어요. 확인해주세요.",
    }


# ─────────────────────────────────────────────
# 지연/중복 관리 + 발송
# ─────────────────────────────────────────────

def _redis():
    try:
        return get_redis_client().redis
    except Exception:
        return None


def _should_send(code: str, category: str) -> bool:
    """지연/중복 체크 → 실제로 알림 보내야 하는지 결정.

    severity=immediate: 중복만 체크.
    severity=delayed: 최초 감지 시각을 저장 → DELAYED_THRESHOLD 지난 뒤에만 발송.
    """
    meta = CATEGORY_META.get(category, {})
    severity = meta.get('severity', 'immediate')
    rc = _redis()
    if rc is None:
        return False
    notified_key = f'payment_alert:notified:{code}:{category}'
    try:
        if rc.exists(notified_key):
            return False  # 24h 안에 이미 알림 발송함
    except Exception:
        pass
    if severity == 'immediate':
        return True
    # delayed: 최초 감지 시각 저장 or 확인
    pending_key = f'payment_alert:pending:{code}:{category}'
    try:
        first_ts = rc.get(pending_key)
        if first_ts is None:
            rc.set(pending_key, str(int(time.time())), ex=NOTIFIED_TTL)
            return False  # 첫 감지 — 다음 폴링에서 지연 판정
        first = int(first_ts if isinstance(first_ts, str) else first_ts.decode())
        elapsed = int(time.time()) - first
        return elapsed >= DELAYED_THRESHOLD
    except Exception as exc:
        logger.warning(f'[PAYMENT_ALERT] pending 체크 실패 ({code}/{category}): {exc}')
        return False


def _mark_notified(code: str, category: str) -> None:
    rc = _redis()
    if rc is None:
        return
    try:
        rc.set(f'payment_alert:notified:{code}:{category}', '1', ex=NOTIFIED_TTL)
        # pending 정리
        rc.delete(f'payment_alert:pending:{code}:{category}')
    except Exception:
        pass


def clear_pending_if_resolved(code: str, category: str) -> None:
    """감지 조건 해제됨 (매니저가 정정함) → pending 삭제해서 재감지 시 다시 카운트."""
    rc = _redis()
    if rc is None:
        return
    try:
        rc.delete(f'payment_alert:pending:{code}:{category}')
    except Exception:
        pass


def send_alert(row: Dict, category: str, detected: Dict, slack_client=None) -> bool:
    """실제 발송 — Ephemeral + DM 병행. 성공 시 mark_notified.

    환경변수 부재 (매니저 user ID 없음) 시 안전하게 skip.
    """
    code = row.get('code', '-')
    manager_user_id = os.getenv('SLACK_MANAGER_SB_USER_ID', '').strip()
    channel = os.getenv('SLACK_PAYMENT_CHANNEL', '').strip()
    bot_token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    if not manager_user_id or not channel or not bot_token:
        logger.debug(f'[PAYMENT_ALERT] 발송 skip ({code}/{category}) — env 미설정')
        return False

    if slack_client is None:
        try:
            from slack_sdk import WebClient
            slack_client = WebClient(token=bot_token)
        except Exception as exc:
            logger.warning(f'[PAYMENT_ALERT] Slack client 초기화 실패: {exc}')
            return False

    meta = CATEGORY_META.get(category, {})
    title = meta.get('title', '알림')
    addr = (row.get('address') or '').strip()[:60]
    body = detected.get('body', '')

    text = (
        f"{title}\n"
        f"• `{code}` {addr}\n"
        f"  {body}"
    )

    sent_any = False
    try:
        slack_client.chat_postEphemeral(channel=channel, user=manager_user_id, text=text)
        sent_any = True
    except Exception as exc:
        logger.warning(f'[PAYMENT_ALERT] ephemeral 실패 ({code}/{category}): {exc}')

    # DM 병행 — conversations.open 으로 IM 채널 열고 postMessage
    #   2026-07-10 safe_slack_call 로 재시도 자동화. 매니저 실수 알림은 놓치면
    #   실질 손실이라 429·5xx 대응 필수.
    try:
        from dashboard.blueprints.slack_helpers import safe_slack_call
        im = safe_slack_call(slack_client.conversations_open, users=manager_user_id)
        if im.get('ok'):
            im_ch = im['channel']['id']
            safe_slack_call(slack_client.chat_postMessage, channel=im_ch, text=text)
            sent_any = True
    except Exception as exc:
        logger.warning(f'[PAYMENT_ALERT] DM 실패 ({code}/{category}): {exc}')

    if sent_any:
        _mark_notified(code, category)
        logger.info(f'[PAYMENT_ALERT] 알림 발송: {code}/{category}')
    return sent_any


# ─────────────────────────────────────────────
# 폴링 훅 (payment_sync 에서 호출)
# ─────────────────────────────────────────────

def check_and_alert(
    row: Dict,
    payments: List[Dict],
    known_initials: set,
    slack_client=None,
) -> List[str]:
    """모든 감지 실행 + 필요 시 발송. 발송한 카테고리 리스트 반환."""
    sent = []
    detectors = [
        ('note_missing',     lambda: detect_note_missing(row, payments)),
        ('amount_typo',      lambda: detect_amount_typo(row, payments)),
        ('required_missing', lambda: detect_required_missing(row, payments)),
        ('complete_untick',  lambda: detect_complete_untick(row, payments)),
        ('unpaid_invalid',   lambda: detect_unpaid_invalid(row)),
        ('unknown_initial',  lambda: detect_unknown_initial(row, known_initials)),
    ]
    for category, fn in detectors:
        try:
            detected = fn()
        except Exception as exc:
            logger.warning(f'[PAYMENT_ALERT] {category} 감지 예외 ({row.get("code", "-")}): {exc}')
            continue
        if detected is None:
            # 조건 해제됨 — pending 정리 (재감지 시 다시 카운트)
            clear_pending_if_resolved(row.get('code', ''), category)
            continue
        if not _should_send(row.get('code', ''), category):
            continue
        if send_alert(row, category, detected, slack_client=slack_client):
            sent.append(category)
    return sent


# ─────────────────────────────────────────────
# 일일 요약 (매일 오전 9시 평일)
# ─────────────────────────────────────────────

def build_daily_summary_text(alerts: List[Dict]) -> str:
    """알림 리스트 → 매니저 친화적 채널 메시지."""
    if not alerts:
        return ":white_check_mark: 오늘은 확인할 이슈가 없어요! 계속 좋은 하루 되세요 :sparkles:"
    lines = [
        f":sunny: *오늘 확인할 것 ({len(alerts)}건)*",
        "─" * 30,
    ]
    grouped = {}
    for a in alerts:
        cat = a.get('category', '기타')
        grouped.setdefault(cat, []).append(a)
    for cat, items in grouped.items():
        meta = CATEGORY_META.get(cat, {})
        title = meta.get('title', cat)
        lines.append(f"\n{title}")
        for it in items:
            addr = (it.get('address') or '')[:50]
            lines.append(f"• `{it['code']}` {addr}")
            lines.append(f"    {it.get('body', '')}")
    return '\n'.join(lines)


def send_daily_summary(alerts: List[Dict]) -> bool:
    """평일 오전 9시 채널 공개 발송."""
    channel = os.getenv('SLACK_PAYMENT_CHANNEL', '').strip()
    bot_token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    if not channel or not bot_token:
        return False
    try:
        from slack_sdk import WebClient
        from dashboard.blueprints.slack_helpers import safe_slack_call
        client = WebClient(token=bot_token)
        text = build_daily_summary_text(alerts)
        # 2026-07-10 safe_slack_call 로 이관 — 일일 요약 놓치면 매니저가 몰라서
        # 하루치 실수 감지 실패. 재시도 자동화 필수.
        safe_slack_call(client.chat_postMessage, channel=channel, text=text)
        logger.info(f'[PAYMENT_ALERT] 일일 요약 발송: {len(alerts)}건')
        return True
    except Exception as exc:
        logger.error(f'[PAYMENT_ALERT] 일일 요약 발송 실패: {exc}', exc_info=True)
        return False
