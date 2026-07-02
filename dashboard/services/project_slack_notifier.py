"""
공사 확정 슬랙 알림 모듈 (별도 봇 "공사 현황 알림 봇" 사용).

- POST /api/projects/auto 에서 호출
- 환경변수: SLACK_PROJECT_BOT_TOKEN, SLACK_PROJECT_CHANNEL
- 별도 토큰이므로 기존 SLACK_BOT_TOKEN과 무관 (인입 알림과 분리)
"""

import json
import os
import urllib.request

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)


def _money_kr(value) -> str:
    """₩1,000 / 1000 / '1,000' 모두 받아 '1,000원' 형태로 정규화. 빈값은 '-'."""
    if value is None or value == '' or value == '-':
        return '-'
    s = str(value).strip()
    # ₩ 와 콤마 제거 후 숫자만 추출
    digits = ''.join(ch for ch in s if ch.isdigit() or ch == '-')
    if not digits or digits == '-':
        return s if s else '-'
    try:
        num = int(digits)
    except ValueError:
        return s
    return f"{num:,}원"


def _val(data: dict, key: str) -> str:
    """data[key]를 표시용 문자열로. 빈값은 '-'."""
    v = data.get(key)
    if v is None:
        return '-'
    s = str(v).strip()
    return s if s else '-'


def _build_message(data: dict, code: str) -> str:
    """공사 확정 알림 메시지 본문 생성.

    13줄 양식 (등록자/사업자/담당자/공사 구분/기계 분류/브랜드/계산서 제외).
    """
    code_safe = code or '-'

    # 부가세 처리 — TRUE/'TRUE'/True 모두 별도 표기
    vat_raw = data.get('부가세')
    vat_is_separate = (
        vat_raw is True
        or (isinstance(vat_raw, str) and vat_raw.strip().upper() in ('TRUE', 'Y', 'YES', '1'))
        or vat_raw == 1
    )
    amount_str = _money_kr(data.get('총액 1'))
    if amount_str != '-' and vat_is_separate:
        amount_line = f"{amount_str} (VAT 별도)"
    else:
        amount_line = amount_str

    lines = [
        f":bell: *[공사 확정 알림]*  `{code_safe}`",
        "--------------------------------------------",
        f":inbox_tray: 유입 구분 : {_val(data, '유입 구분')}",
        f":office: 사업자명 : {_val(data, '사업자명')}",
        f":round_pushpin: 현장 주소 : {_val(data, '현장 주소')}",
        f":bust_in_silhouette: 발주처 담당자 : {_val(data, '발주처 담당자')}",
        f":telephone_receiver: 발주처 연락처 : {_val(data, '발주처 연락처')}",
        f":envelope: 발주처 이메일 : {_val(data, '발주처 이메일')}",
        f":clipboard: 공사 내용 : {_val(data, '공사 내용')}",
        f":hammer_and_wrench: 도급 구분 : {_val(data, '도급 구분')}",
        f":construction_worker: 시공자 : {_val(data, '시공자')}",
        f":heavy_dollar_sign: 공사 금액 : {amount_line}",
        f":date: 공사 시작 : {_val(data, '공사 시작')}",
        f":date: 공사 종료 : {_val(data, '공사 종료')}",
        "--------------------------------------------",
    ]
    return '\n'.join(lines)


def send_project_created_notification(data: dict, code: str) -> bool:
    """공사 확정 알림을 #공사_확정 채널에 발송.

    Args:
        data: 프로젝트 dict (시트 한국어 헤더 키 — 거래처/사업자명/현장 주소 등)
        code: 프로젝트 코드 (예: 'G3764-MJ')

    Returns:
        True 성공, False 실패 (호출자는 무시해도 무방 — 시트 등록은 이미 완료).
    """
    token = os.getenv('SLACK_PROJECT_BOT_TOKEN', '').strip()
    channel = os.getenv('SLACK_PROJECT_CHANNEL', '').strip()
    if not token:
        logger.debug('[PROJECT/SLACK] SLACK_PROJECT_BOT_TOKEN 미설정 — 알림 스킵')
        return False
    if not channel:
        logger.debug('[PROJECT/SLACK] SLACK_PROJECT_CHANNEL 미설정 — 알림 스킵')
        return False

    text = _build_message(data, code)
    payload = {
        'channel': channel,
        'text': text,
        'unfurl_links': False,
    }

    try:
        req = urllib.request.Request(
            'https://slack.com/api/chat.postMessage',
            data=json.dumps(payload, ensure_ascii=False).encode('utf-8'),
            headers={
                'Content-Type': 'application/json; charset=utf-8',
                'Authorization': f'Bearer {token}',
            },
        )
        with urllib.request.urlopen(req, timeout=5) as r:
            resp = json.loads(r.read())
        if resp.get('ok'):
            logger.info(f"[PROJECT/SLACK] 공사 확정 알림 발송 완료: {code} (ts={resp.get('ts')})")
            return True
        logger.warning(f"[PROJECT/SLACK] 슬랙 API 실패 ({code}): {resp.get('error')}")
        return False
    except Exception as exc:
        logger.warning(f"[PROJECT/SLACK] 발송 예외 ({code}): {exc}")
        return False
