"""미처리 정산 핀 리마인드 — 매일 아침 9시 #영업_관리.

경영지원(황샛별)이 입금내역·세금계산서를 올리며 "어떤 프로젝트 비용인지" 확인 요청.
미처리 건은 고정(pin), 처리되면 고정 해제 → **고정된 것 = 미처리**. 매일 아침 요약 알림.

타입 구분:
  - 입금내역: ITG 통장 계좌번호(452/255) 감지 (payment_sync 정규식 재사용)
  - 세금계산서: 나머지 (세금계산서 이미지 첨부)

봇: 세금계산서 관리 알림 봇 (SLACK_INVOICE_BOT_TOKEN, #영업_관리 담당).
필요 scope: pins:read.
"""
from __future__ import annotations

import logging
import os
import re
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

_SEP = '--------------------------------------------'
_BLANK = '⠀'


def _client():
    from slack_sdk import WebClient
    tok = os.getenv('SLACK_INVOICE_BOT_TOKEN', '').strip()
    return WebClient(token=tok) if tok else None


def _channel() -> str:
    return os.getenv('SLACK_INVOICE_CHANNEL_ID', '').strip()


def _msg_full_text(m: dict) -> str:
    """핀 메시지에서 분류·표시용 텍스트 통합 (text + attachments + blocks)."""
    parts: List[str] = []
    if m.get('text'):
        parts.append(str(m['text']))
    for at in m.get('attachments', []) or []:
        t = at.get('text') or at.get('fallback')
        if t:
            parts.append(str(t))
    return '\n'.join(parts)


def _is_deposit(text: str) -> bool:
    """ITG 통장 계좌번호(452 기업 / 255 하나) 포함 → 입금내역."""
    try:
        from dashboard.services.payment_sync import _ACCT_G_RE, _ACCT_R_RE
        return bool(_ACCT_G_RE.search(text) or _ACCT_R_RE.search(text))
    except Exception:
        # fallback — 마스킹 계좌 패턴
        return bool(re.search(r'\b(?:452|255)[\d*\-]{6,}', text))


def _summary_line(text: str) -> str:
    """한 줄 요약 — 계좌번호·마크다운·개행·꼬리 문구 제거."""
    s = text.replace('[Web발신]', '')
    s = re.sub(r'(?:452|255)[\d*\-]{5,}', '', s)   # 계좌번호 제거 (마스킹 깨짐·노이즈 방지)
    s = re.sub(r'[*_]', '', s)                       # bold 마크다운 제거
    s = re.sub(r'\s*\n\s*', ' ', s).strip()
    s = re.sub(r'\s*회신\s*부탁드립니다\.?\s*$', '', s).strip()
    s = re.sub(r'\s{2,}', ' ', s)
    return s[:80] or '(내용 없음)'


def collect_pending_pins() -> Optional[dict]:
    """#영업_관리 고정 메시지 조회 → 입금/세금계산서 분류.

    Returns None (client/channel 미설정·조회 실패) 또는
      {'deposits': [{ts,summary,permalink}], 'invoices': [...], 'total': int}
    """
    c = _client()
    ch = _channel()
    if not c or not ch:
        logger.warning('[PIN] SLACK_INVOICE_BOT_TOKEN/CHANNEL 미설정')
        return None
    try:
        resp = c.pins_list(channel=ch)
    except Exception as exc:
        logger.warning(f'[PIN] pins.list 실패: {exc}')
        return None

    deposits: List[dict] = []
    invoices: List[dict] = []
    for it in resp.get('items', []) or []:
        m = it.get('message', {})
        if not m:
            continue
        ts = m.get('ts', '')
        text = _msg_full_text(m)
        # permalink
        permalink = ''
        try:
            pl = c.chat_getPermalink(channel=ch, message_ts=ts)
            permalink = (pl or {}).get('permalink', '') or ''
        except Exception:
            pass
        entry = {'ts': ts, 'summary': _summary_line(text), 'permalink': permalink}
        (deposits if _is_deposit(text) else invoices).append(entry)

    return {'deposits': deposits, 'invoices': invoices,
            'total': len(deposits) + len(invoices)}


def build_pin_remind_text(data: dict) -> str:
    """리마인드 카드 텍스트 조립."""
    deposits = data.get('deposits', [])
    invoices = data.get('invoices', [])
    total = len(deposits) + len(invoices)

    def _line(e: dict) -> str:
        link = f'  |  <{e["permalink"]}|바로가기>' if e.get('permalink') else ''
        return f'• {e["summary"]}{link}'

    lines = [
        _BLANK,
        f':pushpin: *미처리 정산 {total}건 — 확인 부탁드립니다*',
        _SEP,
    ]
    if deposits:
        lines.append(f':moneybag: *입금내역 ({len(deposits)}건)*')
        lines += [_line(e) for e in deposits]
        lines.append('')
    if invoices:
        lines.append(f':receipt: *세금계산서 ({len(invoices)}건)*')
        lines += [_line(e) for e in invoices]
        lines.append('')
    while lines and lines[-1] == '':
        lines.pop()
    lines.append(_SEP)
    lines.append(':information_source: 처리 완료 건은 고정(:pushpin:)에서 내려주세요.')
    lines.append(_BLANK)
    return '\n'.join(lines)


def send_pin_remind() -> dict:
    """미처리 정산 핀 리마인드 발송 — 매일 아침 9시 스케줄러 진입점.

    주말·공휴일 skip (경영지원 근무일만). 고정 0건이면 발송 skip.
    """
    from datetime import date
    try:
        from dashboard.services.absent_remind import _is_business_day
        if not _is_business_day(date.today()):
            return {'ok': True, 'total': 0, 'reason': 'non_business_day'}
    except Exception:
        pass

    data = collect_pending_pins()
    if data is None:
        return {'ok': False, 'total': 0, 'reason': 'collect_failed'}
    if data['total'] == 0:
        logger.info('[PIN] 미처리 정산 0건 — 발송 skip')
        return {'ok': True, 'total': 0, 'reason': None}

    c = _client()
    ch = _channel()
    text = build_pin_remind_text(data)
    try:
        r = c.chat_postMessage(channel=ch, text=text,
                               unfurl_links=False, unfurl_media=False)
        if r.get('ok'):
            logger.info(f"[PIN] 리마인드 발송 완료 (total={data['total']}, ts={r.get('ts')})")
            return {'ok': True, 'total': data['total'], 'ts': r.get('ts', '')}
        return {'ok': False, 'total': data['total'], 'reason': r.get('error')}
    except Exception as exc:
        logger.error(f'[PIN] 리마인드 발송 예외: {exc}', exc_info=True)
        return {'ok': False, 'total': data['total'], 'reason': str(exc)}
