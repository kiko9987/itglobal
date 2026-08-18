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


# 세금계산서(지출) 매칭 요청 — "MM/DD G/R 금액원 거래처 …" (샛별 표준 양식)
# 이미지 유무 무관. '계약서 진행 여부'·'계산서 미발행건'·정정요청 등은 미매치로 제외.
_INVOICE_RE = re.compile(r'^\s*\d{1,2}/\d{1,2}\s+[GRNgrn]\s+[\d,]+\s*원\s+\S')


# ITG 통장 계좌 3종 (마스킹 * 필수 — 전화번호 등 오탐 방지). 기업452=G / 하나255=R / 농협352=N.
_DEPOSIT_ACCTS = [
    (re.compile(r'452[\d\-]*\*[\d*\-]*'), 'G'),  # 452***38801011
    (re.compile(r'255[\d\-]*\*[\d*\-]*'), 'R'),  # 255***31304
    (re.compile(r'352[\d\-]*\*[\d*\-]*'), 'N'),  # 352-****-1682-33
]


def _deposit_grn(text: str) -> str:
    """입금 계좌 → G/R/N (없으면 '')."""
    for pat, code in _DEPOSIT_ACCTS:
        if pat.search(text or ''):
            return code
    return ''


def is_deposit(text: str) -> bool:
    """ITG 통장 계좌번호(452 기업 / 255 하나 / 352 농협) 포함 → 입금내역."""
    return bool(_deposit_grn(text))


def is_invoice_request(text: str) -> bool:
    """세금계산서 매칭 요청 양식(MM/DD G/R 금액원 거래처) → 세금계산서."""
    return bool(_INVOICE_RE.match(text or ''))


def is_settlement_message(text: str) -> bool:
    """자동 고정 대상 — 입금내역 or 세금계산서 매칭 요청."""
    return is_deposit(text) or is_invoice_request(text)


# 하위 호환 alias
_is_deposit = is_deposit


def _summary_line(text: str) -> str:
    """한 줄 요약 — 계좌번호·마크다운·개행·꼬리 문구 제거."""
    s = text.replace('[Web발신]', '')
    s = re.sub(r'(?:452|255)[\d*\-]{5,}', '', s)   # 계좌번호 제거 (마스킹 깨짐·노이즈 방지)
    s = re.sub(r'[*_]', '', s)                       # bold 마크다운 제거
    s = re.sub(r'\s*\n\s*', ' ', s).strip()
    s = re.sub(r'\s*회신\s*부탁드립니다\.?\s*$', '', s).strip()
    s = re.sub(r'\s{2,}', ' ', s)
    return s[:80] or '(내용 없음)'


# 은행 SMS 필드 라벨 — 하나 라벨형('일시 …/계좌번호 …/적요 거래처')에서 거래처에
# 섞여 누출되던 단어들 (2026-08-12 백테스트: 146건 중 44건 누출 확인 → 제거)
_SMS_LABEL_RE = re.compile(r'일시|적요|계좌번호')


def _format_deposit_summary(text: str) -> str:
    """입금내역 → 'MM/DD G/R/N 금액원 거래처' (수금관리 표기와 통일).

    2026-08-12: 검증된 파서(_parse_memo_block)를 재사용해 하나 라벨형
    (일시/계좌번호/적요)에서도 거래처만 깔끔히 추출. 파싱 실패 시 기존 regex 방식으로
    fallback(라벨 단어 제거 추가). G/R/N 은 계좌번호 기반 판정 유지(신뢰도 높음).
    """
    grn = _deposit_grn(text)  # 원본(마스킹 *)에서 G/R/N 판정
    # 1) 검증된 파서 우선 — 라벨형 포함 대부분 정상 추출
    try:
        from dashboard.services.sms_intake import strip_balance
        from dashboard.services.payment_sync import _parse_memo_block
        res = _parse_memo_block(strip_balance(text)) or {}
    except Exception:
        res = {}
    amount = res.get('amount') or 0
    partner = (res.get('partner') or '').strip()
    date = (res.get('date_md') or '').strip()
    if amount > 0 and partner and partner != '-' and not _SMS_LABEL_RE.search(partner):
        parts = [x for x in [date, grn, f'{amount:,}원', partner] if x]
        return ' '.join(parts)

    # 2) fallback — 기존 regex 방식 (라벨 단어 제거 추가)
    s = re.sub(r'[*]', '', text).replace('[Web발신]', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    am = re.search(r'입금\s*([\d,]+)\s*원', s)
    amount_s = am.group(1) if am else ''
    dm = re.search(r'(?:\d{4}[/.])?(\d{1,2})[/.](\d{1,2})', s)
    date_s = f'{int(dm.group(1)):02d}/{int(dm.group(2)):02d}' if dm else ''
    p = re.sub(r'입금\s*[\d,]+\s*원', ' ', s)
    p = re.sub(r'(?:\d{4}[/.])?\d{1,2}[/.]\d{1,2}', ' ', p)
    p = re.sub(r'\d{1,2}:\d{2}', ' ', p)
    p = re.sub(r'\d{3}[\d*\-]{4,}', ' ', p)   # 계좌번호
    p = re.sub(r'(기업|하나|국민|신한|우리|농협|카카오|토스|SC|씨티)', ' ', p)
    p = _SMS_LABEL_RE.sub(' ', p)             # 라벨 단어 제거 (일시/적요/계좌번호)
    p = p.replace('입금', ' ')
    partner_s = re.sub(r'\s+', ' ', p).strip(' -·,')[:30]
    parts = [x for x in [date_s, grn, (amount_s + '원' if amount_s else ''), partner_s] if x]
    return ' '.join(parts) or _summary_line(text)


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
    others: List[dict] = []   # 계좌·양식 미매치 수동 핀 (계약서 등)
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
        if is_deposit(text):
            summary = _format_deposit_summary(text)
            deposits.append({'ts': ts, 'summary': summary, 'permalink': permalink,
                             'key': _deposit_key(text)})
        elif is_invoice_request(text):
            invoices.append({'ts': ts, 'summary': _summary_line(text), 'permalink': permalink})
        else:
            others.append({'ts': ts, 'summary': _summary_line(text), 'permalink': permalink})

    return {'deposits': deposits, 'invoices': invoices, 'others': others,
            'total': len(deposits) + len(invoices) + len(others)}


def _payment_client():
    from slack_sdk import WebClient
    tok = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    return WebClient(token=tok) if tok else None


def _payment_channel() -> str:
    return os.getenv('SLACK_PAYMENT_INTAKE_CHANNEL', '').strip()


def _fmt_deposit_line(pv: dict) -> str:
    """preview(dict) → 'MM/DD G/R/N 금액원 거래처'."""
    grn = {'기업': 'G', '하나': 'R', '농협': 'N'}.get(pv.get('bank', ''), '')
    parts = [x for x in [
        pv.get('date_md', ''), grn,
        (f"{pv['amount']:,}원" if pv.get('amount') else ''),
        pv.get('partner', ''),
    ] if x]
    return ' '.join(parts) if parts else '(입금 내역)'


def _deposit_key(text: str) -> str:
    """입금 중복 판별 키 '금액|거래처6자' — #영업_관리↔#입금_관리 교차 dedup용.

    검증된 파서(_parse_memo_block)로 금액·거래처 추출해 채널 무관 동일 키 생성.
    양쪽 문자 표기가 조금 달라도 금액이 강한 식별자라 매칭된다.
    """
    try:
        from dashboard.services.sms_intake import strip_balance
        from dashboard.services.payment_sync import _parse_memo_block
        r = _parse_memo_block(strip_balance(text)) or {}
        amt = r.get('amount') or 0
        partner = re.sub(r'\s+', '', (r.get('partner') or ''))[:6]
        if amt:
            return f'{amt}|{partner}'
    except Exception:
        pass
    return ''


def _oam_deposit_keys() -> set:
    """#영업_관리 최근 입금 메시지 키 집합 — 겸용 중복 제거용.

    거기 올라온 입금(고정=처리 예정 / ✅ 체크=처리 완료 — 둘 다 이력에 남음)은
    #입금_관리 인입 쪽에서 리마인드 제외. 조회 실패(scope 등) 시 빈 set → dedup skip.
    """
    c = _client()
    ch = _channel()
    if not c or not ch:
        return set()
    try:
        resp = c.conversations_history(channel=ch, limit=200)
    except Exception as exc:
        logger.warning(f'[PIN] 영업관리 history 조회 실패(겸용 dedup skip): {exc}')
        return set()
    keys = set()
    for m in resp.get('messages', []) or []:
        text = _msg_full_text(m)
        if is_deposit(text):
            k = _deposit_key(text)
            if k:
                keys.add(k)
    return keys


def collect_intake_pending() -> List[dict]:
    """#입금_관리 고정(미처리) 인입 카드 조회 → 요약 리스트 (미지정+확인대기).

    수금봇 토큰(pins:read)으로 #입금_관리 pins 조회. 인입 카드만(버튼 action_id 판별)
    골라 상태(미지정=payment_intake_open / 확인대기=payment_intake_confirm) 표기.
    Returns [{ts, summary, permalink, state}] (client/channel 미설정·실패 시 []).
    """
    c = _payment_client()
    ch = _payment_channel()
    if not c or not ch:
        return []
    try:
        resp = c.pins_list(channel=ch)
    except Exception as exc:
        logger.warning(f'[PIN] 입금관리 pins.list 실패: {exc}')
        return []
    out: List[dict] = []
    for it in resp.get('items', []) or []:
        m = it.get('message', {})
        if not m:
            continue
        aid = intake_id = None
        for b in m.get('blocks', []) or []:
            if b.get('type') != 'actions':
                continue
            for e in b.get('elements', []):
                if e.get('action_id') in ('payment_intake_open', 'payment_intake_confirm'):
                    aid, intake_id = e.get('action_id'), e.get('value')
                    break
        if not aid:
            continue   # 인입 카드 아님(수동 핀 등) 제외
        ts = m.get('ts', '')
        # Redis 1회 로드 → 요약 + 겸용 dedup 키
        summary, key = '(입금 내역)', ''
        try:
            import json
            from dashboard.utils.redis_client import get_redis_client
            raw = get_redis_client().redis.get(f'sms_intake:{intake_id}')
            if raw:
                d = json.loads(raw)
                summary = _fmt_deposit_line(d.get('preview') or {})
                key = _deposit_key(d.get('text') or '')
        except Exception:
            pass
        permalink = ''
        try:
            permalink = (c.chat_getPermalink(channel=ch, message_ts=ts) or {}).get('permalink', '') or ''
        except Exception:
            pass
        out.append({
            'ts': ts, 'summary': summary, 'key': key, 'permalink': permalink,
            'state': '미지정' if aid == 'payment_intake_open' else '확인대기',
        })
    return out


def build_pin_remind_text(data: dict) -> str:
    """리마인드 카드 텍스트 조립."""
    deposits = data.get('deposits', [])
    invoices = data.get('invoices', [])
    others = data.get('others', [])
    intakes = data.get('intakes', [])
    total = len(deposits) + len(invoices) + len(others) + len(intakes)

    def _line(e: dict) -> str:
        link = f'  |  <{e["permalink"]}|바로가기>' if e.get('permalink') else ''
        return f'• {e["summary"]}{link}'

    def _section(header: str, items: list) -> str:
        return '\n'.join([header] + [_line(e) for e in items])

    def _intake_line(e: dict) -> str:
        link = f'  |  <{e["permalink"]}|바로가기>' if e.get('permalink') else ''
        badge = ':hourglass_flowing_sand:' if e.get('state') == '확인대기' else ':link:'
        return f'• {badge} {e["summary"]}  _{e.get("state", "")}_{link}'

    sections = []
    if intakes:
        # #입금_관리 미처리 인입 (프로젝트 지정/확인 필요) — 다른 채널이라 바로가기 링크로
        sections.append('\n'.join(
            [f':inbox_tray: *미처리 입금 — 지정/확인 필요 ({len(intakes)}건)*']
            + [_intake_line(e) for e in intakes]))
    if deposits:
        sections.append(_section(f':moneybag: *입금내역 ({len(deposits)}건)*', deposits))
    if invoices:
        sections.append(_section(f':receipt: *세금계산서 ({len(invoices)}건)*', invoices))
    if others:
        sections.append(_section(f':pushpin: *기타 ({len(others)}건)*', others))

    body = f'\n{_BLANK}\n'.join(sections)   # 섹션 간 빈 줄 한 개
    return (
        f'{_BLANK}\n'
        f':pushpin: *미처리 정산·입금 {total}건 — 확인 후 처리(지정/댓글) 부탁드립니다*\n'
        f'{_SEP}\n'
        f'{body}\n'
        f'{_SEP}\n'
        f'{_BLANK}'
    )


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
    # #입금_관리 미처리 인입 카드 합류 (다른 채널 — 수금봇 토큰으로 조회)
    intakes = collect_intake_pending()
    # 겸용 중복 제거 — 인입(#입금_관리)이 있는 입금은 #영업_관리 입금내역에서 제거해
    # **인입 링크 우선**(#입금_관리로 연결). 앞으로 입금=#입금_관리 일원화 방침 (2026-08-18).
    intake_keys = {i.get('key') for i in intakes if i.get('key')}
    if intake_keys:
        data['deposits'] = [d for d in data.get('deposits', []) if d.get('key') not in intake_keys]
    data['intakes'] = intakes
    data['total'] = (len(data.get('deposits', [])) + len(data.get('invoices', []))
                     + len(data.get('others', [])) + len(intakes))
    if data['total'] == 0:
        logger.info('[PIN] 미처리 정산·입금 0건 — 발송 skip')
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
