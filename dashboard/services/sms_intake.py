# -*- coding: utf-8 -*-
"""은행 입금 SMS 인입 — 순수 로직 (Flask/Slack/Redis 미접촉, 단위 테스트 대상).

폰(안드로이드 2 + 아이폰 1)의 SMS 포워딩 앱 → POST /sms/inbound 로 들어온
은행 입금 문자를 (1) 잔액 라인 제거 (2) 입금 문자 여부 판별 (3) 미리보기 파싱
하는 순수 함수 모음. 실제 라우트·슬랙 게시·시트 기록은 blueprints/sms_inbound.py.

배경(잔액 제거): 은행 SMS 하단/중간에 통장 잔액이 표시됨.
  - 기업(공백형) : '잔액 144,153,377원'  (입금액 바로 아래)
  - 하나(무공백형): '잔액239,612,486원'   (맨 끝줄)
통장 잔고는 전 직원이 볼 필요가 없으므로 서버 문턱에서 제거한 뒤에만
슬랙·시트로 넘어가게 한다. 폰 쪽 가공에 의존하지 않는다(3대·2 OS 일관성 리스크).
"""

import hashlib
import re

# 잔액/잔고 라인 — 기업(공백) '잔액 144,153,377원' / 하나(무공백) '잔액239,612,486원'.
# 방어적으로 잔고·현재잔액·출금가능·이체후잔액 도 포함. 반드시 '금액(+원)' 이 따라와야
# 삭제하여 거래처명 오삭제를 막는다.
_BALANCE_LINE_RE = re.compile(
    r'^\s*(?:잔액|잔고|현재\s*잔액|출금\s*가능(?:금?액)?|이체\s*후\s*잔액)\s*[:：]?\s*[\d,]+\s*원?\s*$'
)
# 은행 단체문자 머리말 — '[Web발신]', '[국외발신]', '[국제발신]' 등. 줄 앞에서만 제거.
_WEB_HEADER_RE = re.compile(r'^\s*\[[^\]]*발신\]\s*')
# 입금 금액 — '입금 407,000원' / '입금5,115,000원'
_DEPOSIT_RE = re.compile(r'입금\s*[\d,]+\s*원')


def strip_balance(text: str) -> str:
    """은행 SMS에서 **잔액 라인만** 제거. '[Web발신]' 머리말 포함 나머지는 원문 유지.

    잔액(=통장 잔고)만 전 직원 노출 방지로 삭제하고, '[…발신]' 머리말 등은 원문
    그대로 보존한다(사용자 요청 2026-08-13). 파서(_parse_memo_block)가 머리말은
    내부에서 알아서 무시하므로 보존해도 파싱 영향 없음. 앞뒤·연속 빈 줄만 정리.
    """
    if not text:
        return ''
    out = [raw for raw in text.splitlines() if not _BALANCE_LINE_RE.match(raw)]
    cleaned = '\n'.join(out)
    cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)  # 3줄 이상 공백 → 2줄
    return cleaned.strip()


def looks_like_payment(text: str) -> bool:
    """입금 문자 판별 — '입금 X원' 패턴 존재 여부. (광고·인증·택배 문자 배제)"""
    if not text:
        return False
    return bool(_DEPOSIT_RE.search(text))


def strip_web_header(text: str) -> str:
    """'[…발신]' 머리말 라인 제거 — **시트 메모 기록용**(카드 표시엔 유지).

    카드는 실제 포워딩 문자처럼 '[Web발신]'을 보여주되, 시트 노트에는 SB 수동 기록
    관행(머리말 없이 붙여넣음)과 맞춰 머리말을 뺀다. 파싱엔 영향 없음.
    """
    if not text:
        return text
    out = [_WEB_HEADER_RE.sub('', ln) for ln in text.splitlines()]
    return '\n'.join(out).strip()


# 농협 2줄 압축 양식 → 하나 카드처럼 재구성용.
#   원본 : '농협 입금350,000원' + '08/13 09:16 352-****-1682-33 오인석'
#   변환 : '[Web발신]' / '농협, 08/13 09:16' / '352-****-1682-33' / '입금 350,000원' / '오인석'
_NH_AMT_RE = re.compile(r'농협\s*입금\s*([\d,]+)\s*원')
_NH_DETAIL_RE = re.compile(r'(\d{1,2}/\d{1,2})\s+(\d{1,2}:\d{2})\s+(\S+)\s+(.+)')


def normalize_deposit_layout(text: str) -> str:
    """은행별 압축 양식을 하나 카드 양식으로 재구성 (현재 농협만). 실패 시 원문 유지.

    농협 SMS는 '[Web발신]'이 없고 날짜·시각·계좌·입금자가 한 줄에 몰려 기업/하나와
    달라 보인다. 하나('하나,MM/DD, HH:MM')와 동일하게 '[Web발신]' + '농협, MM/DD HH:MM'
    헤더로 재배치. 파서는 '농협,' skip 패턴(payment_sync)으로 거래처 오인을 막는다.
    """
    if not text:
        return text
    m_amt = _NH_AMT_RE.search(text)
    m_det = _NH_DETAIL_RE.search(text)
    if not (m_amt and m_det):
        return text
    amount = m_amt.group(1)
    date, tm, acct, partner = (
        m_det.group(1), m_det.group(2), m_det.group(3), m_det.group(4).strip())
    return f"[Web발신]\n농협, {date} {tm}\n{acct}\n입금 {amount}원\n{partner}"


# ITG 사업자 통장 3종 — 개인 계좌 입금 배제용(같은 은행이라도 개인 건은 계좌 tail 이 다름).
#   기업 452***38801011 / 하나 255******31304 / 농협 352-****-1682-33
_BIZ_ACCT_RES = [
    re.compile(r'452\*+38801011'),
    re.compile(r'255\*+31304'),
    re.compile(r'352-\*+-1682-33'),
]


def has_business_account(text: str) -> bool:
    """ITG 사업자 통장 계좌번호(마스킹) 포함 여부 — 개인 입금(다른 계좌) 배제용."""
    return any(p.search(text or '') for p in _BIZ_ACCT_RES)


# 은행 예금 이자·결산 입금 — 프로젝트 입금 아님(적요 '2026년결산' 등). 인입 제외.
# '이자'는 '이자카야' 등 상호 오탐 방지로 단독 제외 — '예금이자'·'결산'만.
_BANK_INTEREST_RE = re.compile(r'결산|예금\s*이자|이자\s*입금')


def is_bank_interest(text: str) -> bool:
    """은행 예금 이자/결산 입금 여부 — 프로젝트 입금이 아니므로 인입 카드 생성 제외."""
    return bool(_BANK_INTEREST_RE.search(text or ''))


def dedup_hash(sender: str, text: str) -> str:
    """중복 판별 키 — 문자 본문 기준(정규화). 앞 16자.

    같은 은행 문자를 여러 폰(YG·JW·SB)이 동시에 포워딩 → 본문은 동일하므로
    앱/플랫폼별 공백·'[…발신]' 머리말 차이만 정규화하면 같은 키 → 첫 1건만 카드,
    나머지는 duplicate 로 skip. 발신번호는 앱마다 표기가 달라 키에서 제외해도
    본문에 시각·금액·잔액이 있어 서로 다른 입금끼리는 이미 충분히 고유하다.
    (sender 인자는 호환성 위해 유지하되 키에는 미사용)
    """
    norm = re.sub(r'\[[^\]]*발신\]', ' ', text or '')   # '[Web발신]' 등 머리말 제거
    norm = re.sub(r'\s+', ' ', norm).strip()            # 공백 정규화
    return hashlib.md5(norm.encode('utf-8')).hexdigest()[:16]


def _normalize_sms_display(text: str) -> str:
    """표시 전용 정규화 — 시트에 기록되는 원문엔 영향 없음.

    - '입금5,000원' → '입금 5,000원' (하나 무공백 양식 가독성)
    - 마스킹 별표(*) → ∗ (slack 마크다운 오해석 방지)
    """
    s = re.sub(r'(입금)\s*([\d,])', r'\1 \2', (text or '').strip())
    return s.replace('*', '∗')


def active_display(text: str) -> str:
    """활성(미완료) 카드용 SMS 표시 — 회색 코드블록 대신 인용(blockquote).

    완료 카드(회색 코드블록)와 시각적으로 구분해 '아직 처리 전'임을 보이게 한다.
    """
    s = _normalize_sms_display(text)
    return f">>> {s}" if s else ""


# 인입 카드 3종(활성·확인대기·완료) 공통 구분선 — 입금 문자 폭에 맞춘 25자.
INTAKE_SEP = '-' * 25


def quoted_body(text: str) -> list:
    """SMS 본문 → '>' 인용 라인 리스트 (표시 정규화 적용).

    카드 3종 공통 레이아웃 — 헤더·구분선·본문을 한 섹션에 전부 '>' 로 감싸
    섹션 간 여백 없이 온라인/방문 카드와 동일하게 보이게 한다.
    """
    norm = _normalize_sms_display(text)
    return [f">{ln}" for ln in norm.split('\n')]


def parse_preview(stripped_text: str) -> dict:
    """미리보기용 best-effort 파싱 (금액/은행/입금자/일시).

    실패해도 카드는 원문을 그대로 보여주므로 예외는 삼키고 빈 dict 반환.
    기존 payment_sync 파서를 재사용한다.
    """
    try:
        from dashboard.services.payment_sync import _parse_memo_block
        blk = _parse_memo_block(stripped_text) or {}
        return {
            'amount': blk.get('amount') or 0,
            'partner': (blk.get('partner') or '').strip(),
            'bank': (blk.get('bank') or '').strip(),
            'date_md': (blk.get('date_md') or '').strip(),
        }
    except Exception:
        return {}
