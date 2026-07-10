"""수금 관리 자동 알림 — 공사 현황 시트의 U/V/W 입금 메모 변경 감지 → #수금_관리 채널 발송.

흐름:
1. 시트 폴링 (값만 가벼움): R/S/T/U/V/W/X/Y/Z/AA + A/F
2. Redis에 이전 (U, V, W, AA) 상태 저장 후 변경 감지
3. 변경된 행만 셀 노트(U/V/W) 추가 fetch
4. 노트 파싱 → 누적 history 메시지 빌더 → 슬랙 발송

메모 파싱 — 4가지 양식 (사용자 데이터 기준):
  양식 A: yyyy/mm/dd HH:MM / 입금 X원 / 거래처 / 계좌 / 은행
  양식 B: 일시 mm/dd, HH:MM / 입금 X원 / 계좌 / 적요 거래처
  양식 C: 하나,mm/dd, HH:MM / 계좌 / 입금X원 / 거래처
  양식 D: yyyy/mm/dd HH:MM / 입금 X원 / 거래처 / 계좌 / 은행 (양식 A와 동일)
빈 줄로 블록 구분, 한 블록 = 한 입금.

표기 매핑:
  Y열='N입금'  → 'N' (현금)
  Y열='카드결제' or 그 외 + 메모 은행 = '기업' → 'G' (글로벌)
                                  + 메모 은행 = '하나' → 'R' (글로벌그룹)
"""

import hashlib
import os
import re
from typing import List, Dict, Optional

from dashboard.utils.logging_config import get_logger
from dashboard.utils.redis_client import get_redis_client

logger = get_logger(__name__)

# 시트 컬럼 (공사 현황 시트 1-based 컬럼)
COL_PROJECT = 'A'    # 프로젝트 코드-이니셜 (예: G3491-YG)
COL_ADDRESS = 'F'    # 현장 주소
COL_TOTAL_R = 'R'    # 총액 1 (부가세 미포함)
COL_VAT = 'S'        # 부가세 체크박스
COL_TOTAL_T = 'T'    # 총액 2 (R + VAT)
COL_DEPOSIT = 'U'    # 계약금
COL_MIDDLE = 'V'     # 중도금
COL_BALANCE = 'W'    # 잔금
COL_UNPAID = 'X'     # 미수금
COL_INVOICE = 'Y'    # 계산서 (N입금/카드결제/미발행/잔금/중도금/계약금)
COL_PAYDATE = 'Z'    # 수금 날짜
COL_CONFIRM = 'AA'   # 수금 확인 체크박스

REDIS_KEY_PREFIX = 'payment_sync:row:'
REDIS_TTL = 60 * 60 * 24 * 90  # 90일

# 정식 프로젝트 코드 패턴 — 예: G3491-YG, R3625-JK
# 비표준(옛 프로젝트 1035, "0 중고...") 행은 알림 skip
_VALID_PROJECT_RE = re.compile(r'^[GR]\d{4}-[A-Z]+$')

# ─────────────────────────────────────────────
# 메모 파싱
# ─────────────────────────────────────────────

# 금액 추출: "입금 12,100,000원" or "금액 : 12,100,000원" 등
_AMOUNT_RE = re.compile(r'(?:입금|금액\s*[:：])\s*([\d,]+)\s*원')
# 한글 금액: "272만원", "272만 원", "1,700만원" 등 → 만 단위 (2026-07-10)
_KO_AMOUNT_RE = re.compile(r'([\d,]+)\s*만\s*원')
# 한글 날짜: "6월15일", "6월 15일" (연도 없으면 현재 년도) (2026-07-10)
_KO_DATE_RE = re.compile(r'(\d{1,2})\s*월\s*(\d{1,2})\s*일')
# 프로젝트 코드 패턴 — partner 파싱 시 skip 대상 (2026-07-10)
_PROJECT_CODE_RE = re.compile(r'^[GRN]\d{4}-[A-Z]{1,3}$')
# 매니저 이체 표기 — "2026-07-10 R>g", "2026-07-10 G>N" (2026-07-10)
# 첫 번째 문자 = 원 계좌, 두 번째 문자 = 이체 후 계좌 (대소문자 무관)
_TRANSFER_RE = re.compile(r'\d{4}-\d{1,2}-\d{1,2}\s+([GRNP])\s*>\s*([grnpGRNP])')
# 라벨 양식 (앱스크립트 자동 + 매니저 수기 덮어쓰기)
_LABEL_DATE_RE = re.compile(r'^입금일\s*:\s*(.+)$')
_LABEL_PAYER_RE = re.compile(r'^입금자\s*:\s*(.+)$')
# 매니저 카톡 양식 (분할 입금 history를 V 셀에 그대로 복사한 케이스)
# 예: "02/02 G 22,308,000원 프레임플러스"
_KATOK_LINE_RE = re.compile(
    r'^(\d{1,2})/(\d{1,2})\s+([GRN]|현금|박C기업)\s+([\d,]+)\s*원\s*(.*?)\s*(?:수금중|수금완료)?\s*$'
)
# 단일 숫자(콤마 OK) 라인 — 매니저가 입금액만 적은 경우
_BARE_AMOUNT_RE = re.compile(r'^([\d,]+)(?:\s*원)?$')
# 날짜: "2026/03/19", "06/25", "5/27" 등 — MM/DD 추출
_DATE_RE = re.compile(r'(?:(\d{4})[/.-])?(\d{1,2})[/.-](\d{1,2})')
# 은행명 추출 — 라인 또는 첫줄 시작
_BANK_RE = re.compile(r'(기업|하나|국민|신한|우리|농협|카카오|토스)')
# ITG 통장 계좌번호 — 카드 결제 시 카드사 약자보다 우선
# 기업 452-039388-01-011 (글로벌 G), 하나 255-910014-31304 (글로벌그룹 R)
_ACCT_G_RE = re.compile(r'452[\*\-]+0?3?9?\d*[\*\-]+\d')  # 기업 452***38801011
_ACCT_R_RE = re.compile(r'255[\*\-]+9?1?0?0?\d*[\*\-]+\d')  # 하나 255******31304


def _parse_memo_block(block: str, fallback_amount: int = 0) -> Optional[Dict]:
    """입금 메모 한 블록 → {'date_md': 'MM/DD', 'amount': int, 'partner': str, 'bank': str}
    fallback_amount: 메모에 금액 라인 없으면 시트 단계 값 사용 (옛 양식 처리).
    """
    lines = [l.strip() for l in block.splitlines() if l.strip()]
    if not lines:
        return None
    # 은행 단체 문자 머리말 제거 (예: '[Web발신]')
    lines = [l for l in lines if not l.startswith('[Web발신]')]
    # 이메일 답장 prefix 제거 ('RE:하나,...' → '하나,...')
    lines = [re.sub(r'^RE:', '', l) for l in lines]
    # 매출이동 표기 라인 제거 ('2026-02-12 R>N 매출이동')
    lines = [l for l in lines if '매출이동' not in l]
    if not lines:
        return None

    # 날짜 — 모든 라인에서 첫 매치 (한글 날짜 우선, 없으면 숫자 형식)
    date_md = ''
    for ln in lines:
        m = _KO_DATE_RE.search(ln)  # "6월15일", "6월 15일"
        if m:
            mm, dd = int(m.group(1)), int(m.group(2))
            date_md = f"{mm:02d}/{dd:02d}"
            break
        m = _DATE_RE.search(ln)
        if m:
            mm, dd = int(m.group(2)), int(m.group(3))
            date_md = f"{mm:02d}/{dd:02d}"
            break

    # 금액
    amount = 0
    for ln in lines:
        m = _AMOUNT_RE.search(ln)
        if m:
            amount = int(m.group(1).replace(',', ''))
            break
    # "입금 X원" 패턴 없으면 단일 숫자 라인에서 추출 (매니저 수기 양식)
    if amount == 0:
        for ln in lines:
            m = _BARE_AMOUNT_RE.match(ln)
            if m:
                amount = int(m.group(1).replace(',', ''))
                break
    # 한글 금액 (272만원, 1,700만원 등) — "만원" 단위, 만 배수
    if amount == 0:
        for ln in lines:
            m = _KO_AMOUNT_RE.search(ln)
            if m:
                amount = int(m.group(1).replace(',', '')) * 10000
                break
    # 옛 양식 fallback — 메모에 금액 없으면 시트 단계 값 사용
    if amount == 0 and fallback_amount > 0:
        amount = fallback_amount
    if amount == 0:
        return None  # 금액 + fallback 모두 없으면 입금 블록 아님
    # 통합 입금 분담 케이스: 메모 amount가 시트 단계값보다 크면
    # → 매니저가 큰 통합 입금 메모를 여러 행에 복사한 것 → 시트 단계값(실제 부담분) 사용
    # 카드 결제(메모 amount < 시트값) 케이스는 보존
    if fallback_amount > 0 and amount > fallback_amount:
        amount = fallback_amount

    # 매니저 수기 양식 — "입금일: YYYY-MM-DD" / "입금자: 현금" 등
    label_payer = ''
    for ln in lines:
        m = _LABEL_DATE_RE.match(ln)
        if m:
            datestr = m.group(1).strip()
            md = _DATE_RE.search(datestr)
            if md and not date_md:
                mm, dd = int(md.group(2)), int(md.group(3))
                date_md = f"{mm:02d}/{dd:02d}"
        m = _LABEL_PAYER_RE.match(ln)
        if m:
            label_payer = m.group(1).strip()

    # 은행 — ITG 통장 계좌번호 우선 (카드사 약자보다 정확)
    bank = ''
    text_join = '\n'.join(lines)
    if _ACCT_G_RE.search(text_join):
        bank = '기업'
    elif _ACCT_R_RE.search(text_join):
        bank = '하나'
    else:
        # 계좌 마스킹 없으면 메모 텍스트에서 은행명 추출
        for ln in lines:
            m = _BANK_RE.search(ln)
            if m:
                bank = m.group(1)
                break

    # 거래처(partner) — 휴리스틱
    # 양식 B: "적요 (주)..." → "적요 " 제거
    # 양식 C(하나): 첫 줄 "하나,..." 이후 — 입금 줄 외에 거래처
    # 양식 A: 입금 다음 줄, 또는 은행/계좌 외 줄
    # 양식 E(농협): 한 라인에 날짜+시간+계좌+거래처 다 들어감 → 토큰 단위 분리 필요
    partner = label_payer  # 매니저 수기 라벨 있으면 우선
    skip_patterns = [
        re.compile(r'^\d{4}[/.-]\d{1,2}[/.-]\d{1,2}'),  # 날짜
        re.compile(r'^일시\s'),
        re.compile(r'^하나[,\s]'),
        re.compile(r'^입금'),
        re.compile(r'^입금일\s*:'),
        re.compile(r'^입금자\s*:'),
        re.compile(r'^계좌번호'),
        re.compile(r'^\d{3}[*\d]+$'),  # 계좌번호 (전체 숫자/* — 끝 문자 있으면 거래처)
        re.compile(r'^[\d,]+(?:\s*원)?$'),  # 단일 숫자 라인 (이미 amount로 사용)
        # 은행명 단독 라인 또는 "은행 입금X원" 라인만 skip — "하나90242344" 같은 카드 승인번호 보존
        re.compile(r'^(기업|하나|국민|신한|우리|농협|카카오|토스)(?:\s+입금|\s*$)'),
    ]
    # 매니저 수기 요약 라인 skip 패턴 (2026-07-10)
    #   "수령완료", "수금완료입니다", "입금완료", "6월15일 272만원 현금 YG 수령완료" 등
    manager_summary_re = re.compile(r'(수령완료|수금완료|입금완료|정산완료)')

    if not partner:
        for ln in lines:
            if any(p.search(ln) for p in skip_patterns):
                continue
            # 프로젝트 코드 라인 skip (2026-07-10 — 매니저가 노트 첫줄에 코드 붙여넣는 케이스)
            if _PROJECT_CODE_RE.match(ln.strip()):
                continue
            # 매니저 요약 라인 skip (2026-07-10)
            if manager_summary_re.search(ln):
                continue
            cleaned = ln
            cleaned = re.sub(r'\d{1,2}/\d{1,2}\b', '', cleaned)
            cleaned = re.sub(r'\d{1,2}:\d{2}\b', '', cleaned)
            cleaned = re.sub(r'\d{2,}[\*\-][\d\*\-]+', '', cleaned)
            cleaned = re.sub(r'^적요\s*', '', cleaned)
            cleaned = re.sub(r'\s+', ' ', cleaned).strip()
            if cleaned and not re.match(r'^[\d*\-,.\s]+$', cleaned):
                partner = cleaned
                break

    # 노트 전체에 "현금" 키워드 있고 partner 아직 없으면 partner='현금' fallback (2026-07-10)
    # → _resolve_payment_code 에서 'N' 코드 반환
    if not partner and '현금' in '\n'.join(lines):
        partner = '현금'

    # 박C 표기 — 대표님 개인 기업통장(추적용 구분)
    note_label = ''
    if '박C' in block:
        note_label = '박C'

    # 이체 표기 (2026-07-10) — "2026-07-10 R>g" 등
    transfer_to = ''
    for ln in lines:
        m = _TRANSFER_RE.search(ln)
        if m:
            transfer_to = m.group(2).upper()
            break

    return {
        'date_md': date_md,
        'amount': amount,
        'partner': partner,
        'bank': bank,
        'note_label': note_label,
        'transfer_to': transfer_to,
    }


def _hash_payments(payments: List[Dict]) -> str:
    """payments 리스트 → 짧은 hash. 메모 변경 감지용."""
    if not payments:
        return ''
    s = '|'.join(
        f"{p.get('stage','')}:{p.get('date_md','')}:{p.get('amount',0)}:{p.get('partner','')}"
        for p in payments
    )
    return hashlib.md5(s.encode('utf-8')).hexdigest()[:16]


def _parse_notes(notes: List[str],
                 stage_vals: Optional[Dict[str, int]] = None) -> List[Dict]:
    """U/V/W 셀 노트 3개를 모두 파싱 → [{stage:'계약금', ...}, {stage:'중도금', ...}, ...]
    파싱 실패한 블록 또는 partner 빈 블록은 직전 블록과 합쳐서 재시도
    (매니저가 한 입금 안에서 빈 줄 입력해서 잘못 분리되는 케이스 보정).
    stage_vals: 단계별 시트 값(U/V/W) — 옛 양식(금액 라인 없음) 처리용 fallback.
    """
    results = []
    stages = ['계약금', '중도금', '잔금']
    for note, stage in zip(notes, stages):
        stage_val = (stage_vals or {}).get(stage, 0)
        if not note:
            # 메모 없는데 단계 값 있으면 fallback payment 추가 (옛 데이터)
            if stage_val > 0:
                results.append({
                    'date_md': '-', 'amount': stage_val, 'partner': '-',
                    'bank': '', 'note_label': '', 'stage': stage,
                })
            continue

        # 매니저 카톡 양식 우선 검사 — 분할 입금 history를 셀에 복사한 케이스
        # 예: "02/02 G 22,308,000원 프레임플러스" 같은 라인이 여러 개
        katok_payments = []
        for ln in note.strip().splitlines():
            m = _KATOK_LINE_RE.match(ln.strip())
            if m:
                mm, dd = int(m.group(1)), int(m.group(2))
                amount = int(m.group(4).replace(',', ''))
                partner = m.group(5).strip()
                # 은행 추정 — 코드만 보고 한국 ITG 기본
                bank = '기업' if m.group(3) in ('G', '현금', '박C기업') else (
                    '하나' if m.group(3) == 'R' else ''
                )
                katok_payments.append({
                    'date_md': f'{mm:02d}/{dd:02d}',
                    'amount': amount,
                    'partner': partner or '-',
                    'bank': bank,
                    'note_label': '박C' if m.group(3) == '박C기업' else '',
                    'stage': stage,
                })
        # 카톡 양식 라인이 1건이고 (실결제/수수료) 같은 카드 정보 라인이 함께 있으면 카드 결제 케이스
        # 예: "06/18 G 1,394,400원 비씨카드\n(실결제 1,400,000원 수수료 5,600원)"
        if len(katok_payments) == 1 and re.search(r'\(.*실결제.*수수료.*\)', note):
            results.extend(katok_payments)
            continue
        # 카톡 양식 라인이 2건 이상이면 카톡 양식으로 처리 (정상 양식과 명확히 다름)
        if len(katok_payments) >= 2:
            # 매니저가 history 누적 — 이전 단계와 중복되는 첫 라인들 제외
            # 예: V 메모의 첫 라인이 U값과 동일하면 계약금 입금 중복
            prev_amounts = {p.get('amount') for p in results}
            while katok_payments and katok_payments[0].get('amount') in prev_amounts:
                katok_payments.pop(0)
            results.extend(katok_payments)
            continue
        # 1차: 빈 줄로 블록 분리
        raw_blocks = re.split(r'\n\s*\n', note.strip())
        # 2차: 한 블록 안에 날짜 패턴이 2번 이상이면 추가 분리 (빈 줄 없는 분할 입금)
        blocks = []
        for rb in raw_blocks:
            # 날짜 패턴 (YYYY/MM/DD HH:MM) 위치 찾기
            date_starts = [m.start() for m in re.finditer(r'\d{4}/\d{1,2}/\d{1,2}\s+\d{1,2}:\d{2}', rb)]
            if len(date_starts) >= 2:
                # 각 날짜 시작점으로 분리
                for i, ds in enumerate(date_starts):
                    end = date_starts[i+1] if i+1 < len(date_starts) else len(rb)
                    blocks.append(rb[ds:end].strip())
            else:
                blocks.append(rb)
        prev_block = ''
        block_count = len(blocks)
        any_parsed = False
        for bidx, block in enumerate(blocks):
            # 블록 하나뿐이고 옛 양식이면 fallback amount 적용
            fb = stage_val if block_count == 1 else 0
            parsed = _parse_memo_block(block, fallback_amount=fb)
            if parsed:
                any_parsed = True
            if not parsed and prev_block:
                # 파싱 실패 — 직전 블록과 합쳐 재시도
                merged = prev_block + '\n' + block
                reparse = _parse_memo_block(merged, fallback_amount=fb)
                if reparse and results and results[-1].get('stage') == stage:
                    # 직전 결과 대체 (partner 등 보강된 정보)
                    reparse['stage'] = stage
                    results[-1] = reparse
                    prev_block = merged
                    continue
            if parsed:
                # 직전 결과가 같은 단계 + partner 비어있으면 합쳐서 재시도
                if (results and results[-1].get('stage') == stage
                        and not results[-1].get('partner') and prev_block):
                    merged = prev_block + '\n' + block
                    reparse = _parse_memo_block(merged)
                    if (reparse and reparse.get('partner')
                            and reparse['amount'] == results[-1]['amount']):
                        reparse['stage'] = stage
                        results[-1] = reparse
                        prev_block = merged
                        continue
                parsed['stage'] = stage
                results.append(parsed)
                prev_block = block
        # 메모 있지만 어떤 블록도 파싱 실패 + 단계 값 있으면 fallback
        if not any_parsed and stage_val > 0:
            results.append({
                'date_md': '-', 'amount': stage_val, 'partner': '-',
                'bank': '', 'note_label': '', 'stage': stage,
            })
    # 날짜순 정렬 (월/일 기준) — 양식에 연도 없으면 동일 연도 가정
    def _date_key(p):
        mm_dd = p.get('date_md', '0/0')
        try:
            mm, dd = mm_dd.split('/')
            return (int(mm), int(dd))
        except Exception:
            return (0, 0)

    # 같은 단계 안에서만 날짜순 정렬 (단계 순서는 계약금→중도금→잔금 유지)
    stage_order = {'계약금': 0, '중도금': 1, '잔금': 2}
    results.sort(key=lambda p: (stage_order.get(p.get('stage', ''), 9), _date_key(p)))
    return results


# ─────────────────────────────────────────────
# 한 글자 표시 (G/N/R) 결정
# ─────────────────────────────────────────────

# 카드 승인번호 — 영문 또는 숫자 필수 (한글만 단독은 카드 아님)
# 예: "하나90242344" / "NH15415440" / "745389850B" / "현108017094" → 카드
#     "미사역파라곤아파" / "프레임플러스" → 일반 입금자
_CARD_PARTNER_RE = re.compile(r'^(?=.*[A-Za-z0-9])[0-9A-Z가-힣]{6,18}$')


_CARD_BRAND_RE = re.compile(
    r'(?:비씨|BC|삼성|현대|롯데|신한|하나|국민|KB|NH|SH|SHC|우리|씨티|카카오)\s*카드'
)


def _is_card_payment(invoice_value: str, partner: str) -> bool:
    """Y열 + 거래처 패턴으로 카드 결제 여부 판별.
    Y='카드결제'/'혼합' 케이스도 단계별로 partner 패턴 확인 — 한 프로젝트에서
    일부 단계만 카드 결제일 수도 있음(예: 계약금 일반, 잔금 카드).
    """
    iv = (invoice_value or '').strip()
    if iv in ('카드결제', '혼합'):
        if not partner:
            return False
        p = partner.strip()
        # 1) 카드 승인번호 패턴 (영숫자/한글+숫자)
        if _CARD_PARTNER_RE.match(p):
            return True
        # 2) 카드사 이름 직접 매칭 (예: "비씨카드", "삼성카드")
        if _CARD_BRAND_RE.search(p):
            return True
    return False


def _resolve_payment_code(invoice_value: str, bank: str, partner: str = '') -> str:
    """Y열 값 + 메모 은행 + 거래처 → 'G' / 'N' / 'R'"""
    iv = (invoice_value or '').strip()
    if iv == 'N입금':
        return 'N'
    # 입금자가 '현금'이면 N으로 강제 (혼합 케이스에서 매니저가 수기 입력)
    if partner and partner.strip() == '현금':
        return 'N'
    # 카드결제 또는 기타 → 메모 은행으로 판별
    if bank == '기업':
        return 'G'
    if bank == '하나':
        return 'R'
    if bank == '농협':
        return 'N'
    return 'G'


# ─────────────────────────────────────────────
# 메시지 빌더
# ─────────────────────────────────────────────

_STAGE_EMOJI = {
    '계약금': ':moneybag:',
    '중도금': ':moneybag:',
    '잔금': ':moneybag:',
}
_SEP = '--------------------------------------------'
_SEP_HARD = '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'


def _build_stage_message(
    stage: str, project: str, address: str,
    payment: Dict, invoice_value: str,
    total_r: int, total_t: int, unpaid: int,
    stage_sheet_val: int = 0,
    construction: str = '',
) -> str:
    """단계별 입금 알림 (계약금/중도금/잔금)."""
    emoji = _STAGE_EMOJI.get(stage, ':moneybag:')
    is_card = _is_card_payment(invoice_value, payment.get('partner', ''))
    code = _resolve_payment_code(invoice_value, payment.get('bank', ''), payment.get('partner', ''))
    note_label = payment.get('note_label', '')
    # 매니저 이체 표기 반영 (2026-07-10): "R>G" 감지 시 원 코드 뒤에 " → 이체 코드" 부기
    transfer_to = payment.get('transfer_to', '')
    if transfer_to and transfer_to != code:
        code = f"{code} → {transfer_to}"
    code_display = f"{code}, {note_label}" if note_label else code
    amount = payment.get('amount', 0)
    partner = payment.get('partner', '-') or '-'
    bank = payment.get('bank', '-') or '-'
    date_md = payment.get('date_md', '-')

    date_label = '결제일' if is_card else '입금일'
    amount_label = '결제금액' if is_card else '입금액'
    partner_label = '승인번호' if is_card else '입금자'
    header_action = '결제' if is_card else '입금'

    lines = [
        '⠀',
        f"{emoji} *{stage} {header_action}* — :id: *{project}*",
        _SEP,
        f"주소 : {address or '-'}",
    ]
    if construction:
        lines.append(f"공사내용 : {construction}")
    lines.extend([
        f"{date_label} : {date_md}",
        f"{amount_label} : {amount:,}원",
        f"{partner_label} : {partner}",
        f"은행 : {bank} ({code_display})",
    ])
    if is_card:
        real_payment = _resolve_real_payment(stage_sheet_val, amount, total_t)
        extra_3pct = round(real_payment * 0.03 / 1.03)
        fee = real_payment - amount
        if fee > 0 and extra_3pct > 0:
            parts = [
                f"실결제 {real_payment:,}원",
                f"3% {extra_3pct:,}원",
                f"카드 수수료 {fee:,}원",
            ]
            lines.append(f"({' / '.join(parts)})")
    lines.append(_SEP)
    lines.append(f"미수금 : {unpaid:,}원")
    lines.append('⠀')
    return '\n'.join(lines)


def _resolve_real_payment(stage_val: int, amount: int, total_t: int) -> int:
    """카드 결제 실결제 자동 판별.
    - stage_val >= amount: 시트 값 = 실결제 (G2530-TH 양식, 매니저가 카드 부담분 반영)
    - stage_val < amount: 시트 값 = R (R3282-MJ 양식), 실결제 = 시트 값 × 1.03
    - stage_val 없으면 total_t × 1.03 fallback
    """
    if stage_val <= 0:
        return round(total_t * 1.03)
    if stage_val >= amount:
        return stage_val
    return round(stage_val * 1.03)


def _build_stage_with_history_message(
    stage: str, project: str, address: str,
    last_payment: Dict, all_payments: List[Dict], invoice_value: str,
    total_r: int, total_t: int, unpaid: int,
    stage_sheet_vals: Optional[Dict[str, int]] = None,
    construction: str = '',
) -> str:
    """단계 카드 + 누적 이력 (중도금 입금 시 사용)."""
    emoji = _STAGE_EMOJI.get(stage, ':moneybag:')
    is_card = _is_card_payment(invoice_value, last_payment.get('partner', ''))
    code = _resolve_payment_code(invoice_value, last_payment.get('bank', ''), last_payment.get('partner', ''))
    note_label = last_payment.get('note_label', '')
    transfer_to = last_payment.get('transfer_to', '')
    if transfer_to and transfer_to != code:
        code = f"{code} → {transfer_to}"
    code_display = f"{code}, {note_label}" if note_label else code
    amount = last_payment.get('amount', 0)
    partner = last_payment.get('partner', '-') or '-'
    bank = last_payment.get('bank', '-') or '-'
    date_md = last_payment.get('date_md', '-')

    date_label = '결제일' if is_card else '입금일'
    amount_label = '결제금액' if is_card else '입금액'
    partner_label = '승인번호' if is_card else '입금자'
    header_action = '결제' if is_card else '입금'

    lines = [
        '⠀',
        f"{emoji} *{stage} {header_action}* — :id: *{project}*",
        _SEP,
        f"주소 : {address or '-'}",
    ]
    if construction:
        lines.append(f"공사내용 : {construction}")
    lines.extend([
        f"{date_label} : {date_md}",
        f"{amount_label} : {amount:,}원",
        f"{partner_label} : {partner}",
        f"은행 : {bank} ({code_display})",
    ])
    if is_card:
        stage_val = (stage_sheet_vals or {}).get(stage, 0)
        real_payment = _resolve_real_payment(stage_val, amount, total_t)
        extra_3pct = round(real_payment * 0.03 / 1.03)
        fee = real_payment - amount
        if fee > 0 and extra_3pct > 0:
            lines.append(
                f"(실결제 {real_payment:,}원 / 3% {extra_3pct:,}원 / 카드 수수료 {fee:,}원)"
            )
    lines.append('')
    lines.append('[누적 이력]')
    # 현재 알림 단계까지만 표시 — 매니저가 다음 단계 메모를 미리 입력한 경우 차단
    _STAGE_ORDER = {'계약금': 0, '중도금': 1, '잔금': 2}
    cur_idx = _STAGE_ORDER.get(stage, 99)
    history = [p for p in all_payments if _STAGE_ORDER.get(p.get('stage'), 99) <= cur_idx]
    for p in history:
        st = p.get('stage', '-')
        d = p.get('date_md', '-')
        c = _resolve_payment_code(invoice_value, p.get('bank', ''), p.get('partner', ''))
        # 이체 표기 반영 (2026-07-10)
        tr = p.get('transfer_to', '')
        if tr and tr != c:
            c = f"{c} → {tr}"
        nl = p.get('note_label', '')
        bk = p.get('bank', '') or ''
        inner = f"{c}, {nl}" if nl else c
        c_disp = f"{bk} ({inner})" if bk else inner
        a = p.get('amount', 0)
        pt = p.get('partner', '-') or '-'
        is_c = _is_card_payment(invoice_value, pt)
        suffix = ' (카드)' if is_c else ''
        lines.append(f"{st}  {d}  {c_disp}  {a:,}원  {pt}{suffix}")
    lines.append(_SEP)
    lines.append(f"미수금 : {unpaid:,}원")
    lines.append(' ')
    lines.append('⠀')
    return '\n'.join(lines)


def _build_complete_message(
    project: str, address: str,
    payments: List[Dict], invoice_value: str, total_t: int,
    stage_sheet_vals: Optional[Dict[str, int]] = None,
    construction: str = '',
) -> str:
    """수금완료 알림 — 전체 history 취합."""
    lines = [
        '⠀',
        f":white_check_mark: *수금완료* — :id: *{project}*",
        _SEP,
        f"주소 : {address or '-'}",
    ]
    if construction:
        lines.append(f"공사내용 : {construction}")
    lines.extend([
        '',
        '[입금 이력]',
    ])
    for p in payments:
        stage = p.get('stage', '-')
        date_md = p.get('date_md', '-')
        code = _resolve_payment_code(invoice_value, p.get('bank', ''), p.get('partner', ''))
        # 이체 표기 반영 (2026-07-10)
        transfer_to = p.get('transfer_to', '')
        if transfer_to and transfer_to != code:
            code = f"{code} → {transfer_to}"
        note_label = p.get('note_label', '')
        bank = p.get('bank', '') or ''
        inner = f"{code}, {note_label}" if note_label else code
        code_display = f"{bank} ({inner})" if bank else inner
        amount = p.get('amount', 0)
        partner = p.get('partner', '-') or '-'
        is_card = _is_card_payment(invoice_value, partner)
        suffix = ' (카드)' if is_card else ''
        lines.append(
            f"{stage}  {date_md}  {code_display}  {amount:,}원  {partner}{suffix}"
        )
        # 카드 결제 단계 — 부가 정보 라인 추가 (단계 시트 값 = 실결제)
        if is_card and stage_sheet_vals:
            # 같은 단계 카드 결제 중 마지막 입금에서만 부가 정보 1회 표시 (분할 입금 합산)
            same_stage_cards = [
                pp for pp in payments
                if pp.get('stage') == stage
                and _is_card_payment(invoice_value, pp.get('partner', ''))
            ]
            if same_stage_cards and same_stage_cards[-1] is p:
                stage_val = stage_sheet_vals.get(stage, 0)
                same_stage_card_total = sum(c['amount'] for c in same_stage_cards)
                real_payment = _resolve_real_payment(stage_val, same_stage_card_total, total_t)
                if real_payment > 0:
                    fee = real_payment - same_stage_card_total
                    three_pct = round(real_payment * 0.03 / 1.03)
                    if fee > 0 and three_pct > 0:
                        lines.append(
                            f"  (실결제 {real_payment:,}원 / 3% {three_pct:,}원 / 카드 수수료 {fee:,}원)"
                        )
    lines.append(_SEP)
    lines.append(f"총액 : {total_t:,}원")
    lines.append('⠀')
    return '\n'.join(lines)


# ─────────────────────────────────────────────
# 시트 폴링 + 변경 감지 + 발송
# ─────────────────────────────────────────────

_payment_service = None
_payment_service_lock = None


def _get_sheets_manager():
    """[기존 호환] 다른 모듈과 같은 manager"""
    from dashboard.services.lead_service import get_sheets_manager
    return get_sheets_manager()


def _get_payment_service():
    """payment_sync 전용 Sheets API service (다른 폴링과 SSL 충돌 방지용)."""
    global _payment_service
    if _payment_service is not None:
        return _payment_service
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        cred_file = os.getenv('GOOGLE_CREDENTIALS_FILE', 'credentials.json').strip()
        if not os.path.isabs(cred_file):
            cred_file = os.path.join(os.getcwd(), cred_file)
        creds = Credentials.from_service_account_file(
            cred_file, scopes=['https://www.googleapis.com/auth/spreadsheets'],
        )
        _payment_service = build(
            'sheets', 'v4', credentials=creds, cache_discovery=False,
        )
        return _payment_service
    except Exception as exc:
        logger.error(f"[PAYMENT] service 초기화 실패: {exc}", exc_info=True)
        return None


def _reset_payment_service():
    """SSL 에러 시 service 재생성용"""
    global _payment_service
    _payment_service = None


def _fetch_row_notes(spreadsheet_id: str, sheet_name: str, row: int) -> List[str]:
    """U/V/W 셀 노트 3개 fetch."""
    service = _get_payment_service()
    if not service:
        return ['', '', '']
    try:
        resp = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            ranges=[f"'{sheet_name}'!U{row}:W{row}"],
            fields='sheets.data.rowData.values.note',
            includeGridData=True,
        ).execute()
        sheets = resp.get('sheets', [])
        if not sheets:
            return ['', '', '']
        data = sheets[0].get('data', [])
        if not data:
            return ['', '', '']
        row_data = data[0].get('rowData', [])
        if not row_data:
            return ['', '', '']
        values = row_data[0].get('values', [])
        notes = []
        for v in values:
            notes.append(v.get('note', '') or '')
        # 부족하면 빈 문자열로 패딩
        while len(notes) < 3:
            notes.append('')
        return notes[:3]
    except Exception as exc:
        logger.error(f"[PAYMENT] 노트 fetch 실패 (row {row}): {exc}", exc_info=True)
        return ['', '', '']


def _to_int_won(s) -> int:
    """'12,100,000' / '₩12,100,000' / '12100000' / 0 → 12100000"""
    if s is None or s == '':
        return 0
    if isinstance(s, (int, float)):
        return int(s)
    s = str(s).strip().replace('₩', '').replace(',', '').replace(' ', '')
    if not s:
        return 0
    try:
        return int(float(s))
    except Exception:
        return 0


def _to_bool(s) -> bool:
    """체크박스 — True/'TRUE'/'true' → True, 그 외 False"""
    if isinstance(s, bool):
        return s
    return str(s).strip().lower() == 'true'


def sync_payments() -> Dict:
    """공사 현황 시트 폴링 → U/V/W 변경 감지 + AA 체크 변경 감지 → 알림 발송."""
    result = {'processed': 0, 'sent': 0, 'errors': 0}

    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    channel = os.getenv('SLACK_PAYMENT_CHANNEL', '').strip()
    bot_token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    if not all([sheet_id, sheet_name, channel, bot_token]):
        logger.debug('[PAYMENT] 필수 환경변수 미설정 — skip')
        return result

    service = _get_payment_service()
    if not service:
        return result
    try:
        # 한 번에 A:AA 가져옴 (광폭 범위) — values만 가벼움
        resp = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!A2:AA10000",
            valueRenderOption='UNFORMATTED_VALUE',
        ).execute()
    except Exception as exc:
        err_l = str(exc).lower()
        is_ssl = any(k in err_l for k in ('ssl', 'wrong_version', 'decryption', 'handshake'))
        if is_ssl:
            logger.warning(f"[PAYMENT] SSL 에러 — service 재생성: {exc}")
            _reset_payment_service()
        else:
            logger.error(f"[PAYMENT] 시트 값 fetch 실패: {exc}", exc_info=True)
        return result

    rows = resp.get('values', [])
    if not rows:
        return result

    try:
        rc = get_redis_client().redis
    except Exception as exc:
        logger.error(f"[PAYMENT] Redis 연결 실패: {exc}")
        return result

    # 슬랙 client
    try:
        from slack_sdk import WebClient
        slack = WebClient(token=bot_token)
    except Exception as exc:
        logger.error(f"[PAYMENT] Slack client 초기화 실패: {exc}")
        return result

    # 컬럼 인덱스 (0-based)
    def col_idx(col: str) -> int:
        if len(col) == 1:
            return ord(col) - ord('A')
        return (ord(col[0]) - ord('A') + 1) * 26 + (ord(col[1]) - ord('A'))

    IDX_A = col_idx('A')
    IDX_F = col_idx('F')
    IDX_L = col_idx('L')
    IDX_R = col_idx('R')
    IDX_T = col_idx('T')
    IDX_U = col_idx('U')
    IDX_V = col_idx('V')
    IDX_W = col_idx('W')
    IDX_X = col_idx('X')
    IDX_Y = col_idx('Y')
    IDX_AA = col_idx('AA')

    # 첫 폴링 감지 — Redis에 키 하나라도 있는지
    try:
        baseline_done = bool(rc.exists('payment_sync:baseline_done'))
    except Exception:
        baseline_done = False

    # 첫 폴링이면 모든 행의 U/V/W 노트를 한 번에 fetch (phash baseline 구축)
    all_phash_by_row: Dict[int, str] = {}
    if not baseline_done:
        try:
            resp_notes = service.spreadsheets().get(
                spreadsheetId=sheet_id,
                ranges=[f"'{sheet_name}'!U2:W10000"],
                fields='sheets.data.rowData.values.note',
                includeGridData=True,
            ).execute()
            sheets_data = resp_notes.get('sheets', [])
            if sheets_data:
                data_list = sheets_data[0].get('data', [])
                if data_list:
                    row_data_list = data_list[0].get('rowData', [])
                    for offset_n, rd in enumerate(row_data_list):
                        vals = rd.get('values', [])
                        notes_3 = []
                        for v in vals:
                            notes_3.append(v.get('note', '') or '')
                        while len(notes_3) < 3:
                            notes_3.append('')
                        payments_n = _parse_notes(notes_3[:3])
                        if payments_n:
                            all_phash_by_row[offset_n + 2] = _hash_payments(payments_n)
            logger.info(
                f"[PAYMENT] baseline 노트 fetch 완료: {len(all_phash_by_row)}개 행"
            )
        except Exception as exc:
            logger.error(f"[PAYMENT] baseline 노트 fetch 실패: {exc}", exc_info=True)

    # 발송 처리할 변경 행만 모음 (한 폴링당 최대 N건 — SSL 동시 호출 방지)
    MAX_PER_TICK = 5
    changed_rows = []

    for offset, row in enumerate(rows):
        sheet_row = offset + 2  # 1-based + 헤더 1행
        # 행 길이 부족 시 패딩 (Sheets API trailing trim 방지)
        while len(row) < 27:
            row.append('')
        def _get(i):
            return row[i] if i < len(row) else ''
        u_val = _to_int_won(_get(IDX_U))
        v_val = _to_int_won(_get(IDX_V))
        w_val = _to_int_won(_get(IDX_W))
        aa_chk = _to_bool(_get(IDX_AA))
        project = str(_get(IDX_A)).strip()
        if not project:
            continue
        # 비표준 코드 (옛 프로젝트 1035, "0 중고..." 등) — 알림 skip
        if not _VALID_PROJECT_RE.match(project):
            continue

        # Redis 옛 상태 비교
        key = f"{REDIS_KEY_PREFIX}{sheet_row}"
        try:
            prev = rc.hgetall(key)
        except Exception:
            prev = {}
        prev_u = int(prev.get('u', 0) or 0)
        prev_v = int(prev.get('v', 0) or 0)
        prev_w = int(prev.get('w', 0) or 0)
        prev_aa = str(prev.get('aa', '')).lower() == 'true'

        # 첫 폴링은 baseline만 저장하고 발송 X
        if not prev:
            init_phash = all_phash_by_row.get(sheet_row, '')
            try:
                rc.hset(key, mapping={
                    'u': u_val, 'v': v_val, 'w': w_val,
                    'aa': 'true' if aa_chk else 'false',
                    'phash': init_phash,
                })
                rc.expire(key, REDIS_TTL)
            except Exception:
                pass
            continue

        # 값 변경/AA 변경 자체는 baseline 갱신 트리거
        any_change = (u_val != prev_u) or (v_val != prev_v) or (w_val != prev_w) \
            or (aa_chk != prev_aa)
        if not any_change:
            continue

        # 발송 트리거 — 새 입금(값 증가) 또는 AA 신규 체크만
        # 값 감소(정정/취소), AA 해제는 baseline만 갱신하고 skip
        new_payment = (u_val > prev_u) or (v_val > prev_v) or (w_val > prev_w)
        aa_newly_checked = aa_chk and not prev_aa
        if not (new_payment or aa_newly_checked):
            try:
                rc.hset(key, mapping={
                    'u': u_val, 'v': v_val, 'w': w_val,
                    'aa': 'true' if aa_chk else 'false',
                })
                rc.expire(key, REDIS_TTL)
            except Exception:
                pass
            continue

        changed_rows.append({
            'row': sheet_row, 'project': project,
            'u': u_val, 'v': v_val, 'w': w_val, 'aa': aa_chk,
            'prev_u': prev_u, 'prev_v': prev_v, 'prev_w': prev_w, 'prev_aa': prev_aa,
            'address': str(_get(IDX_F)).strip(),
            'construction': str(_get(IDX_L)).strip(),
            'invoice': str(_get(IDX_Y)).strip(),
            'total_r': _to_int_won(_get(IDX_R)),
            'total_t': _to_int_won(_get(IDX_T)),
            'unpaid': _to_int_won(_get(IDX_X)),
            'key': key,
            'phash': prev.get('phash', '') if prev else '',
        })

    if not changed_rows:
        return result

    # 안전장치 — 한 폴링에 너무 많은 변경 = 비정상. baseline 재구축만 하고 skip.
    if len(changed_rows) > MAX_PER_TICK * 4:
        logger.warning(
            f"[PAYMENT] 변경 행이 비정상적으로 많음 ({len(changed_rows)}) — baseline 재구축, 발송 skip"
        )
        for c in changed_rows:
            try:
                rc.hset(c['key'], mapping={
                    'u': c['u'], 'v': c['v'], 'w': c['w'],
                    'aa': 'true' if c['aa'] else 'false',
                })
                rc.expire(c['key'], REDIS_TTL)
            except Exception:
                pass
        return result

    # 변경 행 발송 — 한 폴링당 MAX_PER_TICK까지만 (SSL 동시 호출 방지)
    # 분기:
    #   - 단계 양수 증가 + 메모 변경 → 해당 단계 입금 알림
    #   - AA: false→true + X==0 → 수금완료 알림 (전체 history)
    for c in changed_rows[:MAX_PER_TICK]:
        sheet_row = c['row']
        project = c['project']
        result['processed'] += 1
        try:
            prev_u, prev_v, prev_w = c['prev_u'], c['prev_v'], c['prev_w']
            prev_aa = c['prev_aa']
            u_val, v_val, w_val, aa_chk = c['u'], c['v'], c['w'], c['aa']
            notes = _fetch_row_notes(sheet_id, sheet_name, sheet_row)
            payments = _parse_notes(
                notes,
                stage_vals={'계약금': u_val, '중도금': v_val, '잔금': w_val},
            )
            new_phash = _hash_payments(payments)
            prev_phash = c.get('phash', '')
            note_changed = new_phash != prev_phash

            # 단계별 발송 분기 (양수 증가 + 메모 변경된 단계만):
            #   계약금 → 단일 카드
            #   중도금 → 단일 카드 + 누적 이력
            #   잔금   → 수금완료 + 전체 이력 (잔금 입금 = 수금완료 의미)
            sent_this_row = False
            if note_changed and payments:
                stages_increased = []
                if u_val > prev_u:
                    stages_increased.append('계약금')
                if v_val > prev_v:
                    stages_increased.append('중도금')
                if w_val > prev_w:
                    stages_increased.append('잔금')

                # 노트 미완성 감지 (2026-07-10 · 2026-07-10 조건 완화)
                # 매니저 실제 워크플로: U/V/W 금액 먼저 입력 → 노트 나중 입력.
                # 30초 폴링이 그 사이에 걸리면 노트 없는 상태로 발송돼 partner/date 가 비어서 나감.
                #
                # 초기 fix: date_md='-' AND partner='-' (완전 fallback) 만 감지
                # 강화 (G3702-MS 재발 관측): partner 또는 date 하나만 빠져도 skip.
                # 매니저가 노트 부분만 입력한 중간 상태도 감지 대상.
                incomplete = any(
                    (not p.get('partner') or p.get('partner') == '-')
                    or (not p.get('date_md') or p.get('date_md') == '-')
                    for stage in stages_increased
                    for p in payments
                    if p.get('stage') == stage
                )
                if incomplete:
                    logger.info(
                        f"[PAYMENT] 노트 미저장 감지 → skip ({project}, row {sheet_row}). "
                        f"다음 폴링에서 노트 확인 후 재시도"
                    )
                    continue

                # Backfill 검증 flag 상태면 stages_increased 전체를 skip + phash 저장 안 함
                # (P0-1, 2026-07-10 G3717-SH 관측)
                if os.getenv('PAYMENT_SLACK_DISABLED', '').strip() in ('1', 'true', 'True') and stages_increased:
                    logger.info(
                        f"[PAYMENT] 발송 disabled — phash 유지로 재감지 대상 "
                        f"({project}, row {sheet_row}, stages={stages_increased})"
                    )
                    continue

                for stage in stages_increased:
                    stage_payments = [p for p in payments if p.get('stage') == stage]
                    if not stage_payments:
                        continue
                    last_payment = stage_payments[-1]
                    stage_vals = {
                        '계약금': c['u'], '중도금': c['v'], '잔금': c['w'],
                    }
                    if stage == '잔금' and c['unpaid'] == 0:
                        # 잔금 + 미수금 0 → 수금완료
                        text = _build_complete_message(
                            project=project, address=c['address'],
                            payments=payments, invoice_value=c['invoice'],
                            total_t=c['total_t'],
                            stage_sheet_vals=stage_vals,
                            construction=c.get('construction', ''),
                        )
                    elif stage in ('중도금', '잔금'):
                        # 중도금, 또는 잔금 부족 입금(수금중) → 단계 카드 + history
                        text = _build_stage_with_history_message(
                            stage=stage, project=project, address=c['address'],
                            last_payment=last_payment, all_payments=payments,
                            invoice_value=c['invoice'],
                            total_r=c['total_r'], total_t=c['total_t'],
                            unpaid=c['unpaid'],
                            stage_sheet_vals=stage_vals,
                            construction=c.get('construction', ''),
                        )
                    else:  # 계약금
                        text = _build_stage_message(
                            stage=stage, project=project, address=c['address'],
                            payment=last_payment, invoice_value=c['invoice'],
                            total_r=c['total_r'], total_t=c['total_t'],
                            unpaid=c['unpaid'],
                            stage_sheet_val=stage_vals.get(stage, 0),
                            construction=c.get('construction', ''),
                        )
                    resp = slack.chat_postMessage(channel=channel, text=text)
                    result['sent'] += 1
                    logger.info(
                        f"[PAYMENT] {stage} 입금 발송: {project} (row {sheet_row})"
                    )
                    # 발송 성공 시 ts 저장 → 나중에 chat.update 정정 가능 (P0-2, 2026-07-10)
                    try:
                        if resp and resp.get('ok'):
                            ts = resp.get('ts', '')
                            if ts:
                                rc.set(
                                    f'payment_slack:ts:{project}:{stage}',
                                    ts,
                                    ex=60 * 60 * 24 * 90,  # 90일 보관
                                )
                    except Exception as _exc:
                        logger.debug(f'[PAYMENT] ts 저장 실패 ({project}/{stage}): {_exc}')
                    sent_this_row = True

            if not sent_this_row:
                logger.debug(
                    f"[PAYMENT] skip ({project}, row {sheet_row}) — note_changed={note_changed}"
                )

            # 매니저 실수 감지 + 알림 훅 (2026-07-10)
            # 카테고리: 노트 미저장/자릿수 오타/필수 항목 누락/미체크/미수금 이상/이니셜 오타
            try:
                from dashboard.services.payment_alert import check_and_alert
                from dashboard.blueprints.slack_helpers import _load_initials_from_config
                alert_row = {
                    'code': project,
                    'address': c.get('address', ''),
                    'total_t': c['total_t'],
                    'unpaid': c['unpaid'],
                    'aa': aa_chk,
                    '계약금': u_val, '중도금': v_val, '잔금': w_val,
                }
                known_initials = set(_load_initials_from_config().values())
                check_and_alert(alert_row, payments, known_initials, slack_client=slack)
            except Exception as exc:
                logger.warning(f"[PAYMENT_ALERT] 감지 훅 예외 ({project}): {exc}")

            # 새 상태 저장
            rc.hset(c['key'], mapping={
                'u': u_val, 'v': v_val, 'w': w_val,
                'aa': 'true' if aa_chk else 'false',
                'phash': new_phash,
            })
            rc.expire(c['key'], REDIS_TTL)
        except Exception as exc:
            logger.error(
                f"[PAYMENT] 행 처리 실패 ({project}, row {sheet_row}): {exc}",
                exc_info=True,
            )
            result['errors'] += 1

    # baseline 완료 마커 — 첫 폴링이었으면 표시 (다음 폴링부터 includeGridData 안 호출)
    if not baseline_done:
        try:
            rc.set('payment_sync:baseline_done', '1', ex=REDIS_TTL)
        except Exception:
            pass

    # 일일 통계 누적 (오늘 발송 건수)
    if result['sent'] > 0:
        try:
            from datetime import datetime
            today_key = f"payment_sync:daily_stats:{datetime.now().strftime('%Y-%m-%d')}"
            rc.hincrby(today_key, 'sent', result['sent'])
            rc.hincrby(today_key, 'errors', result['errors'])
            rc.expire(today_key, 60 * 60 * 24 * 7)  # 일주일 보관
        except Exception:
            pass

    return result


# ─────────────────────────────────────────────
# 일일 요약 / 미수금 장기 체류 / 검색 (운영 보조)
# ─────────────────────────────────────────────


def daily_payment_summary() -> Optional[str]:
    """오늘 발송한 수금 알림 요약 메시지 빌드 → 슬랙 발송용 text."""
    from datetime import datetime
    try:
        rc = get_redis_client().redis
        today = datetime.now().strftime('%Y-%m-%d')
        stats = rc.hgetall(f"payment_sync:daily_stats:{today}")
        if not stats:
            return None
        sent = int(stats.get('sent', 0) or 0)
        errors = int(stats.get('errors', 0) or 0)
        lines = [
            f":bar_chart: *수금 알림 일일 요약* — {today}",
            '--------------------------------------------',
            f"발송 완료 : {sent}건",
            f"오류 : {errors}건",
            '--------------------------------------------',
        ]
        return '\n'.join(lines)
    except Exception as exc:
        logger.error(f"[PAYMENT] 일일 요약 실패: {exc}", exc_info=True)
        return None


def find_overdue_unpaid(days: int = 30) -> List[Dict]:
    """미수금 ≠ 0 + 최종 입금일이 days일 이상 경과한 프로젝트 리스트.

    반환: [{'project', 'address', 'unpaid', 'last_date'}, ...]
    """
    from datetime import datetime, timedelta
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    if not sheet_id or not sheet_name:
        return []
    service = _get_payment_service()
    if not service:
        return []
    try:
        resp = service.spreadsheets().values().get(
            spreadsheetId=sheet_id, range=f"'{sheet_name}'!A2:AA10000",
            valueRenderOption='UNFORMATTED_VALUE',
        ).execute()
    except Exception as exc:
        logger.error(f"[PAYMENT] overdue 시트 fetch 실패: {exc}")
        return []
    rows = resp.get('values', [])
    cutoff = datetime.now() - timedelta(days=days)
    overdue = []
    for row in rows:
        while len(row) < 27:
            row.append('')
        a = str(row[0]).strip()
        if not a or not _VALID_PROJECT_RE.match(a):
            continue
        unpaid = _to_int_won(row[23])
        if unpaid == 0:
            continue
        # 수금 확인 체크 → 미수금 있어도 정리된 케이스
        if _to_bool(row[26]):
            continue
        # 최종 입금일 — Z열(수금 날짜) 시리얼 → 날짜
        z_val = row[25]
        last_date = None
        if isinstance(z_val, (int, float)) and z_val > 0:
            # 구글 시리얼 (1899-12-30 기준)
            try:
                last_date = datetime(1899, 12, 30) + timedelta(days=int(z_val))
            except Exception:
                last_date = None
        if last_date is None or last_date >= cutoff:
            continue
        overdue.append({
            'project': a,
            'address': str(row[5]).strip(),
            'unpaid': unpaid,
            'last_date': last_date.strftime('%Y-%m-%d'),
        })
    overdue.sort(key=lambda x: x['last_date'])
    return overdue


def build_overdue_message(days: int = 30, limit: int = 20) -> Optional[str]:
    """미수금 장기 체류 슬랙 알림 메시지."""
    items = find_overdue_unpaid(days=days)
    if not items:
        return None
    lines = [
        f":warning: *미수금 {days}일 이상 경과* — {len(items)}건",
        '--------------------------------------------',
    ]
    for it in items[:limit]:
        lines.append(
            f"`{it['project']}` {it['address'][:30]} : "
            f"미수금 {it['unpaid']:,}원 (마지막 입금 {it['last_date']})"
        )
    if len(items) > limit:
        lines.append(f"... 외 {len(items) - limit}건")
    lines.append('--------------------------------------------')
    return '\n'.join(lines)


def search_project(project_code: str) -> Optional[str]:
    """특정 프로젝트의 전체 수금 history 조회 (슬래시 명령용)."""
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    if not sheet_id or not sheet_name:
        return None
    service = _get_payment_service()
    if not service:
        return None
    try:
        resp = service.spreadsheets().values().get(
            spreadsheetId=sheet_id, range=f"'{sheet_name}'!A2:AA10000",
            valueRenderOption='UNFORMATTED_VALUE',
        ).execute()
    except Exception as exc:
        logger.error(f"[PAYMENT] 검색 시트 fetch 실패: {exc}")
        return None
    rows = resp.get('values', [])
    code = project_code.strip().upper()
    for off, row in enumerate(rows):
        while len(row) < 27:
            row.append('')
        if str(row[0]).strip().upper() != code:
            continue
        sheet_row = off + 2
        notes = _fetch_row_notes(sheet_id, sheet_name, sheet_row)
        u_val = _to_int_won(row[20])
        v_val = _to_int_won(row[21])
        w_val = _to_int_won(row[22])
        stage_vals = {'계약금': u_val, '중도금': v_val, '잔금': w_val}
        payments = _parse_notes(notes, stage_vals=stage_vals)
        if not payments:
            return f"`{code}` — 수금 메모 없음 (시트 메모 미입력)"
        address = str(row[5]).strip()
        construction = str(row[11]).strip() if len(row) > 11 else ''
        total_t = _to_int_won(row[19])
        unpaid = _to_int_won(row[23])
        invoice = str(row[24]).strip()
        msg = _build_complete_message(
            code, address, payments, invoice, total_t,
            stage_sheet_vals=stage_vals,
            construction=construction,
        )
        if unpaid != 0:
            msg += f"\n_미수금 : {unpaid:,}원_"
        return msg
    return f"`{code}` — 시트에 없음"
