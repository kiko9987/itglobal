"""수금 관리 자동 알림 — 공사 현황 시트의 U/V/W 입금 메모 변경 감지 → #수금_관리 채널 발송.

흐름:
1. 시트 폴링 (값만 가벼움): R/S/T/U/V/W/X/Y/Z/AA + A/F
2. Redis에 이전 (U, V, W, AA) 상태 저장 후 변경 감지
3. 변경된 행만 셀 메모(U/V/W) 추가 fetch
4. 메모 파싱 → 누적 history 메시지 빌더 → 슬랙 발송

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
COL_UNPAID = 'X'     # 미수금 — 시트 수식: = IF(ABS(T-U-V-W)<2, 0, T-U-V-W)  (2026-07-16 재확인)
                     #   X > 0 : 정상 미납 (총액 대비 입금 부족)
                     #   X = 0 : 완납 (2원 미만 부동소수 오차 포함)
                     #   X < 0 : 초과 입금 (매니저 실수 or 정정 필요)
                     # 슬랙 카드 build 는 -abs 처리로 항상 음수 표시 (매니저 시각 통일).
                     # 이전 (U+V+W)-T 주석·수식 (2026-07-10) 은 부호 반대 오류였음.
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
_AMOUNT_RE = re.compile(r'(?:입금|금액\s*[:：])\s*([\d,]+)(?:\s*원)?')
# 한글 금액: "272만원", "272만 원", "1,700만원" 등 → 만 단위 (2026-07-10)
# "만원권" (지폐 종류) 은 제외. 예: "5만원권 246장" 은 금액 아님.
_KO_AMOUNT_RE = re.compile(r'([\d,]+)\s*만\s*원(?!권)')
# 한글 날짜: "6월15일", "6월 15일" (연도 없으면 현재 년도) (2026-07-10)
_KO_DATE_RE = re.compile(r'(\d{1,2})\s*월\s*(\d{1,2})\s*일')
# 프로젝트 코드 패턴 — partner 파싱 시 skip 대상 (2026-07-10)
_PROJECT_CODE_RE = re.compile(r'^[GRN]\d{4}-[A-Z]{1,3}$')
# 매니저 이체 표기 — "2026-07-10 R>g", "2026-07-10 G>N", "2026-07-10 r>g" 등 (2026-07-10)
# 첫 번째 문자 = 원 계좌, 두 번째 문자 = 이체 후 계좌. 양쪽 대소문자 무관.
_TRANSFER_RE = re.compile(r'\d{4}-\d{1,2}-\d{1,2}\s+([grnpGRNP])\s*>\s*([grnpGRNP])')
# 라벨 양식 (앱스크립트 자동 + 매니저 수기 덮어쓰기)
_LABEL_DATE_RE = re.compile(r'^입금일\s*:\s*(.+)$')
_LABEL_PAYER_RE = re.compile(r'^입금자\s*:\s*(.+)$')
# 매니저 카톡 양식 (분할 입금 history를 V 셀에 그대로 복사한 케이스)
# 예: "02/02 G 22,308,000원 프레임플러스"
_KATOK_LINE_RE = re.compile(
    r'^(\d{1,2})/(\d{1,2})\s+([GRN]|현금|박C기업)\s+([\d,]+)\s*원\s*(.*?)\s*(?:수금중|수금완료)?\s*$'
)
# 매니저 요약 양식 (2026-07-10 R2951-TH 관측)
# 예: "2025-12-05 300,000원 김연종 농협"  → date/amount/partner/bank 순서
_MANAGER_LINE_RE = re.compile(
    r'^\d{4}-(\d{1,2})-(\d{1,2})\s+([\d,]+)\s*원\s+(\S.*?)\s+(기업|하나|농협|국민|신한|우리|카카오|토스)\s*$'
)
# 단일 숫자(콤마 OK) 라인 — 매니저가 입금액만 적은 경우
_BARE_AMOUNT_RE = re.compile(r'^([\d,]+)(?:\s*원)?$')
# 지폐 수량 표기 — "5만원권 74장" = 5 × 10000 × 74 (2026-07-15 R3791-MJ 관측)
_BILL_COUNT_RE = re.compile(r'(\d+)\s*만\s*원권\s*(\d+)\s*장')
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
    # 지폐 수량 (5만원권 74장) — X * 10000 * Y (2026-07-15 R3791-MJ)
    if amount == 0:
        for ln in lines:
            m = _BILL_COUNT_RE.search(ln)
            if m:
                amount = int(m.group(1)) * 10000 * int(m.group(2))
                break
    # 옛 양식 fallback — 메모에 금액 없으면 시트 단계 값 사용
    if amount == 0 and fallback_amount > 0:
        amount = fallback_amount
    # amount 없어도 date_md 나 label 있으면 metadata payment 로 반환 (2026-07-10)
    # _parse_notes 후처리에서 시트값 - 다른 블록 amount 합 = 유추 amount 세팅
    _has_meta_only = False
    if amount == 0:
        # 라벨 양식 감지 (입금일:, 입금자:) or 표준 양식의 date 라인 만 있는 케이스
        if any(_LABEL_DATE_RE.match(ln) or _LABEL_PAYER_RE.match(ln) for ln in lines):
            _has_meta_only = True
        elif date_md:
            _has_meta_only = True
        else:
            return None  # 금액도 metadata 도 없으면 입금 블록 아님
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
    # 매니저 수기 요약/설명 라인 skip 패턴 (2026-07-10)
    #   "수령완료", "수령 완료", "수금완료입니다", "입금완료", "6월15일 272만원 현금 YG 수령완료"
    #   "550만원 VAT포함에서 500만원 VAT별도로 변경", "현재 400만원 현금으로 수금하였음"
    #   "잔금 100만원 남음", "400만원 수금 으로만 체크 부탁드립니다"
    manager_summary_re = re.compile(
        r'(수령\s*완료|수금\s*완료|입금\s*완료|정산\s*완료|VAT|변경|하였음|남음|부탁드립니다|이체|취소|체크\s*부탁|만원권)'
    )

    if not partner:
        for ln in lines:
            if any(p.search(ln) for p in skip_patterns):
                continue
            # 프로젝트 코드 라인 skip (2026-07-10 — 매니저가 메모 첫줄에 코드 붙여넣는 케이스)
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

    # 다른 프로젝트 흡수 케이스 (2026-07-10 G2016-YG 관측)
    # 메모에 '{프로젝트코드} 일부' 같은 참조 있으면 partner 로 세팅
    # 예: "2025-06-16 G2172/2173 일부 140,000원 부가세 포함 수금"
    if not partner:
        _ref_re = re.compile(r'([GRN]\d{4}(?:/\d{4})?)')
        joined_text = '\n'.join(lines)
        m = _ref_re.search(joined_text)
        if m:
            partner = f'참조 {m.group(1)}'

    # 메모 전체에 "현금" 키워드 있고 partner 아직 없으면 partner='현금' fallback (2026-07-10)
    # → _resolve_payment_code 에서 'N' 코드 반환
    if not partner and '현금' in '\n'.join(lines):
        partner = '현금 수령'

    # 현금 수령 free text 케이스는 매니저가 날짜 명시 안 하는 경우 많음.
    # date_md 없으면 오늘 날짜 fallback (매니저가 지금 입력 = 오늘 수령).
    # incomplete 감지에서 skip 되지 않도록 (2026-07-15 R3791-MJ)
    if partner == '현금 수령' and not date_md:
        from datetime import datetime as _dt
        date_md = _dt.now().strftime('%m/%d')

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
        '_meta_only': _has_meta_only,  # 시트값 후처리로 amount 유추 대상
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
    """U/V/W 셀 메모 3개를 모두 파싱 → [{stage:'계약금', ...}, {stage:'중도금', ...}, ...]
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
        # 카톡 양식 라인 1건 — 서술 라인들과 섞인 케이스 (2026-07-11 R3166-YM 관측)
        # 노트에 프로젝트 코드·주소·설명 라인이 섞여있으면 blocks 파서가 이걸
        # meta-only 로 잡아 amount 를 반으로 분배. katok payment 를 채택해서
        # 나머지 서술 블록 skip.
        if len(katok_payments) == 1:
            _kp = katok_payments[0]
            _partner = (_kp.get('partner') or '').strip()
            # partner 노이즈 (숫자·연산자·괄호만) sanitize → 은행 코드 기반 대체
            if _partner and re.match(r'^[\d,()\s*+\-x×.]+$', _partner):
                _bank_code = _kp.get('bank', '')
                if _bank_code == '현금' or _kp.get('note_label') != '박C':
                    _kp['partner'] = '현금 수령' if '현금' in note else ''
            results.append(_kp)
            continue

        # 카톡 양식 라인이 2건 이상이면 카톡 양식으로 처리 (정상 양식과 명확히 다름)
        if len(katok_payments) >= 2:
            # 매니저가 history 누적 — 이전 단계와 중복되는 첫 라인들 제외
            # 예: V 메모의 첫 라인이 U값과 동일하면 계약금 입금 중복
            prev_amounts = {p.get('amount') for p in results}
            while katok_payments and katok_payments[0].get('amount') in prev_amounts:
                katok_payments.pop(0)
            # 이체 중복 제거 (2026-07-10 G3240-HS) — "매출이동" 언급 있으면 같은 date+amount 는 하나만
            if '매출이동' in note or '이체' in note:
                seen = set()
                dedup = []
                for p in katok_payments:
                    key = (p.get('date_md'), p.get('amount'), p.get('partner'))
                    if key not in seen:
                        seen.add(key)
                        dedup.append(p)
                katok_payments = dedup
            results.extend(katok_payments)
            continue

        # 매니저 요약 양식 (2026-07-10 R2951-TH) — "2025-12-05 300,000원 김연종 농협"
        manager_payments = []
        for ln in note.strip().splitlines():
            m = _MANAGER_LINE_RE.match(ln.strip())
            if m:
                mm, dd = int(m.group(1)), int(m.group(2))
                amount = int(m.group(3).replace(',', ''))
                partner_v = m.group(4).strip()
                bank_v = m.group(5)
                manager_payments.append({
                    'date_md': f'{mm:02d}/{dd:02d}',
                    'amount': amount,
                    'partner': partner_v or '-',
                    'bank': bank_v,
                    'note_label': '',
                    'stage': stage,
                })
        if manager_payments:
            prev_amounts = {p.get('amount') for p in results}
            while manager_payments and manager_payments[0].get('amount') in prev_amounts:
                manager_payments.pop(0)
            results.extend(manager_payments)
            continue
        # 1차: 빈 줄로 블록 분리
        raw_blocks = re.split(r'\n\s*\n', note.strip())
        # 2차: 한 블록 안에 헤더 패턴이 2번 이상이면 추가 분리 (빈 줄 없는 분할 입금)
        # 지원 헤더:
        #   - YYYY/MM/DD HH:MM  (은행 SMS full-date 양식)
        #   - 일시 MM/DD, HH:MM  (은행 SMS 압축 양식, 2026-07-11 R3692-MJ)
        #   - 은행,MM/DD, HH:MM  (하나/기업/농협 등 축약, 2026-07-11 R3589-MW)
        #   - 입금일: (매니저 라벨 양식, 2026-07-11 G0278/G1866/G1997)
        #   - MM/DD HH:MM 계좌…  (헤더 없이 날짜만, 2026-07-11 G0323)
        _BLOCK_HEADER_RE = re.compile(
            r'(?:'
            r'\d{4}/\d{1,2}/\d{1,2}\s+\d{1,2}:\d{2}'
            r'|\d{4}[-.]\d{1,2}[-.]\d{1,2}(?=\s|$)'    # YYYY-MM-DD (G0278 관측)
            r'|일시\s+\d{1,2}/\d{1,2}'
            r'|(?:하나|기업|국민|신한|우리|농협|카카오|토스)[,\s]\s*\d{1,2}/\d{1,2}'
            r'|입금일\s*:'
            r'|(?<![\d/])\d{1,2}/\d{1,2}\s+\d{1,2}:\d{2}'
            r')'
        )
        # 추가 감지 — 은행 SMS 양식이지만 헤더 완전 생략, "입금 X원" 반복 (G1897-MW).
        # 규칙: 각 header 뒤 첫 amount 는 그 header 에 소속 (skip). 같은 header 뒤
        # 두 번째 이후 amount 는 새 블록 시작. header 없으면 각 amount 를 시작으로.
        _AMOUNT_LINE_RE = re.compile(r'(?:^|\n)입금\s*[\d,]+원')
        blocks = []
        for rb in raw_blocks:
            header_starts = [m.start() for m in _BLOCK_HEADER_RE.finditer(rb)]
            amt_matches = list(_AMOUNT_LINE_RE.finditer(rb))
            starts = set(header_starts)
            prev_header_seen = None
            for m in amt_matches:
                pos = m.start()
                nearest_header = max(
                    (h for h in header_starts if h <= pos), default=None
                )
                if nearest_header is not None and nearest_header != prev_header_seen:
                    prev_header_seen = nearest_header
                else:
                    starts.add(pos)
            starts = sorted(starts)
            if len(starts) >= 2:
                for i, ds in enumerate(starts):
                    end = starts[i+1] if i+1 < len(starts) else len(rb)
                    blocks.append(rb[ds:end].strip())
            else:
                blocks.append(rb)
        prev_block = ''
        block_count = len(blocks)
        any_parsed = False
        for bidx, block in enumerate(blocks):
            # 매출이동/계좌이체 블록 skip (2026-07-12 G2965-TH 관측)
            #   "2025-12-05 G>N \n 농협 입금2,450,000원..." 은 앞 블록 입금액을
            #   G(기업)→N(농협) 계좌로 옮긴 것. 새 입금 아님. 파싱 결과에서 제외.
            if _TRANSFER_RE.search(block):
                continue
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
                # 단, 두 블록의 amount 가 같아야만 병합 (같은 payment 를 매니저가
                # 빈 줄로 잘못 분리한 케이스). amount 다르면 서로 다른 입금 →
                # 별도 payment 로 추가. (2026-07-11 G2835-JW 회귀 방지)
                if (results and results[-1].get('stage') == stage
                        and not results[-1].get('partner') and prev_block
                        and parsed.get('amount') == results[-1].get('amount')):
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
            # 2026-07-11 서술식 노트 fallback 강화 (R3779, R3520, G3662 등 관측)
            # 노트 텍스트에 '현금' 언급 → partner='현금'
            # 'N통장' 언급 → partner='N통장'
            # 'R>N 매출이동' 언급 → partner='매출이동'
            # 노트에서 date 추출 시도 (YYYY-MM-DD, YYYY/MM/DD, MM/DD, N월 M일)
            _joined = note
            _partner = '-'
            if re.search(r'현금\s*(?:수?령|수?금)?', _joined):
                _partner = '현금 수령'
            elif re.search(r'N\s*통장', _joined):
                _partner = 'N통장'
            elif re.search(r'매출\s*이동|[RG]\s*>\s*N', _joined):
                _partner = '매출이동'
            _date = '-'
            _m_ko = _KO_DATE_RE.search(_joined)
            if _m_ko:
                _date = f'{int(_m_ko.group(1)):02d}/{int(_m_ko.group(2)):02d}'
            else:
                # YYYY-MM-DD 우선 매치 (매니저 노트에 흔한 양식)
                _m_full = re.search(r'(?<!\d)(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?!\d)', _joined)
                if _m_full:
                    _date = f'{int(_m_full.group(2)):02d}/{int(_m_full.group(3)):02d}'
                else:
                    # MM/DD (앞뒤에 숫자 없어야, 예: '03/25')
                    _m_md = re.search(r'(?<!\d)(\d{1,2})/(\d{1,2})(?![:/\d])', _joined)
                    if _m_md:
                        _date = f'{int(_m_md.group(1)):02d}/{int(_m_md.group(2)):02d}'
            results.append({
                'date_md': _date, 'amount': stage_val, 'partner': _partner,
                'bank': '', 'note_label': '', 'stage': stage,
            })

        # 이 stage 의 meta-only 블록에 amount 유추 세팅 (2026-07-10 R2205-YM)
        # 매니저가 라벨 양식만 쓴 블록 (입금일:/입금자: 만) 의 amount 를
        # 시트값 - 정상 파싱 블록 합 으로 계산.
        stage_results = [p for p in results if p.get('stage') == stage]
        meta_only = [p for p in stage_results if p.get('_meta_only')]
        confirmed = [p for p in stage_results if not p.get('_meta_only')]
        if meta_only and stage_val > 0:
            confirmed_sum = sum(p.get('amount', 0) for p in confirmed)
            diff = stage_val - confirmed_sum
            if diff > 0 and len(meta_only) == 1:
                # 단일 meta-only 블록 → 차액을 그대로 세팅
                meta_only[0]['amount'] = diff
                meta_only[0]['_meta_only'] = False
            elif diff > 0 and len(meta_only) > 1:
                # 여러 개 → 균등 분배 (매우 드문 케이스)
                per = diff // len(meta_only)
                for p in meta_only:
                    p['amount'] = per
                    p['_meta_only'] = False
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

    # 매니저 설명 블록 필터 (2026-07-10 G1865-YG 관측)
    # 매니저가 메모에 설명 문장을 여러 줄 쓰면 각 블록이 별도 payment 로 오인됨.
    # 유효 payment 는 date 또는 bank 하나 이상 있어야 함.
    # 새 파서 결과의 date_md='' (빈값) 는 파싱 실패 → 매니저 설명 블록으로 간주.
    # 옛 fallback payment 는 date_md='-' 라 skip 안 됨 (옛 데이터 보존).
    _noise_re = re.compile(r'(VAT|변경|하였음|남음|부탁드립니다|이체|취소|체크\s*부탁|수령\s*완료|수금\s*완료|입금\s*완료|만원권)')
    # 2026-07-11 partner sanitize — amount/원 포함된 partner 는 파싱 오류.
    #   노트에 '현금'/'N통장' 언급 있으면 partner 를 그것으로 교체.
    _bad_partner_re = re.compile(r'[\d,]+\s*원|[GRN]\s*[\d,]+')
    filtered = []
    for p in results:
        partner = p.get('partner') or ''
        has_date = bool(p.get('date_md'))
        has_bank = bool(p.get('bank'))
        # amount 유추 실패한 meta-only 는 skip
        if p.get('_meta_only') and p.get('amount', 0) == 0:
            continue
        if not has_date and not has_bank:
            continue  # 완전 설명 블록 or 파싱 실패 → skip
        if _noise_re.search(partner):
            p['partner'] = ''  # 유효 블록의 오인된 partner → 비움
        elif _bad_partner_re.search(partner):
            # partner 에 amount 표기 섞임 → sanitize
            # 노트 원본에서 '현금'/'N통장' 언급 있으면 그걸로 교체
            _joined_notes = '\n'.join(n or '' for n in notes)
            if re.search(r'현금', _joined_notes):
                p['partner'] = '현금 수령'
            elif re.search(r'N\s*통장', _joined_notes):
                p['partner'] = 'N통장'
            else:
                p['partner'] = ''
        # 내부 필드 제거 (외부 노출 안 함)
        p.pop('_meta_only', None)
        filtered.append(p)

    # 2026-07-11 filter 후 fallback — stage 별로 payment 없으면 노트 있고 stage_val > 0 이면
    #   서술식 노트 (현금 수령 등) 로 판단해서 fallback 추가.
    for stage, stage_val in [
        ('계약금', (stage_vals or {}).get('계약금', 0)),
        ('중도금', (stage_vals or {}).get('중도금', 0)),
        ('잔금', (stage_vals or {}).get('잔금', 0)),
    ]:
        if stage_val <= 0:
            continue
        if any(p.get('stage') == stage for p in filtered):
            continue
        # 이 stage 노트가 있는지
        stage_idx = ['계약금', '중도금', '잔금'].index(stage)
        note = notes[stage_idx] if stage_idx < len(notes) else ''
        if not note:
            continue
        _partner = '-'
        if re.search(r'현금', note):
            _partner = '현금 수령'
        elif re.search(r'N\s*통장', note):
            _partner = 'N통장'
        elif re.search(r'매출\s*이동|[RG]\s*>\s*N', note):
            _partner = '매출이동'
        else:
            continue  # 서술 키워드 없으면 fallback 안 함
        _date = '-'
        _m_ko = _KO_DATE_RE.search(note)
        if _m_ko:
            _date = f'{int(_m_ko.group(1)):02d}/{int(_m_ko.group(2)):02d}'
        else:
            _m_full = re.search(r'(?<!\d)(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?!\d)', note)
            if _m_full:
                _date = f'{int(_m_full.group(2)):02d}/{int(_m_full.group(3)):02d}'
            else:
                _m_md = re.search(r'(?<!\d)(\d{1,2})/(\d{1,2})(?![:/\d])', note)
                if _m_md:
                    _date = f'{int(_m_md.group(1)):02d}/{int(_m_md.group(2)):02d}'
        filtered.append({
            'date_md': _date, 'amount': stage_val, 'partner': _partner,
            'bank': '', 'note_label': '', 'stage': stage,
        })

    # 재정렬
    stage_order2 = {'계약금': 0, '중도금': 1, '잔금': 2}
    def _dk(p):
        mm_dd = p.get('date_md', '0/0')
        try:
            mm, dd = mm_dd.split('/')
            return (int(mm), int(dd))
        except Exception:
            return (0, 0)
    filtered.sort(key=lambda p: (stage_order2.get(p.get('stage', ''), 9), _dk(p)))
    return filtered


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
    # 입금자가 '현금' 또는 '현금 수령' 이면 N으로 강제 (혼합 케이스 매니저 수기 입력)
    if partner and partner.strip() in ('현금', '현금 수령'):
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


def _build_unified_stage_message(
    stage: str,
    projects: List[Dict],           # [{'code','address','sheet_val','construction'}, ...]
    payments: List[Dict],           # 그룹 공통 payments (모든 지점 메모가 동일)
    invoice_value: str,
    is_complete: bool,
) -> str:
    """통합 입금 카드 — 여러 프로젝트가 같은 메모로 통합 입금됐을 때.

    2026-07-11 신규. SK텔레콤 4지점 통합 입금 사례 대응:
    - 매니저가 4개 지점 각 메모에 통합 입금 이력(잔금 353만+209만) 복사
    - 시트 잔금은 개별 배분값 (1,045/1,826/1,045/1,705)
    - 개별 카드 4건 대신 통합 카드 1건으로 발송 → 매니저 즉시 인지
    - 그룹 전체 phash 갱신 → 중복 발송 방지
    """
    n = len(projects)
    codes = [p['code'] for p in projects]
    headline = ':white_check_mark: *수금완료 (통합 입금)*' if is_complete else f':moneybag: *{stage} 통합 입금*'
    # 코드 4개 이하는 전부 나열, 초과분은 "외 N건" 으로 축약. 총 개수는 (총 N건) 로 명시.
    if n <= 4:
        codes_join = ', '.join(codes)
    else:
        codes_join = ', '.join(codes[:4]) + f' 외 {n-4}건'
    lines = [
        '⠀',
        f"{headline} — :id: *{codes_join}* (총 {n}건)",
        _SEP,
    ]

    # 대표 파트너
    partners = sorted({(p.get('partner') or '').strip() for p in payments if p.get('partner')})
    if partners:
        lines.append(f'입금자 : {", ".join(partners)}')

    # 입금 이력 (공통 payments)
    lines.extend(['', '[입금 이력]'])
    for p in payments:
        stg = p.get('stage', '-')
        date_md = p.get('date_md', '-')
        code = _resolve_payment_code(invoice_value, p.get('bank', ''), p.get('partner', ''))
        transfer_to = p.get('transfer_to', '')
        if transfer_to and transfer_to != code:
            code = f"{code} → {transfer_to}"
        note_label = p.get('note_label', '')
        bank = p.get('bank', '') or ''
        inner = f"{code}, {note_label}" if note_label else code
        code_display = f"{bank} ({inner})" if bank else inner
        amount = p.get('amount', 0)
        partner = p.get('partner', '-') or '-'
        lines.append(
            f"{stg}  {date_md}  {code_display}  {amount:,}원  {partner}"
        )

    # 개별 지점 (짧은 주소)
    lines.extend(['', '[개별 지점 배분]'])
    for pj in projects:
        addr = (pj.get('address') or '-')[:50]
        sv = int(pj.get('sheet_val') or 0)
        lines.append(f"• `{pj['code']}`  {addr}  {sv:,}원")

    # 총액 (그룹 시트 합계)
    sheet_sum = sum(int(pj.get('sheet_val') or 0) for pj in projects)
    lines.append(_SEP)
    lines.append(f"통합 입금 총액 : {sheet_sum:,}원")
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
    """U/V/W 셀 메모 3개 fetch."""
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
        logger.error(f"[PAYMENT] 메모 fetch 실패 (row {row}): {exc}", exc_info=True)
        return ['', '', '']


def _to_int_won(s) -> int:
    """'12,100,000' / '₩12,100,000' / '12100000' / 0 → 12100000

    Google Sheets UNFORMATTED_VALUE 는 float 저장 (예: 3499999.9999).
    int() 는 floor 라 3499999 되어 1 부족. round() 로 정수 반올림.
    """
    if s is None or s == '':
        return 0
    if isinstance(s, (int, float)):
        return int(round(s))
    s = str(s).strip().replace('₩', '').replace(',', '').replace(' ', '')
    if not s:
        return 0
    try:
        return int(round(float(s)))
    except Exception:
        return 0


def _to_bool(s) -> bool:
    """체크박스 — True/'TRUE'/'true' → True, 그 외 False"""
    if isinstance(s, bool):
        return s
    return str(s).strip().lower() == 'true'


_SYNC_MUTEX_KEY = 'payment_sync:running'
_SYNC_MUTEX_TTL = 60  # 안전 상한 — sync 는 5-10초 예상, 정상 종료 시 즉시 해제


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

    # Redis mutex — APScheduler 자동 폴링 + 수동 스크립트 동시 실행 방지 (중복 발송 방지).
    #   2026-07-15 관측: 카톡 순서 재발송 스크립트 실행 중 서비스 자동 sync 병렬 진입
    #   → G3635-YG, G3805-YG 등 카드 중복 발송. SET NX 로 동시 진입 차단.
    try:
        _mutex_rc = get_redis_client().redis
        acquired = _mutex_rc.set(_SYNC_MUTEX_KEY, '1', nx=True, ex=_SYNC_MUTEX_TTL)
        if not acquired:
            logger.debug('[PAYMENT] 이미 다른 프로세스에서 sync 실행 중 — skip')
            return result
    except Exception as exc:
        logger.warning(f'[PAYMENT] mutex 획득 실패, sync 진행: {exc}')
        _mutex_rc = None

    try:
        return _sync_payments_locked(result, sheet_id, sheet_name, channel, bot_token)
    finally:
        if _mutex_rc is not None:
            try:
                _mutex_rc.delete(_SYNC_MUTEX_KEY)
            except Exception:
                pass


def _sync_payments_locked(result, sheet_id, sheet_name, channel, bot_token):
    """sync_payments 본체 — mutex 안에서 실행. 원본 함수 시그니처 유지."""
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

    # 2026-07-11 매 폴링마다 전체 메모 fetch. 통합 입금 그룹 감지에 다른 프로젝트
    #   payments 접근 필요. baseline flag 는 유지하되 매번 fetch 실행 (기존 phash 저장
    #   로직 그대로 사용). 시트 API batch 호출 1번 추가 (~300ms).
    all_phash_by_row: Dict[int, str] = {}
    all_payments_by_row: Dict[int, List[Dict]] = {}  # 통합 그룹 감지용
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
                        row_num = offset_n + 2
                        all_phash_by_row[row_num] = _hash_payments(payments_n)
                        all_payments_by_row[row_num] = payments_n
        if not baseline_done:
            logger.info(
                f"[PAYMENT] baseline 메모 fetch 완료: {len(all_phash_by_row)}개 행"
            )
    except Exception as exc:
        logger.error(f"[PAYMENT] 메모 fetch 실패: {exc}", exc_info=True)

    # 2026-07-11 통합 입금 그룹 감지용 signature 인덱스.
    #   (stage, sorted_payments_signature) → [(row, code, sheet_val, address, construction)]
    #   같은 stage 에 완전 동일한 payments 를 가진 프로젝트가 여러 개면 그룹 후보.
    def _payment_signature(payments_: List[Dict], stage_: str) -> tuple:
        return tuple(sorted(
            (p.get('amount', 0), (p.get('partner') or '').strip(), (p.get('date_md') or '').strip())
            for p in payments_ if p.get('stage') == stage_
        ))

    sig_index: Dict[tuple, List[Dict]] = {}
    for _r_num, _ps in all_payments_by_row.items():
        # 시트에서 이 행의 U/V/W 값·코드·주소·공사내용 추출
        _off = _r_num - 2
        if _off < 0 or _off >= len(rows):
            continue
        _row = rows[_off]
        while len(_row) < 27:
            _row.append('')
        _code = str(_row[IDX_A] if IDX_A < len(_row) else '').strip()
        if not _code or not _VALID_PROJECT_RE.match(_code):
            continue
        _u = _to_int_won(_row[IDX_U] if IDX_U < len(_row) else '')
        _v = _to_int_won(_row[IDX_V] if IDX_V < len(_row) else '')
        _w = _to_int_won(_row[IDX_W] if IDX_W < len(_row) else '')
        _addr = str(_row[IDX_F] if IDX_F < len(_row) else '').strip()
        _constr = str(_row[IDX_L] if IDX_L < len(_row) else '').strip()
        for _stage, _sval in (('계약금', _u), ('중도금', _v), ('잔금', _w)):
            _sig = _payment_signature(_ps, _stage)
            if not _sig:
                continue
            sig_index.setdefault((_stage, _sig), []).append({
                'row': _r_num, 'code': _code, 'sheet_val': _sval,
                'address': _addr, 'construction': _constr,
            })

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

        # 값 변경/AA 변경/메모 변경 자체는 baseline 갱신 트리거.
        # 메모 변화 감지 (2026-07-15): 매니저가 값 먼저 저장 → 폴링 → 메모 나중
        # 저장 시나리오. 값·AA 변화 없어서 skip 되던 문제.
        cur_phash = all_phash_by_row.get(sheet_row, '')
        prev_phash_raw = prev.get('phash', '')
        memo_changed = cur_phash != prev_phash_raw
        any_change = (u_val != prev_u) or (v_val != prev_v) or (w_val != prev_w) \
            or (aa_chk != prev_aa) or memo_changed
        if not any_change:
            continue

        # 발송 트리거 — 새 입금(값 증가), AA 신규 체크, 또는 메모 신규 저장 (값 있고 phash 빈→채워짐).
        # 값 감소(정정/취소), AA 해제는 baseline만 갱신하고 skip.
        new_payment = (u_val > prev_u) or (v_val > prev_v) or (w_val > prev_w)
        aa_newly_checked = aa_chk and not prev_aa
        # 메모 신규 저장 감지: 값 이미 있고, 이전 phash 빈값 → 지금 phash 있음
        memo_newly_added = (
            not prev_phash_raw and cur_phash
            and (u_val > 0 or v_val > 0 or w_val > 0)
        )
        if not (new_payment or aa_newly_checked or memo_newly_added):
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
            'memo_newly_added': memo_newly_added,  # 값 있고 memo 방금 저장
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

                # 메모 나중 저장 시나리오 (2026-07-15):
                # 매니저가 값(u/v/w) 먼저 입력 → 폴링(baseline 저장) → 메모 나중 저장.
                # 값 변화 없어 stages_increased 빈 리스트 → skip 발생.
                # 값 있는 stage 를 발송 대상으로 (payments 로부터 어느 stage 에
                # 실제 payment 있는지 확인).
                if c.get('memo_newly_added') and not stages_increased:
                    _stages_with_payment = {p.get('stage') for p in payments}
                    if '계약금' in _stages_with_payment and u_val > 0:
                        stages_increased.append('계약금')
                    if '중도금' in _stages_with_payment and v_val > 0:
                        stages_increased.append('중도금')
                    if '잔금' in _stages_with_payment and w_val > 0:
                        stages_increased.append('잔금')
                    logger.info(
                        f"[PAYMENT] memo_newly_added → stages_increased fallback: "
                        f"{project} row {sheet_row} stages={stages_increased}"
                    )

                # 메모 미완성 감지 (2026-07-10 · 2026-07-10 조건 완화)
                # 매니저 실제 워크플로: U/V/W 금액 먼저 입력 → 메모 나중 입력.
                # 30초 폴링이 그 사이에 걸리면 메모 없는 상태로 발송돼 partner/date 가 비어서 나감.
                #
                # 초기 fix: date_md='-' AND partner='-' (완전 fallback) 만 감지
                # 강화 (G3702-MS 재발 관측): partner 또는 date 하나만 빠져도 skip.
                # 매니저가 메모 부분만 입력한 중간 상태도 감지 대상.
                incomplete = any(
                    (not p.get('partner') or p.get('partner') == '-')
                    or (not p.get('date_md') or p.get('date_md') == '-')
                    for stage in stages_increased
                    for p in payments
                    if p.get('stage') == stage
                )
                if incomplete:
                    logger.info(
                        f"[PAYMENT] 메모 미저장 감지 → skip ({project}, row {sheet_row}). "
                        f"다음 폴링에서 메모 확인 후 재시도"
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

                    # 2026-07-11 통합 입금 그룹 감지
                    #   sig_index 에 같은 stage 시그니처 가진 프로젝트가 여러개면 그룹.
                    #   시트합 = 메모합 이면 통합 입금 확정 → 통합 카드 1건 발송.
                    #   대표(sig_index 첫 프로젝트)만 발송하고 나머지는 phash 만 갱신.
                    _sig = _payment_signature(payments, stage)
                    _group = sig_index.get((stage, _sig), [])
                    _is_unified = False
                    _unified_projects: List[Dict] = []
                    if len(_group) >= 2:
                        _note_sum = sum(p.get('amount', 0) for p in stage_payments)
                        _sheet_sum = sum(int(g['sheet_val'] or 0) for g in _group)
                        if _sheet_sum > 0 and abs(_note_sum - _sheet_sum) / max(_sheet_sum, 1) <= 0.05:
                            _is_unified = True
                            _unified_projects = _group

                    # 2026-07-11 grace 통합 승격 — 이전 폴링에서 발송된 카드가 아직 pending
                    #   (10분 TTL) 이고 자기 시그니처와 일치하면 그 카드에 자기를 추가해
                    #   chat.update 로 통합 카드로 승격. 매니저가 30초 안에 4지점 메모를 다
                    #   못 채우고 순차 입력하는 실사용 패턴 대응.
                    import hashlib as _hl
                    import json as _json
                    _sig_hash = _hl.md5(str(_sig).encode()).hexdigest()[:16]
                    _group_key = f'payment_pending_group:{stage}:{_sig_hash}'
                    _pending: Optional[Dict] = None
                    if not _is_unified and _sig:
                        try:
                            _raw = rc.get(_group_key)
                            if _raw:
                                _pending = _json.loads(_raw if isinstance(_raw, str) else _raw.decode())
                        except Exception:
                            _pending = None

                    if _pending:
                        # 이미 자기 포함된 그룹이면 skip
                        _existing_codes = {p['code'] for p in _pending.get('projects', [])}
                        if project in _existing_codes:
                            logger.debug(
                                f"[PAYMENT] grace group 이미 포함 → skip ({project} row {sheet_row})"
                            )
                            # phash 만 갱신 (다음 폴링 재감지 방지)
                            try:
                                rc.hset(key, mapping={
                                    'u': u_val, 'v': v_val, 'w': w_val,
                                    'aa': 'true' if aa_chk else 'false',
                                    'phash': new_phash,
                                })
                                rc.expire(key, REDIS_TTL)
                            except Exception:
                                pass
                            continue
                        # 자기 정보 추가 → 통합 카드 재렌더 → chat.update
                        _pending['projects'].append({
                            'code': project,
                            'sheet_val': int(stage_vals.get(stage, 0)),
                            'address': c['address'],
                            'construction': c.get('construction', ''),
                        })
                        _updated_text = _build_unified_stage_message(
                            stage=stage,
                            projects=_pending['projects'],
                            payments=_pending.get('payments', stage_payments),
                            invoice_value=_pending.get('invoice_value', c['invoice']),
                            is_complete=(stage == '잔금' and c['unpaid'] == 0),
                        )
                        try:
                            from dashboard.blueprints.slack_helpers import safe_slack_call
                            safe_slack_call(
                                slack.chat_update,
                                channel=_pending['channel'],
                                ts=_pending['ts'],
                                text=_updated_text,
                            )
                            # 그룹 저장 갱신 (TTL 10분 재갱신)
                            rc.set(_group_key, _json.dumps(_pending), ex=10 * 60)
                            # 자기 phash / ts 매핑 저장
                            rc.hset(key, mapping={
                                'u': u_val, 'v': v_val, 'w': w_val,
                                'aa': 'true' if aa_chk else 'false',
                                'phash': new_phash,
                            })
                            rc.expire(key, REDIS_TTL)
                            rc.set(
                                f'payment_slack:ts:{project}:{stage}',
                                _pending['ts'],
                                ex=60 * 60 * 24 * 90,
                            )
                            logger.info(
                                f"[PAYMENT] grace 통합 승격: {project} → 그룹 {len(_pending['projects'])}건 "
                                f"(원 카드 ts={_pending['ts']})"
                            )
                            sent_this_row = True
                            continue
                        except Exception as _upd_exc:
                            logger.warning(
                                f"[PAYMENT] grace chat.update 실패, 개별 발송 fallback "
                                f"({project}): {_upd_exc}"
                            )
                            # 실패 시 개별 발송으로 진행 (아래 로직)

                    if _is_unified:
                        # 대표 아니면 발송 skip (대표가 통합 카드로 4건 다 처리)
                        _rep_row = _unified_projects[0]['row']
                        if sheet_row != _rep_row:
                            # phash 저장 후 skip — 다음 폴링에서 반복 감지 방지.
                            try:
                                rc.hset(key, mapping={
                                    'u': u_val, 'v': v_val, 'w': w_val,
                                    'aa': 'true' if aa_chk else 'false',
                                    'phash': new_phash,
                                })
                                rc.expire(key, REDIS_TTL)
                            except Exception:
                                pass
                            logger.info(
                                f"[PAYMENT] 통합 그룹 대표 아님 → phash 갱신 후 skip "
                                f"({project} row {sheet_row}, 대표 row {_rep_row})"
                            )
                            continue
                        # 통합 카드 발송
                        text = _build_unified_stage_message(
                            stage=stage,
                            projects=_unified_projects,
                            payments=stage_payments,
                            invoice_value=c['invoice'],
                            is_complete=(stage == '잔금' and c['unpaid'] == 0),
                        )
                    elif stage == '잔금' and c['unpaid'] == 0:
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
                    # 2026-07-11 통합 입금 힌트 — 메모 amount 합이 시트 stage 값보다
                    #   훨씬 크면 (1.5배 이상) 다른 프로젝트가 통합 입금 대상일 가능성.
                    #   실사례: SK텔레콤 4지점 잔금 통합 입금 시 G2855 카드에 메모 562만
                    #   기록되지만 총액은 개별 배분값 104만. 매니저가 카드만 봐선 오해.
                    #   카드 하단에 힌트 부착해 통합 입금 여부 즉시 확인 유도.
                    #   통합 카드로 발송하는 경우엔 이미 명확하니 힌트 불필요.
                    if not _is_unified:
                        try:
                            _note_sum = sum(p.get('amount', 0) for p in stage_payments)
                            _sheet_val = int(stage_vals.get(stage, 0))
                            if _sheet_val > 0 and _note_sum >= _sheet_val * 1.5:
                                text += (
                                    '\n\n:warning: *_총액은 시트 개별 배분값입니다._*\n'
                                    f'_메모 입금 합 {_note_sum:,}원 vs 시트 값 {_sheet_val:,}원._\n'
                                    '_같은 사업자의 다른 프로젝트가 통합 입금 대상일 수 있어요._'
                                )
                        except Exception:
                            pass

                    # 2026-07-11 특이사항 라인 표시 (G1019-MW 관측) — 노트에 '제외/차감/
                    #   채권추심/반환/안분' 키워드가 있는 라인은 매니저가 정상 payment 로 잡히지
                    #   않는 이유를 서술한 것. `[입금 이력]` 블록 아래, 마지막 구분선 위에
                    #   `[특이사항]` 섹션으로 표시해서 매니저가 이력 합·시트 총액 차이 이유를
                    #   즉시 파악하도록. 앞의 날짜/시간 부분은 제거하고 stage 접두어 추가.
                    #   2026-07-12 확장 — 매출이동/계좌이체 라인 ("YYYY-MM-DD G>N") 도
                    #   특이사항으로 표시. 파서는 이체 블록을 skip 하지만 매니저 카드에는
                    #   회계 이력으로 남아야 함 (G2965-TH 관측).
                    try:
                        _stage_by_idx = ('계약금', '중도금', '잔금')
                        _date_prefix_re = re.compile(
                            r'^(?:\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2})\s*'
                            r'(?:\d{1,2}:\d{2}\s*)?'
                        )
                        _special_re = re.compile(r'(?:제외|차감|채권추심|반환|안분|상계|매입|환불|재입금|정정)')
                        _transfer_md_re = re.compile(r'^\d{4}-(\d{1,2})-(\d{1,2})')
                        _special_entries = []  # [(stage, cleaned_line)]
                        for _idx, _note in enumerate(notes):
                            if not _note:
                                continue
                            _stage_name = _stage_by_idx[_idx] if _idx < 3 else ''
                            for _ln in _note.splitlines():
                                _ln = _ln.strip()
                                if not _ln or _ln.startswith('입금 '):
                                    continue
                                _tr = _TRANSFER_RE.search(_ln)
                                if _tr:
                                    _md_m = _transfer_md_re.match(_ln)
                                    _md = f'{int(_md_m.group(1)):02d}/{int(_md_m.group(2)):02d}' if _md_m else ''
                                    _a, _b = _tr.group(1).upper(), _tr.group(2).upper()
                                    _cleaned = f'{_md} {_a}>{_b} 매출이동'.strip()
                                elif _special_re.search(_ln):
                                    # 날짜/시간 prefix 제거
                                    _cleaned = _date_prefix_re.sub('', _ln).strip()
                                else:
                                    continue
                                if not _cleaned:
                                    continue
                                _entry = (_stage_name, _cleaned)
                                if _entry not in _special_entries:
                                    _special_entries.append(_entry)
                        if _special_entries:
                            _special_block = ['', '[특이사항]']
                            for _stg, _cln in _special_entries[:5]:
                                _prefix = f'{_stg} ' if _stg else ''
                                _special_block.append(f'{_prefix}{_cln}')
                            _special_text = '\n'.join(_special_block)
                            # 마지막 구분선 위에 삽입
                            _parts = text.rsplit(_SEP, 1)
                            if len(_parts) == 2:
                                text = _parts[0].rstrip() + '\n' + _special_text + '\n' + _SEP + _parts[1]
                            else:
                                text += '\n' + _special_text
                    except Exception:
                        pass

                    # 2026-07-15 rollback: 초기 설계 (122a5bb, 각 stage 완전 독립 카드)
                    #   로 복귀. P2-4 (fd68d48) 스레드 답글 로직 제거 — 프로젝트별 이력
                    #   통합은 카드 자체 [누적 이력] 섹션이 이미 담당.
                    # 2026-07-10 safe_slack_call — 결제 알림 놓치면 매니저가 실입금 인지
                    #   실패. 429·5xx 자동 재시도 필수.
                    from dashboard.blueprints.slack_helpers import safe_slack_call
                    resp = safe_slack_call(
                        slack.chat_postMessage, channel=channel, text=text,
                    )
                    result['sent'] += 1
                    logger.info(
                        f"[PAYMENT] {stage} 입금 발송: {project} (row {sheet_row})"
                    )
                    # 발송 성공 시 ts 저장 → 나중에 chat.update 정정 + 스레드 연결 (P0-2 · P2-4)
                    try:
                        if resp and resp.get('ok'):
                            ts = resp.get('ts', '')
                            if ts:
                                rc.set(
                                    f'payment_slack:ts:{project}:{stage}',
                                    ts,
                                    ex=60 * 60 * 24 * 90,  # 90일 보관
                                )
                                # 2026-07-11 통합 카드였으면 그룹 전체 프로젝트에도 동일 ts 저장.
                                #   각 지점 코드로 카드 조회하는 후속 처리 (편집 답글 등) 가
                                #   같은 카드를 참조하도록 → 그룹 어느 코드든 lookup 시 매칭.
                                #   phash 도 그룹 전체 갱신해서 이후 폴링 중복 발송 방지.
                                if _is_unified:
                                    for _g in _unified_projects:
                                        _row_num = _g['row']
                                        _code_g = _g['code']
                                        if _code_g == project:
                                            continue  # 대표는 이미 처리됨
                                        try:
                                            rc.set(
                                                f'payment_slack:ts:{_code_g}:{stage}',
                                                ts,
                                                ex=60 * 60 * 24 * 90,
                                            )
                                            # 그룹 전체 phash 갱신 (개별 발송 중복 방지)
                                            _row_key = f"{REDIS_KEY_PREFIX}{_row_num}"
                                            rc.hset(_row_key, mapping={
                                                'phash': all_phash_by_row.get(_row_num, ''),
                                            })
                                            rc.expire(_row_key, REDIS_TTL)
                                        except Exception:
                                            pass
                                    logger.info(
                                        f"[PAYMENT] 통합 카드 발송 후 그룹 phash/ts 갱신: "
                                        f"{len(_unified_projects)}건 (대표 {project})"
                                    )

                                # 2026-07-11 grace 그룹 pending 저장 — 이후 10분 안에 다른
                                #   지점에서 같은 시그니처 감지되면 이 카드에 chat.update.
                                #   개별 발송이든 통합 발송이든 저장 (통합이면 이미 그룹 완성이지만
                                #   신규 지점이 추가로 뜰 수도 있으니).
                                try:
                                    _pending_projects = (
                                        [{'code': _g['code'], 'sheet_val': int(_g['sheet_val'] or 0),
                                          'address': _g.get('address', ''),
                                          'construction': _g.get('construction', '')}
                                         for _g in _unified_projects]
                                        if _is_unified
                                        else [{'code': project,
                                               'sheet_val': int(stage_vals.get(stage, 0)),
                                               'address': c['address'],
                                               'construction': c.get('construction', '')}]
                                    )
                                    _pending_data = {
                                        'ts': ts,
                                        'channel': channel,
                                        'invoice_value': c['invoice'],
                                        'projects': _pending_projects,
                                        'payments': stage_payments,
                                    }
                                    rc.set(
                                        _group_key,
                                        _json.dumps(_pending_data, ensure_ascii=False),
                                        ex=10 * 60,  # 10분 grace
                                    )
                                except Exception as _p_exc:
                                    logger.debug(
                                        f'[PAYMENT] grace 저장 실패 ({project}/{stage}): {_p_exc}'
                                    )
                    except Exception as _exc:
                        logger.debug(f'[PAYMENT] ts 저장 실패 ({project}/{stage}): {_exc}')
                    sent_this_row = True

            if not sent_this_row:
                logger.debug(
                    f"[PAYMENT] skip ({project}, row {sheet_row}) — note_changed={note_changed}"
                )

            # 매니저 실수 감지 + 알림 훅 (2026-07-10)
            # 카테고리: 메모 미저장/자릿수 오타/필수 항목 누락/미체크/미수금 이상/이니셜 오타
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
            # Phash 이력 저장 (P2-2, 2026-07-10) — 디버깅용, 최근 20개만 유지
            # payment_sync:row:{row}:phash_history LIST — "{ts}|{phash}|sent={0/1}"
            try:
                import time as _t
                hist_key = f'{c["key"]}:phash_history'
                entry = f'{int(_t.time())}|{new_phash}|sent={1 if sent_this_row else 0}'
                rc.lpush(hist_key, entry)
                rc.ltrim(hist_key, 0, 19)
                rc.expire(hist_key, REDIS_TTL)
            except Exception:
                pass
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
