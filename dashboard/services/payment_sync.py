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

# 금액 추출: "입금 12,100,000원" or "입금12,100,000원"
_AMOUNT_RE = re.compile(r'입금\s*([\d,]+)\s*원')
# 라벨 양식 (앱스크립트 자동 + 매니저 수기 덮어쓰기)
_LABEL_DATE_RE = re.compile(r'^입금일\s*:\s*(.+)$')
_LABEL_PAYER_RE = re.compile(r'^입금자\s*:\s*(.+)$')
# 단일 숫자(콤마 OK) 라인 — 매니저가 입금액만 적은 경우
_BARE_AMOUNT_RE = re.compile(r'^([\d,]+)(?:\s*원)?$')
# 날짜: "2026/03/19", "06/25", "5/27" 등 — MM/DD 추출
_DATE_RE = re.compile(r'(?:(\d{4})[/.-])?(\d{1,2})[/.-](\d{1,2})')
# 은행명 추출 — 라인 또는 첫줄 시작
_BANK_RE = re.compile(r'(기업|하나|국민|신한|우리|농협|카카오|토스)')


def _parse_memo_block(block: str) -> Optional[Dict]:
    """입금 메모 한 블록 → {'date_md': 'MM/DD', 'amount': int, 'partner': str, 'bank': str}"""
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

    # 날짜 — 모든 라인에서 첫 매치
    date_md = ''
    for ln in lines:
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
    if amount == 0:
        return None  # 금액 없으면 입금 블록 아님

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

    # 은행
    bank = ''
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
        re.compile(r'^(기업|하나|국민|신한|우리|농협|카카오|토스)(\s+입금)?'),
    ]
    if not partner:
        for ln in lines:
            if any(p.search(ln) for p in skip_patterns):
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

    # 박C 표기 — 대표님 개인 기업통장(추적용 구분)
    note_label = ''
    if '박C' in block:
        note_label = '박C'

    return {
        'date_md': date_md,
        'amount': amount,
        'partner': partner,
        'bank': bank,
        'note_label': note_label,
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


def _parse_notes(notes: List[str]) -> List[Dict]:
    """U/V/W 셀 노트 3개를 모두 파싱 → [{stage:'계약금', ...}, {stage:'중도금', ...}, ...]"""
    results = []
    stages = ['계약금', '중도금', '잔금']
    for note, stage in zip(notes, stages):
        if not note:
            continue
        # 빈 줄로 블록 분리
        blocks = re.split(r'\n\s*\n', note.strip())
        for block in blocks:
            parsed = _parse_memo_block(block)
            if parsed:
                parsed['stage'] = stage
                results.append(parsed)
    # 날짜순 정렬 (월/일 기준) — 양식에 연도 없으면 동일 연도 가정
    def _date_key(p):
        mm_dd = p.get('date_md', '0/0')
        try:
            mm, dd = mm_dd.split('/')
            return (int(mm), int(dd))
        except Exception:
            return (0, 0)
    results.sort(key=_date_key)
    return results


# ─────────────────────────────────────────────
# 한 글자 표시 (G/N/R) 결정
# ─────────────────────────────────────────────

_CARD_PARTNER_RE = re.compile(r'^[0-9A-Z가-힣]{6,18}$')


def _is_card_payment(invoice_value: str, partner: str) -> bool:
    """Y열 + 거래처 패턴으로 카드 결제 여부 판별."""
    iv = (invoice_value or '').strip()
    if iv == '카드결제':
        return True
    if iv == '혼합':
        # 혼합 — 거래처가 영숫자 코드(승인번호 패턴)면 카드
        return bool(partner and _CARD_PARTNER_RE.match(partner.strip()))
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
    return 'G'


# ─────────────────────────────────────────────
# 메시지 빌더
# ─────────────────────────────────────────────

_STAGE_EMOJI = {
    '계약금': ':moneybag:',
    '중도금': ':moneybag:',
    '잔금': ':moneybag:',
}
_SEP = '---------------------------------------------'


def _build_stage_message(
    stage: str, project: str, address: str,
    payment: Dict, invoice_value: str,
    total_r: int, total_t: int, unpaid: int,
) -> str:
    """단계별 입금 알림 (계약금/중도금/잔금)."""
    emoji = _STAGE_EMOJI.get(stage, ':moneybag:')
    is_card = _is_card_payment(invoice_value, payment.get('partner', ''))
    code = _resolve_payment_code(invoice_value, payment.get('bank', ''), payment.get('partner', ''))
    note_label = payment.get('note_label', '')
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
        f"{emoji} *{stage} {header_action}* — `{project}`",
        _SEP,
        f"주소 : {address or '-'}",
        f"{date_label} : {date_md}",
        f"{amount_label} : {amount:,}원",
        f"{partner_label} : {partner}",
        f"은행 : {bank} ({code_display})",
        f"미수금 : {unpaid:,}원",
    ]
    if is_card:
        # T = R × 1.1 (S 체크 시 부가세 10% 포함 금액)
        # 카드 결제 시:
        #   실결제 = T × 1.03 (고객에게 3% 추가 부담시킴, 우리 수익원)
        #   3% = T × 0.03 (고객 부담 추가분)
        #   카드 수수료 = 실결제 - 입금 (우리가 카드사에 내는 실제 수수료 ≈ 2.N%)
        real_payment = round(total_t * 1.03)
        extra_3pct = real_payment - total_t
        fee = real_payment - amount
        if fee > 0 and extra_3pct > 0:
            parts = [
                f"실결제 {real_payment:,}원",
                f"3% {extra_3pct:,}원",
                f"카드 수수료 {fee:,}원",
            ]
            lines.append(f"({' / '.join(parts)})")
    lines.append(_SEP)
    return '\n'.join(lines)


def _build_stage_with_history_message(
    stage: str, project: str, address: str,
    last_payment: Dict, all_payments: List[Dict], invoice_value: str,
    total_r: int, total_t: int, unpaid: int,
) -> str:
    """단계 카드 + 누적 이력 (중도금 입금 시 사용)."""
    emoji = _STAGE_EMOJI.get(stage, ':moneybag:')
    is_card = _is_card_payment(invoice_value, last_payment.get('partner', ''))
    code = _resolve_payment_code(invoice_value, last_payment.get('bank', ''), last_payment.get('partner', ''))
    note_label = last_payment.get('note_label', '')
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
        f"{emoji} *{stage} {header_action}* — `{project}`",
        _SEP,
        f"주소 : {address or '-'}",
        f"{date_label} : {date_md}",
        f"{amount_label} : {amount:,}원",
        f"{partner_label} : {partner}",
        f"은행 : {bank} ({code_display})",
        f"미수금 : {unpaid:,}원",
    ]
    if is_card:
        real_payment = round(total_t * 1.03)
        extra_3pct = real_payment - total_t
        fee = real_payment - amount
        if fee > 0 and extra_3pct > 0:
            lines.append(
                f"(실결제 {real_payment:,}원 / 3% {extra_3pct:,}원 / 카드 수수료 {fee:,}원)"
            )
    lines.append('')
    lines.append('[누적 이력]')
    for p in all_payments:
        st = p.get('stage', '-')
        d = p.get('date_md', '-')
        c = _resolve_payment_code(invoice_value, p.get('bank', ''), p.get('partner', ''))
        nl = p.get('note_label', '')
        c_disp = f"{c}({nl})" if nl else c
        a = p.get('amount', 0)
        pt = p.get('partner', '-') or '-'
        is_c = _is_card_payment(invoice_value, pt)
        suffix = ' (카드)' if is_c else ''
        lines.append(f"{st}  {d}  {c_disp}  {a:,}원  {pt}{suffix}")
    lines.append(_SEP)
    return '\n'.join(lines)


def _build_complete_message(
    project: str, address: str,
    payments: List[Dict], invoice_value: str, total_t: int,
) -> str:
    """수금완료 알림 — 전체 history 취합."""
    lines = [
        f":white_check_mark: *수금완료* — `{project}`",
        _SEP,
        f"주소 : {address or '-'}",
        '',
        '[입금 이력]',
    ]
    for p in payments:
        stage = p.get('stage', '-')
        date_md = p.get('date_md', '-')
        code = _resolve_payment_code(invoice_value, p.get('bank', ''), p.get('partner', ''))
        note_label = p.get('note_label', '')
        code_display = f"{code}({note_label})" if note_label else code
        amount = p.get('amount', 0)
        partner = p.get('partner', '-') or '-'
        is_card = _is_card_payment(invoice_value, partner)
        suffix = ' (카드)' if is_card else ''
        lines.append(
            f"{stage}  {date_md}  {code_display}  {amount:,}원  {partner}{suffix}"
        )
    lines.append(_SEP)
    lines.append(f"총액 : {total_t:,}원")
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
            notes = _fetch_row_notes(sheet_id, sheet_name, sheet_row)
            payments = _parse_notes(notes)
            new_phash = _hash_payments(payments)
            prev_phash = c.get('phash', '')
            note_changed = new_phash != prev_phash

            prev_u, prev_v, prev_w = c['prev_u'], c['prev_v'], c['prev_w']
            prev_aa = c['prev_aa']
            u_val, v_val, w_val, aa_chk = c['u'], c['v'], c['w'], c['aa']

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

                for stage in stages_increased:
                    stage_payments = [p for p in payments if p.get('stage') == stage]
                    if not stage_payments:
                        continue
                    last_payment = stage_payments[-1]
                    if stage == '잔금' and c['unpaid'] == 0:
                        # 잔금 + 미수금 0 → 수금완료
                        text = _build_complete_message(
                            project=project, address=c['address'],
                            payments=payments, invoice_value=c['invoice'],
                            total_t=c['total_t'],
                        )
                    elif stage in ('중도금', '잔금'):
                        # 중도금, 또는 잔금 부족 입금(수금중) → 단계 카드 + history
                        text = _build_stage_with_history_message(
                            stage=stage, project=project, address=c['address'],
                            last_payment=last_payment, all_payments=payments,
                            invoice_value=c['invoice'],
                            total_r=c['total_r'], total_t=c['total_t'],
                            unpaid=c['unpaid'],
                        )
                    else:  # 계약금
                        text = _build_stage_message(
                            stage=stage, project=project, address=c['address'],
                            payment=last_payment, invoice_value=c['invoice'],
                            total_r=c['total_r'], total_t=c['total_t'],
                            unpaid=c['unpaid'],
                        )
                    slack.chat_postMessage(channel=channel, text=text)
                    result['sent'] += 1
                    sent_this_row = True
                    logger.info(
                        f"[PAYMENT] {stage} 입금 발송: {project} (row {sheet_row})"
                    )

            if not sent_this_row:
                logger.debug(
                    f"[PAYMENT] skip ({project}, row {sheet_row}) — note_changed={note_changed}"
                )

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

    return result
