"""
당근 비즈프로필 리드폼 엑셀 자동 처리

- 사업자번호로 암호화된 .xlsx 복호화 (msoffcrypto)
- 7열 (응답일시/이름/연락처/장소/기기/주소/문의) → 우리 시트 15열 매핑
- 증분 추출: 시트의 당근 리드 마지막 응답일시 이후만
- 연락처 정규화 + 만료(`만료됨`) 감지 + 주소 의심 감지
- 시트에 일괄 등록 (리드 No 자동 발번)
"""

import io
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import msoffcrypto

from dashboard.services.lead_service import (
    load_leads_data,
    _get_sheet_config,
    get_sheets_manager,
    LEAD_COLUMN_ORDER,
    invalidate_leads_cache,
)
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)


# ─────────────────────────────────────────────────────────────
# 상수
# ─────────────────────────────────────────────────────────────
KARROT_PLATFORM = '당근'
DEFAULT_STATUS = '상담 대기'

# 당근 엑셀 컬럼명 (원본 그대로)
KCOL_CONSULT = '응답 일시'
KCOL_NAME = '이름'
KCOL_PHONE = '연락처'
KCOL_PLACE = '설치 희망하시는 장소를 선택해주세요'
KCOL_DEVICE = '설치 원하시는 기기 종류를 선택해 주세요'
KCOL_ADDRESS = '방문 견적 받으실 주소를 입력해 주세요'
KCOL_INQUIRY = '문의 내용을 간단하게 남겨주세요'


# ─────────────────────────────────────────────────────────────
# 1. 복호화
# ─────────────────────────────────────────────────────────────
def decrypt_karrot_xlsx(file_bytes: bytes, business_number: str) -> io.BytesIO:
    """사업자번호로 암호화된 xlsx 복호화. 하이픈/공백 자동 제거."""
    pw = re.sub(r'\D', '', str(business_number))
    if not pw:
        raise ValueError('사업자번호가 비어있거나 숫자가 없습니다')

    src = io.BytesIO(file_bytes)
    out = io.BytesIO()
    of = msoffcrypto.OfficeFile(src)
    of.load_key(password=pw)
    of.decrypt(out)
    out.seek(0)
    return out


def parse_karrot_excel(decrypted_buf: io.BytesIO) -> pd.DataFrame:
    """복호화된 엑셀 → DataFrame (첫 번째 시트)"""
    return pd.read_excel(decrypted_buf, engine='openpyxl')


# ─────────────────────────────────────────────────────────────
# 2. 정규화 함수들
# ─────────────────────────────────────────────────────────────
def normalize_phone(raw: Any) -> Tuple[str, bool]:
    """연락처 정규화. 반환: (정규화_문자열, is_expired)"""
    s = str(raw).strip() if raw is not None else ''
    if s in ('', 'nan', 'NaN', 'None', '만료됨'):
        return '', (s == '만료됨')

    digits = re.sub(r'\D', '', s)
    if len(digits) == 11 and digits.startswith('010'):
        return f'{digits[:3]}-{digits[3:7]}-{digits[7:]}', False
    if len(digits) == 10 and digits.startswith('02'):
        return f'{digits[:2]}-{digits[2:6]}-{digits[6:]}', False
    if len(digits) == 10:
        return f'{digits[:3]}-{digits[3:6]}-{digits[6:]}', False
    return s, False


def is_suspicious_address(addr: str) -> bool:
    """주소 칸에 주소답지 않은 텍스트가 들어왔는지 추정"""
    if not addr or addr.strip() in ('', 'nan', 'NaN'):
        return True
    # 한국 주소 키워드
    keywords = ['시', '구', '군', '동', '읍', '면', '리', '로', '길', '번지', '아파트', '빌딩', '층', '호']
    has_keyword = any(k in addr for k in keywords)
    # 너무 짧거나(주소는 보통 10자 이상) 너무 길면(상황설명) 의심
    if len(addr) > 80:
        return True
    return not has_keyword


def _parse_consult_dt(s: Any) -> Optional[datetime]:
    """당근 형식 '2026.03.23. 13:31' → datetime. 실패 시 None."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    if isinstance(s, datetime):
        return s
    s = str(s).strip()
    if not s or s in ('nan', 'NaN'):
        return None
    for fmt in ('%Y.%m.%d. %H:%M', '%Y.%m.%d %H:%M',
                '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M',
                '%Y/%m/%d %H:%M'):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


# ─────────────────────────────────────────────────────────────
# 3. 시트 기존 데이터 인덱싱 (증분 + 중복 처리용)
# ─────────────────────────────────────────────────────────────
def get_last_processed_consult_time(sheet_df: pd.DataFrame) -> Optional[datetime]:
    """시트의 당근 리드 중 가장 최근 응답 일시 (디버그/통계용으로 유지)"""
    if sheet_df is None or sheet_df.empty or '플랫폼' not in sheet_df.columns:
        return None
    karrot_rows = sheet_df[sheet_df['플랫폼'].astype(str).str.strip() == KARROT_PLATFORM]
    if karrot_rows.empty or '상담 시간' not in karrot_rows.columns:
        return None
    times = [_parse_consult_dt(t) for t in karrot_rows['상담 시간'].dropna().astype(str)]
    times = [t for t in times if t is not None]
    return max(times) if times else None


def get_karrot_existing_keys(sheet_df: pd.DataFrame) -> set:
    """
    시트의 당근 리드들의 (정규화 연락처, 응답일시) 튜플 set.
    dedup 핵심 — 같은 (연락처, 시간) 조합이면 이미 등록된 것.
    당근 플랫폼만 대상 (같은 고객이 다른 채널로 별도 문의했으면 별개 리드로 유지).
    """
    keys = set()
    if sheet_df is None or sheet_df.empty:
        return keys
    if '플랫폼' not in sheet_df.columns:
        return keys
    karrot_rows = sheet_df[sheet_df['플랫폼'].astype(str).str.strip() == KARROT_PLATFORM]
    if karrot_rows.empty:
        return keys

    phone_col = '고객 연락처' if '고객 연락처' in karrot_rows.columns else (
        '연락처' if '연락처' in karrot_rows.columns else None
    )
    time_col = '상담 시간' if '상담 시간' in karrot_rows.columns else None
    if phone_col is None or time_col is None:
        return keys

    for _, row in karrot_rows.iterrows():
        phone_raw = row.get(phone_col, '')
        time_raw = row.get(time_col, '')
        phone_digits = re.sub(r'\D', '', str(phone_raw)) if pd.notna(phone_raw) else ''
        time_str = str(time_raw).strip() if pd.notna(time_raw) else ''
        if phone_digits or time_str:
            keys.add((phone_digits, time_str))
    return keys


# ─────────────────────────────────────────────────────────────
# 4. 행 단위 매핑 (당근 7열 → 우리 15열)
# ─────────────────────────────────────────────────────────────
def map_karrot_row_to_lead(row: pd.Series) -> Dict[str, Any]:
    """당근 한 행 → 우리 시트 형식 dict (+ 내부 메타)"""
    def _s(key):
        v = row.get(key, '')
        s = str(v).strip() if v is not None else ''
        return '' if s in ('nan', 'NaN', 'None') else s

    consult_time = _s(KCOL_CONSULT)
    name = _s(KCOL_NAME)
    raw_phone = row.get(KCOL_PHONE, '')
    place = _s(KCOL_PLACE)
    device = _s(KCOL_DEVICE)
    address = _s(KCOL_ADDRESS)
    inquiry = _s(KCOL_INQUIRY)

    phone, is_expired = normalize_phone(raw_phone)

    # 상담 내용: 장소 + 기기 + 문의 합치기
    parts = []
    if place: parts.append(f'장소: {place}')
    if device: parts.append(f'기기: {device}')
    if inquiry: parts.append(f'문의: {inquiry}')
    content = ' / '.join(parts) if parts else ''

    # 키워드: 기기 + 첫 장소
    kw = []
    if device: kw.append(device)
    if place:
        first_place = place.split(' / ')[0].strip()
        if first_place: kw.append(first_place)
    if is_expired:
        kw.insert(0, '만료')
    keyword = ', '.join(kw)

    return {
        '리드 No': '',  # 시트 등록 시 자동 발번
        '상담 시간': consult_time,
        '플랫폼': KARROT_PLATFORM,
        '상태': DEFAULT_STATUS,
        '방문 예정일': '-',
        '고객 연락처': phone,
        '이메일': '-',
        '고객명': name,
        '방문 주소': address,
        '상담 내용': content,
        '키워드': keyword,
        '온라인 상담자': '',
        '영업 담당자': '',
        '마지막 연락일': '',
        '피드백': '',
        # 메타 — 슬랙 메시지 양식용 (시트 등록 시 LEAD_COLUMN_ORDER 필터링으로 자동 제외)
        '_meta_place': place,
        '_meta_device': device,
        '_meta_inquiry': inquiry,
    }


# ─────────────────────────────────────────────────────────────
# 5. 전체 파이프라인
# ─────────────────────────────────────────────────────────────
def process_karrot_excel(file_bytes: bytes, business_number: str) -> Dict[str, Any]:
    """
    당근 엑셀 전체 처리 (복호화 → 매핑 → 증분 추출 → dedup).
    아직 시트에 등록은 안 함. append_leads_to_sheet() 별도 호출 필요.

    Returns:
        {
            'total': 430,                       # 엑셀 전체 행 수
            'new_count': 12,                    # 신규 등록 대상 수
            'new': [...lead dicts...],          # 신규 리드 목록 (메타 제거됨)
            'duplicates': 343,                  # 이미 있는 연락처 (스킵)
            'old_skipped': 75,                  # 마지막 처리 시점 이전 (증분 스킵)
            'expired_count': 1,                 # 만료된 연락처 (신규 중)
            'suspicious_count': 2,              # 주소 의심 (신규 중)
            'last_processed': '2026.05.15 14:23' or None,
        }
    """
    decrypted = decrypt_karrot_xlsx(file_bytes, business_number)
    df = parse_karrot_excel(decrypted)
    total = len(df)

    sheet_df = load_leads_data(force_refresh=True)
    last_dt = get_last_processed_consult_time(sheet_df)  # 디버그/통계용 (필터링에는 사용 안 함)
    karrot_existing_count = 0
    if sheet_df is not None and not sheet_df.empty and '플랫폼' in sheet_df.columns:
        karrot_existing_count = int(
            (sheet_df['플랫폼'].astype(str).str.strip() == KARROT_PLATFORM).sum()
        )

    # dedup 키 = (정규화 연락처 digits, 응답 일시 문자열)
    existing_keys = get_karrot_existing_keys(sheet_df)

    new_leads: List[Dict[str, Any]] = []
    duplicates = 0
    expired_count = 0
    suspicious_count = 0

    for _, row in df.iterrows():
        # 행의 식별 키 생성 (시트 형식과 동일하게)
        phone_raw = row.get(KCOL_PHONE, '')
        time_raw = row.get(KCOL_CONSULT, '')
        time_str = str(time_raw).strip() if pd.notna(time_raw) else ''
        phone_digits = re.sub(r'\D', '', str(phone_raw)) if pd.notna(phone_raw) else ''

        # 연락처 만료된 경우 phone_digits='' 이라 시간으로만 비교
        key = (phone_digits, time_str)

        if key in existing_keys:
            duplicates += 1
            continue

        # 같은 엑셀 내 중복도 차단
        existing_keys.add(key)

        lead = map_karrot_row_to_lead(row)
        # _meta_* 는 슬랙 메시지에서 사용. 시트 등록 시 LEAD_COLUMN_ORDER 필터로 자동 제외됨.

        new_leads.append(lead)

    logger.info(
        f'[KARROT] 처리 결과: total={total} new={len(new_leads)} dup={duplicates} '
        f'sheet_karrot_count={karrot_existing_count}'
    )

    # 응답 시각 오름차순 정렬 → 슬랙에 보낼 때 가장 최신이 마지막 메시지가 되도록
    new_leads.sort(key=lambda l: _parse_consult_dt(l.get('상담 시간')) or datetime.min)

    return {
        'total': total,
        'new_count': len(new_leads),
        'new': new_leads,
        'duplicates': duplicates,
        'sheet_karrot_count': karrot_existing_count,
        'last_processed': last_dt.strftime('%Y.%m.%d %H:%M') if last_dt else None,
    }


# ─────────────────────────────────────────────────────────────
# 6. 시트 일괄 등록
# ─────────────────────────────────────────────────────────────
def append_leads_to_sheet(leads: List[Dict[str, Any]]) -> List[str]:
    """
    여러 리드를 시트에 일괄 추가. 자동 발번된 리드 No 리스트 반환.
    내부적으로 append_row 반복 (Google Sheets API).
    """
    if not leads:
        return []

    cfg = _get_sheet_config()
    if cfg is None:
        raise RuntimeError('ONLINE_LEADS_SHEET_ID 환경변수 미설정')

    manager = get_sheets_manager()

    # 마지막 리드 No 조회 → 시퀀스 계산
    df = load_leads_data(force_refresh=True)
    max_num = 0
    if df is not None and not df.empty and '리드 No' in df.columns:
        for ln in df['리드 No'].dropna().astype(str):
            s = ln.strip()
            if not s.upper().startswith('L'):
                continue
            digits = re.sub(r'\D', '', s)
            try:
                max_num = max(max_num, int(digits))
            except ValueError:
                continue

    lead_nos = []
    for i, lead in enumerate(leads, start=1):
        new_no = f'L-{max_num + i:05d}'
        lead['리드 No'] = new_no
        lead_nos.append(new_no)

    # 헤더 순서대로 행 데이터 + 시트 append
    for lead in leads:
        row = [lead.get(col, '') for col in LEAD_COLUMN_ORDER]
        manager.append_row(cfg['sheet_id'], cfg['sheet_name'], row)

    invalidate_leads_cache()
    logger.info(f'[KARROT] 시트 일괄 등록 완료: {len(leads)}건 (시작 No: {lead_nos[0]})')
    return lead_nos
