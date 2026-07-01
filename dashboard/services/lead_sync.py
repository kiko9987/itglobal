"""
외부 인입 채널 → 메인 시트 + 슬랙 자동 동기화

- 당근 자동 연동 시트 폴링 → 신규 행을 메인 시트에 추가 + 슬랙 알림
- 향후: Gmail 인입 (홈페이지 문의), 기타 플랫폼 동일 패턴

연락처 dedup 정책:
  - 메인 시트의 모든 연락처(digits만) set에 모음
  - 인입 시트의 각 행 연락처가 set에 없으면 신규 등록
  - 만료 연락처(빈값)는 자동으로 set에 안 들어가서 매번 중복 카운트 X
    → 사용자 결정: 만료된 옛 데이터는 자연스럽게 한 번 등록되고 끝
"""

import os
import re
import textwrap
import threading
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

import pandas as pd

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

# 당근 시트 컬럼명 (실제 새 시트 헤더 그대로)
KCOL_CONSULT = '응답 일시'
KCOL_NAME = '이름'
KCOL_PHONE = '연락처'
KCOL_PLACE = '설치 희망 장소를 선택해주세요'
KCOL_DEVICE = '설치 희망 기기 종류를 선택해 주세요'
KCOL_ADDRESS = '방문 견적 받으실 주소를 입력해 주세요'
KCOL_INQUIRY = '문의 내용을 간단하게 남겨주세요'


# ─────────────────────────────────────────────────────────────
# 시간·전화번호 정규화
# ─────────────────────────────────────────────────────────────
def _parse_consult_dt(s: Any) -> Optional[datetime]:
    """다양한 datetime 포맷 파싱"""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    if isinstance(s, datetime):
        return s
    s = str(s).strip()
    if not s or s in ('nan', 'NaN'):
        return None
    for fmt in (
        '%Y-%m-%d %H:%M:%S',  # 2026-06-10 18:21:57 (당근 자동 시트)
        '%Y-%m-%d %H:%M',
        '%Y.%m.%d. %H:%M',    # 2026.05.29. 09:42 (메인 시트 형식)
        '%Y.%m.%d %H:%M',
        '%Y/%m/%d %H:%M',
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _format_main_dt(dt: Optional[datetime]) -> str:
    """메인 시트 형식: 2026.05.29. 09:42"""
    if dt is None:
        return ''
    return dt.strftime('%Y.%m.%d. %H:%M')


def normalize_phone(raw: Any) -> str:
    """연락처 정규화 — 공용 lead_helpers로 위임 (휴대폰·서울·지역·070 모두)"""
    from dashboard.services.lead_helpers import normalize_phone as _h
    return _h(raw)


# ─────────────────────────────────────────────────────────────
# 당근 시트 한 행 매핑
# ─────────────────────────────────────────────────────────────
def map_karrot_row_to_lead(row: pd.Series) -> Dict[str, Any]:
    """당근 자동 연동 시트 한 행 → 우리 메인 시트 형식 dict (+ 메타)"""
    def _s(key):
        v = row.get(key, '')
        s = str(v).strip() if v is not None else ''
        return '' if s in ('nan', 'NaN', 'None') else s

    consult_dt = _parse_consult_dt(row.get(KCOL_CONSULT))
    consult_time = _format_main_dt(consult_dt)

    name = _s(KCOL_NAME)
    phone = normalize_phone(row.get(KCOL_PHONE, ''))
    place = _s(KCOL_PLACE)
    device = _s(KCOL_DEVICE)
    address_raw = _s(KCOL_ADDRESS)
    inquiry = _s(KCOL_INQUIRY)

    # 당근 폼은 자유 입력이라 카카오 검증 + 정규화 필요
    address, address_level = '', ''
    if address_raw:
        from dashboard.services import address_resolver as _ar
        from dashboard.services import lead_helpers as _lh
        _regex = _lh.extract_korean_address(address_raw)
        _regex_addr = _regex[0] if _regex else None
        _regex_lv = _regex[1] if _regex else ''
        address, address_level = _ar.resolve_address(address_raw, _regex_addr, _regex_lv)
        if not address:
            address = address_raw  # 최후 fallback

    # 상담 내용: 순수 문의 본문만 (장소/기기는 별도 컬럼/키워드로)
    content = inquiry

    # 키워드: KEYWORD_VOCAB 매칭 (device + place + inquiry)
    from dashboard.services.lead_helpers import extract_keywords_from_sources
    # K열(키워드) = 폼에서 선택한 device 값만 (옵션 B). vocab 매칭으로 정규화.
    # 추가 vocab(중고/매입/세척/소상공인)은 전화 통화 후 수동 입력용.
    keyword = extract_keywords_from_sources(device)

    return {
        '리드 No': '',  # 시트 등록 시 자동 발번
        '상담 시간': consult_time,
        '플랫폼': KARROT_PLATFORM,
        '상태': DEFAULT_STATUS,
        '방문 예정일': '-',
        '고객 연락처': phone,
        '이메일': '-',
        '고객명': name,
        '방문 주소': address or '-',
        '상담 내용': content,
        '키워드': keyword,
        '온라인 상담자': '',
        '영업 담당자': '',
        '마지막 연락일': '',
        '피드백': '',
        # 메타 (시트 등록 시 LEAD_COLUMN_ORDER 필터로 자동 제외)
        '_meta_place': place,
        '_meta_device': device,
        '_meta_inquiry': inquiry,
        '_meta_consult_dt': consult_dt,
        '_meta_address_level': address_level,  # verified면 정확, level5~7/raw면 _(추정)_ 표시
    }


# ─────────────────────────────────────────────────────────────
# 메인 시트 dedup용 인덱스
# ─────────────────────────────────────────────────────────────
def _get_existing_phones(main_df: Optional[pd.DataFrame]) -> set:
    """레거시 — set 형태로 연락처만 반환 (단순 dedup 용)"""
    phones = set()
    if main_df is None or main_df.empty:
        return phones
    col = '고객 연락처' if '고객 연락처' in main_df.columns else None
    if col is None:
        return phones
    for p in main_df[col].dropna().astype(str):
        digits = re.sub(r'\D', '', p)
        if len(digits) >= 10:
            phones.add(digits)
    return phones


def _normalize_address_key(addr: str) -> str:
    """주소에서 도로명 + 번지까지 추출 (건물명/층/호 제외) — 동일 건물 매칭 키.

    예: '남양주 진건읍 진건오남로 77 심미에셈빌 1층 오헤어'
        → '남양주 진건읍 진건오남로 77'
    """
    if not addr:
        return ''
    addr = addr.strip()
    if addr in ('-', ''):
        return ''
    # 도로명/길 + 번지 패턴
    m = re.search(r'[가-힣]+(?:로|길)\s*\d+(?:-\d+)?', addr)
    if not m:
        # 도로명 없으면 지번 형태 시도: 동 + 번지
        m = re.search(r'[가-힣]+동\s*\d+(?:-\d+)?', addr)
    if not m:
        return ''
    # 매칭된 끝까지 (앞 행정구역 포함, 뒤 건물명 제외)
    return re.sub(r'\s+', ' ', addr[:m.end()]).strip()


def _get_existing_address_lookup(main_df: Optional[pd.DataFrame]) -> dict:
    """주소 정규화 키 → entries (재문의 감지: 같은 건물 다른 사람)."""
    lookup: dict = {}
    if main_df is None or main_df.empty:
        return lookup
    for _, row in main_df.iterrows():
        addr_raw = str(row.get('방문 주소', '') or '')
        key = _normalize_address_key(addr_raw)
        if not key:
            continue
        consult_dt = _parse_consult_dt(row.get('상담 시간'))
        entry = {
            'lead_no': str(row.get('리드 No', '') or ''),
            'consult_dt': consult_dt,
            'consult_time': str(row.get('상담 시간', '') or ''),
            'status': str(row.get('상태', '') or ''),
            'feedback': str(row.get('피드백', '') or ''),
            'inquiry': str(row.get('상담 내용', '') or ''),
            'platform': str(row.get('플랫폼', '') or ''),
            'consultant': str(row.get('온라인 상담자', '') or ''),
            'sales_rep': str(row.get('영업 담당자', '') or ''),
            'customer_name': str(row.get('고객명', '') or ''),
            'phone': str(row.get('고객 연락처', '') or ''),
        }
        lookup.setdefault(key, []).append(entry)
    for key in lookup:
        lookup[key].sort(
            key=lambda e: e['consult_dt'] or datetime.min, reverse=True,
        )
    return lookup


def _get_existing_phone_lookup(main_df: Optional[pd.DataFrame]) -> dict:
    """연락처 → [{lead_no, consult_time(datetime), status, feedback, platform}, ...] 시간 내림차순.

    재문의 감지 + 옛 이력 카드 표시에 사용.
    """
    lookup: dict = {}
    if main_df is None or main_df.empty:
        return lookup
    for _, row in main_df.iterrows():
        phone_raw = str(row.get('고객 연락처', '') or '')
        digits = re.sub(r'\D', '', phone_raw)
        if len(digits) < 10:
            continue
        consult_dt = _parse_consult_dt(row.get('상담 시간'))
        entry = {
            'lead_no': str(row.get('리드 No', '') or ''),
            'consult_dt': consult_dt,
            'consult_time': str(row.get('상담 시간', '') or ''),
            'status': str(row.get('상태', '') or ''),
            'feedback': str(row.get('피드백', '') or ''),
            'inquiry': str(row.get('상담 내용', '') or ''),
            'platform': str(row.get('플랫폼', '') or ''),
            'consultant': str(row.get('온라인 상담자', '') or ''),
            'sales_rep': str(row.get('영업 담당자', '') or ''),
        }
        lookup.setdefault(digits, []).append(entry)
    # 각 연락처별 시간 내림차순
    for digits in lookup:
        lookup[digits].sort(
            key=lambda e: e['consult_dt'] or datetime.min,
            reverse=True,
        )
    return lookup


def _find_previous_leads(phone_lookup: dict, phone_digits: str,
                         new_dt: Optional[datetime] = None) -> list:
    """같은 연락처의 옛 lead 이력 (현 신규 시각보다 이른 것만)"""
    if not phone_digits or phone_digits not in phone_lookup:
        return []
    entries = phone_lookup[phone_digits]
    if new_dt:
        entries = [e for e in entries if e['consult_dt'] and e['consult_dt'] < new_dt]
    return entries


def _format_time_ago(now_dt: Optional[datetime], prev_dt: Optional[datetime]) -> str:
    """시간 차이를 사람이 읽는 형태로 ('3개월 전', '2년 전' 등)"""
    if not now_dt or not prev_dt:
        return ''
    diff_days = (now_dt - prev_dt).days
    if diff_days < 0:
        return ''
    if diff_days < 1:
        return '오늘'
    if diff_days < 7:
        return f'{diff_days}일 전'
    if diff_days < 30:
        return f'{diff_days // 7}주 전'
    if diff_days < 365:
        return f'{diff_days // 30}개월 전'
    return f'{diff_days // 365}년 전'


def _wrap_quoted(text: str, width: int = 60) -> str:
    """텍스트를 width자 단위로 wrap하고 각 줄에 '>' prefix 부착.

    슬랙 mrkdwn에서 한 줄이 너무 길면 화면 wrap 시 '>' 마커가 끊겨 인용 그룹을 이탈함.
    미리 명시적 줄바꿈을 넣어 매 줄에 '>' 보장.
    """
    if not text:
        return ''
    out_lines = []
    for raw in text.split('\n'):
        if not raw:
            out_lines.append('>')
            continue
        # 한국어 텍스트는 단어 경계가 모호하므로 break_long_words=False보다는
        # break_on_hyphens=False로 두고 textwrap 기본 처리
        wrapped = textwrap.fill(raw, width=width, break_long_words=True,
                                break_on_hyphens=False) or raw
        for ln in wrapped.split('\n'):
            out_lines.append('>' + ln)
    return '\n'.join(out_lines)


def _build_repeat_section(lead: dict) -> str:
    """재문의 컨텍스트 섹션 (이전 No + 문의 내용 + 상태 + 상담자).

    이전 응대자(상담자)와 상태가 보이면 인수인계/우선순위 판단에 도움.
    빈값은 '-'로 통일.
    """
    prev_leads = lead.get('_meta_previous_leads') or []
    if not prev_leads:
        return ''
    most_recent = prev_leads[0]
    prev_no = most_recent.get('lead_no', '')
    prev_time = (most_recent.get('consult_time') or '').strip()
    prev_inquiry = (most_recent.get('inquiry') or '').strip() or '-'
    prev_status = (most_recent.get('status') or '').strip() or '-'
    # 상담자/방문자: 한국 이름 → 이니셜 통일
    from dashboard.blueprints.slack_bot import _to_initial
    prev_consultant_raw = (most_recent.get('consultant') or '').strip()
    prev_consultant = _to_initial(prev_consultant_raw) or '-'
    # 방문자: 이전 상태가 "방문 예약"일 때만 영업 담당자 값, 그 외 '-'
    if prev_status == '방문 예약':
        prev_sales_rep_raw = (most_recent.get('sales_rep') or '').strip()
        prev_sales_rep = _to_initial(prev_sales_rep_raw) or '-'
    else:
        prev_sales_rep = '-'
    # 길이 제한 (200자)
    if len(prev_inquiry) > 200:
        prev_inquiry = prev_inquiry[:200] + '...'

    # 이전 문의 식별: "시간 / 리드No" 형태 (시간 없으면 리드No만)
    prev_label = f"{prev_time} / {prev_no}" if prev_time else prev_no

    # 문의 내용은 60자 wrap + 각 줄 '>' prefix (긴 라인 화면 wrap 시 인용 이탈 방지)
    prev_inquiry_quoted = _wrap_quoted(prev_inquiry, width=60)
    return (
        f">:repeat: *재문의 감지*\n"
        f">*이전 문의* : {prev_label}\n"
        f">*문의 내용* :\n{prev_inquiry_quoted}\n"
        f">*상태* : {prev_status}\n"
        f">*상담자* : {prev_consultant}\n"
        f">*방문자* : {prev_sales_rep}\n"
        f">---------------------------------------------\n"
    )


# ─────────────────────────────────────────────────────────────
# 메인 sync 함수
# ─────────────────────────────────────────────────────────────
def sync_karrot() -> Dict[str, Any]:
    """
    당근 자동 연동 시트 폴링 → 신규 행을 메인 시트에 추가 + 슬랙 알림.
    APScheduler가 주기적으로 호출.
    """
    karrot_sheet_id = os.getenv('KARROT_AUTO_SHEET_ID', '').strip()
    karrot_tab = os.getenv('KARROT_AUTO_SHEET_TAB', '시트1').strip()

    if not karrot_sheet_id:
        logger.warning('[SYNC/karrot] KARROT_AUTO_SHEET_ID 미설정 - 스킵')
        return {'error': 'KARROT_AUTO_SHEET_ID 미설정'}

    mgr = get_sheets_manager()

    try:
        karrot_df = mgr.get_sheet_data(karrot_sheet_id, f'{karrot_tab}!A:Z')
    except Exception as exc:
        logger.error(f'[SYNC/karrot] 당근 시트 읽기 실패: {exc}', exc_info=True)
        return {'error': str(exc)}

    if karrot_df.empty:
        logger.debug('[SYNC/karrot] 당근 시트 빈 데이터')
        return {'total': 0, 'new_count': 0, 'duplicates': 0}

    # 메인 시트 (운영 시트) 로드 + dedup 인덱스
    main_df = load_leads_data(force_refresh=True)
    phone_lookup = _get_existing_phone_lookup(main_df)
    # 같은 sync 내 중복 폴링 방지용 (연락처+시간 윈도우)
    seen_in_sync: dict = {}  # phone_digits → consult_dt
    DEDUP_WINDOW = timedelta(hours=1)  # 1시간 이내 같은 번호 = 재폴링으로 간주

    new_leads: List[Dict[str, Any]] = []
    duplicates = 0
    for _, row in karrot_df.iterrows():
        lead = map_karrot_row_to_lead(row)
        phone_digits = re.sub(r'\D', '', lead['고객 연락처'])
        new_dt = lead.get('_meta_consult_dt')

        # 1) 같은 sync 내 중복 폴링 (1시간 이내) → skip
        if phone_digits and phone_digits in seen_in_sync:
            if new_dt and abs((new_dt - seen_in_sync[phone_digits]).total_seconds()) < 3600:
                duplicates += 1
                continue

        # 2) 메인 시트의 가장 최근 lead와 1시간 이내 → skip (재폴링 가능성)
        if phone_digits and phone_digits in phone_lookup and new_dt:
            latest = phone_lookup[phone_digits][0]
            if latest['consult_dt']:
                time_diff = new_dt - latest['consult_dt']
                if abs(time_diff.total_seconds()) < 3600:
                    duplicates += 1
                    continue
            # 1시간 이상 차이 → 재문의로 간주, 옛 이력 메타에 저장
            lead['_meta_previous_leads'] = phone_lookup[phone_digits]

        new_leads.append(lead)
        if phone_digits and new_dt:
            seen_in_sync[phone_digits] = new_dt

    # 응답 시각 오름차순 정렬 (가장 최신이 채널의 마지막 메시지로)
    new_leads.sort(key=lambda l: l.get('_meta_consult_dt') or datetime.min)

    # 폭주 안전망 — 한 sync에 신규 lead가 비정상적으로 많으면(Redis 장애 등)
    # 슬랙 발송 skip + 알림. 정상은 보통 0~2건, 5건 이상이면 dedup 문제 의심.
    MAX_NEW_PER_SYNC = int(os.getenv('KARROT_MAX_NEW_PER_SYNC', '5'))
    if len(new_leads) > MAX_NEW_PER_SYNC:
        logger.error(
            f'[SYNC/karrot] 신규 lead 비정상 다수 ({len(new_leads)}건 > {MAX_NEW_PER_SYNC}) '
            f'— Redis dedup 장애 의심, 발송 skip'
        )
        return {
            'total': len(karrot_df),
            'new_count': 0,
            'duplicates': duplicates,
            'lead_nos': [],
            'skipped_runaway': len(new_leads),
        }

    lead_nos = []
    if new_leads:
        lead_nos = _append_leads_to_main(new_leads)
        # 슬랙 발송 — 함수 자체가 SSL/네트워크 에러로 raise 가능 → 외부 try/except 안전망
        # raise되면 sent=빈 set으로 처리 → 모든 lead가 pending 큐에 들어감 (자동 재시도)
        try:
            sent = _send_slack_notifications(new_leads, lead_nos, source='당근')
        except Exception as exc:
            logger.error(
                f'[SYNC/karrot] 슬랙 발송 함수 자체 실패 → 전체 pending 큐 등록: {exc}'
            )
            sent = set()
        # 미발송 lead — Redis pending 큐에 저장 (SSL 에러 등으로 누락된 lead 추적용)
        try:
            from dashboard.utils.redis_client import get_redis_client
            rc = get_redis_client().redis
            for ln in lead_nos:
                if ln not in sent:
                    rc.set(f'pending_slack_notify:{ln}', '당근', ex=60 * 60 * 24 * 7)
        except Exception:
            pass

    result = {
        'total': len(karrot_df),
        'new_count': len(new_leads),
        'duplicates': duplicates,
        'lead_nos': lead_nos,
    }
    logger.info(
        f'[SYNC/karrot] total={result["total"]} '
        f'new={result["new_count"]} dup={result["duplicates"]}'
    )
    return result


_APPEND_LEAD_LOCK = threading.Lock()


def _append_leads_to_main(leads: List[Dict[str, Any]]) -> List[str]:
    """
    메인 시트에 일괄 추가 + 리드No 자동 발번.

    Google Sheets API의 spreadsheets.values.append() 직접 호출.
    (mgr.append_row()는 시트명이 '공사 현황의 사본'으로 하드코딩돼 있어서 사용 불가)

    threading.Lock으로 직렬화 — 동시 호출 시 lead_no 발번 race condition 방지.
    """
    if not leads:
        return []

    cfg = _get_sheet_config()
    if cfg is None:
        raise RuntimeError('ONLINE_LEADS_SHEET_ID 환경변수 미설정')

    # 락 — 동시 lead_no 발번 충돌 방지 (블로킹 acquire)
    with _APPEND_LEAD_LOCK:
        return _append_leads_to_main_locked(leads, cfg)


def _append_leads_to_main_locked(leads: List[Dict[str, Any]], cfg) -> List[str]:
    """lock 잡힌 후 실제 append 처리 (내부용)."""
    mgr = get_sheets_manager()
    df = load_leads_data(force_refresh=True)

    # 다음 리드 No 시퀀스
    max_num = 0
    if df is not None and not df.empty and '리드 No' in df.columns:
        for ln in df['리드 No'].dropna().astype(str):
            digits = re.sub(r'\D', '', ln.strip())
            try:
                max_num = max(max_num, int(digits))
            except ValueError:
                continue

    # 리드 No 발번 + row 데이터 구성 (15열, LEAD_COLUMN_ORDER 순서)
    lead_nos = []
    rows = []
    for i, lead in enumerate(leads, start=1):
        ln = f'L-{max_num + i:05d}'
        lead['리드 No'] = ln
        lead_nos.append(ln)
        rows.append([lead.get(col, '') for col in LEAD_COLUMN_ORDER])

    # values.append() 사용 — 자동으로 grid 확장 + 다음 빈 행에 추가
    # range를 'A1:O1'로 한정해 시트의 다른 컬럼 영향 받지 않게 (헤더 영역만 참조)
    sheet_name = cfg['sheet_name']
    range_name = f"'{sheet_name}'!A1:O1"

    result = mgr.service.spreadsheets().values().append(
        spreadsheetId=cfg['sheet_id'],
        range=range_name,
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body={'values': rows},
    ).execute()

    updates = result.get('updates', {})
    updated_range = updates.get('updatedRange', '?')
    invalidate_leads_cache()
    logger.info(
        f'[SYNC] 메인 시트 등록 완료: {len(leads)}건 '
        f'({lead_nos[0]} ~ {lead_nos[-1]}, '
        f'range={updated_range}, '
        f'updatedCells={updates.get("updatedCells", "?")})'
    )

    # 등록된 행의 배경색 설정
    # - 방문 예약: 연한 노란색 (#fff2cc)
    # - 그 외: 흰색 (INSERT_ROWS의 위 행 색 자동 상속 방지)
    try:
        statuses = [str(lead.get('상태', '') or '') for lead in leads]
        _reset_row_background(mgr, cfg['sheet_id'], sheet_name, updated_range,
                              num_cols=len(LEAD_COLUMN_ORDER), statuses=statuses)
    except Exception as exc:
        logger.warning(f'[SYNC] 신규 행 배경색 설정 실패 (등록은 정상): {exc}')

    return lead_nos


_SHEET_GID_CACHE: Dict[str, int] = {}


def _get_sheet_gid(mgr, spreadsheet_id: str, sheet_name: str) -> Optional[int]:
    """시트명 → 시트 gid (서식 batchUpdate 용)"""
    cache_key = f'{spreadsheet_id}::{sheet_name}'
    if cache_key in _SHEET_GID_CACHE:
        return _SHEET_GID_CACHE[cache_key]
    info = mgr.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for s in info.get('sheets', []):
        props = s.get('properties', {})
        if props.get('title') == sheet_name:
            gid = props.get('sheetId')
            _SHEET_GID_CACHE[cache_key] = gid
            return gid
    return None


def _reset_row_background(mgr, spreadsheet_id: str, sheet_name: str,
                          updated_range: str, num_cols: int,
                          statuses: Optional[List[str]] = None):
    """방금 등록된 행들의 배경색 설정.

    - 상태 '방문 예약' → 연한 노란색 (#fff2cc)
    - 그 외 → 흰색 (위 행 색 상속 방지)
    """
    m = re.search(r'![A-Z]+(\d+):[A-Z]+(\d+)', updated_range or '')
    if not m:
        return
    start_row = int(m.group(1))
    end_row = int(m.group(2))
    num_rows = end_row - start_row + 1
    if num_rows <= 0:
        return

    gid = _get_sheet_gid(mgr, spreadsheet_id, sheet_name)
    if gid is None:
        logger.warning(f'[SYNC] 시트 gid 못 찾음: {sheet_name}')
        return

    white = {'red': 1.0, 'green': 1.0, 'blue': 1.0}
    yellow = {'red': 1.0, 'green': 242/255, 'blue': 204/255}  # #fff2cc 연한 노란색
    statuses = statuses or []

    rows_payload = []
    for i in range(num_rows):
        status = statuses[i] if i < len(statuses) else ''
        color = yellow if status == '방문 예약' else white
        rows_payload.append({
            'values': [{'userEnteredFormat': {'backgroundColor': color}}] * num_cols
        })

    request = {
        'updateCells': {
            'range': {
                'sheetId': gid,
                'startRowIndex': start_row - 1,
                'endRowIndex': end_row,
                'startColumnIndex': 0,
                'endColumnIndex': num_cols,
            },
            'fields': 'userEnteredFormat.backgroundColor',
            'rows': rows_payload,
        }
    }
    mgr.service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': [request]},
    ).execute()
    yellow_count = sum(1 for s in statuses[:num_rows] if s == '방문 예약')
    logger.info(f'[SYNC] 신규 행 배경색 설정 완료 ({num_rows}행, 방문 예약 {yellow_count}건 노란색)')


# ─────────────────────────────────────────────────────────────
# Slack 알림
# ─────────────────────────────────────────────────────────────
_CHANNEL_ID_CACHE: Dict[str, str] = {}


def _resolve_channel_id(client, channel_name_or_id: str) -> str:
    """채널명 → 채널 ID 변환 (캐시). 이미 ID(C로 시작 11자)면 그대로 반환."""
    if not channel_name_or_id:
        return ''

    cn = channel_name_or_id.lstrip('#').strip()
    # 이미 채널 ID 형식이면 그대로
    if len(cn) >= 9 and cn[0] in ('C', 'G', 'D') and cn[1:].replace('_', '').isalnum():
        return cn

    # 캐시 hit
    if cn in _CHANNEL_ID_CACHE:
        return _CHANNEL_ID_CACHE[cn]

    # 슬랙 API로 채널 목록 조회
    try:
        cursor = None
        while True:
            resp = client.conversations_list(
                types='public_channel,private_channel',
                limit=200,
                cursor=cursor,
            )
            for ch in resp.get('channels', []):
                if ch.get('name') == cn:
                    _CHANNEL_ID_CACHE[cn] = ch['id']
                    logger.info(f'[SLACK] 채널 lookup: #{cn} → {ch["id"]}')
                    return ch['id']
            cursor = resp.get('response_metadata', {}).get('next_cursor')
            if not cursor:
                break
    except Exception as exc:
        logger.error(f'[SLACK] 채널 목록 조회 실패: {exc}')

    logger.warning(f'[SLACK] 채널 #{cn} 못 찾음 — 봇이 채널에 초대됐는지 확인')
    return cn  # 그대로 시도 (실패 시 명확한 에러)


def _send_slack_notifications(leads: List[Dict[str, Any]], lead_nos: List[str],
                              source: str = '당근') -> set:
    """각 신규 리드를 사용자 양식으로 채널에 개별 메시지 전송.

    Returns: 발송 성공한 lead_no의 set (실패는 미포함).
             Slack 미구성 / 클라이언트 초기화 실패 시 빈 set.
    """
    sent: set = set()
    bot_token = os.getenv('SLACK_BOT_TOKEN', '').strip()
    channel_setting = os.getenv('SLACK_LEAD_CHANNEL', '').strip()

    if not bot_token or 'your' in bot_token.lower():
        logger.warning(f'[SYNC/{source}] SLACK_BOT_TOKEN 미설정 - 슬랙 알림 스킵')
        return sent
    if not channel_setting or '여기에' in channel_setting:
        logger.warning(f'[SYNC/{source}] SLACK_LEAD_CHANNEL 미설정 - 슬랙 알림 스킵')
        return sent

    try:
        from slack_sdk import WebClient
        client = WebClient(token=bot_token)
    except Exception as exc:
        logger.error(f'[SYNC/{source}] Slack 클라이언트 초기화 실패: {exc}')
        return sent

    # 채널명 → ID 자동 변환 (한글 채널명 문제 회피)
    channel = _resolve_channel_id(client, channel_setting)

    # 봇이 채널에 안 들어가 있으면 자동 가입 시도 (public 채널만 가능)
    try:
        client.conversations_join(channel=channel)
    except Exception:
        pass  # 이미 가입돼있거나 private 채널 (그래도 시도)

    for lead, ln in zip(leads, lead_nos):
        try:
            blocks, fallback = build_inquiry_blocks(lead, ln, source)
            resp = client.chat_postMessage(
                channel=channel,
                text=fallback,
                blocks=blocks,
                unfurl_links=False,
            )
            if resp and resp.get('ok'):
                sent.add(ln)
                # 카드 ts 저장 — [기존 lead 연결] 시 그 카드에 ✅ reaction 추가용
                try:
                    from dashboard.utils.redis_client import get_redis_client
                    rc = get_redis_client().redis
                    msg_ts = resp.get('ts', '')
                    if msg_ts:
                        rc.set(
                            f'lead_card_msg:{ln}',
                            f'{channel}|{msg_ts}',
                            ex=60 * 60 * 24 * 180,  # 180일 보존
                        )
                except Exception:
                    pass
            else:
                logger.error(f'[SYNC/{source}] 슬랙 전송 응답 not ok ({ln}): {resp}')
        except Exception as exc:
            logger.error(f'[SYNC/{source}] 슬랙 전송 실패 ({ln}): {exc}')
    return sent


# ─────────────────────────────────────────────────────────────
# 인입 알림 블록 (홈페이지/당근/기타 플랫폼 공용) - 동적 타이틀
# ─────────────────────────────────────────────────────────────
def build_inquiry_blocks(lead: dict, lead_no: str, source: str = '당근') -> tuple:
    """
    Returns (blocks, fallback_text). 양식:

        *접수번호:* `L-02345`
        :bell: *새 문의 접수 알림 - 당근*
        ---------------------------------------------

    플랫폼별 타이틀: `새 문의 접수 알림 - {source}` 형식으로 통일.
    새 플랫폼 추가 시 source 인자만 바꾸면 자동 적용.
        >*문의시간* : 2026.06.10. 18:21
        >*이름 / 상호* : 신호현
        >*연락처* : 010-2977-1698
        >*이메일* : -
        >*설치 희망 장소* : 사무실 / 관공서
        >*설치 희망 기기* : 천장형
        >*상세 문의 내용* :
        실평9평. 15-20평형 천장형으로 견적원합니다...
        ---------------------------------------------
        [방문 요청] [가격 문의]
    """
    consult_time = (lead.get('상담 시간') or '').strip() or '-'
    name = (lead.get('고객명') or '').strip() or '-'
    phone = (lead.get('고객 연락처') or '').strip() or '-'
    email = (lead.get('이메일') or '').strip() or '-'
    place = (lead.get('_meta_place') or '').strip() or '-'
    device = (lead.get('_meta_device') or '').strip() or '-'
    inquiry = (lead.get('_meta_inquiry') or lead.get('상담 내용') or '').strip() or '-'
    # 슬랙 section text 3000자 한도 — 메타데이터 여유분 고려 안전선 2400자
    if len(inquiry) > 2400:
        inquiry = inquiry[:2400] + '\n…(내용이 길어 일부만 표시 — 시트 참조)'

    # 방문 주소 (자동 추출 또는 시트에 등록된 값) — 신뢰도 레벨에 따라 표시 분기
    address = (lead.get('방문 주소') or '').strip()
    addr_level = (lead.get('_meta_address_level') or '').strip()
    if address:
        # 신뢰도별 표시:
        # - verified (카카오 매칭): 마커 없음 (정확)
        # - level1~4 (정규식 풀 패턴): 마커 없음 (정확)
        # - level5~7 / regex / raw: _(추정)_ 마커 (영업 검증 필요)
        if addr_level in ('verified', 'level1', 'level2', 'level3', 'level3b', 'level4', ''):
            address_display = address
        else:
            address_display = f'{address}  _(추정)_'
    else:
        address_display = '-'

    title = f"새 문의 접수 알림 - {source}"

    # 재문의 감지 — 같은 번호로 이전 lead가 있으면 섹션 추가 (타이틀은 그대로 유지)
    repeat_section = _build_repeat_section(lead)

    # 헤더 + 구분선 + 본문 + 구분선 모두 한 `>` blockquote 그룹으로 통일.
    # 같은 blockquote 안 라인은 슬랙이 복사 시 줄바꿈을 정상 보존함.
    # 멀티라인 inquiry는 60자 wrap + 각 줄마다 `>` prefix.
    inquiry_quoted = _wrap_quoted(inquiry.rstrip(), width=60)
    main_text = (
        "⠀\n"
        f">:bell: *{title}*  `{lead_no}`\n"
        f">---------------------------------------------\n"
        + repeat_section
        + f">*문의시간* : {consult_time}\n"
        f">*이름 / 상호* : {name}\n"
        f">*연락처* : {phone}\n"
        f">*이메일* : {email}\n"
        f">*설치 희망 장소* : {place}\n"
        f">*설치 희망 기기* : {device}\n"
        f">*방문 주소* : {address_display}\n"
        f">*상세 문의 내용* : \n{inquiry_quoted}\n"
        f">---------------------------------------------"
    )

    blocks = [
        {"type": "section", "text": {"type": "mrkdwn", "text": main_text}},
        {
            "type": "actions",
            "elements": [
                {
                    "type": "button",
                    "text": {"type": "plain_text", "text": ":point_right: 상담하기", "emoji": True},
                    "style": "primary",
                    "value": lead_no,
                    "action_id": "button_consult",
                },
            ],
        },
        {"type": "context", "elements": [{"type": "mrkdwn", "text": "⠀"}]},
    ]
    fallback_text = f"[{source}] {lead_no} {name} / {phone}"
    return blocks, fallback_text


# ─────────────────────────────────────────────────────────────
# 슬랙 워크플로우가 직접 추가한 전화 lead 보정 (B안 폴링)
# ─────────────────────────────────────────────────────────────
_MONTH_ABBR = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}


def _parse_english_longform_dt(raw: str) -> Optional[datetime]:
    """영문 long form 시간 인식 (locale 무관).

    예: 'Jun 23, 2026, 9:10:33 AM' / 'Jun 23, 2026, 9:10 AM' / 'Jun 23, 2026'
    """
    m = re.match(
        r'(\w{3,9})\s+(\d{1,2}),\s+(\d{4})'
        r'(?:,?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?',
        raw, re.IGNORECASE,
    )
    if not m:
        return None
    month = _MONTH_ABBR.get(m.group(1)[:3].lower())
    if not month:
        return None
    day = int(m.group(2))
    year = int(m.group(3))
    hour = int(m.group(4)) if m.group(4) else 0
    minute = int(m.group(5)) if m.group(5) else 0
    second = int(m.group(6)) if m.group(6) else 0
    ampm = (m.group(7) or '').upper()
    if ampm == 'PM' and hour < 12:
        hour += 12
    elif ampm == 'AM' and hour == 12:
        hour = 0
    try:
        return datetime(year, month, day, hour, minute, second)
    except ValueError:
        return None


def _normalize_workflow_datetime(raw: str) -> str:
    """슬랙 워크플로의 시간 변수를 'YYYY.MM.DD. HH:MM' 한국식 점 표기로 정규화.

    인식 포맷:
      - 이미 점 표기 ('2026.06.23. 17:19') — 그대로
      - 타임스탬프 (epoch seconds: 10자리 / milliseconds: 13자리)
      - ISO 8601 / 슬래시
      - 영문 long form ('Jun 23, 2026, 9:10:33 AM') — locale 무관 패턴
      - 인식 실패 시 원본 그대로 반환 (수동 정리 가능하도록)
    """
    if not raw:
        return ''
    raw = raw.strip()
    if re.match(r'^\d{4}\.\d{2}\.\d{2}\.', raw):
        return raw

    # epoch 타임스탬프 (10자리 = seconds, 13자리 = milliseconds)
    if raw.isdigit():
        try:
            n = int(raw)
            if len(raw) == 13:
                n = n // 1000
            if 946684800 <= n <= 32503680000:  # 2000-01-01 ~ 3000-01-01
                return datetime.fromtimestamp(n).strftime('%Y.%m.%d. %H:%M')
        except (ValueError, OverflowError):
            pass

    formats = [
        '%Y-%m-%dT%H:%M:%S.%fZ', '%Y-%m-%dT%H:%M:%SZ',
        '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M',
        '%Y/%m/%d %H:%M:%S', '%Y/%m/%d %H:%M',
    ]
    for fmt in formats:
        try:
            return datetime.strptime(raw, fmt).strftime('%Y.%m.%d. %H:%M')
        except ValueError:
            continue

    # 영문 long form
    dt = _parse_english_longform_dt(raw)
    if dt:
        return dt.strftime('%Y.%m.%d. %H:%M')

    return raw


def _normalize_workflow_date(raw: str) -> str:
    """방문 예정일 정규화 — Google Sheets 시리얼 숫자도 ISO로 복원.

    슬랙 워크플로의 datepicker는 ISO('2026-06-25')를 보내지만 시트가 USER_ENTERED
    모드에서 시리얼 정수(예: 46197)로 자동 변환할 수 있음.

    Returns:
        '2026-06-25' (ISO 그대로) — 호출 측이 시트에 escape prefix 붙여 저장.
        시리얼/ISO 둘 다 인식, 인식 실패 시 빈 문자열.
    """
    if not raw:
        return ''
    raw = raw.strip()
    if re.match(r'^\d{4}-\d{2}-\d{2}$', raw):
        return raw
    if raw.isdigit():
        try:
            serial = int(raw)
            if 1 <= serial <= 100000:  # 1900-01-01 ~ 2173-10-15 범위
                base = datetime(1899, 12, 30)
                dt = base + timedelta(days=serial)
                return dt.strftime('%Y-%m-%d')
        except (ValueError, OverflowError):
            pass
    return ''


def retry_pending_slack_notifications() -> Dict[str, Any]:
    """SSL 에러 등으로 슬랙 발송 누락된 lead 자동 재발송.

    Redis 'pending_slack_notify:{lead_no}'에 마커 있는 lead를 메인 시트에서 찾아
    _send_slack_notifications 재호출. 성공한 lead의 마커는 삭제.
    """
    result = {'attempted': 0, 'sent': 0}
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        keys = list(rc.scan_iter(match='pending_slack_notify:*', count=100))
        if not keys:
            return result
        # 메인 시트 fetch
        main_df = load_leads_data(force_refresh=True)
        if main_df is None or main_df.empty:
            return result
        pending_leads = []
        pending_nos = []
        pending_sources = []
        for k in keys:
            ks = k if isinstance(k, str) else k.decode()
            lead_no = ks.split(':')[-1]
            source_raw = rc.get(k) or '당근'
            source = source_raw if isinstance(source_raw, str) else source_raw.decode()
            row = main_df[main_df['리드 No'].astype(str).str.strip() == lead_no]
            if row.empty:
                rc.delete(k)
                continue
            lead = row.iloc[0].to_dict()
            # _meta 필드 채움 (build_inquiry_blocks 의존)
            lead['_meta_place'] = ''
            lead['_meta_device'] = str(lead.get('키워드', '') or '')
            lead['_meta_inquiry'] = str(lead.get('상담 내용', '') or '')
            pending_leads.append(lead)
            pending_nos.append(lead_no)
            pending_sources.append(source)
        if not pending_leads:
            return result
        result['attempted'] = len(pending_leads)
        # 출처별 그룹 발송 (당근/홈페이지 등)
        from collections import defaultdict
        groups = defaultdict(list)
        for lead, ln, src in zip(pending_leads, pending_nos, pending_sources):
            groups[src].append((lead, ln))
        for src, items in groups.items():
            leads = [l for l, _ in items]
            nos = [n for _, n in items]
            sent = _send_slack_notifications(leads, nos, source=src)
            for ln in nos:
                if ln in sent:
                    rc.delete(f'pending_slack_notify:{ln}')
                    result['sent'] += 1
        if result['sent']:
            logger.info(
                f"[SYNC/retry] 미발송 슬랙 알림 재발송: {result['sent']}/{result['attempted']}"
            )
    except Exception as exc:
        logger.error(f"[SYNC/retry] 재발송 처리 실패: {exc}", exc_info=True)
    return result


def sync_workflow_phone_leads() -> Dict[str, Any]:
    """슬랙 워크플로우가 메인 시트에 직접 추가한 전화 lead 자동 보정.

    스캔 대상: 리드 No 빈 + 플랫폼='전화' 행
    처리:
      1. lead_no 발번 (max + 1)
      2. 상담 시간 ISO → 'YYYY.MM.DD. HH:MM'
      3. 방문 예정일 ISO → 시트 escape ('YYYY-MM-DD)
      4. 키워드 vocab 매칭
      5. 행 색 분기 (상태별)
      6. 상태='방문 예약' 시 #방문_일정 카드 발송
    """
    result = {'processed': 0, 'visit_notified': 0, 'errors': 0}
    try:
        main_df = load_leads_data(force_refresh=True)
        if main_df is None or main_df.empty:
            return result

        # 빈 lead_no + 플랫폼='전화' 행
        is_empty_no = main_df['리드 No'].astype(str).str.strip() == ''
        is_phone = main_df['플랫폼'].astype(str).str.strip() == '전화'
        candidates = main_df[is_empty_no & is_phone].copy()
        if candidates.empty:
            return result

        cfg = _get_sheet_config()
        if not cfg:
            logger.warning('[SYNC/전화WF] ONLINE_LEADS_SHEET_ID 미설정')
            return result

        manager = get_sheets_manager()

        # 다음 lead_no — 시트 max + 1
        existing_nos = main_df['리드 No'].astype(str).str.extract(r'L-(\d+)')[0]
        existing_nos = pd.to_numeric(existing_nos, errors='coerce').dropna()
        next_no_int = int(existing_nos.max()) + 1 if len(existing_nos) > 0 else 1

        from dashboard.services.lead_helpers import extract_keywords_from_sources
        notify_visit = []

        for idx, row in candidates.iterrows():
            sheet_row = idx + 2  # 헤더 1행 + 0-based
            # 매니저 입력 진행 중 — 필수 필드(고객명 or 연락처) 없으면 다음 sync 대기
            # (슬랙 워크플로우가 시트에 여러 셀 순차 write하는 동안 sync가 잡으면 빈 카드 발송됨)
            _name = str(row.get('고객명', '') or '').strip()
            _phone = str(row.get('고객 연락처', '') or '').strip()
            if not _name and not _phone:
                logger.debug(
                    f'[SYNC/전화WF] 매니저 입력 진행 중 skip (row {sheet_row})'
                )
                continue
            new_lead_no = f"L-{next_no_int:05d}"
            next_no_int += 1

            update_cells = [(f"A{sheet_row}", new_lead_no)]

            # 상담 시간 정규화
            consult_raw = str(row.get('상담 시간', '')).strip()
            consult_norm = _normalize_workflow_datetime(consult_raw)
            if consult_norm and consult_norm != consult_raw:
                update_cells.append((f"B{sheet_row}", consult_norm))

            # 방문 예정일 정규화 — 단일 ISO은 escape 형태, 범위는 양식 정돈
            # 슬랙 워크플로우 빌더에서 datepicker 2개 결합 시 raw ISO 범위 ("2026-06-23~2026-06-24")로 들어옴
            visit_date_raw = str(row.get('방문 예정일', '')).strip()
            visit_date = visit_date_raw
            if visit_date_raw and '~' in visit_date_raw:
                # 범위 — 양식 정돈 (같은 달이면 ~DD, 다른 달이면 ~MM-DD)
                parts = visit_date_raw.split('~', 1)
                start_iso = parts[0].lstrip("'").strip()
                end_iso = parts[1].strip() if len(parts) > 1 else ''
                from dashboard.blueprints.slack_helpers import _format_visit_date_range
                formatted = _format_visit_date_range(start_iso, end_iso)
                if formatted and formatted != visit_date_raw:
                    # 범위 텍스트는 escape prefix 없이 (Google Sheets가 텍스트로 인식)
                    update_cells.append((f"E{sheet_row}", formatted))
                    visit_date = formatted
            elif visit_date_raw and not visit_date_raw.startswith("'"):
                iso = _normalize_workflow_date(visit_date_raw)
                if iso:
                    update_cells.append((f"E{sheet_row}", f"'{iso}"))
                    visit_date = iso

            # 키워드 vocab 매칭
            keyword_raw = str(row.get('키워드', '')).strip()
            keyword_norm = extract_keywords_from_sources(keyword_raw) or '-'
            if keyword_norm != keyword_raw:
                update_cells.append((f"K{sheet_row}", keyword_norm))

            # batchUpdate — SSL/transient 에러 시 즉시 1회 재시도
            import time as _t
            batch_body = {
                'valueInputOption': 'USER_ENTERED',
                'data': [
                    {'range': f"'{cfg['sheet_name']}'!{r}", 'values': [[v]]}
                    for r, v in update_cells
                ],
            }
            try:
                try:
                    manager.service.spreadsheets().values().batchUpdate(
                        spreadsheetId=cfg['sheet_id'], body=batch_body,
                    ).execute()
                except Exception as exc1:
                    err_l = str(exc1).lower()
                    if any(k in err_l for k in ('ssl', 'wrong_version',
                                                  'timeout', 'connection')):
                        logger.warning(
                            f"[SYNC/전화WF] 일시 에러 재시도 ({new_lead_no}): {exc1}"
                        )
                        _t.sleep(2)
                        manager.service.spreadsheets().values().batchUpdate(
                            spreadsheetId=cfg['sheet_id'], body=batch_body,
                        ).execute()
                    else:
                        raise

                # 행 색
                status = str(row.get('상태', '')).strip()
                try:
                    updated_range = (
                        f"'{cfg['sheet_name']}'!A{sheet_row}:O{sheet_row}"
                    )
                    _reset_row_background(
                        manager, cfg['sheet_id'], cfg['sheet_name'],
                        updated_range, num_cols=len(LEAD_COLUMN_ORDER),
                        statuses=[status],
                    )
                except Exception as exc_bg:
                    logger.warning(f"[SYNC/전화WF] 행 색 실패 ({new_lead_no}): {exc_bg}")

                result['processed'] += 1
                logger.info(f"[SYNC/전화WF] 보정 완료: {new_lead_no} ({status})")

                # 방문 예약 시 #방문_일정 카드 + 슬랙 List 후보
                if status == '방문 예약':
                    # 등록자: 영업 담당자 > 온라인 상담자 (둘 다 한국 이름 가정)
                    user_name = (str(row.get('영업 담당자', '')).strip()
                                 or str(row.get('온라인 상담자', '')).strip())
                    notify_visit.append({
                        'lead_no': new_lead_no,
                        'name': str(row.get('고객명', '')).strip(),
                        'phone': str(row.get('고객 연락처', '')).strip(),
                        'email': str(row.get('이메일', '')).strip(),
                        'consult_time': consult_norm or consult_raw,
                        'address': str(row.get('방문 주소', '')).strip(),
                        'visit_date': visit_date,
                        'inquiry': str(row.get('상담 내용', '')).strip(),
                        'keyword': keyword_norm,
                        'user_name': user_name,
                    })
            except Exception as exc:
                logger.error(f"[SYNC/전화WF] 보정 실패 (row {sheet_row}): {exc}",
                             exc_info=True)
                result['errors'] += 1

        # 방문 예약 카드 발송 + 슬랙 List 등록 — 슬랙 client 가져오기 (순환 import 회피)
        if notify_visit:
            try:
                from dashboard.blueprints.slack_bot import (
                    _slack_app, _post_visit_notice, _post_to_slack_list,
                )
                client = _slack_app.client if _slack_app else None
                if client:
                    for lead in notify_visit:
                        notice_channel, notice_ts = '', ''
                        try:
                            vd = lead['visit_date']
                            if vd.startswith("'"):
                                vd = vd[1:]
                            notice_channel, notice_ts = _post_visit_notice(
                                client, lead_no=lead['lead_no'], category='전화',
                                user_id='', visit_date=vd,
                                name=lead['name'], contact=lead['phone'],
                                visit_address=lead['address'],
                                consultation=lead['inquiry'],
                                user_name=lead.get('user_name', ''),
                            )
                            result['visit_notified'] += 1
                        except Exception as exc:
                            logger.error(
                                f"[SYNC/전화WF] #방문_일정 카드 실패 ({lead['lead_no']}): {exc}",
                                exc_info=True,
                            )

                        # 슬랙 List 워크플로우 등록 — 방문 예약 한정
                        # 원본 메시지 링크 = 방금 발송한 #방문_일정 카드
                        try:
                            lead_data = {
                                '리드 No': lead['lead_no'],
                                '고객명': lead['name'],
                                '고객 연락처': lead['phone'],
                                '이메일': lead.get('email', ''),
                                '상담 시간': lead.get('consult_time', ''),
                                '방문 주소': lead['address'],
                                '상담 내용': lead['inquiry'],
                                '키워드': lead.get('keyword', ''),
                                '플랫폼': '전화',
                            }
                            _post_to_slack_list(
                                client, lead_data,
                                modal_fields={
                                    'visit_date': vd,
                                    'visit_address': lead['address'],
                                    'consultation': lead['inquiry'],
                                    'estimate': '',
                                },
                                channel=notice_channel, message_ts=notice_ts,
                                action='visit',
                            )
                        except Exception as exc:
                            logger.error(
                                f"[SYNC/전화WF] 슬랙 List 등록 실패 ({lead['lead_no']}): {exc}",
                                exc_info=True,
                            )
            except Exception as exc:
                logger.warning(f"[SYNC/전화WF] 슬랙 client 로드 실패: {exc}")

        if result['processed'] > 0:
            invalidate_leads_cache()
        return result
    except Exception as exc:
        logger.error(f"[SYNC/전화WF] 동기화 실패: {exc}", exc_info=True)
        return result
