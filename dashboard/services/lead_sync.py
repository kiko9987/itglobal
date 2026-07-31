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

import json
import os
import re
import textwrap
import threading
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

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
        '문의 내용': content,
        '상담 내용': '',
        '키워드': keyword,
        '온라인 상담자': '',
        '영업 담당자': '',
        '마지막 연락일': '',
        # 메타 (시트 등록 시 LEAD_COLUMN_ORDER 필터로 자동 제외)
        '_meta_place': place,
        '_meta_device': device,
        '_meta_inquiry': inquiry,
        '_meta_consult_dt': consult_dt,
        '_meta_address_level': address_level,  # verified면 정확, level5~7/raw면 _(추정)_ 표시
        '_meta_address_raw': address_raw,      # 원본 주소 표시용 (자동 정리 전)
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
    """주소 정규화 키 → entries (재문의 감지: 같은 건물 다른 사람).

    ETC-xxxxxx 리드(기타 방문 pseudo lead) 는 재문의 대상 아님 → skip.
    """
    lookup: dict = {}
    if main_df is None or main_df.empty:
        return lookup
    for _, row in main_df.iterrows():
        lead_no_str = str(row.get('리드 No', '') or '').strip()
        if lead_no_str.startswith('ETC-'):
            continue  # 기타 방문 pseudo lead 재문의 감지 대상 X (2026-07-15)
        addr_raw = str(row.get('방문 주소', '') or '')
        key = _normalize_address_key(addr_raw)
        if not key:
            continue
        consult_dt = _parse_consult_dt(row.get('상담 시간'))
        entry = {
            'lead_no': lead_no_str,
            'consult_dt': consult_dt,
            'consult_time': str(row.get('상담 시간', '') or ''),
            'status': str(row.get('상태', '') or ''),
            'feedback': str(row.get('상담 내용', '') or ''),  # 새 K열 (매니저 상담 결과)
            'inquiry': str(row.get('문의 내용', '') or ''),   # 새 J열 (인입 원본)
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
    ETC-xxxxxx 리드(기타 방문 pseudo lead) 는 재문의 대상 아님 → skip (2026-07-15).
    """
    lookup: dict = {}
    if main_df is None or main_df.empty:
        return lookup
    for _, row in main_df.iterrows():
        lead_no_str = str(row.get('리드 No', '') or '').strip()
        if lead_no_str.startswith('ETC-'):
            continue
        phone_raw = str(row.get('고객 연락처', '') or '')
        digits = re.sub(r'\D', '', phone_raw)
        if len(digits) < 10:
            continue
        consult_dt = _parse_consult_dt(row.get('상담 시간'))
        entry = {
            'lead_no': lead_no_str,
            'consult_dt': consult_dt,
            'consult_time': str(row.get('상담 시간', '') or ''),
            'status': str(row.get('상태', '') or ''),
            'feedback': str(row.get('상담 내용', '') or ''),  # 새 K열 (매니저 상담 결과)
            'inquiry': str(row.get('문의 내용', '') or ''),   # 새 J열 (인입 원본)
            'address': str(row.get('방문 주소', '') or ''),   # I열 — 재문의 시 이전 방문지 노출용
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
    prev_address = (most_recent.get('address') or '').strip() or '-'
    prev_feedback = (most_recent.get('feedback') or '').strip() or '-'
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
    if len(prev_feedback) > 200:
        prev_feedback = prev_feedback[:200] + '...'

    # 이전 문의 식별: "시간 / 리드No" 형태 (시간 없으면 리드No만)
    prev_label = f"{prev_time} / {prev_no}" if prev_time else prev_no

    # 방문 주소·상담 내용의 개행은 flatten (blockquote 구조 유지)
    prev_address = re.sub(r'\s*\n\s*', ' ', prev_address).strip() or '-'
    prev_feedback_oneline = re.sub(r'\s*\n\s*', ' ', prev_feedback).strip() or '-'

    # 문의 내용은 60자 wrap + 각 줄 '>' prefix (긴 라인 화면 wrap 시 인용 이탈 방지)
    prev_inquiry_quoted = _wrap_quoted(prev_inquiry, width=60)
    return (
        f">:repeat: *재문의 감지*\n"
        f">*이전 문의* : {prev_label}\n"
        f">*방문 주소* : {prev_address}\n"
        f">*문의 내용* :\n{prev_inquiry_quoted}\n"
        f">*상태* : {prev_status}\n"
        f">*상담자* : {prev_consultant}\n"
        f">*방문자* : {prev_sales_rep}\n"
        f">*상담 내용* : {prev_feedback_oneline}\n"
        f">--------------------------------------------\n"
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
        # 2-1) 상담시간 정확 매치도 skip (몇 개월 지난 원본이 재감지되는 케이스 방어)
        #      2026-07-15 관측: L-03269~L-03274 이기현 06/16 15:27 문의가 4번 append.
        #      원인은 phone_lookup 순간적 갱신 지연 (write-behind, force_refresh race).
        #      상담시간이 정확히 동일하면 동일 문의로 확정 → 무조건 skip.
        if phone_digits and phone_digits in phone_lookup and new_dt:
            _exact_match = any(
                e['consult_dt'] == new_dt for e in phone_lookup[phone_digits] if e['consult_dt']
            )
            if _exact_match:
                duplicates += 1
                continue
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
        # 순서 중요: 시트 등록 성공 직후 → pending 큐 선등록 → Slack 발송 → 성공분 큐에서 삭제
        # (SSL 에러가 Slack 발송 함수 내부 어디서 터지든 안전망 확보)
        # payload에 원본 lead dict(_meta_place/_meta_device 포함) JSON 저장 → 재발송 시 원본 그대로 복원
        rc = None
        try:
            from dashboard.utils.redis_client import get_redis_client
            rc = get_redis_client().redis
            for lead, ln in zip(new_leads, lead_nos):
                rc.set(
                    f'pending_slack_notify:{ln}',
                    _serialize_pending_payload('당근', lead),
                    ex=60 * 60 * 24 * 7,
                )
        except Exception as exc:
            logger.warning(f'[SYNC/karrot] pending 큐 선등록 실패 (계속 진행): {exc}')

        try:
            sent = _send_slack_notifications(new_leads, lead_nos, source='당근')
        except Exception as exc:
            logger.error(
                f'[SYNC/karrot] 슬랙 발송 함수 자체 실패 → pending 큐가 자동 재시도: {exc}'
            )
            sent = set()

        # 성공한 lead만 pending 큐에서 삭제
        if rc is not None and sent:
            try:
                for ln in sent:
                    rc.delete(f'pending_slack_notify:{ln}')
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
    (mgr.append_row()는 시트명이 '공사 현황'으로 하드코딩돼 있어서 사용 불가)

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

    # 다음 리드 No 시퀀스 — L-XXXXX 만 카운트 (ETC-xxxxxx 는 hex 라 제외)
    max_num = 0
    if df is not None and not df.empty and '리드 No' in df.columns:
        for ln in df['리드 No'].dropna().astype(str):
            m = re.match(r'^L-(\d+)$', ln.strip())
            if m:
                try:
                    max_num = max(max_num, int(m.group(1)))
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
    # range를 'A1:P1'로 한정해 시트의 다른 컬럼 영향 받지 않게 (헤더 영역만 참조)
    sheet_name = cfg['sheet_name']
    range_name = f"'{sheet_name}'!A1:P1"

    result = mgr.service.spreadsheets().values().append(
        spreadsheetId=cfg['sheet_id'],
        range=range_name,
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body={'values': rows},
    ).execute()

    # API 응답 유효성 검증 — SSL/네트워크 순간 오류로 partial 응답 방어
    # updates.updatedRange 없거나 updatedRows=0이면 실제 쓰기 실패로 간주 → raise
    updates = result.get('updates') or {}
    updated_range = updates.get('updatedRange')
    updated_rows = updates.get('updatedRows', 0)
    expected_rows = len(rows)
    if not updated_range or updated_rows < expected_rows:
        raise RuntimeError(
            f'Google Sheets append 응답 이상 — 쓰기 실패 추정 '
            f'(expected_rows={expected_rows}, got_rows={updated_rows}, '
            f'range={updated_range!r}, raw={result!r})'
        )

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
            # 2026-07-10 safe_slack_call — 리드 카드 놓치면 매니저가 신규 문의 인지 실패.
            from dashboard.blueprints.slack_helpers import safe_slack_call
            resp = safe_slack_call(
                client.chat_postMessage,
                channel=channel,
                text=fallback,
                blocks=blocks,
                unfurl_links=False,
            )
            if resp and resp.get('ok'):
                sent.add(ln)
                # 카드 ts 저장 — [기존 lead 연결] 시 그 카드에 ✅ reaction 추가용
                # (실패 시 orphan recovery가 중복 재발송할 수 있어 반드시 로그로 남긴다.
                #  이전엔 except pass라 조용히 실패 → 2026-07-08 L-03167 중복 발송 사고 유발)
                msg_ts = resp.get('ts', '')
                if msg_ts:
                    try:
                        from dashboard.utils.redis_client import get_redis_client
                        rc = get_redis_client().redis
                        rc.set(
                            f'lead_card_msg:{ln}',
                            f'{channel}|{msg_ts}',
                            ex=60 * 60 * 24 * 180,  # 180일 보존
                        )
                    except Exception as _exc:
                        logger.error(
                            f'[SYNC/{source}] lead_card_msg 저장 실패 ({ln}) — '
                            f'orphan recovery가 중복 재발송할 위험: {_exc}',
                            exc_info=True,
                        )
                else:
                    logger.warning(
                        f'[SYNC/{source}] 슬랙 응답에 ts 없음 ({ln}) — '
                        f'lead_card_msg 저장 스킵. orphan recovery 오탐 가능.'
                    )
            else:
                logger.error(f'[SYNC/{source}] 슬랙 전송 응답 not ok ({ln}): {resp}')
        except Exception as exc:
            logger.error(f'[SYNC/{source}] 슬랙 전송 실패 ({ln}): {exc}')
    return sent


def _post_phone_lead_completed_card(lead: dict, lead_no: str) -> bool:
    """전화 lead 시트 등록 후 온라인_문의 채널에 회색 처리 카드 + [재상담] 발송.

    2026-07-21 도입 — 매니저 통화 완료 후 등록 시맨틱 반영.
      - 인입 카드 원본을 code block 안에 담고
      - 회색 헤더 (상담 완료 - {상태}, 처리자, 처리 시간, 상담 내용) 씌우고
      - [재상담] 버튼 (action_id=button_consult) 부착
      - 재통화 시 매니저가 [재상담] 눌러 회차 이력 append (기존 online lead flow 재사용)
    """
    import re
    bot_token = os.getenv('SLACK_BOT_TOKEN', '').strip()
    channel_setting = os.getenv('SLACK_LEAD_CHANNEL', '').strip()
    if not bot_token or not channel_setting:
        logger.warning('[SYNC/전화WF/카드] SLACK_BOT_TOKEN/CHANNEL 미설정')
        return False
    try:
        from slack_sdk import WebClient
        client = WebClient(token=bot_token)
        channel = _resolve_channel_id(client, channel_setting)
    except Exception as exc:
        logger.error(f'[SYNC/전화WF/카드] slack 초기화 실패: {exc}')
        return False

    # lead dict 정규화 (2026-07-22): sync 흐름상 batchUpdate 이전 row 원본이
    # 넘어옴. 상담 시간 영문 포맷 그대로 카드에 들어가는 이슈 방지.
    lead = dict(lead)  # shallow copy — 호출자 dict 변형 방지
    _ct_raw = str(lead.get('상담 시간') or '').strip()
    _ct_norm = _normalize_workflow_datetime(_ct_raw) or _ct_raw
    if _ct_norm:
        lead['상담 시간'] = _ct_norm

    # 1. 인입 카드 blocks 로 원본 텍스트 조립 (기존 헬퍼 재사용)
    inquiry_blocks, _fallback = build_inquiry_blocks(lead, lead_no, source='전화')
    original_text = ''
    for blk in inquiry_blocks:
        if blk.get('type') == 'section':
            original_text = (blk.get('text') or {}).get('text', '')
            break

    # 2. blockquote/마크다운 제거 → code block 안 표시용 정리
    cleaned = [ln.lstrip('>').lstrip().replace('*', '') for ln in original_text.split('\n')]
    clean_text = re.sub(r'^[\s⠀]+|[\s⠀]+$', '', '\n'.join(cleaned))
    try:
        from dashboard.blueprints.slack_helpers import _normalize_shortcodes_to_unicode
        clean_text = _normalize_shortcodes_to_unicode(clean_text)
    except Exception:
        pass

    # 3. 헤더 조립 — 매니저 통화 완료 후 등록이므로 처음부터 회색 처리
    status = str(lead.get('상태') or '').strip() or '유선 상담'
    _register = str(lead.get('온라인 상담자') or '').strip().lstrip('@').strip()
    try:
        # _to_initial 헬퍼 재사용 — KiKO 예외 등 이니셜 정규화 로직 통일 (2026-07-22)
        from dashboard.blueprints.slack_helpers import _to_initial
        _processor = _to_initial(_register) or _register or '-'
        # DB lookup 결과가 'KIKO' 로 오는 케이스 후처리 (사용자 예외 표기)
        if _processor.upper() == 'KIKO':
            _processor = 'KiKO'
    except Exception:
        _processor = _register or '-'
    _consult_time = str(lead.get('상담 시간') or '').strip() or '-'
    _consult = str(lead.get('상담 내용') or '').strip() or '-'

    # 리드번호 배지 — 온라인 완료 카드(slack_bot.py 상담완료/취소 헤더)와 동일 컨벤션
    #   ':white_check_mark: *상담 완료 - {status}*  `{lead_no}`' (두 칸 + 백틱)
    _hdr_lno = f"  `{lead_no}`" if lead_no else ""
    header_lines = [
        '⠀',
        f':white_check_mark: *상담 완료 - {status}*{_hdr_lno}',
        f'처리자 : {_processor}',
        f'처리 시간 : {_consult_time}',
        f'상담 내용 : {_consult}',
    ]
    new_text = '\n'.join(header_lines) + f'\n\n```\n{clean_text}\n```'
    new_blocks = [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': new_text}},
        {'type': 'actions', 'elements': [{
            'type': 'button',
            'text': {'type': 'plain_text', 'text': '✏️ 재상담', 'emoji': True},
            'value': lead_no,
            'action_id': 'button_consult',
        }]},
    ]

    try:
        client.conversations_join(channel=channel)
    except Exception:
        pass
    try:
        from dashboard.blueprints.slack_helpers import safe_slack_call
        resp = safe_slack_call(
            client.chat_postMessage,
            channel=channel, text=new_text, blocks=new_blocks,
            unfurl_links=False,
        )
        if resp and resp.get('ok'):
            msg_ts = resp.get('ts', '')
            if msg_ts:
                try:
                    from dashboard.utils.redis_client import get_redis_client
                    rc = get_redis_client().redis
                    rc.set(
                        f'lead_card_msg:{lead_no}',
                        f'{channel}|{msg_ts}',
                        ex=60 * 60 * 24 * 180,
                    )
                except Exception as _exc:
                    logger.warning(
                        f'[SYNC/전화WF/카드] lead_card_msg 저장 실패 ({lead_no}): {_exc}'
                    )
                logger.info(f'[SYNC/전화WF/카드] 발송 완료 ({lead_no}, ts={msg_ts})')
                return True
        logger.warning(f'[SYNC/전화WF/카드] 발송 응답 not ok ({lead_no}): {resp}')
    except Exception as exc:
        logger.error(f'[SYNC/전화WF/카드] 예외 ({lead_no}): {exc}', exc_info=True)
    return False


def _post_active_phone_lead_card(lead: dict, lead_no: str) -> bool:
    """전화 lead 중 '견적 요청' 상태 → 회색 완료 카드 대신 활성 인입 카드 발송.

    2026-07-27 도입 — 전화로 "가견적만 달라"는 요청을 접수한 미처리 리드.
      - '견적 제출'(상담 후 결과) 과 달리 '견적 요청'은 아직 상담 전 단계.
      - 회색 완료 카드가 아니라 온라인 새 리드처럼 [상담하기] 버튼 달린 활성
        카드로 발송 → 담당 매니저가 견적 준비 후 [상담하기] 로 처리.
      - build_inquiry_blocks 재사용 (온라인 lead 와 동일 UX·action_id).
    """
    bot_token = os.getenv('SLACK_BOT_TOKEN', '').strip()
    channel_setting = os.getenv('SLACK_LEAD_CHANNEL', '').strip()
    if not bot_token or not channel_setting:
        logger.warning('[SYNC/전화WF/견적요청] SLACK_BOT_TOKEN/CHANNEL 미설정')
        return False
    try:
        from slack_sdk import WebClient
        client = WebClient(token=bot_token)
        channel = _resolve_channel_id(client, channel_setting)
    except Exception as exc:
        logger.error(f'[SYNC/전화WF/견적요청] slack 초기화 실패: {exc}')
        return False

    # 상담 시간 영문 포맷 정규화 (완료 카드 helper 와 동일 방어)
    lead = dict(lead)  # shallow copy — 호출자 dict 변형 방지
    _ct_raw = str(lead.get('상담 시간') or '').strip()
    _ct_norm = _normalize_workflow_datetime(_ct_raw) or _ct_raw
    if _ct_norm:
        lead['상담 시간'] = _ct_norm

    blocks, fallback = build_inquiry_blocks(lead, lead_no, source='전화')

    try:
        client.conversations_join(channel=channel)
    except Exception:
        pass
    try:
        from dashboard.blueprints.slack_helpers import safe_slack_call
        resp = safe_slack_call(
            client.chat_postMessage,
            channel=channel, text=fallback, blocks=blocks,
            unfurl_links=False,
        )
        if resp and resp.get('ok'):
            msg_ts = resp.get('ts', '')
            if msg_ts:
                try:
                    from dashboard.utils.redis_client import get_redis_client
                    rc = get_redis_client().redis
                    rc.set(
                        f'lead_card_msg:{lead_no}',
                        f'{channel}|{msg_ts}',
                        ex=60 * 60 * 24 * 180,
                    )
                except Exception as _exc:
                    logger.warning(
                        f'[SYNC/전화WF/견적요청] lead_card_msg 저장 실패 ({lead_no}): {_exc}'
                    )
                logger.info(f'[SYNC/전화WF/견적요청] 활성 카드 발송 완료 ({lead_no}, ts={msg_ts})')
                return True
        logger.warning(f'[SYNC/전화WF/견적요청] 발송 응답 not ok ({lead_no}): {resp}')
    except Exception as exc:
        logger.error(f'[SYNC/전화WF/견적요청] 예외 ({lead_no}): {exc}', exc_info=True)
    return False


def _post_visit_link_reply_to_lead_card(client, lead_no: str,
                                        visit_channel: str, visit_ts: str,
                                        user_name: str = '') -> None:
    """온라인 lead 카드 thread 에 방문 카드 permalink 링크 reply.

    매니저 수동 상담 완료(방문 예약) flow 의 permalink reply(_process_consult_submission)
    와 UX 통일 — 전화 workflow 자동 등록도 같은 형식으로 반영.

    lead_card_msg:{lead_no} → (channel, ts) 조회 후 thread 에 짧은 reply.
    """
    if not visit_channel or not visit_ts:
        return
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        v = rc.get(f'lead_card_msg:{lead_no}')
    except Exception:
        return
    if not v:
        return
    raw = v.decode() if isinstance(v, bytes) else v
    if '|' not in raw:
        return
    lead_ch, lead_ts = raw.split('|', 1)
    # permalink 조회
    try:
        perm = client.chat_getPermalink(channel=visit_channel, message_ts=visit_ts)
        permalink = (perm or {}).get('permalink', '')
    except Exception as exc:
        logger.debug(f'[SYNC/전화WF] permalink 조회 실패 ({lead_no}): {exc}')
        permalink = ''
    # 등록자 이니셜 (user_name 있으면 이니셜 추출)
    ini = '-'
    if user_name:
        try:
            from dashboard.blueprints.slack_helpers import _to_initial
            ini = _to_initial(user_name) or '-'
        except Exception:
            pass
    if permalink:
        reply_text = (
            f':white_check_mark: *방문 예약 등록* — `{lead_no}` by `{ini}`\n'
            f':round_pushpin: <{permalink}|#방문_일정 카드에서 상세 보기>'
        )
    else:
        reply_text = (
            f':white_check_mark: *방문 예약 등록* — `{lead_no}` by `{ini}`\n'
            f'_(#방문_일정 채널에 카드 발송됨)_'
        )
    try:
        client.chat_postMessage(
            channel=lead_ch, thread_ts=lead_ts, text=reply_text,
            unfurl_links=True,  # 방문 카드 unfurl embed 원함
        )
        logger.info(f'[SYNC/전화WF] permalink reply 완료 ({lead_no})')
    except Exception as exc:
        logger.warning(f'[SYNC/전화WF] permalink reply postMessage 실패 ({lead_no}): {exc}')


# ─────────────────────────────────────────────────────────────
# 인입 알림 블록 (홈페이지/당근/기타 플랫폼 공용) - 동적 타이틀
# ─────────────────────────────────────────────────────────────
def build_inquiry_blocks(lead: dict, lead_no: str, source: str = '당근') -> tuple:
    """
    Returns (blocks, fallback_text). 양식:

        *접수번호:* `L-02345`
        :bell: *새 문의 접수 알림 - 당근*
        --------------------------------------------

    플랫폼별 타이틀: `새 문의 접수 알림 - {source}` 형식으로 통일.
    새 플랫폼 추가 시 source 인자만 바꾸면 자동 적용.
        >*문의시간* : 2026.06.10. 18:21
        >*이름 / 상호* : 신호현
        >*연락처* : 010-2977-1698
        >*이메일* : -
        >*설치 희망 장소* : 사무실 / 관공서
        >*설치 희망 기기* : 천장형
        >*문의 내용* :
        실평9평. 15-20평형 천장형으로 견적원합니다...
        --------------------------------------------
        [방문 요청] [가격 문의]
    """
    consult_time = (lead.get('상담 시간') or '').strip() or '-'
    name = (lead.get('고객명') or '').strip() or '-'
    phone = (lead.get('고객 연락처') or '').strip() or '-'
    email = (lead.get('이메일') or '').strip() or '-'
    place = (lead.get('_meta_place') or '').strip() or '-'
    device = (lead.get('_meta_device') or '').strip() or '-'
    # 인입 원본은 새 J열 '문의 내용' — 옛 K열 '상담 내용'은 fallback (마이그레이션 호환)
    # 전화 유입은 J='-' (인입 텍스트 없음) + K에 통화 내용 → '-' 를 빈값처럼 취급해 K로 넘어감
    def _pick(v):
        s = str(v or '').strip()
        return s if s and s != '-' else ''
    inquiry = (
        _pick(lead.get('_meta_inquiry'))
        or _pick(lead.get('문의 내용'))
        or _pick(lead.get('상담 내용'))
        or '-'
    )
    # 슬랙 section text 3000자 한도 — 메타데이터 여유분 고려 안전선 2400자
    if len(inquiry) > 2400:
        inquiry = inquiry[:2400] + '\n…(내용이 길어 일부만 표시 — 시트 참조)'

    # 방문 주소 (자동 추출 또는 시트에 등록된 값) — 신뢰도 레벨에 따라 표시 분기
    address = (lead.get('방문 주소') or '').strip()
    addr_level = (lead.get('_meta_address_level') or '').strip()
    from dashboard.services.lead_helpers import is_blank_address, ADDRESS_MISSING_LABEL
    if is_blank_address(address):
        # 고객이 주소 미기입 (빈 값 or '()' 등) → 명시적 안내 (2026-07-27)
        address_display = ADDRESS_MISSING_LABEL
    elif address and address != '-':
        # 신뢰도별 표시:
        # - verified (카카오 매칭): 마커 없음 (정확)
        # - level1~4 (정규식 풀 패턴): 마커 없음 (정확)
        # - level5~7: _[추정]_ 마커 (영업 검증 필요)
        # - '' or 'raw' (resolve 완전 실패로 원본 그대로): _[주소 확인 필요]_ 마커
        #   → resolver 가 아무 처리 못 함 = "변환 실패" 이므로 [추정] 아닌 [확인 필요]
        #   → 전화번호/오타/문의글이 주소 필드로 들어온 케이스. 매니저 오방문 방지.
        #   (2026-07-21 raw level 도 [주소 확인 필요] 로 통일 — L-03314 사고)
        if addr_level in ('verified', 'level1', 'level2', 'level3', 'level3b', 'level4'):
            address_display = address
        elif not addr_level or addr_level == 'raw':
            # 2026-07-17 이탤릭 폰트 이질감 제거 → 이모지 + bold 로 강조
            address_display = f'{address}  ⚠️ *[주소 확인 필요]*'
        else:
            address_display = f'{address}  ⚠️ *[추정]*'
    else:
        address_display = '-'

    from dashboard.services.lead_helpers import format_inflow_display
    title = f"새 문의 접수 알림 - {format_inflow_display(source)}"

    # 단일 라인 필드는 개행 flatten — 필드값에 \n이 있으면 slack blockquote 구조가 깨져
    # 뒤 필드/구분선까지 밖으로 튀어나옴. (예: 고객이 폼 주소란에 여러 줄 입력)
    def _oneline(s: str) -> str:
        return re.sub(r'\s*\n\s*', ' ', str(s or '')).strip() or '-'
    consult_time = _oneline(consult_time)
    name = _oneline(name)
    phone = _oneline(phone)
    email = _oneline(email)
    place = _oneline(place)
    device = _oneline(device)
    address_display = _oneline(address_display)

    # 원본 주소 vs 자동 정리 결과 비교 — 다를 때만 두 줄로 표시 (2026-07-13).
    # 매니저가 파싱 결과가 원본에서 어떻게 변형됐는지 즉시 확인 → 오탐/축소 감지.
    # 공백 여러 개·앞뒤 공백만 다른 경우엔 실질 변화 없음 → 한 줄로 통합 (2026-07-21).
    address_raw = _oneline(lead.get('_meta_address_raw') or '')
    def _addr_key(s: str) -> str:
        return re.sub(r'\s+', ' ', (s or '').strip())
    address_field_lines = []
    if is_blank_address(address):
        # 주소 미입력이면 원본/변환 비교 없이 단일 라인 (2026-07-27)
        address_field_lines.append(f">*방문 주소* : {address_display}")
    elif (address_raw and address_raw != '-'
            and _addr_key(address_raw) != _addr_key(address)):
        address_field_lines.append(f">*원본 주소* : {address_raw}")
        address_field_lines.append(f">*변환 주소* : {address_display}")
    else:
        address_field_lines.append(f">*방문 주소* : {address_display}")
    address_field = '\n'.join(address_field_lines)

    # 재문의 감지 — 같은 번호로 이전 lead가 있으면 섹션 추가 (타이틀은 그대로 유지)
    repeat_section = _build_repeat_section(lead)

    # 헤더 + 구분선 + 본문 + 구분선 모두 한 `>` blockquote 그룹으로 통일.
    # 같은 blockquote 안 라인은 슬랙이 복사 시 줄바꿈을 정상 보존함.
    # 멀티라인 inquiry는 60자 wrap + 각 줄마다 `>` prefix.
    inquiry_quoted = _wrap_quoted(inquiry.rstrip(), width=60)
    main_text = (
        "⠀\n"
        f">:bell: *{title}*  `{lead_no}`\n"
        f">--------------------------------------------\n"
        + repeat_section
        + f">*문의시간* : {consult_time}\n"
        f">*이름 / 상호* : {name}\n"
        f">*연락처* : {phone}\n"
        f">*이메일* : {email}\n"
        f">*설치 희망 장소* : {place}\n"
        f">*설치 희망 기기* : {device}\n"
        + f"{address_field}\n"
        + f">*문의 내용* : \n{inquiry_quoted}\n"
        f">--------------------------------------------"
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

    # 영문 long form — 슬랙 워크플로 '{{Current time}}' 전용 포맷.
    # 슬랙이 UTC 기준으로 저장하므로 파싱 후 KST(+9h) 로 변환.
    dt = _parse_english_longform_dt(raw)
    if dt:
        dt = dt + timedelta(hours=9)
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


def _serialize_pending_payload(source: str, lead: Dict[str, Any]) -> str:
    """pending_slack_notify 큐 payload 직렬화.

    형식: JSON {"v": 2, "source": "당근", "lead": {...}}
    _meta_consult_dt 같은 datetime/pd.Timestamp는 ISO 문자열로 변환.
    """
    safe_lead = {}
    for k, v in lead.items():
        if v is None:
            safe_lead[k] = None
        elif isinstance(v, (str, int, float, bool)):
            safe_lead[k] = v
        elif hasattr(v, 'isoformat'):  # datetime, pd.Timestamp
            try:
                safe_lead[k] = v.isoformat()
            except Exception:
                safe_lead[k] = str(v)
        elif isinstance(v, (list, dict)):
            try:
                json.dumps(v, ensure_ascii=False)  # 검증
                safe_lead[k] = v
            except Exception:
                safe_lead[k] = str(v)
        else:
            safe_lead[k] = str(v)
    return json.dumps({'v': 2, 'source': source, 'lead': safe_lead}, ensure_ascii=False)


def _deserialize_pending_payload(raw) -> Optional[Dict[str, Any]]:
    """pending_slack_notify 큐 payload 역직렬화. 실패 시 None."""
    if not raw:
        return None
    if isinstance(raw, bytes):
        raw = raw.decode('utf-8', errors='ignore')
    try:
        obj = json.loads(raw)
        if isinstance(obj, dict) and obj.get('v') == 2 and 'lead' in obj:
            return obj
    except Exception:
        pass
    return None


def retry_pending_slack_notifications() -> Dict[str, Any]:
    """SSL 에러 등으로 슬랙 발송 누락된 lead 자동 재발송.

    Redis 'pending_slack_notify:{lead_no}' 지원 payload 형식:
      - v=2 JSON: {"source": "당근", "lead": {...원본 lead dict...}}  → 원본 그대로 복원
      - 레거시 문자열 (source만): 시트에서 lookup, _meta 필드는 최선 재구성
    성공한 lead의 마커는 삭제.
    """
    result = {'attempted': 0, 'sent': 0}
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        keys = list(rc.scan_iter(match='pending_slack_notify:*', count=100))
        if not keys:
            return result
        # 레거시(시트 lookup 필요) 케이스 대비 메인 시트 캐시 (필요 시에만 fetch)
        main_df = None
        pending_leads = []
        pending_nos = []
        pending_sources = []
        for k in keys:
            ks = k if isinstance(k, str) else k.decode()
            lead_no = ks.split(':')[-1]
            raw = rc.get(k)
            payload = _deserialize_pending_payload(raw)
            if payload:
                # v=2 JSON payload — 원본 lead dict 그대로 사용
                lead = dict(payload.get('lead') or {})
                source = payload.get('source') or '당근'
            else:
                # 레거시 (source 문자열만) — 시트에서 lookup + _meta 최선 재구성
                source = raw if isinstance(raw, str) else (raw.decode() if raw else '당근')
                if main_df is None:
                    main_df = load_leads_data(force_refresh=True)
                    if main_df is None or main_df.empty:
                        continue
                row = main_df[main_df['리드 No'].astype(str).str.strip() == lead_no]
                if row.empty:
                    rc.delete(k)
                    continue
                lead = row.iloc[0].to_dict()
                lead['_meta_place'] = ''
                lead['_meta_device'] = str(lead.get('키워드', '') or '')
                lead['_meta_inquiry'] = str(lead.get('문의 내용', '') or lead.get('상담 내용', '') or '')
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
            # 재시작 타이밍 방어 (2026-07-21) — 첫 발송이 성공했는데 pending 삭제 전
            # 재시작되면 pending 큐에 payload 가 남아 이 함수가 두 번째 카드를 발송함
            # (L-03309 사고). lead_card_msg 이 이미 있으면 pending 정리 후 skip.
            filtered = []
            for lead, ln in items:
                if rc.get(f'lead_card_msg:{ln}'):
                    logger.info(
                        f"[SYNC/retry] {ln} lead_card_msg 이미 있음 — pending 정리 후 skip"
                    )
                    rc.delete(f'pending_slack_notify:{ln}')
                    continue
                filtered.append((lead, ln))
            if not filtered:
                continue
            leads = [l for l, _ in filtered]
            nos = [n for _, n in filtered]
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


# orphan recovery 대상 플랫폼 — 우리 서버가 build_inquiry_blocks로 슬랙 알림을
# 직접 발송하는 플랫폼만 포함해야 함. 채널톡 SaaS가 슬랙에 직접 카드를 밀어넣는
# 카카오톡·채널톡은 lead_card_msg를 저장하지 않으므로 여기 넣으면 매번 오탐지
# 재발송된다 (2026-07-08 L-03167 사고). 전화(수동 입력)도 제외.
_AUTO_SYNC_PLATFORMS = {
    '홈페이지', '당근', '숨고', '큐플레이스', '메일', '모바일',
}


def recover_orphan_lead_notifications() -> Dict[str, Any]:
    """시트에는 등록됐지만 슬랙 카드 흔적이 전혀 없는 고아 리드 재발송.

    시나리오: 시트 append 성공 직후 pending 큐 등록 전에 Flask 재시작 →
    pending 큐도 lead_card_msg도 없어서 retry_pending_slack에 안 걸림.

    조건:
    - 상담 시간 최근 60분 이내 (오래된 것은 의도적 skip으로 간주)
    - 상담 시간 15분 이상 지남 (2026-07-08 L-03167 사고 이후 5분→15분: 첫 발송 성공 후
      lead_card_msg 저장이 지연되는 case 여유. 그 시간 안엔 매니저가 채팅 열기 등으로
      상담 시작할 수 있어 재발송이 명백한 오탐이 되기 때문)
    - 플랫폼이 자동 sync 대상 (전화·수동 입력 제외)
    - Redis에 lead_card_msg:{ln} 없음 (이미 발송된 것 skip)
    - Redis에 pending_slack_notify:{ln} 없음 (retry가 이미 처리 중)
    - 상태가 상담 대기 (스팸/드랍 등 이미 처리된 건 skip)
    - 재발송 직전에 lead_card_msg 한 번 더 확인 (double check)
    """
    from datetime import datetime, timedelta
    result = {'checked': 0, 'recovered': 0}
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        df = load_leads_data(force_refresh=True)
        if df is None or df.empty:
            return result

        now = datetime.now()
        cutoff_new = now - timedelta(minutes=15)  # 이보다 최근 = 아직 sync 발송/저장 중일 수 있음
        cutoff_old = now - timedelta(minutes=60)  # 이보다 오래 = 의도적 skip으로 간주

        # 최근 30건만 훑음 (성능)
        recent = df.tail(30)
        candidates = []
        for _, row in recent.iterrows():
            ln = str(row.get('리드 No', '') or '').strip()
            if not ln.startswith('L-'):
                continue
            platform = str(row.get('플랫폼', '') or '').strip()
            if platform not in _AUTO_SYNC_PLATFORMS:
                continue
            status = str(row.get('상태', '') or '').strip()
            if status and status != '상담 대기':
                continue  # 매니저가 이미 처리한 상태는 skip
            # 상담 시간 파싱
            consult_str = str(row.get('상담 시간', '') or '').strip()
            try:
                dt = datetime.strptime(consult_str, '%Y.%m.%d. %H:%M')
            except Exception:
                continue
            if dt < cutoff_old or dt > cutoff_new:
                continue
            # Redis 상태 확인
            if rc.get(f'lead_card_msg:{ln}'):
                continue  # 이미 카드 존재
            if rc.get(f'pending_slack_notify:{ln}'):
                continue  # pending 큐가 처리 중
            candidates.append((row, ln, platform, dt))

        result['checked'] = len(candidates)
        if not candidates:
            return result

        # 플랫폼별로 그룹핑 (source 인자 정확히 매칭)
        from collections import defaultdict
        groups = defaultdict(list)
        for row, ln, platform, dt in candidates:
            lead = row.to_dict()
            # _meta 필드 최선 재구성 (시트에는 없으니 K열 등에서 유추)
            lead['_meta_place'] = ''
            lead['_meta_device'] = str(lead.get('키워드', '') or '').strip()
            lead['_meta_inquiry'] = str(
                lead.get('문의 내용', '') or lead.get('상담 내용', '') or ''
            )
            lead['_meta_address_level'] = ''
            groups[platform].append((lead, ln))

        for src, items in groups.items():
            # 재발송 직전 lead_card_msg 재확인 — 스캔 시점과 발송 사이의 race 방지
            # (다른 스레드가 방금 성공 마킹을 저장했을 수 있음)
            recheck_items = []
            for lead, ln in items:
                if rc.get(f'lead_card_msg:{ln}'):
                    logger.info(
                        f"[SYNC/recover] {ln} 재발송 직전 재확인에서 lead_card_msg 발견 — "
                        f"오탐 방지 skip"
                    )
                    continue
                recheck_items.append((lead, ln))
            if not recheck_items:
                continue
            leads = [l for l, _ in recheck_items]
            nos = [n for _, n in recheck_items]
            sent = _send_slack_notifications(leads, nos, source=src)
            for ln in nos:
                if ln in sent:
                    result['recovered'] += 1
                    logger.warning(
                        f"[SYNC/recover] 고아 리드 슬랙 카드 재발송: "
                        f"{ln} (플랫폼={src}) — sync 발송 유실 감지"
                    )
    except Exception as exc:
        logger.error(f"[SYNC/recover] 고아 리드 감지 실패: {exc}", exc_info=True)
    return result


def retry_pending_visit_notices() -> Dict[str, Any]:
    """SSL 에러 등으로 방문 카드 발송 누락된 lead 자동 재발송.

    Redis 'pending_visit_notice:{lead_no}'에 payload(JSON) 저장 → 카드 재발송.
    성공한 lead의 마커는 삭제. 시트에 lead가 없으면 큐에서 제거.
    """
    result = {'attempted': 0, 'sent': 0}
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        keys = list(rc.scan_iter(match='pending_visit_notice:*', count=100))
        if not keys:
            return result

        from dashboard.blueprints.slack_bot import _slack_app, _post_visit_notice
        client = _slack_app.client if _slack_app else None
        if client is None:
            logger.debug('[SYNC/retry-visit] slack client 미준비 — 다음 폴링 대기')
            return result

        # 시트 검증용 fetch
        main_df = load_leads_data(force_refresh=True)
        existing_leads = set()
        if main_df is not None and not main_df.empty:
            existing_leads = set(main_df['리드 No'].astype(str).str.strip().tolist())

        for k in keys:
            ks = k if isinstance(k, str) else k.decode()
            lead_no = ks.split(':')[-1]
            payload_raw = rc.get(k)
            if not payload_raw:
                rc.delete(k)
                continue
            if lead_no not in existing_leads:
                rc.delete(k)
                continue
            try:
                p = json.loads(payload_raw if isinstance(payload_raw, str) else payload_raw.decode())
            except Exception:
                rc.delete(k)
                continue

            result['attempted'] += 1
            try:
                lead_platform = p.get('platform', '')
                if lead_platform == '전화':
                    card_category, card_platform = '온라인', '전화'
                else:
                    card_category, card_platform = lead_platform or '기타', ''
                vd = p.get('visit_date', '')
                if vd.startswith("'"):
                    vd = vd[1:]
                _post_visit_notice(
                    client, lead_no=lead_no,
                    category=card_category, platform=card_platform,
                    user_id='', visit_date=vd,
                    name=p.get('name', ''), contact=p.get('phone', ''),
                    visit_address=p.get('address', ''),
                    consultation=p.get('inquiry', ''),
                    user_name=p.get('user_name', ''),
                    addr_note=p.get('addr_note'),
                )
                rc.delete(k)
                result['sent'] += 1
            except Exception as exc:
                logger.warning(f'[SYNC/retry-visit] {lead_no} 재발송 실패, 다음 폴링 대기: {exc}')

        if result['sent']:
            logger.info(
                f"[SYNC/retry-visit] 미발송 방문 카드 재발송: {result['sent']}/{result['attempted']}"
            )
    except Exception as exc:
        logger.error(f"[SYNC/retry-visit] 재발송 처리 실패: {exc}", exc_info=True)
    return result


# ─────────────────────────────────────────────────────────────
# 전화 워크플로 dedup helpers (2026-07-14)
# ─────────────────────────────────────────────────────────────

def _find_existing_lead_by_phone(
    main_df, phone_raw: str, within_days: int = 90,
) -> Optional[Tuple[str, int]]:
    """정규화 연락처 매치 + 최근 within_days 이내 리드 검색.

    반환: (lead_no, sheet_row_1indexed) — 없으면 None.
    여러 개 매치 시 가장 최근 것.
    """
    from datetime import datetime, timedelta
    phone_norm = normalize_phone(phone_raw)
    if not phone_norm:
        return None
    cutoff = datetime.now() - timedelta(days=within_days)

    best = None  # (dt, lead_no, row_idx)
    for idx, r in main_df.iterrows():
        lead_no = str(r.get('리드 No', '') or '').strip()
        if not lead_no.startswith('L-'):
            continue
        this_phone = normalize_phone(str(r.get('고객 연락처', '') or ''))
        if this_phone != phone_norm:
            continue
        # 상담 시간 파싱 — 정규 포맷 + 워크플로 raw 포맷 (콜론→마침표) 지원.
        # L-03241 관측: '2026.07.14.14.18' 처럼 정규화 안 된 채 남은 리드도 매치.
        raw_dt = str(r.get('상담 시간', '')).strip()
        dt = None
        for fmt in ('%Y.%m.%d. %H:%M', '%Y.%m.%d.%H.%M', '%Y-%m-%d %H:%M'):
            try:
                dt = datetime.strptime(raw_dt, fmt)
                break
            except Exception:
                continue
        if dt is None:
            continue
        if dt < cutoff:
            continue
        if best is None or dt > best[0]:
            best = (dt, lead_no, idx + 2)  # sheet row = df idx + header(1) + 1-index
    if best is None:
        return None
    _, ln, sr = best
    return (ln, sr)


def _update_existing_lead_from_workflow(
    manager, cfg: dict, existing_sheet_row: int, workflow_row,
) -> None:
    """워크플로 임시 행의 값을 기존 리드 행에 덮어쓰기.

    정책: 빈 값이나 '-' 는 skip (기존값 보존). 워크플로 매니저가 채운 값만 반영.
    """
    from dashboard.services.lead_helpers import extract_keywords_from_sources
    # 필드 → 컬럼 매핑
    # 2026-07-27 (L-03406 중복 폭주): '상담 시간'(B) 은 최초 문의 시간이자 당근
    #   dedup 의 식별 키. 전화 워크플로 통화시간으로 덮어쓰면 원래 당근 메시지가
    #   기존 리드와 시간 매치 안 돼 (>1h) 재문의로 오판 → 매 폴링 중복 생성.
    #   → 상담 시간은 기존값(최초 문의시간) 보존, 워크플로 값으로 덮어쓰지 않음.
    field_col = [
        ('D', '상태'),
        ('E', '방문 예정일'),
        ('F', '고객 연락처'),
        ('G', '이메일'),
        ('H', '고객명'),
        ('I', '방문 주소'),
        ('J', '문의 내용'),
        ('K', '상담 내용'),
        ('L', '키워드'),
        ('M', '온라인 상담자'),
        ('N', '영업 담당자'),
        ('O', '마지막 연락일'),
    ]
    updates = []
    for col, field in field_col:
        v = str(workflow_row.get(field, '') or '').strip()
        if not v or v == '-':
            continue  # 빈값·'-' 는 기존값 보존
        # 필드별 정규화
        if field == '상담 시간':
            norm = _normalize_workflow_datetime(v)
            if norm:
                v = norm
        elif field == '방문 예정일':
            if '~' in v:
                pass  # 범위는 그대로
            else:
                iso = _normalize_workflow_date(v)
                if iso:
                    v = iso  # E열 셀 서식 '@ 텍스트' 라 escape 불필요
        elif field == '키워드':
            v = extract_keywords_from_sources(v) or v
        elif field == '온라인 상담자' or field == '영업 담당자':
            v = v.lstrip('@').strip()
        updates.append((f'{col}{existing_sheet_row}', v))

    if not updates:
        return
    batch_body = {
        'valueInputOption': 'USER_ENTERED',
        'data': [
            {'range': f"'{cfg['sheet_name']}'!{r}", 'values': [[val]]}
            for r, val in updates
        ],
    }
    manager.service.spreadsheets().values().batchUpdate(
        spreadsheetId=cfg['sheet_id'], body=batch_body,
    ).execute()


def _batch_delete_sheet_rows(manager, cfg: dict, row_numbers: list) -> None:
    """시트 행 batch delete (Google Sheets deleteDimension).

    row_numbers: 1-indexed sheet row 리스트. 앞 행 삭제 시 뒤 행 shift 되므로
    큰 rowIndex 부터 삭제.
    """
    if not row_numbers:
        return
    # 시트 numeric id 조회
    meta = manager.service.spreadsheets().get(
        spreadsheetId=cfg['sheet_id'],
        fields='sheets.properties(sheetId,title)',
    ).execute()
    sheet_num_id = None
    for s in meta.get('sheets', []):
        if s['properties']['title'] == cfg['sheet_name']:
            sheet_num_id = s['properties']['sheetId']
            break
    if sheet_num_id is None:
        raise RuntimeError(f'sheet not found: {cfg["sheet_name"]}')

    # 큰 rowIndex 부터 삭제 (shift 방지)
    requests = []
    for r in sorted(set(row_numbers), reverse=True):
        requests.append({
            'deleteDimension': {
                'range': {
                    'sheetId': sheet_num_id,
                    'dimension': 'ROWS',
                    'startIndex': r - 1,  # 0-indexed
                    'endIndex': r,
                },
            },
        })
    manager.service.spreadsheets().batchUpdate(
        spreadsheetId=cfg['sheet_id'],
        body={'requests': requests},
    ).execute()


def sync_workflow_phone_leads() -> Dict[str, Any]:
    """슬랙 워크플로우/수동 입력 모달이 메인 시트에 직접 추가한 lead 자동 보정.

    스캔 대상: 리드 No 빈 + 플랫폼 in {'전화', '거래처', '기타', '소개'} 행
      - 전화: 온라인_문의 채널 '전화문의 등록하기' 모달 (법인폰 유입)
      - 거래처: 방문_일정 채널 '방문요청 등록하기' 모달 (방문 유형=거래처)
      - 소개: 방문_일정 채널 '방문요청 등록하기' 모달 (방문 유형=소개)
      - 기타: 방문_일정 채널 '방문요청 등록하기' 모달 (방문 유형=기타).

    리드 번호 발번 (2026-07-15 시나리오 D):
      - 정규 (전화/거래처/소개) → L-XXXXX (연속 순번)
      - 기타 → ETC-xxxxxx (랜덤 hex, L 번호 소모 안 함)
      두 경우 모두 시트에 정상 등록. 대시보드 리드 목록은 '플랫폼 != 기타'
      필터로 기타를 시야에서 감춤. 매니저는 시트 안 보고 슬랙+대시보드만 사용.

    동시성 방어:
      - Redis 락 sync_workflow_phone_leads_lock (nx, TTL 30초) — 중복 sync
        실행 방지 (매니저 두 명 동시 등록 시 서버가 겹쳐 도는 것 방지).
    처리:
      1. lead_no 발번 (max + 1)
      2. 상담 시간 ISO → 'YYYY.MM.DD. HH:MM'
      3. 방문 예정일 ISO → 시트 escape ('YYYY-MM-DD)
      4. 키워드 vocab 매칭
      5. 행 색 분기 (상태별)
      6. 상태='방문 예약' 시 #방문_일정 카드 발송 (방문 수정/취소 버튼 자동 포함)
    """
    result = {'processed': 0, 'visit_notified': 0, 'errors': 0}

    # 동시성 방어 (2026-07-15): 매니저 두 명이 거의 동시에 등록 시 sync 가 중복
    # 실행되어 같은 리드를 두 번 처리하는 걸 방지. 5분 폴링 + 여러 워크플로
    # trigger 가 겹칠 때도 유효. TTL 30초 (sync 실제 처리 2~3초 + 여유).
    _sync_lock = None
    try:
        from dashboard.utils.redis_client import get_redis_client as _get_rc_sync
        _sync_lock = _get_rc_sync().redis
        _got_sync_lock = _sync_lock.set(
            'sync_workflow_phone_leads_lock', '1', nx=True, ex=30,
        )
        if not _got_sync_lock:
            logger.debug('[SYNC/전화WF] 다른 인스턴스 진행 중 — 이번 호출 skip')
            result['skipped_locked'] = True
            return result
    except Exception as exc:
        logger.warning(f'[SYNC/전화WF] 락 획득 실패 — 계속 진행: {exc}')
        _sync_lock = None

    try:
        main_df = load_leads_data(force_refresh=True)
        if main_df is None or main_df.empty:
            return result

        # 빈 lead_no + 플랫폼 in {'전화', '거래처', '기타', '소개'}
        WF_PLATFORMS = {'전화', '거래처', '기타', '소개'}
        is_empty_no = main_df['리드 No'].astype(str).str.strip() == ''
        is_target = main_df['플랫폼'].astype(str).str.strip().isin(WF_PLATFORMS)
        candidates = main_df[is_empty_no & is_target].copy()
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
        # 전화 워크플로 dedup 시 삭제할 임시 행 번호 (마지막에 batch delete)
        temp_rows_to_delete = []

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

            _platform_this = str(row.get('플랫폼', '') or '').strip()

            # 전화 워크플로 dedup (2026-07-14): 같은 연락처 90일 이내 리드가
            # 이미 있으면 기존 리드 update + 임시 행 삭제 (매니저 수동 시트
            # 등록 + 워크플로 이중 사용 사고 방지).
            if _platform_this == '전화' and _phone:
                existing = _find_existing_lead_by_phone(main_df, _phone, within_days=90)
                if existing is not None:
                    existing_lead_no, existing_sheet_row = existing
                    try:
                        _update_existing_lead_from_workflow(
                            manager, cfg, existing_sheet_row, row,
                        )
                        temp_rows_to_delete.append(sheet_row)
                        result.setdefault('updated_existing', 0)
                        result['updated_existing'] += 1
                        logger.info(
                            f'[SYNC/전화WF/DEDUP] 기존 리드 {existing_lead_no} 갱신 '
                            f'(임시 row {sheet_row} 삭제 예정, 연락처 매치: {_phone})'
                        )
                        # 상태가 방문 예약이면 갱신 카드 발송 (기존 lead_no)
                        _status = str(row.get('상태', '') or '').strip()
                        if _status == '방문 예약':
                            # 슬랙 카드 값 병합 — 워크플로 새 값 우선, 없으면 기존 리드값.
                            # (dedup 로직이 시트엔 기존값 유지시켰지만 카드 build 는
                            # workflow row 만 참조하면 빈값 카드 발송 이슈)
                            existing_row_data = main_df.loc[existing_sheet_row - 2]
                            def _pick(field):
                                v = str(row.get(field, '') or '').strip()
                                if v and v != '-':
                                    return v
                                v2 = str(existing_row_data.get(field, '') or '').strip()
                                return v2 if v2 and v2 != '-' else ''
                            _visit_date_raw = _pick('방문 예정일')
                            # 전화 문의는 '상담 내용' 우선 (매니저가 통화 후 정리한 것).
                            # 문의 내용은 온라인 리드용이라 전화에서는 fallback.
                            _inquiry = _pick('상담 내용') or _pick('문의 내용')
                            notify_visit.append({
                                'lead_no': existing_lead_no,
                                'name': _pick('고객명') or _name,
                                'phone': _pick('고객 연락처') or _phone,
                                'email': _pick('이메일'),
                                # 상담 시간은 기존 리드 최초 문의시간 우선 (dedup 키 보존, 2026-07-27)
                                'consult_time': (
                                    str(existing_row_data.get('상담 시간', '') or '').strip()
                                    or _pick('상담 시간')
                                ),
                                'address': _pick('방문 주소'),
                                'visit_date': _visit_date_raw,
                                'inquiry': _inquiry,
                                'keyword': _pick('키워드'),
                                'user_name': (
                                    _pick('영업 담당자').lstrip('@').strip()
                                    or _pick('온라인 상담자').lstrip('@').strip()
                                ),
                                # 기존 리드 플랫폼 우선 (당근 리드가 전화WF dedup 갱신돼도
                                # 카드는 원래 유입 채널 표시. 2026-07-27 L-03406)
                                'platform': (
                                    str(existing_row_data.get('플랫폼', '') or '').strip()
                                    or _platform_this
                                ),
                                'is_dedup_update': True,
                            })
                        continue
                    except Exception as exc:
                        logger.error(
                            f'[SYNC/전화WF/DEDUP] 기존 리드 갱신 실패 (신규 발번으로 진행): {exc}',
                            exc_info=True,
                        )
                        # dedup 실패 시 fallthrough → 신규 발번

            # 기타 방문은 L- 대신 ETC-xxxxxx (2026-07-15) — L 번호 연속 유지.
            # 정규 리드 (거래처/소개/전화) 는 L-XXXXX 순번.
            if _platform_this == '기타':
                from dashboard.blueprints.slack_bot import _etc_new_lead_no
                new_lead_no = _etc_new_lead_no()
            else:
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
            # 종료일 미선택 시 "2026-06-23~" 형태로 들어올 수 있음 → 단일로 처리
            visit_date_raw = str(row.get('방문 예정일', '')).strip()
            visit_date = visit_date_raw
            if visit_date_raw and '~' in visit_date_raw:
                parts = visit_date_raw.split('~', 1)
                start_iso = parts[0].lstrip("'").strip()
                end_iso = parts[1].strip() if len(parts) > 1 else ''
                if not end_iso:
                    # 종료일 비어있음 → 단일 날짜로 처리 (escape prefix + ISO)
                    iso = _normalize_workflow_date(start_iso)
                    if iso:
                        update_cells.append((f"E{sheet_row}", iso))
                        visit_date = iso
                else:
                    # 범위 — 양식 정돈 (같은 달이면 ~DD, 다른 달이면 ~MM-DD)
                    from dashboard.blueprints.slack_helpers import _format_visit_date_range
                    formatted = _format_visit_date_range(start_iso, end_iso)
                    if formatted and formatted != visit_date_raw:
                        # 범위 텍스트는 escape prefix 없이 (Google Sheets가 텍스트로 인식)
                        update_cells.append((f"E{sheet_row}", formatted))
                        visit_date = formatted
            elif visit_date_raw and not visit_date_raw.startswith("'"):
                iso = _normalize_workflow_date(visit_date_raw)
                if iso:
                    update_cells.append((f"E{sheet_row}", iso))
                    visit_date = iso
            elif not visit_date_raw:
                # 유선 상담 등 방문 예정일 없는 케이스 — 빈 값 대신 '-'
                update_cells.append((f"E{sheet_row}", '-'))
                visit_date = '-'

            # 키워드 vocab 매칭 (L열, 2026-07 J/K 스왑 이후 K→L 이동)
            # 매칭 실패 시 원본 값 유지 (수금·A/S·세척 등 서비스 값은 vocab 미포함 가능)
            keyword_raw = str(row.get('키워드', '')).strip()
            keyword_norm = extract_keywords_from_sources(keyword_raw) or keyword_raw or '-'
            if keyword_norm != keyword_raw:
                update_cells.append((f"L{sheet_row}", keyword_norm))

            # 온라인 상담자 (M열) — "@고광일" → "고광일" (시트는 실명, 슬랙 카드만 이니셜)
            counselor_raw = str(row.get('온라인 상담자', '')).strip()
            counselor_norm = counselor_raw.lstrip('@').strip() or '-'
            if counselor_norm != counselor_raw:
                update_cells.append((f"M{sheet_row}", counselor_norm))

            # 영업 담당자 (N열) — 2026-07 규칙:
            #   모든 리드에서 N열 = List 배정 방문자. 리드 생성 시점에는 미배정 → '-'.
            #   거래처의 owner 개념은 M열(온라인 상담자)에 있음 — 백엔드가 플랫폼별로 골라 읽음.
            # 2026-07-17 본인 방문 필수 자동 매핑:
            #   O열이 "본인 방문 필수" 이면 → M열 (워크플로 시작자) 을 N열 (영업 담당자) 로 복사.
            #   N열 이미 값 있으면 (매니저 수동 지정) 유지.
            sales_raw = str(row.get('영업 담당자', '') or '').strip()
            self_visit_val = str(row.get('본인 방문 여부', '') or '').strip()
            if '본인 방문 필수' in self_visit_val and (not sales_raw or sales_raw == '-'):
                _online_manager = str(row.get('온라인 상담자', '') or '').strip().lstrip('@')
                if _online_manager and _online_manager != '-':
                    update_cells.append((f"N{sheet_row}", _online_manager))
                else:
                    update_cells.append((f"N{sheet_row}", '-'))
            elif not sales_raw:
                update_cells.append((f"N{sheet_row}", '-'))

            # 방문 주소 자동 정규화 (2026-07-20) — 거래처/기타/소개 방문 예약 한정.
            # 매니저가 워크플로 모달에 raw 주소 (예: "경기도 부천시 오정동 810-1
            # 원일테크노2 4층 402호") 를 붙여넣은 경우 카카오 API로 도로명·건물명
            # 정정 → 시트 I열 update + notify_visit 에 addr_note 첨부 (카드 배지 +
            # 등록자 ephemeral). 전화는 매니저 개입 무의미하므로 대상 아님.
            # 온라인/당근은 별도 정규화 이미 있음 (address_resolver.resolve_address
            # 동일 함수 사용 → 포맷 일관).
            #
            # 특이사항 분리 (2026-07-24 L-03361): 매니저가 방문 주소 필드에 개행 후
            # 특이사항 ("방문 전 연락 요망" 등) 을 함께 입력한 케이스 처리.
            #   → 첫 줄 = 실제 주소, 2번째 줄+ = 특이사항으로 분리.
            #   → 상담 내용 (K열) 뒤에 ' / 특이사항' append. 원본 주소는 첫 줄만 유지.
            addr_raw_full = str(row.get('방문 주소', '') or '').strip()
            _addr_lines_raw = [
                ln.strip() for ln in re.split(r'\r?\n', addr_raw_full) if ln.strip()
            ]
            addr_raw = _addr_lines_raw[0] if _addr_lines_raw else ''
            extra_notes = ' / '.join(_addr_lines_raw[1:]) if len(_addr_lines_raw) > 1 else ''
            if extra_notes and _platform_this in {'거래처', '기타', '소개', '전화'}:
                _cur_consult = str(row.get('상담 내용', '') or '').strip()
                if _cur_consult and _cur_consult != '-':
                    _new_consult = f"{_cur_consult} / {extra_notes}"
                else:
                    _new_consult = extra_notes
                update_cells.append((f"K{sheet_row}", _new_consult))
                # I열도 첫 줄만으로 축소 (정규화 이전 raw 정리)
                update_cells.append((f"I{sheet_row}", addr_raw))
                logger.info(
                    f'[SYNC/방문주소] 특이사항 분리 ({new_lead_no}): '
                    f'주소="{addr_raw}", 특이사항="{extra_notes}"'
                )
            addr_note: Optional[Dict[str, str]] = None
            addr_for_notify = addr_raw
            # 온라인 완료 카드 주소 배지 판정용 — resolve 결과 레벨 캡처 (2026-07-31 L-03476)
            _addr_level_result: Optional[str] = None
            # 2026-07-22: 전화 추가 — 매니저 통화 후 raw 주소 입력 케이스도 자동 정규화
            #   (L-03313: '경기도 김포시 하성면 전류리 산57-22' → '김포 하성면 전류리 산 57-22')
            _addr_target_platforms = {'거래처', '기타', '소개', '전화'}
            _status_early = str(row.get('상태', '') or '').strip()
            if (
                addr_raw
                and _platform_this in _addr_target_platforms
                and _status_early == '방문 예약'
            ):
                try:
                    from dashboard.services import address_resolver as _ar
                    from dashboard.services import lead_helpers as _lh
                    _regex = _lh.extract_korean_address(addr_raw)
                    _regex_addr = _regex[0] if _regex else None
                    _regex_lv = _regex[1] if _regex else ''
                    _norm, _level = _ar.resolve_address(
                        addr_raw, _regex_addr, _regex_lv,
                    )
                    _addr_level_result = _level
                    if _level == 'verified' and _norm and _norm != addr_raw:
                        # 카카오 verified 정정 성공 & 원본과 다름 → 시트 update + 배지
                        update_cells.append((f"I{sheet_row}", _norm))
                        addr_for_notify = _norm
                        addr_note = {
                            'kind': 'normalized',
                            'original': addr_raw,
                            'normalized': _norm,
                        }
                    elif _level == 'verified' and _norm == addr_raw:
                        # verified 성공 & 원본 동일 → 조용히 (배지 없음)
                        pass
                    else:
                        # verified 아닌 모든 경우 (regex/level2/raw/실패) — 신뢰도 낮음.
                        # 2026-07-20 L-03289 관측: 매니저 오타 (충북 아산시 방배읍…) 가
                        # regex level2 fallback 으로 조용히 저장돼 오타 잡을 기회 없었음.
                        # → 원본 유지 + 확인 필요 배지로 매니저 재입력 유도.
                        addr_note = {
                            'kind': 'failed',
                            'original': addr_raw,
                            'normalized': '',
                        }
                except Exception as exc:
                    logger.warning(
                        f'[SYNC/전화WF] 주소 정규화 실패 ({new_lead_no}): {exc}'
                    )

            # 특이사항 이동 사실을 addr_note 에 첨부 (2026-07-24 L-03361)
            # → _post_addr_note_ephemeral 이 매니저에게만 안내 ephemeral 발송.
            if extra_notes and _platform_this in {'거래처', '기타', '소개', '전화'}:
                if addr_note is None:
                    # 정규화 조건 안 걸림 (성공 & 원본 동일 등) → note_only 로 알림만
                    addr_note = {
                        'kind': 'note_only',
                        'original': addr_raw,
                        'normalized': addr_raw,
                        'moved_notes': extra_notes,
                    }
                else:
                    addr_note['moved_notes'] = extra_notes

            # 빈 텍스트 셀 정규화 — 워크플로우가 값 안 넣은 옵션 필드는 '-' 로
            # J(문의 내용)는 전화 시맨틱상 항상 '-' — 워크플로우 매핑도 '-' 고정
            # 방문 주소(I) 는 위에서 별도 처리 (거래처 정규화) — 여기서는 원본 빈 값만 '-'.
            # 고객 연락처(F) 는 기타 방문 요청 시 공란 유입 케이스 대응 (2026-07-20).
            _dash_columns = [
                ('F', '고객 연락처'),
                ('G', '이메일'),
                ('H', '고객명'),
                ('I', '방문 주소'),
                ('J', '문의 내용'),
                ('O', '본인 방문 여부'),
            ]
            for col_letter, field in _dash_columns:
                current = str(row.get(field, '') or '').strip()
                if not current:
                    update_cells.append((f"{col_letter}{sheet_row}", '-'))

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
                        f"'{cfg['sheet_name']}'!A{sheet_row}:P{sheet_row}"
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

                # 전화 lead 온라인_문의 채널 카드 발송 (2026-07-21)
                # 매니저 통화 완료 후 등록 = 이미 상담 발생 → 회색 처리 카드 + [재상담] 버튼.
                # 재통화 시 매니저가 [재상담] 눌러 회차 이력 append (기존 online lead flow 재사용).
                # 단, '견적 요청'(가견적만 접수·상담 전) 은 회색 없이 활성 카드 (2026-07-27).
                _plat_row = str(row.get('플랫폼', '')).strip()
                if _plat_row == '전화':
                    try:
                        _lead_dict = row.to_dict() if hasattr(row, 'to_dict') else dict(row)
                        # 온라인 완료 카드도 방문 카드와 동일 정규화 소스 사용 (2026-07-31 L-03476)
                        #   row.to_dict() 는 시트 raw + _meta 미포함 → build_inquiry_blocks 가
                        #   addr_level='' 로 보고 항상 '주소 확인 필요' 배지를 붙였음. 방문 카드는
                        #   addr_for_notify(정규화본)를 써 배지 없이 원본/변환 2줄로 나오는데
                        #   온라인 완료 카드만 raw+배지라 불일치. 정규화 결과를 주입해 온라인 리드와
                        #   동일하게 원본/변환 표시 — verified=배지없음, 비verified=배지유지(오방문 방지).
                        if _addr_level_result is not None and addr_raw:
                            _lead_dict['방문 주소'] = addr_for_notify
                            _lead_dict['_meta_address_raw'] = addr_raw
                            _lead_dict['_meta_address_level'] = (
                                'verified' if _addr_level_result == 'verified' else ''
                            )
                        # 공백 유무 무관 매칭 (워크플로 '견적요청' / 시트 '견적 요청')
                        if status.replace(' ', '') == '견적요청':
                            _post_active_phone_lead_card(_lead_dict, new_lead_no)
                        else:
                            _post_phone_lead_completed_card(_lead_dict, new_lead_no)
                    except Exception as _exc:
                        logger.warning(
                            f'[SYNC/전화WF] 온라인 카드 발송 실패 ({new_lead_no}): {_exc}'
                        )

                # 방문 예약 시 #방문_일정 카드 + 슬랙 List 후보
                if status == '방문 예약':
                    # 등록자: 영업 담당자 > 온라인 상담자 (둘 다 한국 이름 가정)
                    # 워크플로우가 "@고광일" 형태로 write 하므로 @ 제거
                    user_name = (
                        str(row.get('영업 담당자', '')).strip().lstrip('@').strip()
                        or str(row.get('온라인 상담자', '')).strip().lstrip('@').strip()
                    )
                    notify_visit.append({
                        'lead_no': new_lead_no,
                        'name': str(row.get('고객명', '')).strip(),
                        'phone': str(row.get('고객 연락처', '')).strip(),
                        'email': str(row.get('이메일', '')).strip(),
                        'consult_time': consult_norm or consult_raw,
                        # 방문 주소: 정규화됐으면 정정본, 아니면 원본
                        'address': addr_for_notify,
                        'visit_date': visit_date,
                        'inquiry': (
                            # 전화 유입은 문의 내용='-' + 상담 내용에 통화 내용 → '-'를 빈값 취급
                            (lambda a, b: (a if a and a != '-' else b))(
                                str(row.get('문의 내용', '') or '').strip(),
                                str(row.get('상담 내용', '') or '').strip(),
                            )
                        ),
                        'keyword': keyword_norm,
                        'user_name': user_name,
                        'platform': str(row.get('플랫폼', '')).strip(),
                        # 주소 정규화 배지 (거래처/기타/소개 전용). None 이면 배지 없음.
                        'addr_note': addr_note,
                    })
            except Exception as exc:
                logger.error(f"[SYNC/전화WF] 보정 실패 (row {sheet_row}): {exc}",
                             exc_info=True)
                result['errors'] += 1

        # 전화 워크플로 dedup — 임시 행 batch delete (기존 리드에 이미 반영됨)
        if temp_rows_to_delete:
            try:
                _batch_delete_sheet_rows(manager, cfg, temp_rows_to_delete)
                logger.info(
                    f'[SYNC/전화WF/DEDUP] 임시 행 {len(temp_rows_to_delete)}개 삭제 완료'
                )
            except Exception as exc:
                logger.error(
                    f'[SYNC/전화WF/DEDUP] 임시 행 삭제 실패 (수동 정리 필요): {exc}',
                    exc_info=True,
                )

        # 방문 예약 카드 발송 + 슬랙 List 등록 — 슬랙 client 가져오기 (순환 import 회피)
        if notify_visit:
            # 안전망: 카드 발송 전에 pending 큐 선등록 → 성공 시 삭제
            # (SSL/네트워크 에러가 카드 발송 중 raise 되어도 자동 재시도)
            rc = None
            try:
                from dashboard.utils.redis_client import get_redis_client
                rc = get_redis_client().redis
                for lead in notify_visit:
                    payload = json.dumps({
                        'lead_no': lead['lead_no'],
                        'platform': lead.get('platform', ''),
                        'name': lead.get('name', ''),
                        'phone': lead.get('phone', ''),
                        'email': lead.get('email', ''),
                        'address': lead.get('address', ''),
                        'inquiry': lead.get('inquiry', ''),
                        'visit_date': lead.get('visit_date', ''),
                        'consult_time': lead.get('consult_time', ''),
                        'keyword': lead.get('keyword', ''),
                        'user_name': lead.get('user_name', ''),
                        'addr_note': lead.get('addr_note'),
                    }, ensure_ascii=False)
                    rc.set(
                        f'pending_visit_notice:{lead["lead_no"]}', payload,
                        ex=60 * 60 * 24 * 7,
                    )
            except Exception as exc:
                logger.warning(f'[SYNC/전화WF] 방문 카드 pending 큐 선등록 실패 (계속 진행): {exc}')

            try:
                from dashboard.blueprints.slack_bot import (
                    _slack_app, _post_visit_notice, _post_to_slack_list,
                )
                client = _slack_app.client if _slack_app else None
                if client:
                    for lead in notify_visit:
                        notice_channel, notice_ts = '', ''
                        card_ok = False
                        try:
                            vd = lead['visit_date']
                            if vd.startswith("'"):
                                vd = vd[1:]
                            # 카드 헤더 카테고리 결정:
                            #   전화 → 온라인 (전화)
                            #   거래처 → 거래처
                            #   기타 → 기타
                            lead_platform = lead.get('platform', '')
                            if lead_platform == '전화':
                                card_category, card_platform = '온라인', '전화'
                            else:
                                card_category, card_platform = lead_platform or '기타', ''
                            notice_channel, notice_ts = _post_visit_notice(
                                client, lead_no=lead['lead_no'],
                                category=card_category, platform=card_platform,
                                user_id='', visit_date=vd,
                                name=lead['name'], contact=lead['phone'],
                                visit_address=lead['address'],
                                consultation=lead['inquiry'],
                                user_name=lead.get('user_name', ''),
                                addr_note=lead.get('addr_note'),
                            )
                            result['visit_notified'] += 1
                            card_ok = True
                            # 2026-07-23: 전화 workflow 로 등록된 방문 예약도 온라인 lead
                            # 카드 thread 에 방문 카드 permalink 링크 reply.
                            # (매니저 수동 상담 완료 flow 와 동일 UX — L-03344 사례)
                            try:
                                _post_visit_link_reply_to_lead_card(
                                    client, lead['lead_no'], notice_channel, notice_ts,
                                    lead.get('user_name', ''),
                                )
                            except Exception as _exc:
                                logger.warning(
                                    f"[SYNC/전화WF] permalink reply 실패 ({lead['lead_no']}): {_exc}"
                                )
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
                                '문의 내용': lead['inquiry'],
                                '키워드': lead.get('keyword', ''),
                                '플랫폼': lead.get('platform', '') or '전화',
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

                        # 카드 발송 성공 → pending 큐에서 삭제 (List는 실패해도 카드가 뜨는게 우선)
                        if card_ok and rc is not None:
                            try:
                                rc.delete(f'pending_visit_notice:{lead["lead_no"]}')
                            except Exception:
                                pass
            except Exception as exc:
                logger.warning(f"[SYNC/전화WF] 슬랙 client 로드 실패: {exc}")

        if result['processed'] > 0:
            invalidate_leads_cache()
        return result
    except Exception as exc:
        logger.error(f"[SYNC/전화WF] 동기화 실패: {exc}", exc_info=True)
        return result
    finally:
        if _sync_lock is not None:
            try:
                _sync_lock.delete('sync_workflow_phone_leads_lock')
            except Exception:
                pass
