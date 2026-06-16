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


def _build_repeat_section(lead: dict) -> str:
    """재문의 컨텍스트 섹션 (이전 No + 이전 문의 내용)"""
    prev_leads = lead.get('_meta_previous_leads') or []
    if not prev_leads:
        return ''
    most_recent = prev_leads[0]
    prev_no = most_recent.get('lead_no', '')
    prev_inquiry = (most_recent.get('inquiry') or '').strip() or '-'
    # 길이 제한 (200자)
    if len(prev_inquiry) > 200:
        prev_inquiry = prev_inquiry[:200] + '...'

    return (
        f":repeat: *재문의 감지*\n"
        f">*이전 문의* {prev_no}\n"
        f">*문의 내용* : {prev_inquiry}\n"
        f"---------------------------------------------\n"
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

    lead_nos = []
    if new_leads:
        lead_nos = _append_leads_to_main(new_leads)
        _send_slack_notifications(new_leads, lead_nos, source='당근')

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


def _append_leads_to_main(leads: List[Dict[str, Any]]) -> List[str]:
    """
    메인 시트에 일괄 추가 + 리드No 자동 발번.

    Google Sheets API의 spreadsheets.values.append() 직접 호출.
    (mgr.append_row()는 시트명이 '공사 현황의 사본'으로 하드코딩돼 있어서 사용 불가)
    """
    if not leads:
        return []

    cfg = _get_sheet_config()
    if cfg is None:
        raise RuntimeError('ONLINE_LEADS_SHEET_ID 환경변수 미설정')

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
    return lead_nos


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


def _send_slack_notifications(leads: List[Dict[str, Any]], lead_nos: List[str], source: str = '당근'):
    """각 신규 리드를 사용자 양식으로 채널에 개별 메시지 전송"""
    bot_token = os.getenv('SLACK_BOT_TOKEN', '').strip()
    channel_setting = os.getenv('SLACK_LEAD_CHANNEL', '').strip()

    if not bot_token or 'your' in bot_token.lower():
        logger.warning(f'[SYNC/{source}] SLACK_BOT_TOKEN 미설정 - 슬랙 알림 스킵')
        return
    if not channel_setting or '여기에' in channel_setting:
        logger.warning(f'[SYNC/{source}] SLACK_LEAD_CHANNEL 미설정 - 슬랙 알림 스킵')
        return

    try:
        from slack_sdk import WebClient
        client = WebClient(token=bot_token)
    except Exception as exc:
        logger.error(f'[SYNC/{source}] Slack 클라이언트 초기화 실패: {exc}')
        return

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
            client.chat_postMessage(
                channel=channel,
                text=fallback,
                blocks=blocks,
                unfurl_links=False,
            )
        except Exception as exc:
            logger.error(f'[SYNC/{source}] 슬랙 전송 실패 ({ln}): {exc}')


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

    main_text = (
        f"*접수번호:* `{lead_no}`\n"
        f":bell: *{title}*\n"
        f"---------------------------------------------\n"
        + repeat_section
        + f">*문의시간* : {consult_time}\n"
        f">*이름 / 상호* : {name}\n"
        f">*연락처* : {phone}\n"
        f">*이메일* : {email}\n"
        f">*설치 희망 장소* : {place}\n"
        f">*설치 희망 기기* : {device}\n"
        f">*방문 주소* : {address_display}\n"
        f">*상세 문의 내용* : \n{inquiry}\n"
    )

    blocks = [
        {"type": "section", "text": {"type": "mrkdwn", "text": main_text}},
        {"type": "section", "text": {"type": "mrkdwn",
                                      "text": "---------------------------------------------"}},
        {
            "type": "actions",
            "elements": [
                {
                    "type": "button",
                    "text": {"type": "plain_text", "text": "방문 요청"},
                    "style": "primary",
                    "value": lead_no,
                    "action_id": "button_visit",
                },
                {
                    "type": "button",
                    "text": {"type": "plain_text", "text": "가격 문의"},
                    "style": "danger",
                    "value": lead_no,
                    "action_id": "button_price",
                },
            ],
        },
    ]
    fallback_text = f"[{source}] {lead_no} {name} / {phone}"
    return blocks, fallback_text
