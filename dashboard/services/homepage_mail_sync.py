"""
홈페이지 문의 메일 자동 처리 (Gmail API + 아임웹 메일 파서)

흐름:
1. G Workspace 도메인 위임으로 kiko@itg-aircon.com Gmail INBOX 폴링
2. 키워드로 새 입력폼/게시판 메일 검색 (이미 처리한 것 제외)
3. 본문 파싱 → 7개 필드 추출 (옛 Apps Script 정규식 그대로)
4. 메인 시트 연락처와 dedup → 신규만 추가
5. 슬랙 알림 (source='홈페이지' or '게시판')
6. 처리 완료 라벨 부착 (다음 폴링 시 중복 처리 방지)

아임웹 메일 종류:
- 입력폼: "새 입력폼 응답이 접수되었습니다." 키워드
- 게시판: "위 예시 참고하여 작성 부탁드립니다." 키워드
"""

import os
import re
import base64
from datetime import datetime
from typing import Any, Dict, List, Optional

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from dashboard.services.lead_service import load_leads_data
from dashboard.services.lead_helpers import (
    normalize_phone as _h_normalize_phone,
    extract_korean_address,
    clean_multiline,
    extract_keywords_from_sources,
)
from dashboard.services.address_resolver import resolve_address
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)


# ─────────────────────────────────────────────────────────────
# 상수
# ─────────────────────────────────────────────────────────────
HOMEPAGE_PLATFORM = '홈페이지'
DEFAULT_STATUS = '상담 대기'

GMAIL_SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
]


def _processed_label() -> str:
    return os.getenv('HOMEPAGE_PROCESSED_LABEL', 'Processed-by-Bot').strip()


def _search_query() -> str:
    """
    미처리 메일만 검색 (Processed-by-Bot 라벨 제외 + 최근 N일만)

    HOMEPAGE_MAIL_DAYS_BACK 환경변수로 검색 기간 조절 (기본 3일).
    옛 메일이 매번 재처리되는 폭주 방지 + Gmail 라벨 부착 실패 시 안전망.
    """
    label = _processed_label()
    days_back = os.getenv('HOMEPAGE_MAIL_DAYS_BACK', '3').strip()
    time_filter = f'newer_than:{days_back}d ' if days_back and days_back.isdigit() else ''
    return (
        f'-in:sent -in:trash {time_filter}'
        '("새 입력폼 응답이 접수되었습니다." OR "위 예시 참고하여 작성 부탁드립니다.") '
        f'-label:{label}'
    )


# ─────────────────────────────────────────────────────────────
# Gmail 클라이언트 (도메인 위임)
# ─────────────────────────────────────────────────────────────
def _get_gmail_service():
    """서비스 계정 → kiko@itg-aircon.com 대리 자격으로 Gmail API 빌드"""
    creds_file = (
        os.getenv('GOOGLE_CREDENTIALS_FILE')
        or os.getenv('GOOGLE_APPLICATION_CREDENTIALS')
        or 'credentials.json'
    )
    if not os.path.exists(creds_file):
        raise FileNotFoundError(f'서비스 계정 인증서 없음: {creds_file}')

    delegated_user = os.getenv('HOMEPAGE_MAIL_USER', '').strip()
    if not delegated_user:
        raise ValueError('HOMEPAGE_MAIL_USER 환경변수 미설정')

    credentials = service_account.Credentials.from_service_account_file(
        creds_file, scopes=GMAIL_SCOPES,
    )
    delegated_creds = credentials.with_subject(delegated_user)
    return build('gmail', 'v1', credentials=delegated_creds, cache_discovery=False)


# ─────────────────────────────────────────────────────────────
# 메일 본문 추출 (Gmail message → plain text)
# ─────────────────────────────────────────────────────────────
def _decode_data(data: str) -> str:
    """base64url → string"""
    if not data:
        return ''
    return base64.urlsafe_b64decode(data + '==').decode('utf-8', errors='replace')


def _html_to_text(html: str) -> str:
    """간단한 HTML → text (태그 제거 + 공백 정리)"""
    text = re.sub(r'<br\s*/?>', '\n', html, flags=re.IGNORECASE)
    text = re.sub(r'</p>|</div>|</tr>|</li>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    # HTML entity 디코딩 (자주 쓰이는 것만)
    text = (text.replace('&nbsp;', ' ').replace('&lt;', '<')
            .replace('&gt;', '>').replace('&amp;', '&')
            .replace('&quot;', '"').replace('&#39;', "'"))
    # 연속 공백/줄바꿈 정리
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n+', '\n\n', text)
    return text.strip()


def _get_message_body_text(msg: dict) -> str:
    """Gmail API message dict → plain text 본문 (multipart 재귀)"""
    payload = msg.get('payload', {})

    def _walk(parts) -> str:
        # 1순위: text/plain
        for part in parts or []:
            if part.get('mimeType') == 'text/plain':
                data = part.get('body', {}).get('data')
                if data:
                    return _decode_data(data)
        # 2순위: 재귀 multipart
        for part in parts or []:
            if part.get('parts'):
                r = _walk(part['parts'])
                if r:
                    return r
        # 3순위: text/html → text
        for part in parts or []:
            if part.get('mimeType') == 'text/html':
                data = part.get('body', {}).get('data')
                if data:
                    return _html_to_text(_decode_data(data))
        return ''

    # 단일 body
    body = payload.get('body', {}).get('data')
    if body:
        text = _decode_data(body)
        if payload.get('mimeType') == 'text/html':
            text = _html_to_text(text)
        return text

    return _walk(payload.get('parts', []))


# ─────────────────────────────────────────────────────────────
# 아임웹 메일 파서 (옛 Apps Script 정규식 그대로)
# ─────────────────────────────────────────────────────────────
def _detect_category(body: str) -> str:
    if '새 입력폼 응답이 접수되었습니다.' in body:
        return '홈페이지'
    if '위 예시 참고하여 작성 부탁드립니다.' in body:
        return '게시판'
    return '기타'


def _safe_search(pattern: str, text: str, default: str = '') -> str:
    m = re.search(pattern, text)
    return m.group(1).strip() if m else default


def _clean(text: str) -> str:
    """본문 텍스트 정리 (불필요 마커/안내 문구 제거)"""
    text = re.sub(r'\[도면 및 이미지 붙여넣기 가능\]', '', text)
    text = re.sub(r'\[지역 표기 필수\s*/\s*장문 가능\]', '', text)
    text = re.sub(r'\[중복선택가능\]', '', text)
    text = re.sub(r'\[\*필수는 아니지만 견적을 보내드립니다\.\]', '', text)
    # 폼 라벨에 종종 붙는 안내 문구 (선행 위치만)
    text = re.sub(r'^\s*-\s*가정용은 취급하지 않습니다\.\s*\n?', '', text)
    text = re.sub(r'-\s*가정용은 취급하지 않습니다\.', '', text)
    return text.strip()


def parse_mail_body(body: str) -> Dict[str, Any]:
    """아임웹 메일 본문 → 7개 필드 dict. category에 따라 분기."""
    category = _detect_category(body)
    result = {'category': category}

    if category == '홈페이지':
        result['inquiry_time'] = _safe_search(r'등록시각([\s\S]+?)응답', body)
        result['name'] = _clean(_safe_search(r'상호([\s\S]+?)연락처', body))
        result['phone'] = _clean(_safe_search(r'연락처([\s\S]+?)이메일', body))
        email = _safe_search(
            r'이메일[\s\S]*?\n([\w._%+\-]+@[\w.\-]+\.[a-zA-Z]{2,})', body
        )
        result['email'] = email or ''
        result['place'] = _clean(
            _safe_search(r'선택해주세요\.\s*([\s\S]*?)\s*설치하실', body)
        )
        result['device'] = _clean(
            _safe_search(r'중복선택가능\]\s*([\s\S]+?)\s*문의 내용', body)
        )
        result['details'] = _clean(
            _safe_search(r'문의 내용([\s\S]+?)입력폼 관리하기', body)
        )

    elif category == '게시판':
        result['inquiry_time'] = _safe_search(
            r'등록시각([\s\S]+?)게시물', body
        ).replace('◆', '').strip()
        result['name'] = _safe_search(
            r'◆\s*이름 or 상호\s*:\s*([\s\S]+?)◆\s*연락처', body
        ).replace('◆', '').strip()
        result['phone'] = _safe_search(
            r'◆\s*연락처\s*:\s*([\s\S]+?)◆\s*메일 주소', body
        ).replace('◆', '').strip()
        email_raw = _safe_search(
            r'◆\s*메일 주소\s*:\s*([\s\S]+?)◆\s*설치 희망 장소', body
        ).replace('◆', '').strip()
        email_match = re.search(r'[\w._%+\-]+@[\w.\-]+\.[A-Za-z]+', email_raw)
        result['email'] = email_match.group(0) if email_match else ''
        result['place'] = _safe_search(
            r'◆\s*설치 희망 장소\s*:\s*([\s\S]+?)(?=◆\s*설치 희망 기기|\n|$)', body
        )
        result['device'] = _safe_search(
            r'◆\s*설치 희망 기기\s*:\s*([\s\S]+?)(?=◆\s*문의 내용|\n|$)', body
        )
        body_no_ex = re.sub(
            r'^[\s\S]*위 예시 참고하여 작성 부탁드립니다\.[\s\S]*?\n', '', body
        )
        result['details'] = _clean(_safe_search(
            r'문의 내용([\s\S]+?)게시물 확인하기', body_no_ex
        ))

    return result


# ─────────────────────────────────────────────────────────────
# 시간/연락처 정규화 (lead_sync와 동일)
# ─────────────────────────────────────────────────────────────
def _normalize_consult_time(raw: str) -> str:
    """아임웹 시간 → 메인 시트 형식 2026.05.29. 09:42"""
    raw = (raw or '').strip()
    if not raw:
        return ''
    for fmt in (
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%d %H:%M',
        '%Y.%m.%d %H:%M',
        '%Y.%m.%d. %H:%M',
        '%Y-%m-%d',
    ):
        try:
            return datetime.strptime(raw, fmt).strftime('%Y.%m.%d. %H:%M')
        except ValueError:
            continue
    return raw


def _normalize_phone(raw: str) -> str:
    """공용 lead_helpers의 강화된 정규화 사용 (휴대폰·서울·지역·070 모두)"""
    return _h_normalize_phone(raw)


# ─────────────────────────────────────────────────────────────
# 파싱 결과 → 메인 시트 lead 형식
# ─────────────────────────────────────────────────────────────
def to_lead(parsed: Dict[str, Any]) -> Dict[str, Any]:
    place = (parsed.get('place') or '').strip() or '-'
    # device가 멀티라인(중복선택)이면 콤마로 정리
    device_raw = (parsed.get('device') or '').strip()
    device = clean_multiline(device_raw, sep=', ') if device_raw else '-'
    inquiry = (parsed.get('details') or '').strip()
    name = (parsed.get('name') or '').strip()
    email = (parsed.get('email') or '').strip() or ''

    # 문의 내용에서 한국 주소 추출 (정규식 → 카카오 검증 → 원문 첫 줄 fallback)
    extracted_address = ''
    extract_level = ''
    if inquiry:
        regex_result = extract_korean_address(inquiry)
        regex_addr = regex_result[0] if regex_result else None
        regex_level = regex_result[1] if regex_result else ''
        extracted_address, extract_level = resolve_address(
            inquiry, regex_addr, regex_level
        )

    # 키워드: device + place + inquiry 에서 KEYWORD_VOCAB 매칭
    # K열(키워드) = 폼에서 선택한 device 값만 (옵션 B). vocab 매칭으로 정규화.
    # 추가 vocab(중고/매입/세척/소상공인)은 전화 통화 후 수동 입력용.
    keyword = extract_keywords_from_sources(device)

    # 상담 내용: 순수 문의 본문만 (장소/기기 prefix 제거)
    # 옛 양식: "장소: ... / 기기: ... / 문의: ..." → 문의 본문만
    content = inquiry

    return {
        '리드 No': '',
        '상담 시간': _normalize_consult_time(parsed.get('inquiry_time', '')),
        '플랫폼': HOMEPAGE_PLATFORM,
        '상태': DEFAULT_STATUS,
        '방문 예정일': '-',
        '고객 연락처': _normalize_phone(parsed.get('phone', '')),
        '이메일': email or '-',
        '고객명': name,
        '방문 주소': extracted_address or '-',
        '상담 내용': content,
        '키워드': keyword,
        '온라인 상담자': '',
        '영업 담당자': '',
        '마지막 연락일': '',
        '피드백': '',
        # 슬랙 메시지용 메타
        '_meta_place': place,
        '_meta_device': device,
        '_meta_inquiry': inquiry,
        '_meta_address_level': extract_level,  # 신뢰도 표시용
    }


# ─────────────────────────────────────────────────────────────
# Gmail 라벨 관리 (처리 완료 마킹)
# ─────────────────────────────────────────────────────────────
def _get_or_create_label(service, label_name: str) -> str:
    """라벨 ID 반환 (없으면 생성)"""
    labels = service.users().labels().list(userId='me').execute().get('labels', [])
    for lbl in labels:
        if lbl['name'] == label_name:
            return lbl['id']
    new_label = service.users().labels().create(
        userId='me',
        body={
            'name': label_name,
            'labelListVisibility': 'labelShow',
            'messageListVisibility': 'show',
        },
    ).execute()
    logger.info(f'[SYNC/홈페이지] Gmail 라벨 생성: {label_name}')
    return new_label['id']


def _mark_processed(service, msg_id: str, label_id: str) -> bool:
    """라벨 부착. 성공 여부 반환."""
    try:
        result = service.users().messages().modify(
            userId='me', id=msg_id,
            body={'addLabelIds': [label_id]},
        ).execute()
        # 라벨이 정말 추가됐는지 검증
        if label_id in result.get('labelIds', []):
            return True
        logger.error(f'[SYNC/홈페이지] 라벨 부착 응답이 이상함 ({msg_id}): labelIds={result.get("labelIds")}')
        return False
    except Exception as exc:
        logger.error(f'[SYNC/홈페이지] 라벨 부착 실패 ({msg_id}): {exc}', exc_info=True)
        return False


# ─────────────────────────────────────────────────────────────
# 메인 sync 함수
# ─────────────────────────────────────────────────────────────
def sync_homepage_email() -> Dict[str, Any]:
    """Gmail 폴링 → 메인 시트 추가 + 슬랙 알림. APScheduler가 주기적으로 호출."""
    try:
        service = _get_gmail_service()
    except Exception as exc:
        logger.error(f'[SYNC/홈페이지] Gmail 서비스 초기화 실패: {exc}')
        return {'error': str(exc)}

    try:
        label_id = _get_or_create_label(service, _processed_label())
    except Exception as exc:
        logger.error(f'[SYNC/홈페이지] 라벨 처리 실패: {exc}')
        return {'error': str(exc)}

    # 미처리 메일 검색
    try:
        results = service.users().messages().list(
            userId='me', q=_search_query(), maxResults=30,
        ).execute()
        msg_list = results.get('messages', [])
    except HttpError as exc:
        logger.error(f'[SYNC/홈페이지] 메일 검색 실패: {exc}')
        return {'error': str(exc)}

    if not msg_list:
        logger.debug('[SYNC/홈페이지] 미처리 메일 없음')
        return {'total': 0, 'new_count': 0, 'duplicates': 0, 'unknown': 0}

    # 메인 시트 dedup 인덱스 (연락처+시간 윈도우, 재문의 컨텍스트 포함)
    from dashboard.services.lead_sync import (
        _get_existing_phone_lookup, _append_leads_to_main,
    )
    from datetime import timedelta
    main_df = load_leads_data(force_refresh=True)
    phone_lookup = _get_existing_phone_lookup(main_df)
    seen_in_sync: Dict[str, datetime] = {}

    new_leads: List[Dict[str, Any]] = []
    new_by_source: Dict[str, List[Dict[str, Any]]] = {}
    # 시트엔 이미 있지만 라벨이 없는 메일 = 슬랙 미발송 추정 → 재발송 후보
    resend_leads: List[Dict[str, Any]] = []
    resend_by_source: Dict[str, List[Dict[str, Any]]] = {}
    duplicates = 0
    unknown = 0
    processed_msg_ids: List[str] = []

    for msg_meta in msg_list:
        msg_id = msg_meta['id']
        try:
            msg = service.users().messages().get(
                userId='me', id=msg_id, format='full',
            ).execute()
        except Exception as exc:
            logger.error(f'[SYNC/홈페이지] 메일 가져오기 실패 ({msg_id}): {exc}')
            continue

        body = _get_message_body_text(msg)
        parsed = parse_mail_body(body)
        category = parsed.get('category', '기타')

        if category == '기타':
            unknown += 1
            logger.warning(f'[SYNC/홈페이지] 알 수 없는 메일 형식 (id={msg_id})')
            processed_msg_ids.append(msg_id)
            continue

        lead = to_lead(parsed)
        phone_digits = re.sub(r'\D', '', lead['고객 연락처'])
        # 새 lead 시각 파싱
        try:
            new_dt = datetime.strptime(lead.get('상담 시간', ''), '%Y.%m.%d. %H:%M')
        except (ValueError, TypeError):
            new_dt = None
        lead['_meta_consult_dt'] = new_dt

        # 같은 sync 내 중복 (1시간 윈도우)
        if phone_digits and phone_digits in seen_in_sync and new_dt:
            if abs((new_dt - seen_in_sync[phone_digits]).total_seconds()) < 3600:
                duplicates += 1
                processed_msg_ids.append(msg_id)
                continue

        # 메인 시트 최근 lead와 1시간 이내 → 중복 (재폴링)
        # 라벨이 없는 메일이 여기 잡혔다 = 이전 폴링에서 시트 등록 후 슬랙 발송 실패
        # → 시트 기존 lead_no 그대로 슬랙만 재발송 (라벨은 발송 성공해야 부착)
        if phone_digits and phone_digits in phone_lookup and new_dt:
            latest = phone_lookup[phone_digits][0]
            if latest['consult_dt']:
                if abs((new_dt - latest['consult_dt']).total_seconds()) < 3600:
                    duplicates += 1
                    lead['_meta_msg_id'] = msg_id
                    lead['_meta_resend_lead_no'] = latest.get('lead_no', '')
                    resend_leads.append(lead)
                    resend_by_source.setdefault(category, []).append(lead)
                    continue
            # 1시간 이상 차이 → 재문의로 간주, 옛 이력 메타에 저장
            lead['_meta_previous_leads'] = phone_lookup[phone_digits]

        # msg_id를 lead에 보관 — 슬랙 발송 성공 시에만 라벨 부착 (D안)
        lead['_meta_msg_id'] = msg_id
        new_leads.append(lead)
        new_by_source.setdefault(category, []).append(lead)
        if phone_digits and new_dt:
            seen_in_sync[phone_digits] = new_dt
        # new_leads의 msg_id는 슬랙 발송 결과 보고 나서 라벨 부착
        # (여기선 processed_msg_ids에 추가하지 않음)

    # 응답 시각 오름차순 정렬 (가장 최신이 채널 마지막에)
    def _sort_key(l):
        try:
            return datetime.strptime(l.get('상담 시간', ''), '%Y.%m.%d. %H:%M')
        except ValueError:
            return datetime.min

    new_leads.sort(key=_sort_key)
    for src in new_by_source:
        new_by_source[src].sort(key=_sort_key)

    lead_nos: List[str] = []
    if new_leads:
        lead_nos = _append_leads_to_main(new_leads)

        # source별로 슬랙 알림 — 발송 성공한 lead_no만 set으로 받음
        from dashboard.services.lead_sync import _send_slack_notifications
        sent_lead_nos: set = set()
        for src, leads in new_by_source.items():
            src_lead_nos = []
            for lead in leads:
                idx = new_leads.index(lead)
                src_lead_nos.append(lead_nos[idx])
            sent = _send_slack_notifications(leads, src_lead_nos, source=src) or set()
            sent_lead_nos.update(sent)

        # D안: 슬랙 발송 성공한 new_lead의 msg_id만 라벨 부착 대상에 추가
        # 발송 실패 lead의 msg_id는 라벨 미부착 → 다음 폴링에서 재시도
        # 시트 자체 dedup(연락처+시간)으로 중복 등록은 방지됨
        for lead, ln in zip(new_leads, lead_nos):
            if ln in sent_lead_nos and lead.get('_meta_msg_id'):
                processed_msg_ids.append(lead['_meta_msg_id'])
            elif lead.get('_meta_msg_id'):
                logger.warning(
                    f'[SYNC/홈페이지] 슬랙 발송 실패 → 라벨 보류 ({ln}, msg_id={lead["_meta_msg_id"]})'
                )

    # 재발송 처리 — 시트엔 이미 있지만 슬랙 미발송 추정 케이스
    # (이전 폴링에서 시트 등록 후 발송 실패한 사고 케이스의 자동 복구)
    if resend_leads:
        from dashboard.services.lead_sync import _send_slack_notifications
        sent_resend: set = set()
        for src, leads in resend_by_source.items():
            r_lead_nos = [l.get('_meta_resend_lead_no', '') for l in leads]
            sent = _send_slack_notifications(leads, r_lead_nos, source=src) or set()
            sent_resend.update(sent)

        for lead in resend_leads:
            ln = lead.get('_meta_resend_lead_no', '')
            if ln in sent_resend and lead.get('_meta_msg_id'):
                processed_msg_ids.append(lead['_meta_msg_id'])
                logger.info(f'[SYNC/홈페이지] 재발송 완료 ({ln})')
            elif lead.get('_meta_msg_id'):
                logger.warning(
                    f'[SYNC/홈페이지] 재발송 실패 → 라벨 보류 ({ln}, msg_id={lead["_meta_msg_id"]})'
                )

    # 처리 완료 라벨 부착 (dup/unknown은 즉시, new_lead/resend는 슬랙 성공한 것만)
    for msg_id in processed_msg_ids:
        _mark_processed(service, msg_id, label_id)

    logger.info(
        f'[SYNC/홈페이지] total={len(msg_list)} new={len(new_leads)} '
        f'dup={duplicates} unknown={unknown}'
    )
    return {
        'total': len(msg_list),
        'new_count': len(new_leads),
        'duplicates': duplicates,
        'unknown': unknown,
        'lead_nos': lead_nos,
        'by_source': {src: len(leads) for src, leads in new_by_source.items()},
    }
