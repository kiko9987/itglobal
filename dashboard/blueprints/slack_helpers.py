"""슬랙 봇 공통 유틸 — slack_bot.py에서 추출.

다른 모듈 의존성 적은 pure utility 함수만:
- 모달 state 추출 (_v, _v_multi)
- 시트 날짜 포맷 (_format_date_for_sheet)
- 이니셜 매핑 (_to_initial, _slack_user_to_initial, _slack_user_to_korean_name)
- 시간 표시 (_human_duration)
"""
import json
import os
import re

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)


# ─────────────────────────────────────────────────────────────
# 슬랙 텍스트 길이 제한 대응 (2026-07-10)
# Slack blocks 각 section text 는 3000자 제한. 초과 시 API 400 반환.
# 문의 내용·상담 내용 등 사용자 입력 필드가 넘칠 위험.
# ─────────────────────────────────────────────────────────────
SLACK_SECTION_TEXT_MAX = 3000
SLACK_BLOCK_TEXT_MAX = 3000
SLACK_MSG_TEXT_MAX = 40000  # message.text 는 40KB

_TRUNC_SUFFIX = '\n… _(내용 길이 초과 — 관리 사이트에서 전체 확인)_'


def slack_truncate(text: str, max_len: int = SLACK_SECTION_TEXT_MAX) -> str:
    """슬랙 text 필드용 안전 truncate.

    - None/빈 문자열 → 그대로
    - 길이 이하 → 그대로
    - 초과 → max_len-len(suffix) 까지 자르고 suffix 부착 (전체 확인 안내)
    """
    if not text:
        return text or ''
    if len(text) <= max_len:
        return text
    suffix_len = len(_TRUNC_SUFFIX)
    keep = max_len - suffix_len
    if keep < 100:
        # max_len 이 너무 작으면 suffix 없이 최대한 잘라 반환
        return text[:max_len]
    return text[:keep].rstrip() + _TRUNC_SUFFIX


def safe_slack_call(client_method, *args, max_retries: int = 3, **kwargs):
    """슬랙 API 호출 rate limit·transient 오류 자동 대응 wrapper.

    - 429 → Retry-After 헤더 존중해 대기 후 재시도 (최대 max_retries 회)
    - 5xx / 네트워크 오류 → 지수 백오프 재시도
    - channel_not_found / not_in_channel / is_archived → 즉시 실패 반환 (재시도 무의미)
    - 성공 응답 반환, 최종 실패 시 마지막 예외 raise

    사용:
        response = safe_slack_call(client.chat_postMessage, channel='#foo', text='hi')

    Notes:
    - 지금까지 슬랙 API 호출은 대부분 단일 시도. sweep 같은 대량 발송은
      이미 별도로 sleep(1.1) 적용됨. 이 wrapper 는 신규 코드에서 활용.
    - 기존 41개 chat_postMessage 지점을 일괄 wrapping 하는 건 리스크가 커서
      필요 시점에 하나씩 이전.
    """
    import time as _t
    FATAL_ERRORS = {'channel_not_found', 'not_in_channel', 'is_archived',
                    'not_authed', 'invalid_auth', 'account_inactive'}
    last_exc = None
    for attempt in range(max_retries):
        try:
            resp = client_method(*args, **kwargs)
            return resp
        except Exception as exc:
            last_exc = exc
            err_code = ''
            retry_after = 0
            # SlackApiError 인 경우
            if hasattr(exc, 'response'):
                try:
                    err_code = (exc.response.get('error') or '') if hasattr(exc.response, 'get') else ''
                    retry_after = int(exc.response.headers.get('Retry-After', 0)) if hasattr(exc.response, 'headers') else 0
                except Exception:
                    pass
            # 재시도 무의미한 fatal error
            if err_code in FATAL_ERRORS:
                logger.warning(f'[SLACK/API] Fatal error {err_code} — 재시도 안 함')
                raise
            # rate limit
            if err_code == 'ratelimited' or getattr(exc, 'status', 0) == 429:
                wait = max(retry_after, 1)
                logger.warning(
                    f'[SLACK/API] Rate limited (attempt {attempt+1}/{max_retries}) — '
                    f'{wait}s 후 재시도'
                )
                _t.sleep(wait)
                continue
            # 다른 transient 오류 (네트워크 등) → 지수 백오프
            if attempt < max_retries - 1:
                wait = 2.0 * (attempt + 1)
                logger.warning(
                    f'[SLACK/API] transient 오류 ({type(exc).__name__}) '
                    f'attempt {attempt+1}/{max_retries} — {wait}s 후 재시도: {exc}'
                )
                _t.sleep(wait)
            else:
                logger.error(f'[SLACK/API] {max_retries}회 재시도 실패: {exc}', exc_info=True)
    if last_exc:
        raise last_exc
    return None


# ─────────────────────────────────────────────────────────────
# 시트 날짜 포맷
# ─────────────────────────────────────────────────────────────
def _format_date_for_sheet(value: str) -> str:
    """단일 ISO 날짜("2026-06-25")는 escape prefix 추가, 범위 텍스트는 그대로.

    Google Sheets는 USER_ENTERED 모드에서 ISO 날짜를 시리얼 숫자로 자동 변환.
    작은따옴표 prefix는 Sheets의 텍스트 escape 문자 — UI/API에 표시되지 않음.
    범위 텍스트('~' 포함)는 시트가 텍스트로 인식하므로 escape 불필요.
    """
    if not value:
        return ''
    if '~' in value:
        return value  # 범위 텍스트는 그대로
    return f"'{value}"


def _split_visit_date_range(value: str) -> tuple:
    """범위 양식 → (start_iso, end_iso) 분리. 단일이면 (start, '').

    - "2026-06-23~24" → ("2026-06-23", "2026-06-24")  # 같은 년월
    - "2026-07-01~08-15" → ("2026-07-01", "2026-08-15")  # 같은 년
    - "2026-12-30~2027-01-02" → ("2026-12-30", "2027-01-02")  # 다른 년
    - "2026-06-25" → ("2026-06-25", "")  # 단일
    - "'2026-06-25" → ("2026-06-25", "")  # escape prefix 제거
    """
    if not value:
        return ('', '')
    value = value.lstrip("'").strip()
    if '~' not in value:
        return (value, '')
    start, end = value.split('~', 1)
    start = start.strip()
    end = end.strip()
    if not end:
        return (start, '')
    # end="DD" (2자리 숫자) — 같은 년월
    if len(end) == 2 and end.isdigit():
        parts = start.split('-')
        if len(parts) == 3:
            return (start, f'{parts[0]}-{parts[1]}-{end}')
    # end="MM-DD" (5자, 가운데 '-') — 같은 년
    if len(end) == 5 and end[2] == '-':
        parts = start.split('-')
        if len(parts) == 3:
            return (start, f'{parts[0]}-{end}')
    # end="YYYY-MM-DD" — 다른 년
    return (start, end)


def _format_visit_date_range(start_iso: str, end_iso: str) -> str:
    """방문 예정일 표시 양식 — 시작/종료 ISO 날짜 → 사용자 친화적 양식.

    - end 없거나 start와 동일: "2026-06-25"
    - 같은 년월: "2026-06-23~24"
    - 같은 년, 다른 월: "2026-07-01~07-03"
    - 다른 년: "2026-12-30~2027-01-02"
    """
    if not start_iso:
        return ''
    if not end_iso or end_iso == start_iso:
        return start_iso
    s = start_iso.split('-')
    e = end_iso.split('-')
    if len(s) != 3 or len(e) != 3:
        return f'{start_iso}~{end_iso}'
    if s[0] == e[0] and s[1] == e[1]:
        # 같은 년월 — 일자만
        return f'{start_iso}~{e[2]}'
    if s[0] == e[0]:
        # 같은 년, 다른 월 — MM-DD
        return f'{start_iso}~{e[1]}-{e[2]}'
    # 다른 년 — 전체
    return f'{start_iso}~{end_iso}'


# ─────────────────────────────────────────────────────────────
# 모달 state 추출
# ─────────────────────────────────────────────────────────────
def _v(state, block_id, default=''):
    """모달 state.values에서 안전하게 값 추출 (datepicker / text / select 자동 분기)"""
    try:
        item = state[block_id]["value"]
        return (
            item.get("selected_date")
            or item.get("value")
            or (item.get("selected_option") or {}).get("value")
            or default
        )
    except Exception:
        return default


def _v_multi(state, block_id) -> list:
    """멀티 선택 체크박스/multi_static_select 값 추출"""
    try:
        item = state[block_id]["value"]
        opts = item.get("selected_options") or []
        return [o.get("value") for o in opts if o.get("value")]
    except Exception:
        return []


# ─────────────────────────────────────────────────────────────
# 직원 이니셜 매핑 (시트 A열 수식과 동일)
# ─────────────────────────────────────────────────────────────
# 조회 우선순위:
#   1) Redis 캐시 `user:project_code_suffix:{email}` — 관리자 UI 편집 결과
#      (users.py `_get_user_suffix` 와 동일 소스)
#   2) project_config.json 의 `owner_suffix_map`
#   3) 아래 하드코딩 SALES_INITIALS (부팅 오류 등 fallback)
# 신규 매니저 이니셜 추가 시:
#   - 관리자 UI `/admin/users` 에서 project_code_suffix 편집 (Redis 반영 즉시)
#   - 또는 `dashboard/project_config.json` `owner_suffix_map` 수정 후 재시작
_SALES_INITIALS_FALLBACK = {
    '박용구': 'YG', '박정우': 'JW', '강성환': 'SH', '박민우': 'MW',
    '이근혁': 'GH', '김호중': 'HJ', '아이티': 'IT', '김단이': 'DN',
    '권태훈': 'TH', '주영민': 'YM', '심장원': 'SJW', '빈승정': 'SJ',
    '박민재': 'MJ', '조성헌': 'JSH', '황해승': 'HS', '강민석': 'MS',
    '강정권': 'JK', '이상덕': 'SD', '고광일': 'KiKO',
}


def _load_initials_from_config() -> dict:
    """project_config.json + 하드코딩 fallback 병합. 실패 시 fallback만."""
    merged = dict(_SALES_INITIALS_FALLBACK)
    try:
        from dashboard.services import project_service
        cfg_map = (project_service.PROJECT_CONFIG or {}).get('owner_suffix_map') or {}
        merged.update({str(k): str(v) for k, v in cfg_map.items() if k and v})
    except Exception:
        pass  # fallback 만 사용
    return merged


def _lookup_initial_by_name(name: str) -> str:
    """이름 → 이니셜. Redis(관리자 편집) → config → fallback 순."""
    if not name:
        return ''
    # 1) users 테이블 조회 → email 알아내서 Redis 확인
    try:
        from dashboard.utils.user_database import get_user_database
        from dashboard.utils.redis_client import get_redis_client
        users = get_user_database().get_all_users() or []
        target_email = None
        for u in users:
            if u.get('name', '').strip() == name.strip():
                target_email = u.get('email', '').strip()
                break
        if target_email:
            rc = get_redis_client()
            v = rc.get(f'user:project_code_suffix:{target_email}')
            if v:
                return v.decode() if isinstance(v, bytes) else v
    except Exception:
        pass
    # 2) config + fallback 병합 조회
    return _load_initials_from_config().get(name, '')


# 하위 호환 (기존 코드가 SALES_INITIALS[name] 형태로 참조하는 경우 대비)
SALES_INITIALS = _SALES_INITIALS_FALLBACK


# Slack이 저장 시 unicode 이모지를 :bell: 같은 shortcode로 정규화해
# body에서 다시 읽을 때 shortcode 상태로 돌아옴. 정상 mrkdwn 렌더링에선 문제 없지만
# ``` 코드블록 안에 들어가면 텍스트로 남아 가독성 저하.
# 취소 UI(원본 카드를 회색 코드블록으로 감쌈)에서 재변환 필요.
_SHORTCODE_TO_EMOJI = {
    ':bell:': '🔔',
    ':inbox_tray:': '📥',
    ':office:': '🏢',
    ':round_pushpin:': '📍',
    ':bust_in_silhouette:': '👤',
    ':telephone_receiver:': '📞',
    ':envelope:': '✉️',
    ':email:': '✉️',
    ':clipboard:': '📋',
    ':hammer_and_wrench:': '🛠️',
    ':construction_worker:': '👷',
    ':heavy_dollar_sign:': '💲',
    ':moneybag:': '💰',
    ':date:': '📅',
    ':file_folder:': '📁',
    ':paperclip:': '📎',
    ':white_check_mark:': '✅',
    ':white_large_square:': '⬜',
    ':no_entry_sign:': '🚫',
    ':no_entry:': '⛔',
    ':memo:': '📝',
    ':bulb:': '💡',
    ':warning:': '⚠️',
    ':speech_balloon:': '💬',
    ':point_right:': '👉',
    ':receipt:': '🧾',
    ':pencil2:': '✏️',
    ':x:': '❌',
    ':arrow_backward:': '◀️',
    ':leftwards_arrow_with_hook:': '↩️',
    ':house:': '🏠',
    ':wrench:': '🔧',
    ':gear:': '⚙️',
}


def _normalize_shortcodes_to_unicode(text: str) -> str:
    """`:name:` shortcode → unicode 이모지. 매핑 없는 건 그대로.

    Slack 코드블록(```) 안에서 shortcode는 텍스트로 보여 가독성이 심각히 저하되는데
    이 함수로 정규화하면 코드블록 안에서도 이모지가 정상 렌더된다.
    """
    if not text:
        return text
    for k, v in _SHORTCODE_TO_EMOJI.items():
        if k in text:
            text = text.replace(k, v)
    return text


def _to_initial(name: str) -> str:
    """한국 이름 / 이니셜 / 빈값 → 이니셜 통일.
    - 한국 이름 → Redis(관리자 편집) → project_config.json → fallback dict 순 조회
    - 영문 2~5자 → 대문자 ('KiKO' 예외)
    - 매핑 없으면 원본 그대로
    """
    if not name:
        return ''
    name = name.strip()
    if not name or name in ('-', '미정'):
        return ''
    looked = _lookup_initial_by_name(name)
    if looked:
        return looked
    if re.match(r'^[A-Za-z]{2,5}$', name):
        if name.lower() == 'kiko':
            return 'KiKO'
        return name.upper()
    return name


def _slack_user_to_korean_name(client, user_id: str) -> str:
    """슬랙 user_id → SALES_EMAILS 매핑 한국 이름 (fallback: display_name/real_name)"""
    if not user_id:
        return ''
    try:
        resp = client.users_info(user=user_id)
        if not resp.get("ok"):
            return ''
        profile = resp["user"]["profile"]
        email = (profile.get("email") or '').strip().lower()

        try:
            sales_emails = json.loads(os.getenv("SALES_EMAILS", "{}"))
        except Exception:
            sales_emails = {}
        for name, mapped in sales_emails.items():
            if str(mapped).strip().lower() == email:
                return name

        return (profile.get("display_name")
                or profile.get("real_name")
                or '').strip()
    except Exception as exc:
        logger.warning(f"[SLACK] users_info 실패 ({user_id}): {exc}")
        return ''


def _slack_user_to_initial(client, user_id: str) -> str:
    """슬랙 user_id → 직원 이니셜 (한국 이름 거쳐서 매핑)"""
    if not user_id:
        return ''
    korean = _slack_user_to_korean_name(client, user_id)
    return _to_initial(korean)


# ─────────────────────────────────────────────────────────────
# 시간 표시 (`/청소 5분`, `/청소 1시간` 등에 사용)
# ─────────────────────────────────────────────────────────────
def _human_duration(seconds: int) -> str:
    """초 단위 → 한국어 표시 (`/청소`용 단일 단위)"""
    if seconds >= 86400:
        return f"{seconds // 86400}일"
    if seconds >= 3600:
        return f"{seconds // 3600}시간"
    if seconds >= 60:
        return f"{seconds // 60}분"
    return f"{seconds}초"
