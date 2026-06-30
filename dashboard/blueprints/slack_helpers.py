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
# 시트 날짜 포맷
# ─────────────────────────────────────────────────────────────
def _format_date_for_sheet(iso_date: str) -> str:
    """ISO 형식("2026-06-25")을 시트에 텍스트로 저장하도록 escape.

    Google Sheets는 USER_ENTERED 모드에서 ISO 날짜를 시리얼 숫자로 자동 변환.
    작은따옴표 prefix는 Sheets의 텍스트 escape 문자 — UI/API에 표시되지 않음.
    출력: "2026-06-25" 그대로 (시트의 기존 방문 예정일 형식과 일관)
    """
    if not iso_date:
        return ''
    return f"'{iso_date}"


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
SALES_INITIALS = {
    '박용구': 'YG', '박정우': 'JW', '강성환': 'SH', '박민우': 'MW',
    '이근혁': 'GH', '김호중': 'HJ', '아이티': 'IT', '김단이': 'DN',
    '권태훈': 'TH', '주영민': 'YM', '심장원': 'SJW', '빈승정': 'SJ',
    '박민재': 'MJ', '조성헌': 'JSH', '황해승': 'HS', '강민석': 'MS',
    '강정권': 'JK', '이상덕': 'SD', '고광일': 'KiKO',
}


def _to_initial(name: str) -> str:
    """한국 이름 / 이니셜 / 빈값 → 이니셜 통일.
    - 한국 이름 → SALES_INITIALS 매핑
    - 영문 2~5자 → 대문자 ('KiKO' 예외)
    - 매핑 없으면 원본 그대로
    """
    if not name:
        return ''
    name = name.strip()
    if not name:
        return ''
    if name in SALES_INITIALS:
        return SALES_INITIALS[name]
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
