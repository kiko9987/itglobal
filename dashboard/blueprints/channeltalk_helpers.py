"""채널톡 봇 공통 유틸 — channeltalk.py에서 추출.

순수 유틸 함수만:
- 시간 포맷 (_format_ts, _format_ts_full)
- 스팸 감지 (_is_spam_message)
"""
import re
from datetime import datetime


# ─────────────────────────────────────────────────────────────
# 시간 포맷
# ─────────────────────────────────────────────────────────────
def _format_ts(created_ms: int) -> str:
    """epoch ms → '06.15 14:32' 양식"""
    try:
        dt = datetime.fromtimestamp(created_ms / 1000.0)
        return dt.strftime('%m.%d %H:%M')
    except Exception:
        return ''


def _format_ts_full(created_ms: int) -> str:
    """epoch ms → '2026.06.22. 09:32' (다른 인입 카드와 동일 양식)"""
    try:
        dt = datetime.fromtimestamp(created_ms / 1000.0)
        return dt.strftime('%Y.%m.%d. %H:%M')
    except Exception:
        return ''


# ─────────────────────────────────────────────────────────────
# 스팸 자동 감지 (채널톡 첫 메시지 인입 시)
# ─────────────────────────────────────────────────────────────
_SPAM_KEYWORDS = [
    # 채널업/마케팅 광고
    '무료체험', '무료 체험', '친구추가', '친구 추가', '채널친구', '채널 친구',
    '채널업', 'channelup', '사장님', '사업 번창', '사업번창', '리뷰', '마케팅',
    '카카오맵', '검색 중', '검색중에', 'sns 마케팅', '무료로 받을',
    '무료체험 신청', '체험 신청', '결제 정보 없이',
    # 광고 패턴
    '인사드립니다', '사업주', '대표님', '광고', '홍보', '바이럴', '구글 광고',
    '네이버 광고', '인스타 광고', '쇼핑몰 솔루션', '광고비', '제안 드립니다',
    '제안드립니다', '도움이 되시', '도움 되시', '효과 보장', '전화드릴',
    '연락드릴', '문자드릴', '카톡드릴', '안내드릴', '소개드릴',
    '저렴하게', '저렴한 가격', '특별 할인', '특가', '이벤트 진행',
    '대출', '저금리', '한도', '카드론', '신용', '코인',
    '상위노출', 'seo', '검색 노출', '키워드 노출',
    '컨설팅', '교육 받', '교육받', '플랫폼 운영',
    'dm 환영', 'dm 부탁', 'dm 주세요', '연락 부탁드립니다',
]
_URL_RE = re.compile(r'https?://\S+')


def _is_spam_message(text: str) -> bool:
    """첫 메시지가 스팸 마케팅 패턴인지 판정.
    - URL 2개 이상 + 마케팅 키워드 1개 이상
    - 또는 URL 1개 + 마케팅 키워드 2개 이상
    """
    if not text or len(text) < 50:
        return False
    text_lower = text.lower()
    url_count = len(_URL_RE.findall(text))
    keyword_count = sum(1 for kw in _SPAM_KEYWORDS if kw.lower() in text_lower)
    if url_count >= 2 and keyword_count >= 1:
        return True
    if url_count >= 1 and keyword_count >= 2:
        return True
    return False
