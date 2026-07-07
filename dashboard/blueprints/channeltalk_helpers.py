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
# 스팸 자동 감지 — 가중치 기반 점수 시스템 (채널톡 첫 메시지 인입 시)
# ─────────────────────────────────────────────────────────────
# HIGH (5점) — 명백한 광고/스팸. 1개만 있어도 강한 시그널.
_SPAM_KEYWORDS_HIGH = [
    # 채널업 등 마케팅 서비스
    '채널업', 'channelup', '카카오맵',
    # SEO/광고 솔루션
    '상위노출', 'seo', '키워드 노출', '검색 노출',
    '구글 광고', '네이버 광고', '인스타 광고', '쇼핑몰 솔루션',
    '바이럴', '플랫폼 운영',
    # 금융 스팸
    '카드론', '저금리', '신용', '대출', '코인',
    # 광고 액션
    '무료체험 신청', '체험 신청', '결제 정보 없이',
    '저렴하게', '저렴한 가격', '특별 할인', '특가',
    'dm 환영', 'dm 부탁', 'dm 주세요',
]
# MID (3점) — 광고 표현. 다른 키워드와 조합 시 차단.
_SPAM_KEYWORDS_MID = [
    '무료체험', '무료 체험', '친구추가', '친구 추가', '채널친구', '채널 친구',
    '사업 번창', '사업번창', '사업주', '리뷰', '마케팅', '광고비',
    '제안 드립니다', '제안드립니다', '효과 보장', '컨설팅',
    '교육 받', '교육받', 'sns 마케팅', '무료로 받을',
    '광고', '홍보', '이벤트 진행',
]
# LOW (1점) — 정상 메시지에도 나타날 수 있는 약한 시그널. 누적 시만 의미.
_SPAM_KEYWORDS_LOW = [
    '사장님', '대표님', '인사드립니다',
    '도움이 되시', '도움 되시',
    '전화드릴', '연락드릴', '문자드릴', '카톡드릴', '안내드릴', '소개드릴',
    '검색 중', '검색중에', '한도',
    '연락 부탁드립니다',
]
_URL_RE = re.compile(r'https?://\S+')


def _spam_score(text_lower: str) -> int:
    """텍스트의 스팸 가중치 점수 합산"""
    score = sum(5 for kw in _SPAM_KEYWORDS_HIGH if kw.lower() in text_lower)
    score += sum(3 for kw in _SPAM_KEYWORDS_MID if kw.lower() in text_lower)
    score += sum(1 for kw in _SPAM_KEYWORDS_LOW if kw.lower() in text_lower)
    return score


def _has_jamo_flood(text: str) -> bool:
    """필터 회피용 자모 도배 감지.

    한글 낱자(ㄱ~ㅎ, ㅏ~ㅣ) 및 옛한글 아래아(ᆞ·ᆢ) 같은 단독 자모 문자가
    유의미 문자 대비 25% 이상이면 스팸으로 판정. 정상 한국어는 완성형 음절
    (가~힣) 사용, 자모 단독은 'ㅋㅋ', 'ㅎㅎ' 같은 짧은 리액션에만 등장.
    광고 메시지에서 필터 우회를 위해 "ㄴ L ᆞ ㅅ 로드" 처럼 자모/라틴 단독을
    섞어 넣는 패턴을 잡음.
    """
    if not text:
        return False
    # 공백/일반 구두점 제외한 유의미 문자
    meaningful = [
        c for c in text
        if not c.isspace()
        and c not in '.,-!?~()[]{}\"\'“”‘’…•'
    ]
    if len(meaningful) < 10:
        return False
    noise = 0
    for c in meaningful:
        # Hangul Compatibility Jamo (ㄱ~ㅎ, ㅏ~ㅣ) — 단독 자모
        if 'ㄱ' <= c <= 'ㆎ':
            noise += 1
        # 옛한글 낱자 (ᆞ, ᅠ 등)
        elif 'ᄀ' <= c <= 'ᇿ':
            noise += 1
        elif 'ꥠ' <= c <= '꥿':
            noise += 1
        # 아래아·유사 특수문자
        elif c in 'ㆍᆞᆢ':
            noise += 1
    return noise / len(meaningful) >= 0.25


def _is_meaningless_short(text: str) -> bool:
    """의미 없는 초단문 첫 메시지 감지 (봇 특유 패턴).

    - 낱자/기호로만 구성 (예: 'ㅡ', 'ㄴㄴ', '..')
    - 5자 이하 + 완성형 음절로 시작 안 함 (예: '.연승', '- 하이')
    정상 인입은 대부분 '안녕하세요' (5자+) 로 시작하므로 오탐 최소.
    """
    if not text:
        return False
    stripped = text.strip()
    if not stripped or len(stripped) > 8:
        return False
    syllables = sum(1 for c in stripped if '가' <= c <= '힣')
    if syllables >= 3:
        return False
    if syllables == 0:
        return True  # 낱자·기호만
    if len(stripped) <= 5 and not ('가' <= stripped[0] <= '힣'):
        return True  # 기호·낱자로 시작 + 5자 이하
    return False


def _is_spam_message(text: str) -> bool:
    """첫 메시지가 스팸 마케팅 패턴인지 판정 — 가중치 + 자모 도배 + 무의미 초단문.

    판정 규칙:
    - 무의미 초단문 → 즉시 차단 (낱자/기호로만, 또는 5자 이하+기호로 시작)
    - 자모 도배 → 즉시 차단 (필터 회피용 chaos 텍스트, 10자↑ 대상)
    - 짧은 메시지 (<50자): URL 1개+ + HIGH 키워드 1개+ (점수≥5) → 차단
      ('https://channelup.kr 무료체험' 같은 명백한 짧은 광고)
    - 긴 메시지 (≥50자):
      · URL 2개+ + 점수≥3 → 차단
      · URL 1개+ + 점수≥5 → 차단

    정상 메시지가 약한 키워드 (예: '사장님' + '도움 되시' = 2점) 우연 누적으로
    잘못 차단되는 일 방지 — LOW 키워드만으론 절대 5점 도달 X (총 14개 keyword).
    """
    if not text:
        return False
    # 무의미 초단문 (봇 첫 메시지 패턴) — 즉시 차단
    if _is_meaningless_short(text):
        return True
    # 자모 도배 (자모+라틴 단독 문자 25%↑) — 즉시 차단
    if _has_jamo_flood(text):
        return True

    text_lower = text.lower()
    url_count = len(_URL_RE.findall(text))
    score = _spam_score(text_lower)

    # 짧은 메시지 — URL + HIGH 키워드 명확한 광고만 차단
    if len(text) < 50:
        return url_count >= 1 and score >= 5

    # 긴 메시지 — URL과 점수 조합
    if url_count >= 2 and score >= 3:
        return True
    if url_count >= 1 and score >= 5:
        return True
    return False


# 하위 호환 — 기존 _SPAM_KEYWORDS 참조하던 코드용
_SPAM_KEYWORDS = _SPAM_KEYWORDS_HIGH + _SPAM_KEYWORDS_MID + _SPAM_KEYWORDS_LOW
