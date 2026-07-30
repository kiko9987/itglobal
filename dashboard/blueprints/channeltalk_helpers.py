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

# ─────────────────────────────────────────────────────────────
# 릴레이/포워딩 잡스팸 신호 (재난문자·요금청구서·계좌번호 덤프)
#  — 마케팅 스팸과 별개 계열. 2026-07-30 L-03451(펌킨779) 케이스:
#    폭염 안전문자 + skylife 청구서 + 은행 9곳 계좌덤프 + [Web발신] 이
#    마케팅 키워드가 0점이라 기존 필터를 그대로 통과했음.
# ─────────────────────────────────────────────────────────────
# 대량/릴레이 발신 마커 — 정상 첫 문의엔 사실상 안 나옴
_RELAY_MARKERS = ('web발신', '국외발신', '국제발신', '무료수신거부', '수신거부')
# 요금청구/납부 클러스터
_BILLING_KEYWORDS = (
    '요금청구서', '납부하실', '미납요금', '지로납부', '납기일',
    '청구금액', '무통장 번호', '무통장번호', '전용무통장',
)
# 재난·안전 안내문자 클러스터 ('폭염' 단독은 정상 문의에도 흔해 제외, '폭염경보'만)
_DISASTER_KEYWORDS = (
    '온열질환', '재난문자', '안전안내문자', '행정안전부', '질병관리청',
    '기상특보', '한파경보', '호우경보', '대설경보', '폭염경보',
)
# 은행명 + 계좌번호 패턴 — 은행명 직후 하이픈/공백 섞인 8자+ 숫자열.
# ≥2개면 계좌 덤프(정상 문의가 은행 계좌를 여러 개 나열할 일 없음).
_BANK_ACCOUNT_RE = re.compile(
    r'(?:국민|농협|하나|우리|신한|기업|부산|수협|케이뱅크|카카오뱅크|토스|'
    r'씨티|SC제일|대구|광주|전북|경남|새마을|신협|우체국|산업)'
    r'\s*[:\-]?\s*\d[\d\-\s]{6,}\d'
)
# 필터 회피용 선두 자모 도배에서 제외할 비웃음·감정 낱자 (ㅋㅋ/ㅎㅎ/ㅠㅠ/ㅜㅜ/ㅡㅡ)
_EMOTIVE_JAMO = frozenset('ㅋㅎㅠㅜㅡ')


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


def _bank_account_dump_count(text: str) -> int:
    """은행명+계좌번호 패턴 개수. 정상 문의는 0~1, 스팸 계좌덤프는 여러 개."""
    if not text:
        return 0
    return len(_BANK_ACCOUNT_RE.findall(text))


def _has_leading_jamo_garbage(text: str) -> bool:
    """선두 15자(공백 제외) 안에 비웃음(ㅋㅎㅠㅜㅡ) 제외 단독 자모가 3개+ → 도배 선두.

    긴 본문 뒤에 붙은 자모 도배는 _has_jamo_flood(비율 25%)가 놓치므로 선두를 따로 포착.
    'ㅃㄴ ㅌㅌ폭염…' 은 잡고 'ㅋㅋㅋ 안녕하세요…' 는 통과(비웃음 낱자 제외).
    """
    if not text:
        return False
    count = 0
    checked = 0
    for c in text:
        if c.isspace():
            continue
        checked += 1
        if checked > 15:
            break
        # Hangul Compatibility Jamo (ㄱ~ㅣ) 단독 낱자 — 비웃음·감정 낱자는 제외
        if 'ㄱ' <= c <= 'ㅣ' and c not in _EMOTIVE_JAMO:
            count += 1
    return count >= 3


def _relay_junk_categories(text: str, text_lower: str) -> int:
    """릴레이 잡스팸 카테고리 존재 수 (발신마커/청구서/재난문자/계좌덤프)."""
    cats = 0
    if any(m in text_lower for m in _RELAY_MARKERS):
        cats += 1
    if any(k in text for k in _BILLING_KEYWORDS):
        cats += 1
    if any(k in text for k in _DISASTER_KEYWORDS):
        cats += 1
    if _bank_account_dump_count(text) >= 1:
        cats += 1
    return cats


def _is_spam_message(text: str) -> bool:
    """첫 메시지가 스팸 패턴인지 판정.

    판정 규칙:
    (1) 무의미 초단문 → 즉시 차단 (낱자/기호로만, 또는 5자 이하+기호로 시작)
    (2) 자모 도배(비율 25%↑) / 선두 자모 도배 → 즉시 차단 (필터 회피 chaos)
    (3) 릴레이/포워딩 잡스팸 (마케팅 키워드 0점 회피 계열) — 마케팅 점수와 독립:
        · 은행 계좌덤프 2개+ → 즉시 차단
        · 발신마커/청구서/재난문자/계좌덤프 중 2개 카테고리+ → 차단
    (4) 마케팅 스팸 — 가중치 점수 + URL 조합:
        · 짧은 메시지 (<50자): URL 1개+ + 점수≥5 → 차단
        · 긴 메시지 (≥50자): URL 2개+ + 점수≥3, 또는 URL 1개+ + 점수≥5 → 차단

    스팸 판정돼도 파괴적 동작 없음 — 시트 자동등록만 skip, 카드는 '스팸 의심'으로
    노출 → 오탐이어도 매니저 응답 시 _register_pending_lead_if_any 로 복구 가능.
    정상 메시지가 약한 키워드 우연 누적으로 잘못 차단되는 일 방지(LOW만으론 5점 X).
    """
    if not text:
        return False
    # (1)(2) 즉시 차단 시그널
    if _is_meaningless_short(text):
        return True
    if _has_jamo_flood(text):
        return True
    if _has_leading_jamo_garbage(text):
        return True

    text_lower = text.lower()

    # (3) 릴레이/포워딩 잡스팸 — 마케팅 키워드 점수와 독립 (재난문자·청구서·계좌덤프)
    if _bank_account_dump_count(text) >= 2:
        return True
    if _relay_junk_categories(text, text_lower) >= 2:
        return True

    # (4) 마케팅 스팸 — 가중치 + URL
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
