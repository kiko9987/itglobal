"""
카카오 로컬 API로 한국 주소 검증·정규화.

호출 순서 (verify_address):
1. 정규식 매칭 결과 (lead_helpers.extract_korean_address)
2. 원문 첫 줄 (메일 폼은 보통 첫 줄에 주소)
3. 본문에서 도로명+번지 패턴만 발췌
첫 매칭 성공한 결과를 사용. 모두 실패하면 None → 호출자가 fallback.

카카오 API 미설정 / 비활성 / 5초 타임아웃 → graceful 통과 (None 반환).
"""

import json
import os
import re
import urllib.error
import urllib.parse
import urllib.request
from functools import lru_cache
from typing import List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

KAKAO_ENDPOINT = 'https://dapi.kakao.com/v2/local/search/address.json'


def _kakao_key() -> str:
    return os.getenv('KAKAO_REST_API_KEY', '').strip()


# ─────────────────────────────────────────────────────────────
# 표시용 정규화 (시도/시 prefix 축약, 광주광역시 → "전남 광주")
# ─────────────────────────────────────────────────────────────
_METRO_KEEP = {'인천', '부산', '대구', '대전', '울산', '세종'}  # 그대로 유지
_PROV_SHORT = {
    '경기', '강원', '충북', '충남',
    '전북', '전남', '경북', '경남',
}
_PROV_FULL_TO_SHORT = {
    '경기도': '경기',
    '강원도': '강원', '강원특별자치도': '강원',
    '충청북도': '충북', '충청남도': '충남',
    '전라북도': '전북', '전북특별자치도': '전북', '전라남도': '전남',
    '경상북도': '경북', '경상남도': '경남',
}
_JEJU = {'제주', '제주도', '제주특별자치도'}
_SEOUL = {'서울', '서울시', '서울특별시'}


def normalize_display(addr: str) -> str:
    """
    주소를 시각적으로 깔끔하게 정규화 (시트·슬랙 표시용).

    규칙:
    - 서울 → 완전 제거 (구부터 표기)
    - 광주광역시 (광주 + ○○구) → "전남 광주" (전남광주특별시 출범 대비)
    - 경기 광주시 → "광주" (일반 도 규칙 적용)
    - 다른 광역시 (인천/부산/대구/대전/울산/세종) → 그대로
    - 제주 + 제주시 → "제주" 한 번만
    - 제주 + 서귀포시 → "서귀포"
    - 다른 도 + ○○시 → 도 prefix 제거 + "시" 제거
    - 다른 도 + ○○군 → 도 prefix 제거 + "군" 제거

    >>> normalize_display('서울 강남구 테헤란로 152')
    '강남구 테헤란로 152'
    >>> normalize_display('인천 연수구 갯벌로 36')
    '인천 연수구 갯벌로 36'
    >>> normalize_display('광주 동구 충장로 1')
    '전남 광주 동구 충장로 1'
    >>> normalize_display('경기 광주시 경안로 100')
    '광주 경안로 100'
    >>> normalize_display('경기 수원시 영통구 광교로 145')
    '수원 영통구 광교로 145'
    >>> normalize_display('제주 제주시 노형동 925')
    '제주 노형동 925'
    >>> normalize_display('제주 서귀포시 중문동')
    '서귀포 중문동'
    """
    if not addr:
        return ''
    tokens = addr.split()
    if len(tokens) < 2:
        return addr

    first = tokens[0]
    second = tokens[1]
    rest = tokens[2:]

    # 풀네임 도 → 약식 변환
    if first in _PROV_FULL_TO_SHORT:
        first = _PROV_FULL_TO_SHORT[first]
        tokens = [first, second] + rest

    # 1. 서울 → 완전 제거
    if first in _SEOUL:
        return ' '.join(tokens[1:])

    # 2. 광주광역시 (광주 + ○○구) → "전남 광주" prefix
    if first in ('광주', '광주광역시') and second.endswith('구'):
        return ' '.join(['전남', '광주'] + tokens[1:])

    # 3. 광역시 (그대로)
    if first in _METRO_KEEP:
        return addr
    # 풀네임 광역시 → 약식 (드물지만 fallback)
    if first.endswith('광역시'):
        short = first.replace('광역시', '')
        if short in _METRO_KEEP:
            return ' '.join([short] + tokens[1:])
    if first in ('인천시', '부산시', '대구시', '대전시', '울산시', '세종시'):
        return ' '.join([first[:-1]] + tokens[1:])

    # 4. 제주 — "제주 제주" 중복만 한 번으로 (그 외 케이스는 그대로)
    if first in ('제주도', '제주특별자치도'):
        # 풀네임 → 약식 (그 다음 토큰은 그대로 유지)
        if second == '제주시':
            return ' '.join(['제주'] + rest)
        return ' '.join(['제주'] + tokens[1:])
    if first == '제주' and second == '제주시':
        return ' '.join(['제주'] + rest)

    # 5. 일반 도 + ○○시/군 → prefix·접미 떼기
    if first in _PROV_SHORT:
        if second.endswith('시'):
            return ' '.join([second[:-1]] + rest)
        if second.endswith('군'):
            return ' '.join([second[:-1]] + rest)
        return ' '.join(tokens[1:])

    return addr


@lru_cache(maxsize=512)
def _kakao_search(query: str) -> Optional[dict]:
    """카카오 로컬 검색. lru_cache로 동일 쿼리 재호출 방지."""
    key = _kakao_key()
    if not key or not query.strip():
        return None
    try:
        url = KAKAO_ENDPOINT + '?' + urllib.parse.urlencode({'query': query.strip()})
        req = urllib.request.Request(url, headers={'Authorization': f'KakaoAK {key}'})
        with urllib.request.urlopen(req, timeout=5) as r:
            data = json.loads(r.read())
        docs = data.get('documents', [])
        return docs[0] if docs else None
    except urllib.error.HTTPError as exc:
        if exc.code in (401, 403):
            logger.warning(
                f'[KAKAO] 인증/권한 실패 (HTTP {exc.code}). '
                f'developers.kakao.com 콘솔에서 카카오 맵 서비스 활성화 필요.'
            )
        else:
            logger.debug(f'[KAKAO] HTTP {exc.code} on "{query[:40]}"')
        return None
    except Exception as exc:
        logger.debug(f'[KAKAO] {type(exc).__name__}: {query[:40]}')
        return None


# 도로명+번지 패턴 (예: "갯벌로 36", "테헤란로 152", "강남대로 401-1")
_ROAD_PATTERN = re.compile(r'([가-힣]+(?:로|길)\s+\d+(?:-\d+)?)')


# 건물명/층/호 정보 추출용
# 의미 있는 tail 신호:
# - 숫자 + 동/층/호 (예: "1층", "102호", "302동")
# - prefix + 시설 키워드 (예: "DMC빌딩", "인하대학교", "광장힐스테이트")
_TAIL_SIGNAL = re.compile(
    r'(?:'
    r'\d+\s*(?:동|호|층|관|블록|블럭|단지)'
    r'|[가-힣A-Za-z0-9]+\s*'
    r'(?:아파트|빌딩|타워|오피스텔|마을|클래스원|클래스|센터|'
    r'힐스테이트|자이|푸르지오|아이파크|래미안|롯데캐슬|이편한세상|위브|더샵|'
    r'센트럴파크|학교|학원|병원|교회|공장|창고|연구원|연수원|회관|호텔|모텔|'
    r'대학교|대학|아이클럽|상가동)'
    r')'
)
_TAIL_STOP_WORDS = [
    '신축', '상담', '견적', '문의', '연락', '에어컨', '설치', '예정',
    '냉방', '냉난방', '제품', '면적', '평수', '평형', '시스템',
    '예정입니다', '있습니다', '합니다', '드립니다', '부탁', '바랍니다',
    '필요합니다', '희망', '원합니다',
    # 모호 명사 (의미 없이 뒤따라붙는 단어들)
    '사무실', '카페', '매장', '점포', '집', '식당', '회사', '입니다',
]


def _extract_building_tail(text: str) -> str:
    """
    원본 텍스트에서 도로명/동 + 번지 뒤의 건물·동·층·호 정보 추출.

    예시:
        "인천광역시 연수구 송도로 갯벌로 36, 인하대학교 항공우주융합원 1층 102호\\n..."
        → "인하대학교 항공우주융합원 1층 102호"

        "서울특별시 마포구 상암동 1605번지 DMC 빌딩 5층"
        → "DMC 빌딩 5층"

        "인천 송도동 7-49 신축 건물입니다"
        → "" (종료 키워드 '신축'으로 잘림 + 의미 부족)

    카카오는 도로명+번지까지만 표준화하므로 이 정보를 따로 보존.
    """
    if not text:
        return ''
    first_line = text.strip().split('\n', 1)[0].strip()
    first_line = re.sub(r'\s+', ' ', first_line)

    candidates = []
    # 1. 도로명·길 + 번지 + 뒤 (예: "갯벌로 36, 인하대학교 ...")
    m = re.search(
        r'[가-힣]+(?:로|길)\s+\d+(?:-\d+)?(?:번지|번길)?\s*[,\s]+(.+)',
        first_line,
    )
    if m:
        candidates.append(m.group(1).strip())
    # 2. ○○동 + 번지 + 뒤 (예: "상암동 1605번지 DMC 빌딩 5층")
    m = re.search(
        r'[가-힣]+동\s*\d+(?:-\d+)?(?:번지)?\s*[,\s]+(.+)',
        first_line,
    )
    if m:
        candidates.append(m.group(1).strip())

    for tail in candidates:
        # 종료 키워드까지 자르기
        cut_pos = len(tail)
        for sw in _TAIL_STOP_WORDS:
            p = tail.find(sw)
            if 0 <= p < cut_pos:
                cut_pos = p
        tail = tail[:cut_pos].strip()
        # 트레일링 부호/공백 정리
        tail = re.sub(r'[,.\s]+$', '', tail).strip()

        # 의미 있는 건물·층·호 신호 있는지 검증
        if 2 <= len(tail) <= 60 and _TAIL_SIGNAL.search(tail):
            return tail
    return ''


def _build_candidates(text: str, regex_addr: Optional[str]) -> List[str]:
    """주소 후보 목록 (중복 제거, 우선순위 순)"""
    seen = set()
    out: List[str] = []

    def _push(s: str):
        s = (s or '').strip()
        # 트레일링 콤마/마침표 제거
        s = re.sub(r'[,.]+$', '', s).strip()
        if s and s not in seen:
            seen.add(s)
            out.append(s)

    if regex_addr:
        _push(regex_addr)

    if text:
        # 첫 줄
        first_line = text.strip().split('\n', 1)[0].strip()
        _push(first_line)

        # 첫 줄에 콤마 있으면 콤마 전 부분만 — "갯벌로 36, 인하대..." → "갯벌로 36"
        if ',' in first_line:
            _push(first_line.split(',', 1)[0])

        # 도로명+번지 추출 (본문 어디든)
        for m in _ROAD_PATTERN.finditer(text):
            _push(m.group(1))

    return out


def verify_address(
    text: str, regex_addr: Optional[str] = None
) -> Optional[Tuple[str, str]]:
    """
    카카오 API로 주소 검증.

    Args:
        text: 원본 본문 (문의 내용 전체)
        regex_addr: 정규식 1차 추출 결과 (있으면 최우선 시도)

    Returns:
        (정규화된 도로명/지번 주소, 'verified') or None
        None이면 호출자가 정규식 결과 또는 원문 첫 줄로 fallback.
    """
    if not _kakao_key():
        return None

    # 카카오 verified 주소에 붙일 건물·층·호 정보 (한 번만 추출)
    building_tail = _extract_building_tail(text)

    for cand in _build_candidates(text, regex_addr):
        doc = _kakao_search(cand)
        if not doc:
            continue
        # 도로명 우선, 없으면 지번
        road = doc.get('road_address')
        if road and road.get('address_name'):
            base = normalize_display(road['address_name'])
            full = f'{base} {building_tail}'.strip() if building_tail else base
            return (full, 'verified')
        jibun = doc.get('address')
        if jibun and jibun.get('address_name'):
            base = normalize_display(jibun['address_name'])
            full = f'{base} {building_tail}'.strip() if building_tail else base
            return (full, 'verified')

    return None


def resolve_address(
    text: str, regex_addr: Optional[str] = None, regex_level: str = ''
) -> Tuple[str, str]:
    """
    주소 확정 최종 함수. 호출 우선순위:

    1. 카카오 verified → ('도로명/지번 주소', 'verified')
    2. 정규식 결과 → (정규식 주소, 원래 level)
    3. 원문 첫 줄 (4~100자 + 한글 포함) → (첫 줄, 'raw')
    4. 다 실패 → ('', '')

    Returns:
        (주소, 신뢰도) 튜플. 신뢰도가 ''면 빈 결과.
    """
    # 1. 카카오 검증 시도
    verified = verify_address(text, regex_addr)
    if verified:
        return verified

    # 2. 정규식 결과 (시도 prefix 정규화 적용)
    if regex_addr:
        return (normalize_display(regex_addr), regex_level or 'regex')

    # 3. 원문 첫 줄 fallback (한글 포함 + 적절한 길이)
    if text:
        first_line = text.strip().split('\n', 1)[0].strip()
        first_line = re.sub(r'\s+', ' ', first_line)
        if (
            4 <= len(first_line) <= 100
            and re.search(r'[가-힣]', first_line)
        ):
            return (first_line, 'raw')

    return ('', '')
