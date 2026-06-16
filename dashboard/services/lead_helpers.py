"""
리드 데이터 공용 정규화 헬퍼 (lead_sync, homepage_mail_sync 공용)

제공:
- normalize_phone(): 휴대폰·서울 02·지역번호·070 모두 처리
- extract_korean_address(): 한국 주소 4단계 정규식 추출
- clean_multiline(): 줄바꿈을 콤마로 통일
"""

import re
from typing import List, Optional, Tuple


# ─────────────────────────────────────────────────────────────
# 키워드 vocabulary (시트 K열 드롭다운과 일치)
# ─────────────────────────────────────────────────────────────
KEYWORD_VOCAB: List[str] = [
    # 기기 종류
    '천장형', '스탠드', '매립덕트', '벽걸이', 'FCU', '전열교환기', '가정용',
    # 추가 키워드 (상태/속성)
    '중고', '매입', '세척', '소상공인',
]


def extract_keywords(text: str) -> List[str]:
    """텍스트에서 KEYWORD_VOCAB 매칭. 등장 순서 무관, vocab 순서 유지.

    >>> extract_keywords('천장형, 스탠드 견적 부탁')
    ['천장형', '스탠드']
    >>> extract_keywords('중고 에어컨 매입 + 세척 문의')
    ['중고', '매입', '세척']
    >>> extract_keywords('일반 텍스트')
    []
    """
    if not text:
        return []
    found = []
    for kw in KEYWORD_VOCAB:
        if kw in text and kw not in found:
            found.append(kw)
    return found


def extract_keywords_from_sources(*texts: str) -> str:
    """여러 텍스트에서 키워드 추출, 콤마 join. 시트의 키워드 셀에 그대로 넣을 형태."""
    seen = set()
    result = []
    for t in texts:
        for kw in extract_keywords(t or ''):
            if kw not in seen:
                seen.add(kw)
                result.append(kw)
    return ', '.join(result)


# ─────────────────────────────────────────────────────────────
# 연락처 정규화
# ─────────────────────────────────────────────────────────────
_MOBILE_PREFIXES = ('010', '011', '016', '017', '018', '019')


def normalize_phone(raw) -> str:
    """
    한국 전화번호 정규화. 모든 케이스 처리.

    >>> normalize_phone('010-1234-5678')
    '010-1234-5678'
    >>> normalize_phone('01012345678')
    '010-1234-5678'
    >>> normalize_phone('025581105')   # 서울 9자리
    '02-558-1105'
    >>> normalize_phone('0212345678')  # 서울 10자리
    '02-1234-5678'
    >>> normalize_phone('0317771234')  # 지역 10자리
    '031-777-1234'
    >>> normalize_phone('07012345678') # 070 11자리
    '070-1234-5678'
    >>> normalize_phone('만료됨')
    ''
    """
    s = str(raw or '').strip()
    if not s or s in ('nan', 'NaN', 'None', '-', '입력되지 않음', '만료됨'):
        return ''

    digits = re.sub(r'\D', '', s)

    # 휴대폰 11자리 (010~019)
    if len(digits) == 11 and digits.startswith(_MOBILE_PREFIXES):
        return f'{digits[:3]}-{digits[3:7]}-{digits[7:]}'

    # 인터넷 전화 070 11자리
    if len(digits) == 11 and digits.startswith('070'):
        return f'{digits[:3]}-{digits[3:7]}-{digits[7:]}'

    # 서울 02
    if digits.startswith('02'):
        if len(digits) == 9:   # 02-XXX-XXXX (옛 국번)
            return f'{digits[:2]}-{digits[2:5]}-{digits[5:]}'
        if len(digits) == 10:  # 02-XXXX-XXXX (신규 국번)
            return f'{digits[:2]}-{digits[2:6]}-{digits[6:]}'

    # 지역번호 (031~064 + 080 등 3자리)
    if len(digits) == 10 and digits.startswith('0'):
        return f'{digits[:3]}-{digits[3:6]}-{digits[6:]}'
    if len(digits) == 11 and digits.startswith('0'):
        return f'{digits[:3]}-{digits[3:7]}-{digits[7:]}'

    return s


# ─────────────────────────────────────────────────────────────
# 한국 주소 추출 (4단계 패턴, 우선순위 순)
# ─────────────────────────────────────────────────────────────
_METRO = '서울|부산|대구|인천|광주|대전|울산|세종'
_PROV = '경기|강원|충북|충남|전북|전남|경북|경남|제주'

# 단지·건물·아파트 브랜드 키워드 (주소 확장용)
_BUILDING = (
    r'(?:아파트|빌딩|타워|오피스텔|상가|건물|단지|마을|클래스원|클래스|센터|'
    r'힐스테이트|자이|푸르지오|아이파크|래미안|롯데캐슬|이편한세상|위브|더샵|'
    r'센트럴파크|역|학교|학원|병원|교회|점포|공장|창고)'
)

# 종료 키워드 (이 단어 만나기 직전까지 주소로 확장)
_STOP_WORDS = [
    '설치', '공간', '면적', '평수', '평형', '신축', '상담', '견적',
    '문의', '연락', '예정입니다', '있습니다', '합니다', '드립니다',
    '부탁', '바랍니다', '필요합니다', '희망', '원합니다', '에어컨',
    '냉난방', '냉방', '시스템', '제품',
]

ADDRESS_PATTERNS = [
    # 1. 풀 주소: 광역시 + 시/군/구 + 동/로/길 + 번지 + (괄호) + 단지명 + 동/호/층
    (
        rf'((?:{_METRO})(?:특별시|광역시|특별자치시)?'
        r'\s*[가-힣]+(?:시|군|구)'
        r'\s+[가-힣\d]+(?:동|읍|면|로|길|번지|번길)'
        r'(?:\s*\d+(?:-\d+)?(?:번지|번길|호)?)*'
        r'(?:\s*\([가-힣\d\s/]+\))?'                # (불로동) 같은 괄호
        rf'(?:\s+[가-힣\dA-Za-z]{{1,30}}\s*{_BUILDING})?'  # 신검단중앙역우미린클래스원
        r'(?:\s+[가-힣\d]+(?:동|호|층|번지|번길|관|블록|블럭))?'  # 상가동 2층
        r'(?:\s+\d+(?:호|층|동|번지)?)?'  # 추가 숫자
        r')',
        'level1',
    ),
    # 2. 도 단위 풀 주소 (동일 확장)
    (
        rf'((?:{_PROV})(?:특별자치도|도)?'
        r'\s*[가-힣]+(?:시|군)'
        r'(?:\s+[가-힣]+(?:구|읍|면))?'
        r'\s+[가-힣\d]+(?:동|읍|면|로|길|번지|번길)'
        r'(?:\s*\d+(?:-\d+)?(?:번지|번길|호)?)*'
        r'(?:\s*\([가-힣\d\s/]+\))?'
        rf'(?:\s+[가-힣\dA-Za-z]{{1,30}}\s*{_BUILDING})?'
        r'(?:\s+[가-힣\d]+(?:동|호|층|번지|번길|관|블록|블럭))?'
        r'(?:\s+\d+(?:호|층|동|번지)?)?'
        r')',
        'level2',
    ),
    # 3. 약식: 구/시 + 동 + 번지/로 + 단지/동·호
    (
        r'([가-힣]{1,5}(?:구|시|군)'
        r'\s+[가-힣\d]+(?:동\d*|읍|면|로\s*\d+(?:번길)?|길)'
        r'(?:\s*\([가-힣\d\s/]+\))?'
        rf'(?:\s+[가-힣\dA-Za-z]{{1,30}}\s*{_BUILDING})?'
        r'(?:\s+[가-힣\d]+(?:동|호|층))?'
        r')',
        'level3',
    ),
    # 3b. 구/시 + 단지명만 (동/로 생략) — 예: "광진구 광장힐스테이트 아파트"
    (
        rf'([가-힣]{{1,5}}(?:구|시|군)\s+[가-힣\dA-Za-z]{{1,30}}\s*{_BUILDING}'
        r'(?:\s+[가-힣\d]+(?:동|호|층))?'
        r')',
        'level3b',
    ),
    # 4. 단축: 구/시 + 동 (+ 단지)
    (
        rf'([가-힣]{{1,5}}(?:구|시|군)\s+[가-힣]+동\d*'
        rf'(?:\s+[가-힣\dA-Za-z]{{1,30}}\s*{_BUILDING})?'
        r')',
        'level4',
    ),
    # 5. 광역 + 지역 + (구/시/군/동) — 명시적
    (
        rf'((?:{_METRO}|{_PROV})\s+[가-힣]+(?:구|시|군|동)'
        rf'(?:\s+[가-힣\dA-Za-z]{{1,30}}\s*{_BUILDING})?'
        r')',
        'level5',
    ),
    # 6. 시/도 + 짧은 지역명 + 위치 컨텍스트
    (
        rf'((?:{_METRO}|{_PROV})\s+[가-힣]{{2,5}})'
        r'(?=\s*(?:에\s*(?:위치|있는|자리|소재|사는|거주)|'
        r'\s+(?:사무실|학원|상가|건물|매장|점포|업소|병원|학교|회사|공장|창고|아파트|빌딩|타워|마을|단지|오피스)))',
        'level6',
    ),
    # level7 제거 — "○○ + 회사/공장/건물" 패턴이 일반 명사·회사명을 지역명으로 오인 (예: "세방정유라는 회사")
]


# 추출 주소 끝점에서 확장 가능한 토큰 패턴 (괄호/단지/동·호·층/숫자)
_EXTEND_TOKEN_RE = re.compile(
    rf'^\s*((?:\([가-힣\d\s/]+\)'                      # (불로동)
    rf'|[가-힣\dA-Za-z]{{1,30}}\s*{_BUILDING}'             # 신검단중앙역우미린클래스원
    rf'|{_BUILDING}'                                       # 단독 _BUILDING (예: " 아파트")
    r'|[가-힣\d]+(?:동|호|층|번지|번길|관|블록|블럭)'   # 상가동 / 2층
    r'|(?:동|호|층|번지|번길|관|블록|블럭)\d*'             # 단독 동/호/층 (예: "동", "층")
    r'|\d+(?:호|층|동|번지)?'                          # 숫자
    r')\s*)',
)


def _extend_address(text: str, base_addr: str, end_offset: int) -> str:
    """
    매칭된 base_addr 끝점에서 추가 토큰(괄호/단지명/동·호·층)을 더 확장.

    종료 조건:
    - 종료 키워드 (_STOP_WORDS) 만남
    - 마침표·줄바꿈·쉼표
    - 80자 초과
    - 더 이상 확장 가능한 토큰 없음
    """
    rest = text[end_offset:end_offset + 80]
    # 종료 키워드 위치
    stop_pos = len(rest)
    for sw in _STOP_WORDS:
        p = rest.find(sw)
        if 0 <= p < stop_pos:
            stop_pos = p
    # 마침표·줄바꿈 위치
    for ch in '.,\n。、':
        p = rest.find(ch)
        if 0 <= p < stop_pos:
            stop_pos = p
    rest = rest[:stop_pos]

    # 확장 가능한 토큰을 한 번에 하나씩 추가
    extended_tail = ''
    pos = 0
    while pos < len(rest):
        m = _EXTEND_TOKEN_RE.match(rest, pos)
        if not m:
            break
        extended_tail += m.group(0)
        pos = m.end()

    if extended_tail:
        # base와 extension 사이 공백은 extension의 leading space로 처리
        # (강제 추가하면 "상가" + " " + "동" = "상가 동" 되는 문제)
        return (base_addr + extended_tail).strip()
    return base_addr


def extract_korean_address(text: str) -> Optional[Tuple[str, str]]:
    """
    텍스트에서 한국 주소 추출. 7단계 정규식 + 종료 키워드 fallback 확장.

    Returns:
        (주소, 신뢰도_레벨) 튜플 또는 None
        레벨: 'level1' (최고) ~ 'level7' (추정)

    >>> extract_korean_address('양천구 신정3동')[0]
    '양천구 신정3동'
    >>> extract_korean_address('서울 강남에 위치한 사무실')[0]
    '서울 강남'
    >>> r = extract_korean_address('인천 서구 금정로 11 (불로동) 신검단중앙역우미린클래스원 상가동 2층 건물 중 2층에 설치 예정입니다.')
    >>> '신검단중앙역우미린클래스원' in r[0]
    True
    >>> extract_korean_address('일반 텍스트만 있음') is None
    True
    """
    if not text or not text.strip():
        return None

    for pattern, level in ADDRESS_PATTERNS:
        m = re.search(pattern, text)
        if m:
            addr = m.group(1)
            addr = re.sub(r'\s+', ' ', addr).strip()

            # 끝의 조사 정리 — 단, "상가"의 "가" 같은 명사 일부 제거 방지
            # 마지막 단어가 3자 이상일 때만 조사 제거 시도
            words = addr.split()
            if words:
                last_word = words[-1]
                if len(last_word) >= 3:
                    cleaned_last = re.sub(r'(에|에서|을|를|이|가|은|는)$', '', last_word)
                    if len(cleaned_last) >= 2:
                        words[-1] = cleaned_last
                        addr = ' '.join(words)

            # 옵션 B: 정규식 매칭 끝점에서 종료 키워드까지 추가 확장
            extended = _extend_address(text, addr, m.end())
            extended = re.sub(r'\s+', ' ', extended).strip()

            if len(extended) >= 4:
                return (extended, level)
    return None


# ─────────────────────────────────────────────────────────────
# 텍스트 정리
# ─────────────────────────────────────────────────────────────
def clean_multiline(text: str, sep: str = ', ') -> str:
    """
    텍스트의 연속 공백·줄바꿈을 콤마로 통일.

    >>> clean_multiline('천장형\\n스탠드')
    '천장형, 스탠드'
    >>> clean_multiline('A\\n\\nB\\n  C')
    'A, B, C'
    """
    if not text:
        return ''
    # 연속 공백/줄바꿈 → 단일 \n
    text = re.sub(r'[ \t]*\n[\s]*', '\n', text)
    # 각 줄 trim
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    return sep.join(lines)
