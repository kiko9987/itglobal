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
    # 서비스 카테고리 (거래처/기타 방문 요청 워크플로우)
    'A/S', '수금', '기타',
]


def format_inflow_display(platform: str) -> str:
    """플랫폼(유입 채널) → 카드 헤더/유입 구분 표기.

    - 거래처 → 거래처
    - 소개 → 거래처 (소개)
    - 기타 → 기타
    - 그 외 (홈페이지·카카오톡·채널톡·전화·당근·숨고·큐플레이스·모바일·앱·이메일 등) → 온라인 (X)
    빈값·'-' 은 '-' 반환.
    """
    v = (platform or '').strip()
    if not v or v == '-':
        return '-'
    if v == '거래처':
        return '거래처'
    if v == '소개':
        return '거래처 (소개)'
    if v == '기타':
        return '기타'
    return f'온라인 ({v})'


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

# 유효 한국 전화번호 자릿수 형태 (하이픈 없는 digits 기준).
# 앞자리 0 복원·정합성 감지 양쪽에서 단일 기준으로 사용.
_VALID_PHONE_RE = re.compile(
    r'^(?:'
    r'01[016789]\d{7,8}'                          # 휴대폰 010/011/016~019
    r'|02\d{7,8}'                                 # 서울 02
    r'|0(?:3[1-3]|4[1-4]|5[1-5]|6[1-4])\d{6,7}'   # 지역번호 031~064
    r'|050\d{7,9}'                                # 평생/안심번호 050
    r'|070\d{8}'                                  # 인터넷전화 070
    r'|080\d{6,8}'                                # 수신자부담 080
    r')$'
)


def is_valid_phone_digits(digits: str) -> bool:
    """하이픈 없는 숫자열이 유효한 한국 전화번호 형태인지."""
    return bool(_VALID_PHONE_RE.match(digits or ''))


def is_valid_phone(raw) -> bool:
    """원본 문자열(하이픈·공백 포함)이 유효 한국 전화번호인지. 빈값·'-' 은 False.

    정합성 체크에서 '값은 있는데 형태가 이상한' 연락처 감지용.
    앞자리 0 탈락·국제표기(+82)는 자동복원 후 판정하므로 True (복원 가능 = 문제 없음).
    """
    return is_valid_phone_digits(canonical_phone_digits(re.sub(r'\D', '', str(raw or ''))))


def restore_leading_zero(digits: str) -> str:
    """스프레드시트 숫자 서식으로 앞자리 0 이 탈락한 번호 복원.

    구글시트 셀이 숫자 서식이면 '01091501411' 입력이 1091501411(앞 0 탈락)로 저장됨.
    0 을 붙였을 때 **유효한 한국 전화번호 형태가 되는 경우에만** 복원 — 애매하면 원본 유지.
    (한국 전화번호는 모두 0 으로 시작하므로, 0 으로 시작 안 하는 것만 복원 대상.)

    >>> restore_leading_zero('1091501411')   # 휴대폰 앞 0 탈락
    '01091501411'
    >>> restore_leading_zero('212345678')    # 서울 02 앞 0 탈락
    '0212345678'
    >>> restore_leading_zero('01091501411')  # 이미 정상
    '01091501411'
    >>> restore_leading_zero('12345')        # 복원해도 유효 아님 → 원본
    '12345'
    """
    if digits and not digits.startswith('0') and is_valid_phone_digits('0' + digits):
        return '0' + digits
    return digits


def _strip_kr_country_code(digits: str) -> str:
    """국제표기 국가번호 82(+82 / 0082) 제거 → 국내 숫자열.

    +82 는 국내번호 앞 0 을 떼고 붙이는 형식(+82 10-9150-1411 = 010-9150-1411).
    82 제거 후 0 을 붙였을 때 유효 번호가 되는 경우에만 제거 (오탐 방지).
    국내번호는 82 로 시작하지 않으므로 안전. 0 복원은 restore_leading_zero 가 담당.
    """
    for prefix in ('0082', '82'):
        if digits.startswith(prefix):
            rest = digits[len(prefix):]
            cand = rest if rest.startswith('0') else '0' + rest
            if is_valid_phone_digits(cand):
                return rest
    return digits


def canonical_phone_digits(digits: str) -> str:
    """숫자열 → 표준 국내 숫자열: 국제표기(82) 제거 + 앞자리 0 복원.

    변환 결과가 유효 번호가 되는 경우만 적용, 아니면 원본 유지.
    유입·표시·매칭·정합성 판정의 단일 정규화 진입점.

    >>> canonical_phone_digits('821091501411')    # +82 휴대폰
    '01091501411'
    >>> canonical_phone_digits('1091501411')      # 앞 0 탈락
    '01091501411'
    >>> canonical_phone_digits('01091501411')     # 이미 정상
    '01091501411'
    """
    return restore_leading_zero(_strip_kr_country_code(digits))


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
    >>> normalize_phone('+82 10-9150-1411')  # 국제표기
    '010-9150-1411'
    >>> normalize_phone('만료됨')
    ''
    """
    s = str(raw or '').strip()
    if not s or s in ('nan', 'NaN', 'None', '-', '입력되지 않음', '만료됨'):
        return ''

    digits = re.sub(r'\D', '', s)
    digits = canonical_phone_digits(digits)   # 국제표기(+82) 제거 + 앞자리 0 탈락 복원

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
    r'센트럴파크|역|학교|학원|병원|교회|점포|공장|창고|'
    r'리조트|콘도|호텔|모텔|펜션|게스트하우스|레지던스|'
    r'플라자|몰|마트|시장|타운|파크|가든|스퀘어|허브|컴플렉스|'
    r'프라자|쇼핑몰|백화점|아울렛|마켓|상사|회관|문화회관|체육관)'
)

# 도로명+번지 붙여쓰기 정규화 (2026-08-14 L-03675/L-03678) — 매니저·고객이 도로명과
#   번지를 붙여 입력('동탄반석로172', '상도로 13길4')하면 추출 패턴이 번지를 못 떼내
#   실패·오분리. 세 규칙으로 canonical 스페이싱 복원 (표시·검증 정확도 향상):
_ROAD_GIL_JOIN_RE = re.compile(r'([가-힣]{2,}로)\s+(\d+번?길)(?=\d|\s|$)')      # 상도로 13길 → 상도로13길
_ROAD_BEONJI_SPLIT_RE = re.compile(r'([가-힣]{2,}대?로)(\d+(?:-\d+)?)(?![가-힣0-9번])')  # 동탄반석로172 → 동탄반석로 172
_GIL_BEONJI_SPLIT_RE = re.compile(r'(\d+번?길)(\d+(?:-\d+)?)(?![가-힣0-9])')     # 13길4 → 13길 4

# 붙여쓴 행정구역 접두 분리 (2026-08-19 L-03741) — '경기도고양시덕양구도내동' →
#   '경기도 고양시 덕양구 도내동'. 시/도(short·full) + 시? + 구/군? + 동/읍/면/리, 각
#   greedy·어간 1~4자('서구'·'우동'·'구로구' 모두 커버). 동 필수 + 2단계 이상 + 매치
#   구간 무공백일 때만 → 이미 띄어쓴 것·도로명·단일 동은 미변경. 붙여쓴 지번주소가
#   ADDRESS_PATTERNS(공백 요구)를 못 타 [추정]으로 빠지던 것 방지.
_ADMIN_PROV = (
    r'(?:서울특별시|부산광역시|대구광역시|인천광역시|광주광역시|대전광역시|울산광역시|'
    r'세종특별자치시|경기도|강원특별자치도|강원도|충청북도|충청남도|전라북도|전북특별자치도|'
    r'전라남도|경상북도|경상남도|제주특별자치도|서울|부산|대구|인천|광주|대전|울산|세종|'
    r'경기|강원|충북|충남|전북|전남|경북|경남|제주)'
)
_GLUED_ADMIN_RE = re.compile(
    r'(?<![가-힣])(' + _ADMIN_PROV + r')?'
    r'([가-힣]{2,4}시)?([가-힣]{1,4}[구군])?([가-힣]{1,4}[동읍면리])(?=\s|\d|$)'
)


# 숫자 사이 대시류 → 하이픈 (2026-08-25 L-03769) — 갤럭시 등에서 '-' 대신 'ㅡ'(U+3161
#   한글 모음)·en/em대시·全角하이픈·마이너스를 입력하면 번지(2ㅡ9)가 파싱 실패. 숫자
#   사이(공백 허용)에서만 치환 → 일반 텍스트 대시 오변환 방지.
_DASH_TO_HYPHEN_RE = re.compile(
    r'(?<=\d)\s*[ㅡ‐‑–—―−－]\s*(?=\d)'
)


def _unglue_admin_prefix(s: str) -> str:
    def _repl(m):
        seg = m.group(0)
        if not seg or ' ' in seg:
            return seg
        parts = [g for g in m.groups() if g]
        if len(parts) < 2:                    # 최소 2단계(동 포함) — 단일 동 미변경
            return seg
        return ' '.join(parts)
    return _GLUED_ADMIN_RE.sub(_repl, s)


def _normalize_road_spacing(s: str) -> str:
    """도로명+번지 붙여쓰기 → canonical 스페이싱. 도로세그먼트(N길/번길)는 보존.

    가드: '언주로107길'(로+숫자+길)은 도로명 일부라 분리 안 함(뒤 '번/한글/숫자'
    lookahead). '봉은사로 26길'→'봉은사로26길'(canonical 조인)은 결과 동일(무해).
    """
    if not s:
        return s
    s = _DASH_TO_HYPHEN_RE.sub('-', s)        # 숫자 사이 대시류(ㅡ 등) → '-' (L-03769)
    # 평수(면적) 이후는 상세/문의라 절단 (2026-08-25 L-03774): '…2층 25평 인테리어공사
    #   시작단계' → '…2층'. 공백+숫자+평(뒤 한글 아님) 부터 줄 끝까지 제거. '평택로'·
    #   '평화빌딩'(앞 공백+숫자 없음)은 미매치. 문의 내용은 별도 필드라 영향 없음.
    s = re.sub(r'\s\d+\s*평(?![가-힣]).*$', '', s)
    s = _unglue_admin_prefix(s)               # 붙여쓴 행정구역 분리 (L-03741)
    s = _ROAD_GIL_JOIN_RE.sub(r'\1\2', s)
    s = _ROAD_BEONJI_SPLIT_RE.sub(r'\1 \2', s)
    s = _GIL_BEONJI_SPLIT_RE.sub(r'\1 \2', s)
    return s


# 종료 키워드 (이 단어 만나기 직전까지 주소로 확장)
_STOP_WORDS = [
    '설치', '공간', '면적', '평수', '평형', '신축', '상담', '견적',
    '문의', '연락', '전화', '예정입니다', '있습니다', '합니다', '드립니다',
    '부탁', '바랍니다', '필요합니다', '희망', '원합니다', '에어컨',
    '냉난방', '냉방', '시스템', '제품',
]
# '전화' 추가 (2026-07-30): '구로구 고척동 전화연락 부탁드립니다' 에서 '연락'만
#   잘리고 남은 '전화' 가 상호 fallback 으로 주소에 흡수되던 leak (G5). '전화'
#   substring 이 '전화연락/전화주세요/전화번호' 모두 앞에서 차단.

ADDRESS_PATTERNS = [
    # 1. 풀 주소: 광역시 + 시/군/구 + 동/로/길 + 번지 + (괄호) + 단지명 + 동/호/층
    (
        rf'((?:{_METRO})(?:특별시|광역시|특별자치시)?'
        r'\s*[가-힣]+(?:시|군|구)'
        r'\s+[가-힣\d]+(?:동|읍|면|로|길|번지|번길)'
        r'(?:\s*\d+(?:-\d+)?(?:번지|번길|호)?)*'
        r'(?:\s*\([가-힣\d\s/]+\))?'                # (불로동) 같은 괄호
        rf'(?:\s+[가-힣\dA-Za-z]{{1,30}}?\s*{_BUILDING})?'  # 신검단중앙역우미린클래스원
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
        rf'(?:\s+[가-힣\dA-Za-z]{{1,30}}?\s*{_BUILDING})?'
        r'(?:\s+[가-힣\d]+(?:동|호|층|번지|번길|관|블록|블럭))?'
        r'(?:\s+\d+(?:호|층|동|번지)?)?'
        r')',
        'level2',
    ),
    # 3. 약식: 구/시 + 동 + 번지/로 + 단지/동·호
    #    (선택) {시} + {구|읍|면} 두 단계 행정구역 (예: "화성시 만세구")
    #    2026-07-13: 도로명 `로` 뒤 `지하NNN` 형식 지원 — 지하상가 주소
    #    (예: "강남구 학동로 지하102"). 기존은 `로\s*\d+` 만 매치해서 앞의
    #    "학동" 을 동 이름으로 오인식 → 도로명 놓침.
    (
        r'([가-힣]{1,5}(?:구|시|군)'
        r'(?:\s+[가-힣]{1,5}(?:구|읍|면))?'
        r'\s+[가-힣\d]+(?:동\d*|읍|면|로\s*(?:지하\s*)?\d+(?:번길)?|길)'
        r'(?:\s*\d+(?:-\d+)?(?:번지|번길|호)?)*'
        r'(?:\s*\([가-힣\d\s/]+\))?'
        rf'(?:\s+[가-힣\dA-Za-z]{{1,30}}?\s*{_BUILDING})?'
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
        rf'(?:\s+[가-힣\dA-Za-z]{{1,30}}?\s*{_BUILDING})?'
        r')',
        'level4',
    ),
    # 5. 광역 + 지역 + (구/시/군/동) — 명시적
    (
        rf'((?:{_METRO}|{_PROV})\s+[가-힣]+(?:구|시|군|동)'
        rf'(?:\s+[가-힣\dA-Za-z]{{1,30}}?\s*{_BUILDING})?'
        r')',
        'level5',
    ),
    # 6. 시/도 + 짧은 지역명 + 위치 컨텍스트
    # "인근/근처/주변/일대/부근/앞" 같은 위치 보조 키워드도 컨텍스트로 인정
    # 예: "서울 삼성역 인근 사무실"
    (
        rf'((?:{_METRO}|{_PROV})\s+[가-힣]{{2,5}})'
        r'(?=\s*(?:에\s*(?:위치|있는|자리|소재|사는|거주)|'
        r'\s+(?:사무실|학원|상가|건물|매장|점포|업소|병원|학교|회사|공장|창고|아파트|빌딩|타워|마을|단지|오피스)|'
        r'\s+(?:인근|근처|주변|일대|부근|앞)))',
        'level6',
    ),
    # 7. 행정구역 누락 지번 — "신사동648-23", "삼성동 123-45" (시·구 없이 동 + 번지)
    # 동 뒤에 번지번호가 강제로 와야 매칭 → "운동", "활동" 등 일반 단어 오인 차단
    (
        r'([가-힣]{2,}(?:동|리)\s*\d+(?:-\d+)?)',
        'level7',
    ),
]


# 추출 주소 끝점에서 확장 가능한 토큰 패턴 (괄호/단지/동·호·층/숫자)
# 주의: re.compile(...).match(s, pos)에서 pos > 0일 때 '^' 앵커는 매칭 실패하므로 사용 X
_EXTEND_TOKEN_RE = re.compile(
    rf'\s*((?:\([가-힣\d\s/]+\)'                       # (불로동)
    r'|-\d+(?:-\d+)?(?:호|층|동|번지)?'                # -105 / -105-1 / -105호 (지번 보충)
    rf'|[가-힣\dA-Za-z]{{1,30}}?\s*{_BUILDING}'             # 그로브리조트 / 신검단중앙역우미린클래스원 (lazy로 _BUILDING 우선 매칭)
    rf'|{_BUILDING}'                                       # 단독 _BUILDING (예: " 아파트")
    r'|[가-힣\d]+(?:동|호|층|번지|번길|블록|블럭)'      # 상가동 / 2층
    r'|\d+관'                                             # 2관 (건물 동 표기) — '관'은 숫자
    #   뒤에서만 유닛. '전파관리소·미술관·체육관' 등 단어 끝 '관'을 유닛으로 오인해
    #   '중앙전파관 리소'처럼 쪼개던 버그 방지 (L-03638). 한글+관 건물명은 상호 fallback 유지.
    r'|(?:동|호|층|번지|번길|블록|블럭)\d*'                # 단독 동/호/층 (예: "동", "층")
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
    # 줄바꿈·쉼표 위치 (마침표는 제외 — "20-16.  1층 일미리금계찜닭" 처럼
    # 매니저가 번지 뒤 마침표 찍고 상세정보 이어붙이는 케이스가 흔함.
    # 2026-07-13 L-03201 관측)
    for ch in ',\n。、':
        p = rest.find(ch)
        if 0 <= p < stop_pos:
            stop_pos = p
    rest = rest[:stop_pos]
    # 마침표를 공백으로 치환 — _EXTEND_TOKEN_RE 첫 그룹이 `\s*` 로 시작해서
    # 마침표를 skip 못 하므로 (2026-07-13 L-03201).
    rest = rest.replace('.', ' ')

    # 확장 가능한 토큰을 한 번에 하나씩 추가
    extended_tail = ''
    pos = 0
    while pos < len(rest):
        m = _EXTEND_TOKEN_RE.match(rest, pos)
        if not m:
            break
        extended_tail += m.group(0)
        pos = m.end()

    # 마지막 fallback — 남은 텍스트에 짧은 한글 명사(상호 후보) 1개만 추가.
    # 매니저는 슬랙 카드 상단 방문 주소만 보므로 상호도 함께 붙어야 방문지 파악 가능
    # (2026-07-13 L-03207 소각커피 / L-03201 일미리금계찜닭 관측).
    # 조건:
    #   - _STOP_WORDS 는 이미 rest 자체 잘라놨으므로 별도 필터 불필요
    #   - 최소 2자 (한글 시작) — 조사·오탈자 회귀 방지
    if pos < len(rest):
        m_shop = re.match(
            r'\s*([가-힣][가-힣A-Za-z0-9]{1,15})(?=\s|$|[.,])',
            rest[pos:],
        )
        if m_shop:
            extended_tail += ' ' + m_shop.group(1)

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

    # 도로명+번지 붙여쓰기 정규화 (L-03675/L-03678) — 매니저 워크플로 입력이 붙여쓴
    #   케이스 자동 복원. 대부분의 방문 주소는 매니저가 워크플로로 입력하므로 중요.
    text = _normalize_road_spacing(text)

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


ADDRESS_MISSING_LABEL = '주소 미입력'


def is_blank_address(addr) -> bool:
    """주소가 실질적으로 비었는지 판정.

    빈 문자열, '-', 빈 괄호 '()' / '( )', 구두점·괄호·공백만 남는 경우 True.
    고객이 홈페이지 폼 주소란을 안 채우면 resolver 가 '()' 를 남기는 케이스 대응.

    >>> is_blank_address('')
    True
    >>> is_blank_address('()')
    True
    >>> is_blank_address('-')
    True
    >>> is_blank_address('금천구 가산디지털2로 14')
    False
    """
    core = re.sub(r'[\s()（）\[\]{}.,·\-]', '', str(addr or ''))
    return not core
