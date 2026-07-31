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


# 로마숫자 ↔ 아라비아 등가 (2026-07-20) — dedup 판정 전용.
# 카카오 building_name 이 "원일테크노Ⅱ" 로 오는데 매니저는 "원일테크노2" 로 입력 →
# dedup 실패로 "원일테크노Ⅱ 원일테크노2" 중복 노출. 등가 판정으로 tail 쪽 아라비아 제거.
# 카카오 정식 building_name 이 로마숫자면 최종 결과도 로마숫자 유지 (공식 표기 우선).
# 12까지만 (건물명·차수에서 흔한 범위). Ⅹ = X 오인 방지 위해 명시 매핑.
_ROMAN_TO_ARABIC = {
    'Ⅰ': '1', 'Ⅱ': '2', 'Ⅲ': '3', 'Ⅳ': '4', 'Ⅴ': '5', 'Ⅵ': '6',
    'Ⅶ': '7', 'Ⅷ': '8', 'Ⅸ': '9', 'Ⅹ': '10', 'Ⅺ': '11', 'Ⅻ': '12',
    'ⅰ': '1', 'ⅱ': '2', 'ⅲ': '3', 'ⅳ': '4', 'ⅴ': '5', 'ⅵ': '6',
    'ⅶ': '7', 'ⅷ': '8', 'ⅸ': '9', 'ⅹ': '10', 'ⅺ': '11', 'ⅻ': '12',
}


def _roman_to_arabic(s: str) -> str:
    """건물명·차수 로마숫자를 아라비아로 정규화 (Ⅱ → 2)."""
    if not s:
        return s
    for k, v in _ROMAN_TO_ARABIC.items():
        if k in s:
            s = s.replace(k, v)
    return s


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
        return _post_normalize_display(addr)

    first = tokens[0]
    second = tokens[1]
    rest = tokens[2:]

    # 풀네임 도 → 약식 변환
    if first in _PROV_FULL_TO_SHORT:
        first = _PROV_FULL_TO_SHORT[first]
        tokens = [first, second] + rest

    # 아래 모든 분기 결과를 result 로 모아 마지막에 후처리 통과 (2026-07-22)
    if first in _SEOUL:
        result = ' '.join(tokens[1:])
    elif first in ('광주', '광주광역시') and second.endswith('구'):
        result = ' '.join(['전남', '광주'] + tokens[1:])
    elif first in _METRO_KEEP:
        result = addr
    elif first.endswith('광역시') and first.replace('광역시', '') in _METRO_KEEP:
        result = ' '.join([first.replace('광역시', '')] + tokens[1:])
    elif first in ('인천시', '부산시', '대구시', '대전시', '울산시', '세종시'):
        result = ' '.join([first[:-1]] + tokens[1:])
    elif first in ('제주도', '제주특별자치도'):
        if second == '제주시':
            result = ' '.join(['제주'] + rest)
        else:
            result = ' '.join(['제주'] + tokens[1:])
    elif first == '제주' and second == '제주시':
        result = ' '.join(['제주'] + rest)
    elif first in _PROV_SHORT:
        if second.endswith('시') or second.endswith('군'):
            result = ' '.join([second[:-1]] + rest)
        else:
            result = ' '.join(tokens[1:])
    elif (first.endswith('시') and len(first) >= 3
          and not first.endswith('광역시') and not first.endswith('특별시')
          and (second.endswith('구') or second.endswith('읍') or second.endswith('면'))):
        result = ' '.join([first[:-1]] + tokens[1:])
    else:
        result = addr

    return _post_normalize_display(result)


_SAME_SUFFIX_DEDUP = ('캠퍼스',)


def _post_normalize_display(addr: str) -> str:
    """normalize_display 이후 미세 표기 정정 — 우리 관행에 맞게 (2026-07-22).

    - 임야 지번 앞 '산' 접두어와 숫자 사이 공백 제거 ('산 57-22' → '산57-22')
      카카오 verify 결과는 '산 57' 로 띄어쓰나 국내 관행상 붙여 씀.
    - 인접 유사 단어 dedup ('판교제2테크노밸리 판교제2테크노벨리' 처럼 카카오
      building_name 과 원본 오탈자 tail 이 함께 붙는 케이스, ETC-b626fb 관측):
      길이 4자 이상 & 편집 유사도 ≥ 0.9 인 인접 토큰은 뒤 것 제거.
    - 같은 접미어(캠퍼스 등) 인접 토큰 dedup (ETC-68b96b '서울캠퍼스 안암캠퍼스'):
      카카오 표준(앞) 유지, 통칭(뒤) 제거.
    """
    if not addr:
        return addr
    # 2026-07-24 L-03367: 매니저가 주소에 콤마 계속 넣는 습관 → `콤마+공백` 만 공백으로 치환.
    #   `3,4,7호` (호수 콤마+공백 없음) 는 유지, `한라시그마밸리2차, 1층` 만 정리.
    # 2026-07-27 L-03401: 숫자 사이 콤마 (호수 나열 `305, 306, 408호`) 는 보존.
    #   `31, 서울숲` (번지-건물 구분) 만 제거. 숫자,(공백)숫자 는 placeholder 로 보호 후 복원.
    addr = re.sub(r'(?<=\d),(?=\s*\d)', '\x00', addr)  # 숫자 사이 콤마 보호
    addr = re.sub(r',\s+', ' ', addr)                   # 나머지 콤마+공백 제거
    addr = addr.replace('\x00', ',')                    # 보호분 복원
    # ' 산 (숫자)' → ' 산(숫자)'
    addr = re.sub(r' 산 (\d)', r' 산\1', addr)
    # 인접 유사 단어 dedup
    from difflib import SequenceMatcher
    tokens = addr.split()
    out = []
    for t in tokens:
        if out:
            prev = out[-1]
            # (1) 같은 접미어 dedup — 'X캠퍼스 Y캠퍼스' 시 카카오 표준(앞) 유지
            _same_suffix = next(
                (s for s in _SAME_SUFFIX_DEDUP
                 if prev.endswith(s) and t.endswith(s) and prev != t),
                None,
            )
            if _same_suffix:
                continue  # 뒤 것 skip (앞의 카카오 표준 유지)
            # (2) 유사도 dedup — 오탈자 tail
            if len(t) >= 4 and len(prev) >= 4:
                if SequenceMatcher(None, prev, t).ratio() >= 0.85:
                    out[-1] = t  # 뒤 것 (카카오 표준) 유지
                    continue
        # (3) 다토큰 near-dup — 정규화(공백·· 제거) 후 이미 나온 앞부분의 substring 이면
        #   제거 (2026-07-30 ETC-858578: 카카오 '온수 어르신복지회관 ·보훈회관' 뒤에
        #   고객원문 '온수어르신복지회관' 이 또 붙어 중복. 인접 비교로는 못 잡음).
        #   숫자 포함(번지·동·호·층) 토큰은 제외, 5자 이상만 (짧은 우연 매칭 방지).
        _tn = re.sub(r'[\s·・]', '', t)
        if len(_tn) >= 5 and not re.search(r'\d', t) and out:
            _pn = re.sub(r'[\s·・]', '', ''.join(out))
            if _tn in _pn:
                continue
        out.append(t)
    return ' '.join(out)


class _KakaoTransientError(Exception):
    """카카오 API 일시 실패(rate-limit/timeout/5xx). lru_cache 가 캐시하면 안 됨."""


# 재시도 대상 HTTP — 429 rate-limit(유입 몰릴 때 잦음), 5xx 일시 서버 오류.
_KAKAO_RETRY_STATUS = frozenset({429, 500, 502, 503, 504})


def _kakao_get_json(url: str):
    """카카오 GET → JSON dict. 일시 실패(429/5xx/timeout/network)는 최대 3회 backoff
    재시도, 소진 시 _KakaoTransientError raise (호출측 lru_cache 가 실패를 캐시 안 하도록
    → 순간 장애가 재시작까지 sticky 하며 멀쩡한 주소에 '확인 필요' 오배지 붙던 문제 방지,
    2026-07-31 L-03476). 인증(401/403)·파싱 오류는 None 반환(영구 상태라 캐시 무방)."""
    import http.client as _hc
    import socket
    import time as _t
    key = _kakao_key()
    if not key:
        return None
    for attempt in range(3):
        try:
            req = urllib.request.Request(
                url, headers={'Authorization': f'KakaoAK {key}'})
            with urllib.request.urlopen(req, timeout=5) as r:
                return json.loads(r.read())
        except urllib.error.HTTPError as exc:
            if exc.code in (401, 403):
                logger.warning(
                    f'[KAKAO] 인증/권한 실패 (HTTP {exc.code}). '
                    f'developers.kakao.com 콘솔에서 카카오 맵 서비스 활성화 필요.'
                )
                return None
            if exc.code in _KAKAO_RETRY_STATUS:
                if attempt < 2:
                    _t.sleep(0.4 * (attempt + 1))
                    continue
                raise _KakaoTransientError()  # 재시도 소진
            logger.debug(f'[KAKAO] HTTP {exc.code}')
            return None
        except (socket.timeout, TimeoutError, _hc.IncompleteRead,
                ConnectionError, OSError):
            if attempt < 2:
                _t.sleep(0.4 * (attempt + 1))
                continue
            raise _KakaoTransientError()
        except Exception as exc:
            logger.debug(f'[KAKAO] {type(exc).__name__}')
            return None
    raise _KakaoTransientError()


@lru_cache(maxsize=512)
def _kakao_search_cached(query: str) -> Optional[dict]:
    url = KAKAO_ENDPOINT + '?' + urllib.parse.urlencode({'query': query})
    data = _kakao_get_json(url)  # _KakaoTransientError 는 lru_cache 미캐시
    if data is None:
        return None
    docs = data.get('documents', [])
    return docs[0] if docs else None


def _kakao_search(query: str) -> Optional[dict]:
    """카카오 주소검색 (첫 결과 doc). 성공/유효-빈결과만 캐시(일시 실패는 재시도)."""
    q = (query or '').strip()
    if not q:
        return None
    try:
        return _kakao_search_cached(q)
    except _KakaoTransientError:
        return None


_KAKAO_POI_ENDPOINT = 'https://dapi.kakao.com/v2/local/search/keyword.json'


@lru_cache(maxsize=512)
def _kakao_search_poi_cached(query: str) -> tuple:
    url = _KAKAO_POI_ENDPOINT + '?' + urllib.parse.urlencode(
        {'query': query, 'size': 3})
    data = _kakao_get_json(url)  # _KakaoTransientError 는 lru_cache 미캐시
    if data is None:
        return ()
    docs = data.get('documents', []) or []
    return tuple(
        (d.get('place_name', '') or '', d.get('road_address_name', '') or '')
        for d in docs
    )


def _kakao_search_poi(query: str) -> tuple:
    """카카오 POI(키워드) 검색 — 상호명 → (place_name, road_address_name) 튜플.
    성공/유효-빈결과만 캐시(일시 실패는 재시도, 2026-07-31 L-03476)."""
    q = (query or '').strip()
    if not q:
        return ()
    try:
        return _kakao_search_poi_cached(q)
    except _KakaoTransientError:
        return ()


_NAVER_LOCAL_ENDPOINT = 'https://openapi.naver.com/v1/search/local.json'


@lru_cache(maxsize=512)
def _naver_search_local(query: str) -> tuple:
    """네이버 지역 검색 (2026-07-21 도입) — 카카오 POI 못 잡는 건물명 fallback.

    카카오 대비 강점: 지식산업센터·아파트·상용 건물명 커버리지 넓음.
    L-03316 사례: '한강듀클래스' → 카카오 미매칭, 네이버 '김포한강듀클래스' 정확 매칭.

    반환: ((place_name, road_address_name), ...) 카카오와 동일 인터페이스.
    """
    cid = os.getenv('NAVER_SEARCH_CLIENT_ID', '').strip()
    csec = os.getenv('NAVER_SEARCH_CLIENT_SECRET', '').strip()
    if not cid or not csec or not query.strip():
        return ()
    try:
        url = _NAVER_LOCAL_ENDPOINT + '?' + urllib.parse.urlencode(
            {'query': query.strip(), 'display': 5}
        )
        req = urllib.request.Request(url, headers={
            'X-Naver-Client-Id': cid,
            'X-Naver-Client-Secret': csec,
        })
        with urllib.request.urlopen(req, timeout=5) as r:
            data = json.loads(r.read())
        items = data.get('items', []) or []
        # title 에 <b> 태그 있음 (강조) → 제거
        return tuple(
            (
                re.sub(r'<[^>]+>', '', d.get('title', '') or ''),
                d.get('roadAddress', '') or '',
            )
            for d in items
        )
    except Exception as exc:
        logger.debug(f'[NAVER/POI] {type(exc).__name__}: {query[:40]}')
        return ()


def _search_poi(query: str) -> tuple:
    """POI 검색 통합 — 카카오 우선 + 네이버 병합 (2026-07-21).

    두 소스 결과 concat 반환 → _enrich_with_poi/_try_poi_fallback 의 매칭 루프가
    카카오 결과부터 순차 검증, 실패 시 네이버 결과로 자연 이동.
    """
    return _kakao_search_poi(query) + _naver_search_local(query)


# 도로명+번지 패턴 (예: "갯벌로 36", "테헤란로 152", "강남대로 401-1", "꽃내음1길 19-22",
# "봉은사로 26길 12", "부천로431번길 16")
# - 한글 2글자 이상으로 도로명 시작 (1글자 "번길" 같은 오인 차단)
# - 도로명 중간/끝 숫자 허용 (꽃내음1길, 시흥대로14길)
# - 2단 도로명 옵션 (○○로 N길 12, 봉은사로 26길 12)
# - 도로명 + 숫자번길 옵션 (부천로431번길 16) — "번" 옵션
_ROAD_PATTERN = re.compile(
    r'([가-힣]{2,}\d*(?:로|길)(?:\s*\d+번?(?:로|길))?\s+\d+(?:-\d+)?)'
)


# 건물명/층/호 정보 추출용
# 의미 있는 tail 신호:
# - 숫자 + 동/층/호 (예: "1층", "102호", "302동")
# - prefix + 시설 키워드 (예: "DMC빌딩", "인하대학교", "광장힐스테이트")
_TAIL_SIGNAL = re.compile(
    r'(?:'
    r'\d+\s*(?:동|호|층|관|블록|블럭|단지)'
    # 아파트 부속 시설 (단독으로 인정) — 2026-07-20 ETC-d656f2 관측: "관리사무소" 유실
    r'|관리사무소|경비실|정문|후문|어린이집|커뮤니티센터'
    r'|[가-힣A-Za-z0-9]+\s*'
    r'(?:아파트|빌딩|타워|오피스텔|마을|클래스원|클래스|센터|'
    r'힐스테이트|자이|푸르지오|아이파크|래미안|롯데캐슬|이편한세상|위브|더샵|'
    r'센트럴파크|학교|학원|병원|교회|공장|창고|연구원|연수원|회관|마을회관|호텔|모텔|'
    r'대학교|대학|아이클럽|상가동|사무실|사무소|'
    # 휴양·숙박·복합시설 (lead_helpers의 _BUILDING과 동기화)
    r'리조트|콘도|펜션|게스트하우스|레지던스|플라자|프라자|쇼핑몰|백화점|아울렛|'
    r'마트|시장|타운|파크|가든|스퀘어|허브|컴플렉스|문화회관|체육관|'
    # 한국식 시설/사업장 명사 (○○집/카페/식당 등 prefix 있어야)
    r'집|상회|공방|하우스|빌라|약국|미용실|매점|갤러리|한의원|식당|카페|'
    # 음식점·소매점 업종 키워드 (suffix)
    r'치킨|통닭|분식|곱창|닭갈비|국밥|냉면|고깃집|쌈밥|족발|보쌈|돈까스|초밥|횟집|'
    r'김밥|떡볶이|토스트|햄버거|피자|쌀국수|우동|라면|쭈꾸미|순대|덮밥|'
    r'정육점|베이커리|빵집|도넛|주점|호프|포차|편의점|슈퍼|문구|꽃집|세탁소)'
    # brand 시작 키워드 + 뒤 한글 (예: "김밥천국", "신촌도넛", "치킨마니아", "떡볶이나라")
    r'|(?:김밥|떡볶이|치킨|통닭|국밥|도넛|쌀국수|분식|족발|보쌈|곱창)[가-힣]{1,6}'
    r')'
)

# 한국 성씨 (흔한 것 70개) — tail 끝의 사람 이름 제거용
_KOREAN_SURNAMES = (
    r'김|이|박|최|정|강|조|윤|장|임|한|신|오|서|권|황|안|송|류|전|홍|'
    r'고|문|양|손|배|백|허|유|남|심|노|하|곽|성|차|주|우|구|민|진|지|'
    r'엄|채|천|방|공|함|변|염|여|추|도|소|석|선|설|마|길|연|위|표|명|'
    r'기|반|라|모|음|편|국'
)


# 상호·업종 접미어 — 사람 이름으로 오인돼 잘리면 안 됨 (2026-07-30 L-03475
#   '남양가 양꼬치' → '양꼬치'(양=성씨 오인) 잘려 '남양가'만 남던 사고).
_SHOP_NAME_SUFFIX = (
    '꼬치', '반점', '식당', '국밥', '곱창', '갈비', '분식', '통닭', '치킨',
    '피자', '숯불', '화로', '카페', '커피', '마트', '약국', '의원', '병원',
    '한의원', '학원', '정육', '세탁', '미용', '네일', '횟집', '뷔페', '베이커리',
    '노래방', '당구장', '문구', '철물', '설비', '공인', '부동산', '김밥', '냉면',
    '족발', '보쌈', '떡집', '방앗간', '포차', '주점', '호프', '실내', '완구',
)


def _strip_personal_name(tail: str) -> str:
    """tail 끝의 한국 사람 이름(성씨 + 1~2자, 총 2~3자)을 제거.

    예: '그로브리조트 정승종' → '그로브리조트'
        'ABC빌딩 김지수' → 'ABC빌딩'
    단, 상호·업종 접미어로 끝나면(양꼬치·홍반점 등) 사람 이름이 아니므로 유지.
    """
    if not tail:
        return tail
    m = re.search(rf'\s+((?:{_KOREAN_SURNAMES})[가-힣]{{1,2}})$', tail)
    if not m:
        return tail
    name = m.group(1)
    if any(name.endswith(s) for s in _SHOP_NAME_SUFFIX):
        return tail  # 상호·업종어 → 사람 이름 아님, 유지
    return tail[:m.start()].strip()
_TAIL_STOP_WORDS = [
    '신축', '상담', '견적', '문의', '연락', '전화', '에어컨', '설치', '예정',
    '냉방', '냉난방', '제품', '면적', '평수', '평형', '시스템',
    '예정입니다', '있습니다', '합니다', '드립니다', '부탁', '바랍니다',
    '필요합니다', '희망', '원합니다',
    # 모호 명사 (의미 없이 뒤따라붙는 단어들) — "집/카페/식당"은 빌딩명에 흔히 들어가므로 제외
    '사무실', '매장', '점포', '회사', '입니다', '관심',
]


def _flatten_paren_tail(tail: str) -> str:
    """괄호로 감싸진 tail을 flatten. 첫 요소가 순수 지번(동/가/리)이면 제거.

    괄호 앞/뒤 텍스트도 살려 이어붙이며, 결과의 콤마는 공백으로 정리한다.

    예:
        "(중계동, 건영아파트 유치원상가 1층 103호, 케이)"
        → "건영아파트 유치원상가 1층 103호 케이"

        "(서초동, 타임빌딩) B1, 위플레이스"      (괄호 뒤 텍스트 있는 케이스)
        → "타임빌딩 B1 위플레이스"

        "(가산동, 이앤씨드림타워7차)"
        → "이앤씨드림타워7차"

        "(걸포동)"                                (순수 지번만 → 빈 문자열)
        → ""

        "(건영아파트 유치원상가)"                 (첫 요소가 지번 아님)
        → "건영아파트 유치원상가"

        "건영아파트 유치원상가"                   (괄호 없음)
        → "건영아파트 유치원상가"

    이렇게 벗겨야 카카오 base와 자연스레 이어지고 매니저에게 층/호/상호 모두 노출된다.
    """
    if not tail:
        return tail

    m = re.search(r'\(([^)]*)\)', tail)
    if not m:
        # 괄호 자체가 없으면 원본 그대로
        return tail

    before = tail[:m.start()].strip()
    after = tail[m.end():].strip()
    parts = [p.strip() for p in m.group(1).split(',')]
    if parts and re.match(
        r'^[가-힣]+(?:동|가|리)(?:\s+\d+(?:-\d+)?)?$', parts[0]
    ):
        parts = parts[1:]
    inner_flat = ' '.join(p for p in parts if p).strip()

    result = ' '.join(x for x in (before, inner_flat, after) if x)
    # tail 안의 잔여 콤마는 공백으로 flatten (예: "B1, 위플레이스" → "B1 위플레이스")
    result = re.sub(r'\s*,\s*', ' ', result)
    result = re.sub(r'\s+', ' ', result).strip()
    return result


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
    # 첫 3줄 합쳐서 분석 (당근 주소 칸 멀티라인 입력 케이스 대응)
    # 예: "고양시덕양구 화신로298\n별빛8단지상가101호 코코헤어"
    lines = text.strip().split('\n')[:3]

    # 멀티라인 + 첫 줄에 회사명/시설명, 둘째 줄에 주소가 오는 거꾸로 입력 케이스
    # 예: "(주)창림아이티\n안산시단원구 신원로 28(신길동)1057-3"
    #     → facility_prefix = "(주)창림아이티"
    facility_prefix = ''
    if len(lines) >= 2:
        line0 = lines[0].strip()
        line1 = lines[1].strip()
        line0_has_addr = bool(re.search(
            r'(?:[가-힣]+동|[가-힣]+(?:로|길))\s*\d', line0
        ))
        line1_has_addr = bool(re.search(
            r'(?:[가-힣]+동|[가-힣]+(?:로|길))\s*\d', line1
        ))
        if not line0_has_addr and line1_has_addr and 2 <= len(line0) <= 40:
            # 회사명/시설명 표지 검사
            is_company = bool(re.search(
                r'\(주\)|㈜|\(유\)|주식회사|유한회사|영농조합법인', line0
            ))
            is_facility = bool(_TAIL_SIGNAL.search(line0))
            if is_company or is_facility:
                facility_prefix = line0

    first_line = ' '.join(line.strip() for line in lines if line.strip())
    first_line = re.sub(r'\s+', ' ', first_line)

    candidates = []
    # 1. 도로명·길 + 번지 + 뒤
    #    1단: "갯벌로 36, ...", "꽃내음1길 19-22, ...", "동호로28길11 ..."
    #    2단(공백): "봉은사로 26길 12 ..." (○○로 + 공백 + N길로 두 단으로 띄어쓴 형식)
    #    번길 합성: "부천로431번길 16 ..." (도로명 + 숫자 + 번길 한 단어)
    #    한글 2글자 이상 — 1글자 "번길" 같은 오인 차단
    m = re.search(
        r'[가-힣]{2,}\d*(?:로|길)(?:\s*\d+번?(?:로|길))?\s*\d+(?:-\d+)?(?:번지|번길)?\s*[,\s]+(.+)',
        first_line,
    )
    if m:
        candidates.append(m.group(1).strip())
    # 2. ○○동(○가)? + 번지 + 뒤 (예: "상암동 1605번지 DMC 빌딩 5층", "성수동1가 12-3 신촌도넛")
    m = re.search(
        r'[가-힣]+동(?:\s*\d+가)?\s*\d+(?:-\d+)?(?:번지)?\s*[,\s]+(.+)',
        first_line,
    )
    if m:
        candidates.append(m.group(1).strip())
    # 3. ○○동 + (번지 없이) 건물명 + 층/호 (2026-07-20 L-03292 관측)
    #    예: "중원구 상대원동 크란츠테크노 405호" — 상대원동 뒤 번지 없이 건물명·호수만.
    #    안전을 위해 뒤에 반드시 층/호/관 신호가 있는 케이스만 인정.
    m = re.search(
        r'[가-힣]+동\s+([가-힣A-Za-z0-9][가-힣A-Za-z0-9\s]*?\s*[A-Za-z]?\d+\s*(?:층|호|관|호실))',
        first_line,
    )
    if m:
        candidates.append(m.group(1).strip())

    for tail in candidates:
        # 리딩 구분자/부호 정리 (사용자가 ". ", "/ ", "- " 같은 구분자 쓴 케이스)
        # 예: "완정로 24 . 이성빌딩" → group(1) = ". 이성빌딩" → "이성빌딩"
        tail = re.sub(r'^[,.·\-/\s]+', '', tail).strip()
        # 종료 키워드까지 자르기 — 앞의 여는 괄호도 함께 자름
        # (2026-07-21 L-03317: "중국마사지샵(예정)" → '예정' 앞 '(' 만 남는 이슈)
        cut_pos = len(tail)
        for sw in _TAIL_STOP_WORDS:
            p = tail.find(sw)
            if 0 <= p < cut_pos:
                # 앞에 여는 괄호·공백 있으면 그것도 함께 자르기
                while p > 0 and tail[p-1] in '(（ ':
                    p -= 1
                cut_pos = p
        # 자유 구분자(/, |) 뒤의 추가 정보(평수/상세 등) 차단
        # 2026-07-20: '~' 는 호수 range (A201~205호) 로 흔히 쓰여서 제외.
        # 평수 range (20~30평) 는 '_TAIL_STOP_WORDS' 의 '평수/평형' 이 차단.
        for sep in ('/', '|'):
            p = tail.find(sep)
            if 0 <= p < cut_pos:
                cut_pos = p
        tail = tail[:cut_pos].strip()
        # 트레일링 부호/공백 정리
        tail = re.sub(r'[,.\s]+$', '', tail).strip()
        # 끝의 한국 사람 이름 제거 (예: "그로브리조트 정승종" → "그로브리조트")
        tail = _strip_personal_name(tail)
        # 단위 문자 대문자화 (2026-07-30 L-03475): 'a동/b호' 처럼 동·호 앞 라틴
        #   소문자를 대문자로 (그랑트윈타워a동 → 그랑트윈타워A동). 앞이 라틴이 아닌
        #   단독 1~2자만 (건물명 중간 소문자 e/kt 등은 미대상).
        tail = re.sub(
            r'(?<![A-Za-z])([a-z]{1,2})(?=동|호|층|관|블록|블럭)',
            lambda _m: _m.group(1).upper(), tail,
        )

        # 의미 있는 건물·층·호 신호 있는지 검증
        if 2 <= len(tail) <= 60 and _TAIL_SIGNAL.search(tail):
            # facility_prefix가 있으면 함께 부착 (예: "회사명 + 건물명")
            return f"{tail} {facility_prefix}".strip() if facility_prefix else tail

    # 주소 뒤 시설명 없지만 멀티라인 첫줄 회사명/시설명만 있는 경우
    return facility_prefix


def _build_candidates(text: str, regex_addr: Optional[str]) -> List[str]:
    """주소 후보 목록 (중복 제거, 우선순위 순).

    우선순위: 원본 첫 줄 → 정규식 결과 → 콤마 전 부분 → 도로명+번지 추출.
    원본 첫 줄이 짧고 깔끔한 주소면 그게 가장 정확 (당근 폼 등).
    매칭 실패하면 정규식 결과로 fallback (홈페이지 메일 본문 등).
    """
    seen = set()
    out: List[str] = []

    def _push(s: str):
        s = (s or '').strip()
        # 트레일링 콤마/마침표 제거
        s = re.sub(r'[,.]+$', '', s).strip()
        if s and s not in seen:
            seen.add(s)
            out.append(s)

    first_line = ''
    if text:
        first_line = text.strip().split('\n', 1)[0].strip()
        # 1) 행정구역 단위까지만 (도/광역시 + 시·군(?옵션) + 구/읍/면/동) — 카카오에 깔끔히 던짐
        # 본문 측정 단위(30평/2대)가 번지로 오인되는 것 방지
        # "경기도 화성 봉담읍" 처럼 "시"가 생략된 케이스도 잡기 위해 시/군은 옵션
        m_admin = re.search(
            r'((?:경기|강원|충북|충남|전북|전남|경북|경남|제주'
            r'|서울|부산|대구|인천|광주|대전|울산|세종)'
            r'(?:특별시|광역시|특별자치시|특별자치도|도)?'
            r'\s+[가-힣]+(?:시|군)?'
            r'\s+[가-힣]+(?:구|읍|면|동))',
            first_line,
        )
        if m_admin:
            _push(m_admin.group(1))

    # 도로명+번지 추출 — first_line 통째보다 우선 (2026-07-23 ETC-1ad649)
    # 카카오 verify 는 tail(상호명 등) 붙은 query 에서 번지를 잘못 근사하는 케이스 있음
    # 예: '장승배기로20길 46-3 송학대교회' → 46 리턴 (잘못됨) vs
    #     '장승배기로20길 46-3' → 46-3 리턴 (정확)
    # 순수 도로명+번지 후보를 우선 시도해 첫 성공이 정확하도록 함.
    if text:
        for m in _ROAD_PATTERN.finditer(text):
            _push(m.group(1))

    if first_line:
        _push(first_line)

    if regex_addr:
        _push(regex_addr)

    if first_line and ',' in first_line:
        _push(first_line.split(',', 1)[0])

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

    # 정규식 결과의 끝 "○○층/호" 분리 — 카카오가 "3층"의 "3"을 번지로 잘못 매칭하는 케이스 방지
    # 예: "인천 서구 가좌동 3층" → carry="3층", clean="인천 서구 가좌동"으로 검색
    floor_carry = ''
    clean_regex = regex_addr
    if regex_addr:
        m_floor = re.search(r'\s+(\d+(?:층|호|관))\s*$', regex_addr)
        if m_floor:
            floor_carry = m_floor.group(1)
            clean_regex = regex_addr[:m_floor.start()].strip()

    # 정규식 결과 끝 "상호명/시설명" 분리 — building_tail이 못 잡은 케이스의 fallback
    # 예: "본오동849번지 원치킨" → carry="원치킨", clean="본오동849번지"로 검색
    # building_tail은 도로명/동+번지+(.+) 패턴이 매칭돼야 작동 — regex_addr가 본문에서 잘려 들어오면
    # 패턴 매칭 실패하므로 여기서 보강
    _FACILITY_RE = re.compile(
        r'\s+([가-힣A-Za-z0-9]+'
        r'(?:치킨|통닭|분식|곱창|닭갈비|국밥|냉면|고깃집|쌈밥|족발|보쌈|돈까스|초밥|횟집|'
        r'김밥|떡볶이|토스트|햄버거|피자|쌀국수|우동|라면|쭈꾸미|순대|덮밥|'
        r'정육점|베이커리|빵집|도넛|주점|호프|포차|편의점|마트|슈퍼|문구|꽃집|세탁소|'
        r'카페|식당|약국|미용실|병원|한의원|학원|공방|상회|매점|갤러리|'
        r'헤어|미용|네일|에스테틱|살롱|파티|스튜디오|체육관|헬스|필라테스|요가'
        r'))\s*$',
    )
    facility_carry = ''
    if clean_regex and not building_tail:
        m_fac = _FACILITY_RE.search(clean_regex)
        if m_fac:
            facility_carry = m_fac.group(1)
            clean_regex = clean_regex[:m_fac.start()].strip()
    # 정규식 결과 없는 경우(박민혜: "내손중앙로51 헤어파티" → 정규식 None)에도
    # 원본 텍스트에서 시설명만 추출해 carry에 보관
    if not facility_carry and not building_tail and text:
        first_line = text.strip().split('\n', 1)[0]
        m_fac2 = _FACILITY_RE.search(' ' + first_line)
        if m_fac2:
            facility_carry = m_fac2.group(1)

    # 괄호로 감싸진 building_tail flatten — module-level 함수로 추출 (테스트 용이).
    building_tail = _flatten_paren_tail(building_tail)

    def _compose(base: str) -> str:
        parts = [base]
        # base에 도로명+번지가 없으면 원문에서 추출해 부착
        # (m_admin 후보로 검색되면 행정구역까지만 정규화돼 도로명/번지 손실)
        # 예: base="안양 동안구" + 원문 "관악대로 69" → "안양 동안구 관악대로 69"
        if not re.search(r'(?:로|길)\s+\d+', base):
            m_road = _ROAD_PATTERN.search(text or '')
            if m_road and m_road.group(1) not in base:
                parts.append(m_road.group(1))
        if building_tail:
            # 카카오 building_name이 이미 base에 부착되어 building_tail과 중복되는 경우 dedup
            # 예: base="...강남파이낸스센터" + tail="강남파이낸스센터" → skip
            #     base="...강남파이낸스센터" + tail="강남파이낸스센터 1층" → "1층"만 추가
            #     base="...SR프라자" + tail="sr프라자" → 대소문자 무시로 중복 판정
            # 2026-07-13 확장 (L-03190 관측): tail 첫 단어가 base 어디에든 있으면
            #   연속으로 제거해 중간 중복도 잡음. base 끝이 아닌 중간에 있어도 대응.
            # 2026-07-20 확장 (L-03278 관측): base 안 로마숫자와 tail 아라비아 숫자
            #   등가 판정 — 카카오 "원일테크노Ⅱ" vs 원본 "원일테크노2" 중복 방지.
            base_lower = _roman_to_arabic(base.lower())
            tail_lower = _roman_to_arabic(building_tail.lower())
            if tail_lower in base_lower:
                pass  # 완전 중복 (대소문자·로마숫자 무시) — skip
            else:
                tail_words = building_tail.split()
                # 각 tail 단어가 base 안에 있으면 (연속으로) 제거
                while (
                    tail_words
                    and _roman_to_arabic(tail_words[0].lower()) in base_lower
                ):
                    tail_words.pop(0)
                # 2026-07-27 L-03406: 카카오 building_name 과 원본 tail 이 유사 오탈자인
                #   케이스 (`유스페이스1` vs `유스페리스1`) — 정확 substring dedup 을
                #   못 통과해 `유스페이스1 유스페리스1` 중복. 유사도 높고 (>=0.7) 앞 2자
                #   같고 끝자리 숫자 동일하면 오탈자로 보고 tail 제거 (카카오 표준 유지).
                #   끝자리 숫자 다르면 (유스페이스1 vs 유스페이스2) 다른 건물이라 유지.
                if tail_words:
                    from difflib import SequenceMatcher as _SM
                    _base_words = base.split()
                    def _tnum(s):
                        _m = re.search(r'(\d+)$', s)
                        return _m.group(1) if _m else ''
                    _kept = []
                    for _tw in tail_words:
                        _dup = any(
                            len(_tw) >= 4 and len(_bw) >= 4 and _tw != _bw
                            and _tw[:2] == _bw[:2]
                            and _tnum(_tw) == _tnum(_bw)
                            and _SM(None, _tw, _bw).ratio() >= 0.7
                            for _bw in _base_words
                        )
                        if not _dup:
                            _kept.append(_tw)
                    tail_words = _kept
                remaining = ' '.join(tail_words).strip()
                if remaining:
                    parts.append(remaining)
        else:
            if facility_carry and facility_carry not in base:
                parts.append(facility_carry)
            if floor_carry and floor_carry not in base:
                parts.append(floor_carry)
        result = ' '.join(parts).strip()
        # 시각적 띄어쓰기 보장
        # 1. 한국 주소·시설 단어 다음 한글 — "단지상가" → "단지 상가"
        # 2026-07-20 예외: 다음이 도로명 suffix (대로/로/길/번길) 또는 부위 접미
        #   (동/층/호/관/번지) 면 skip — "인천타워대로" → 그대로, "상가동" → 그대로.
        result = re.sub(
            r'(단지|상가|아파트|빌딩|타워|오피스텔|맨션|빌라|하우스|클래스원)'
            r'(?=[가-힣])(?!대로|로|길|번길|동|층|호|관|번지)',
            r'\1 ', result,
        )
        # 2. 한글 다음 숫자+동/호/층/관 — "○○상가101호" → "○○상가 101호"
        # 영문 제외 (2026-07-23 ETC-678632): "B1층" (지하 1층) 이 "B 1층" 으로 잘못 분리되는 케이스 방지.
        # 영문+숫자+층/호 는 대개 원본에 이미 공백 있음 ("SR타워 3층", "K타워 9층").
        result = re.sub(r'(?<=[가-힣])(\d+(?:동|호|층|관))', r' \1', result)
        # 2-b. 층/호/관 다음 한글 (부가 설명·시설 tail) — "7층복도" → "7층 복도"
        result = re.sub(r'(\d+(?:동|호|층|관|호실))([가-힣])', r'\1 \2', result)
        # 2-c. 순수 지번 부기만 제거 — "(걸포동)", "(걸포동 172-1)"
        # 콤마 뒤에 부가정보가 있는 괄호는 보존:
        #   예 "(중계동, 건영아파트 유치원상가 1층 103호, 케이)" ← 층/호/상호 정보
        #   예 "(가산동, 이앤씨드림타워7차)" ← 건물명 정보
        # 회귀 방지 (7487390 이후): 매니저 방문 시 층/호 정보 손실 방지.
        result = re.sub(
            r'\s*\([가-힣]+(?:동|가|리)(?:\s+\d+(?:-\d+)?)?\)', '', result,
        )
        # 3. 연속 공백 정리
        result = re.sub(r' +', ' ', result).strip()
        return result

    def _try_kakao(cand_text: str):
        doc = _kakao_search(cand_text)
        if not doc:
            return None
        # 행정구역 단위 매칭(도로명·번지 없음)은 확정 금지 — 뒤 후보로 넘어가야 정답 도달
        # 예: "서울 노원구 중계동"(동만) → address_type=REGION → 여기서 확정하면
        #     실제 도로명 "중계로12가길 23"을 놓치고 "노원구 중계동"만 남음
        if doc.get('address_type') == 'REGION':
            return None
        road = doc.get('road_address')
        if road and road.get('address_name'):
            base = normalize_display(road['address_name'])
            building_name = (road.get('building_name') or '').strip()
            # 원본 tail에 건물명/시설명 신호가 있으면 카카오 building_name 스킵 (원본이 더 상세·정확).
            # 원본 tail이 층/호만 있으면 (건물명 없음) 카카오 building_name도 추가 (양쪽 정보 조합).
            #
            # 예: tail="건영아파트 유치원상가 1층 103호 케이" → "아파트"·"상가" 신호 → 스킵 (원본 우선)
            #     tail="지하 1층"                          → 건물 신호 없음 → "마천빌딩" 추가
            #     tail="104호"                            → 건물 신호 없음 → "노블시티프라자" 추가
            _tail_has_building = bool(
                building_tail
                and re.search(
                    r'아파트|빌딩|타워|상가|오피스텔|프라자|플라자|스퀘어|'
                    r'맨션|빌라|하우스|리조트|콘도|레지던스|'
                    r'대학|학교|병원|공장|센터|파크|가든|타운|허브|쇼핑몰|백화점|'
                    r'주식회사|㈜|\(주\)|영농조합|유한회사|'
                    # 브랜드·주상복합·랜드마크 (2026-07-13 L-03190/03194 관측)
                    r'아울렛|프라임|팰리스|시티|프리미어|캐슬|자이|푸르지오|힐스테이트|'
                    r'아이파크|더샵|롯데캐슬|이편한세상|위브|스카이|스테이션|'
                    # 영문 건물 키워드 (2026-07-13 L-03193 SD TOWER Ⅱ 중복 이슈)
                    r'TOWER|TWR|BUILDING|BLDG|PLAZA|MALL|CENTER|CENTRE|SQUARE|'
                    r'PARK|HOUSE|VILLA|OFFICE|APT|HOSPITAL|SCHOOL',
                    building_tail,
                    re.IGNORECASE,
                )
            )
            if (
                building_name
                and building_name.lower() not in base.lower()
                and not _tail_has_building
            ):
                # 2026-07-23 L-03329 관측: 같은 도로명 주소에 여러 건물 (KBS아레나·
                # KBS스포츠월드) 이 실제 존재하는 케이스. 카카오 verify 는 이 주소의
                # 대표 건물명 하나만 리턴 (스포츠월드) 하므로 원본 상호(아레나) 를
                # 대체해버림. → 원본 상호가 실제 POI 로 존재하면 원본 존중.
                _skip_bldg = False
                # base 의 도로명+번지 keyword (POI road 매칭용)
                _road_m = re.search(
                    r'([가-힣A-Za-z]+\d*(?:대로|로|길)\s*\d+(?:-\d+)?)', base
                )
                _road_key = _road_m.group(1) if _road_m else ''
                # 원본 상호 후보 리스트 — building_tail 우선, 실패 시 text 에서 도로명+번지 뒤 첫 단어
                _tail_candidates = []
                if building_tail:
                    _tail_candidates.append(building_tail.strip().split()[0])
                if text and _road_key:
                    _m_tail = re.search(
                        re.escape(_road_key) + r'\s+([^\s,·/\-|.]+)', text,
                    )
                    if _m_tail:
                        _tail_candidates.append(_m_tail.group(1).strip())
                # POI 검증 — 상호가 같은 도로명에 실제 존재하면 원본 유지
                _confirmed_tail = ''
                if _road_key:
                    for _tail_key in _tail_candidates:
                        if not _tail_key or len(_tail_key) < 2:
                            continue
                        # 카카오 building_name 과 같으면 skip 판정 불필요
                        if _tail_key.lower() == building_name.lower():
                            break
                        for _pn, _road in _search_poi(_tail_key):
                            if (_tail_key.lower() in (_pn or '').lower()
                                    and _road_key in (_road or '')):
                                _skip_bldg = True
                                _confirmed_tail = _tail_key
                                break
                        if _skip_bldg:
                            break
                if _skip_bldg and _confirmed_tail and _confirmed_tail.lower() not in base.lower():
                    # 원본 상호를 base 에 부착 (카카오 bldg 대신)
                    base = f"{base} {_confirmed_tail}"
                elif not _skip_bldg:
                    base = f"{base} {building_name}"
            return (_compose(base), 'verified')
        jibun = doc.get('address')
        if jibun and jibun.get('address_name'):
            return (_compose(normalize_display(jibun['address_name'])), 'verified')
        return None

    for cand in _build_candidates(text, clean_regex):
        result = _try_kakao(cand)
        if result:
            return result
        # 도로명 약식 보정 — "도신4길" → "도신로4길" 등 사용자 약식 입력 자동 정정
        # 패턴: 한글{2,}+ 숫자 + 길 → 한글 + "로" + 숫자 + 길
        fixed = re.sub(
            r'([가-힣]{2,})(\d+길)', r'\1로\2', cand,
        )
        if fixed != cand:
            result = _try_kakao(fixed)
            if result:
                return result

    return None


def _strip_redundant_legal_dong(addr: str) -> str:
    """도로명주소에 법정동이 중복으로 낀 경우 제거 (카카오 region_3depth 기준).

    홈페이지 주소검색기(Daum/카카오)가 `성수일로12길 52 (성수동2가, 건물명)` 형태로
    법정동을 딸려 보내면 resolver 가 인라인으로 남기던 이슈 (2026-07-27 L-03400).
    도로명주소엔 법정동이 불필요 → 카카오가 알려주는 정확한 법정동만 standalone 제거.

    - 도로명+번지가 있을 때만 동작 (지번주소는 법정동이 본질이라 건드리지 않음)
    - 카카오 region_3depth_name (예: '성수동2가') 과 정확히 일치하는 토큰만 제거
    - lookbehind/lookahead 로 건물명 안 부분매칭 방지 ('성수동 롯데캐슬' 의 성수동 보존)
    """
    if not addr:
        return addr
    m_road = re.search(r'[가-힣\d]+(?:로|길)\s*\d+(?:-\d+)?', addr)
    if not m_road:
        return addr  # 도로명+번지 없음 (지번주소 등) → skip
    doc = _kakao_search(m_road.group(0))  # 도로명+번지로 조회 (lru_cache)
    if not doc:
        return addr
    road = doc.get('road_address') or {}
    dong = (road.get('region_3depth_name') or '').strip()  # 예: 성수동2가
    if not dong or dong not in addr:
        return addr
    stripped = re.sub(
        rf'(?<![가-힣]){re.escape(dong)}(?![가-힣])\s*', '', addr,
    )
    stripped = re.sub(r'\s+', ' ', stripped).strip()
    return stripped or addr


def _enrich_verified_address(
    verified_addr: str, original_text: str, regex_addr: Optional[str]
) -> str:
    """카카오 verified 결과에 원본/정규식에 있는 누락 정보 보강.

    1. {도로명} N번길 M 패턴 — 카카오가 "N"까지만 인식하고 "번길 M" 누락 케이스
    2. verified에 도로명+번지가 없는데 원본/정규식엔 있는 케이스 (행정구역만 매칭)
    3. 층/호 tail 부착 — verified 에 층/호 없는데 원본에 있으면 뒤에 부착 (2026-07-20 L-03292)
    """
    if not verified_addr or not original_text:
        return verified_addr

    # 3. 층/호 tail 부착 (verified 에 층/호 없으면 원본에서 추출) — L-03292 관측
    #    지번동+건물명 (번지 없음) 케이스에서 카카오 verified 실패 후 regex fallback
    #    시 층/호 정보가 유실되는 것 방지. 원본에 "405호"·"A201~205호" 있으면 뒤에 부착.
    if not re.search(r'\d+\s*(?:층|호|호실|관)', verified_addr):
        # 뒤에 한글 오면 skip — '55 관리사무소' 의 '55 관' 을 별개 부속번호로 오인 방지 (2026-07-24 L-03362)
        m_floor = re.search(
            r'([A-Za-z]?\d+(?:~\d+)?\s*(?:층|호|호실|관))(?![가-힣])', original_text,
        )
        if m_floor:
            cand = m_floor.group(1).strip()
            if cand and cand not in verified_addr:
                verified_addr = f'{verified_addr} {cand}'

    # 4. "○○ ○○동" 형태 중복 제거 (2026-07-20 L-03294)
    #    카카오 building_name = "센트럴 레드" + 원본 tail = "레드동" → "센트럴 레드 레드동"
    #    정확 조건: 뒤 단어 == 앞 단어 + '동' 만 매치 (김포 vs 김포한강 같은 부분 매칭 회귀 방지)
    _words = verified_addr.split()
    _out = []
    for w in _words:
        if _out and w == _out[-1] + '동':
            _out[-1] = w  # 짧은 쪽 (레드) → 상세한 쪽 (레드동) 로 승격
            continue
        _out.append(w)
    verified_addr = ' '.join(_out)

    # 1. {도로명} N번길 M 보강 (송미나 케이스)
    m = re.search(
        r'([가-힣]+(?:로|길))\s*(\d+)\s*번길[\s,]*(\d+(?:-\d+)?)',
        original_text,
    )
    if m and '번길' not in verified_addr:
        road, num, ext = m.group(1), m.group(2), m.group(3)
        target = f"{road} {num}"
        target_alt = f"{road}{num}"
        replacement = f"{road} {num}번길 {ext}"
        if target in verified_addr:
            verified_addr = verified_addr.replace(target, replacement)
        elif target_alt in verified_addr:
            verified_addr = verified_addr.replace(target_alt, replacement)

    # 1-b. {호|층|-번지} 다음 단어 = 상호명 자동 부착
    #   예: "105호 베베드피노" / "1층 일미리금계찜닭" / "205-5 소각커피"
    #   2026-07-13 L-03201·L-03207 확장 — 호/층/번지 없어도 마지막 상호 후보 잡음.
    last_line = original_text.split('\n')[-1].strip().rstrip('.')
    m_shop = re.search(
        r'(?:\d+(?:-\d+)?(?:호|층|번지)?)\s+([가-힣][가-힣A-Za-z0-9]{1,15})\s*$',
        last_line,
    )
    if m_shop:
        shop = m_shop.group(1)
        # 도로명·행정구역 접미사·층/호/번지 로 끝나면 상호 아님
        # (2026-07-13 L-03222 관측: '수은빌딩3층' 이 상호로 오인식되어 중복 부착)
        if not re.search(r'(?:로|길|구|시|군|동|읍|면|층|호|번지)$', shop) \
                and shop not in verified_addr \
                and not re.search(rf'\b{re.escape(shop)}\b', verified_addr):
            verified_addr = f"{verified_addr} {shop}".strip()

    # 도로명 끝 + 숫자 사이 공백 보강 ("세월길2" → "세월길 2"). 단 "12길" 같은 도로명 일부는 제외
    # 2026-07-24 L-03374 fix: lookahead 에 `[가-힣]+길` 추가 — `번안길`, `안길` 등
    #   한글 접두어가 있는 도로명 suffix 도 분리 대상에서 제외. 이전엔 `번안길`
    #   같은 케이스가 `성현로 135번안길` 로 잘려 `_road_key` 파싱 오차 발생.
    verified_addr = re.sub(
        r'([가-힣]+(?:로|길))(\d+)(?![0-9]|[가번]?길|[가-힣]+길)',
        r'\1 \2', verified_addr,
    )

    # 2. verified에 도로명+번지 없으면 원본/정규식에서 추출 부착 (황경철 케이스)
    if not re.search(r'(?:로|길)\s*\d', verified_addr):
        # 정규식 결과 우선 검사
        candidates = []
        if regex_addr:
            m_r = re.search(r'([가-힣]+(?:로|길)\s*\d+(?:-\d+)?)', regex_addr)
            if m_r:
                candidates.append(m_r.group(1))
        m_t = re.search(r'([가-힣]+(?:로|길)\s*\d+(?:-\d+)?)', original_text)
        if m_t:
            candidates.append(m_t.group(1))
        for cand in candidates:
            if cand not in verified_addr:
                verified_addr = f"{verified_addr} {cand}".strip()
                verified_addr = re.sub(r'\s+', ' ', verified_addr)
                break

    # 최종 — 도로명 + 숫자 사이 공백 보강 (부착 후에도 적용)
    # 2026-07-24 L-03374 fix: lookahead 에 `[가-힣]+길` 추가 — `번안길`, `안길` 등
    #   한글 접두어가 있는 도로명 suffix 도 분리 대상에서 제외. 이전엔 `번안길`
    #   같은 케이스가 `성현로 135번안길` 로 잘려 `_road_key` 파싱 오차 발생.
    verified_addr = re.sub(
        r'([가-힣]+(?:로|길))(\d+)(?![0-9]|[가번]?길|[가-힣]+길)',
        r'\1 \2', verified_addr,
    )

    # 3. 카카오 POI 검색으로 상호 지점명 부착 (2026-07-13).
    #   원문에 상호 후보가 있고, POI 결과의 도로명이 verified 와 일치하면
    #   지점명(place_name) 을 verified 뒤 부기 → 매니저가 지점 정확 파악.
    #   예: "마성떡볶이" + "학동로 지하 102" → "마성떡볶이 논현역점" 부기
    verified_addr = _enrich_with_poi(verified_addr, original_text)

    return verified_addr


# _STOP_WORDS 를 lead_helpers 재사용 (import 순환 방지 — 지연 import)
# 2026-07-24 L-03374 fix: '호|층' 제거 — '지호창호' 같이 상호가 '호' 로 끝나는 케이스가
#   후보에서 제외되던 버그. 아파트 부속 표기 (101호, 3층) 은 한글 시작 필터
#   ([가-힣][가-힣A-Za-z0-9]{1,14}) 로 이미 배제되므로 여기 유지 불필요.
_ADMIN_SUFFIX_RE = re.compile(r'(로|길|구|시|군|동|읍|면|리|번지|가|동로|번길)$')


def _extract_region_hint(verified_addr: str) -> str:
    """verified 주소에서 지역 힌트 추출 (첫 시/군/구/광역시)."""
    for w in verified_addr.split():
        if re.search(r'(?:시|군|구|도)$', w):
            return w
    return ''


def _road_key(addr: str) -> str:
    """주소에서 도로명+번지 정규화 키 추출 (`학동로 지하 102`, `지산2길 20-16`)."""
    m = re.search(
        r'([가-힣\d]+(?:로|길)\s+(?:지하\s*)?\d+(?:-\d+)?)',
        addr,
    )
    return m.group(1).replace(' ', '') if m else ''


def _extract_shop_candidates(text: str) -> list:
    """원문에서 상호 후보(짧은 한글 명사) 뽑기 — POI 검색용.

    - 한글 2-15자
    - 행정구역·도로명 접미사(로/길/구/시/군/동/읍/면/번지/호/층/리/가) 로 끝나는 단어 제외
    - _STOP_WORDS 포함하면 제외
    - 중복 제거
    """
    try:
        from dashboard.services.lead_helpers import _STOP_WORDS
    except Exception:
        _STOP_WORDS = []
    seen, out = set(), []
    for w in re.findall(r'[가-힣][가-힣A-Za-z0-9]{1,14}', text):
        if w in seen:
            continue
        if _ADMIN_SUFFIX_RE.search(w):
            continue
        if any(sw in w for sw in _STOP_WORDS):
            continue
        # 2026-07-25 ETC-841a3c: 층 표기 접두어 (`지하1층`, `옥탑2층` 등) exclude.
        #   `호|층` 을 _ADMIN_SUFFIX_RE 에서 뺀 후 (지호창호 fix) 층 표기가 shop
        #   후보로 잘못 뽑히던 이슈. POI 검색 시 우연히 매칭되는 다른 건물 부착
        #   위험 방지 (예: `하나은행365 엘타워빌딩 지하1층` 오부착).
        if re.match(r'^(지하|옥탑|옥상|지상)', w):
            continue
        seen.add(w)
        out.append(w)
    return out


def _joined_shop_candidates(text: str) -> list:
    """인접 상호 단어 2개를 결합한 후보 리스트 (2026-07-30 G1, L-03475).

    개별 단어가 행정동 접미('가' 등)로 _extract_shop_candidates 에서 제외돼도
    ('남양가' + '양꼬치'), 결합하면 유효 상호('남양가양꼬치')일 수 있어 POI
    place_name 채택용으로 별도 생성. 주소 토큰(도로/번지/행정구역/단위 동·호·층)은
    결합 대상에서 제외. 반환: [(joined_nospace, 'a b'(원본 공백형)), ...]
    """
    words = re.findall(r'[가-힣][가-힣A-Za-z0-9]{1,14}', text)

    def _is_addr_token(w: str) -> bool:
        if re.search(r'\d', w):                          # 번지·동번호 등 숫자 포함
            return True
        if re.search(r'(?:로|길|구|시|군|읍|면)$', w):      # 도로·행정구역 접미
            return True
        if re.search(r'(?:[A-Za-z]동|번지|번길|호|층)$', w):  # 단위 a동/번지/호/층
            return True
        return False

    out = []
    for i in range(len(words) - 1):
        a, b = words[i], words[i + 1]
        if _is_addr_token(a) or _is_addr_token(b):
            continue
        joined = a + b
        if len(joined) >= 4:
            out.append((joined, f'{a} {b}'))
    return out


# POI place_name 후행 단어가 부속시설이면 상호 지점 아님 → 부착 skip.
_POI_FACILITY_EXACT = frozenset({'지하', '옥상', '별관', '분관', '사무실'})
_POI_FACILITY_SUBSTR = (
    '주차장', '화장실', '엘리베이터', '승강기', '경비실', '관리실',
    '관리사무소', '관리소', '관리단', '기계실', '전기실', '방재실',
    '충전소', '정문', '후문', '출입구',
    # 2026-07-30 근본 수정: 카카오 POI 가 시설 POI 를 먼저 반환하면 junk 부착됨.
    #   'ATM'(현금인출기)·'옥외'(앞_옥외 등 옥외 주차/공간).
    'atm', '옥외',
)


def _poi_has_facility(place_name: str) -> bool:
    """place_name 의 상호(첫 단어) 이후 **모든** 단어에 부속시설어가 있으면 True.

    2026-07-30 근본 수정: 기존엔 두 번째 단어(place_words[1])만 검사해서
    '롯데마트 고양점 주차장'(주차장=3번째)·'한국화훼농협 ATM 본점' 같은 시설 POI 가
    통과해 junk 부착 (신규 등록에도 발생하던 버그, backfill 아님). 상호 뒤 전 단어 검사.
    """
    words = place_name.split()
    for w in words[1:]:
        wl = w.lower()
        if w in _POI_FACILITY_EXACT:
            return True
        if any(f in wl for f in _POI_FACILITY_SUBSTR):
            return True
    return False


def _replace_compound_with_tenant(verified_addr, original_text, candidates, v_key, region):
    """복합 등록명(A ·B) 건물 → 고객이 지목한 시설의 POI 정식명으로 재구성.

    2026-07-30 ETC-45cab9: 카카오 주소 API 건물 등록명이 '온수 어르신복지회관 ·보훈회관'
    (한 건물 내 별개 시설 병기)인데 고객은 '온수어르신복지회관'(테넌트)만 방문. 보훈회관은
    층·구역이 다른 별개 시설이라 방문 주소 노이즈. 고객이 지목한 테넌트가 POI place_name
    ('온수어르신복지관', 도로 일치)과 유사하면 도로+번지 + POI 정식명 + 층/호 로 재구성.
    반환: 재구성된 주소 or None(적용 안 함 — 안전 fallback).
    """
    from difflib import SequenceMatcher
    for cand in candidates:
        if len(cand) < 4:
            continue
        for P, road in _search_poi(f'{cand} {region}'.strip()):
            if not P or not road or _road_key(road) != v_key or _poi_has_facility(P):
                continue
            _pn, _cn = P.replace(' ', ''), cand.replace(' ', '')
            # 같은 시설 판정 — 유사도 0.8+ or 5자+ 공통 접두 (복지관↔복지회관 사소 차이 허용)
            same = (SequenceMatcher(None, _cn, _pn).ratio() >= 0.8
                    or (len(_cn) >= 5 and _pn.startswith(_cn[:5])))
            if not same:
                continue
            # 도로+번지 prefix 추출 (로/길 + 번지). 없으면 재구성 포기(안전).
            m = re.match(
                r'(.*?[가-힣\dA-Za-z]+(?:로|길)\s+(?:지하\s*)?\d+(?:-\d+)?)', verified_addr)
            if not m:
                return None
            prefix = m.group(1).strip()
            details = re.findall(r'(?:지하\s*)?\d+(?:-\d+)?\s*(?:층|호|동)', verified_addr)
            tail = ' '.join(d.strip() for d in details)
            return (prefix + ' ' + P + ((' ' + tail) if tail else '')).strip()
    return None


def _enrich_with_poi(verified_addr: str, original_text: str) -> str:
    """POI 검색으로 상호 지점명 부착.

    조건:
      - POI road_address_name 이 verified 도로명·번지와 일치 (다른 위치 배제)
      - place_name 이 후보 상호 + 공백 + 지점명 형태 (예: "마성떡볶이 논현역점")
        → 상호 뒤에 실제 지점명이 붙어있을 때만 유의미. 아파트 이름 확장·유사 상호는 배제.
      - verified 에 이미 후보 상호 부착돼있으면 지점명 포함 형태로 replace
    """
    if not verified_addr or not original_text:
        return verified_addr
    v_key = _road_key(verified_addr)
    if not v_key:
        return verified_addr
    region = _extract_region_hint(verified_addr)
    candidates = _extract_shop_candidates(original_text)
    if not candidates:
        return verified_addr

    # (B) 복합 등록명(·) 건물 → 고객 지목 시설 POI 정식명 재구성 (2026-07-30 ETC-45cab9)
    if '·' in verified_addr or '・' in verified_addr:
        _b = _replace_compound_with_tenant(
            verified_addr, original_text, candidates, v_key, region)
        if _b:
            return _b

    # 다단어 상호 → 카카오 정식 place_name 채택 (2026-07-30 G1, L-03475 남양가 양꼬치)
    #   verified 에 실재하는 공백형 상호('남양가 양꼬치')만, 카카오 place_name 이
    #   그 상호로 시작('남양가양꼬치 마곡점')하고 도로명 일치할 때만 replace.
    #   가드: 공백형 verified 존재 + place_name startswith 결합형 + 도로 일치 +
    #        지점 suffix 有 + 부속시설 아님 → 오탐 최소.
    for _jc, _spaced in _joined_shop_candidates(original_text):
        if _spaced not in verified_addr:
            continue
        for _pname, _road in _search_poi(f'{_jc} {region}'.strip()):
            if not _pname or not _road or _road_key(_road) != v_key:
                continue
            _pn_ns = _pname.replace(' ', '')
            if not _pn_ns.startswith(_jc) or _pn_ns == _jc:
                continue  # 정식명이 상호로 시작 + 지점명 추가된 경우만
            if _poi_has_facility(_pname):
                continue  # 부속시설(주차장·ATM 등) POI
            return verified_addr.replace(_spaced, _pname, 1)

    # verified 에 이미 있는 후보는 우선순위 낮춤 (원문 신규 상호 먼저 시도)
    priority = (
        [c for c in candidates if c not in verified_addr]
        + [c for c in candidates if c in verified_addr]
    )
    for cand in priority[:5]:
        results = _search_poi(f'{cand} {region}'.strip())
        for place_name, road_name in results:
            if not place_name or not road_name:
                continue
            if _road_key(road_name) != v_key:
                continue
            # 매칭 조건 (2026-07-21 L-03316 확장):
            #   - 정확 매치 or "cand + 공백 + 지점명" (기존, 예: 마성떡볶이 논현역점)
            #   - place_name.endswith(cand) — 로컬 정식명이 접두 지역명 포함하는 케이스
            #     (예: 한강듀클래스 → 김포한강듀클래스). endswith 만 허용해 오탐 방지
            #     (L-03306 위례포레샤인 → 위례포레샤인23단지아파트 로 replace 되는 오탐 차단)
            _match = (
                place_name == cand
                or place_name.startswith(cand + ' ')
                or (len(cand) >= 4 and place_name.endswith(cand))
            )
            if not _match:
                continue
            # 부속시설 blacklist (2026-07-20 L-03299~) — 상호(첫 단어) 이후 어느
            # 단어든 부속시설이면 상호 지점 아니라 시설 표시. skip.
            # 2026-07-30 근본 수정: 기존 `_place_words[1]`(두 번째 단어)만 검사 →
            #   '롯데마트 고양점 주차장'(주차장=3번째)·'한국화훼농협 ATM 본점' 통과하던
            #   버그. _poi_has_facility 로 후행 전 단어 검사 + ATM·옥외 추가.
            if _poi_has_facility(place_name):
                continue
            if cand in verified_addr:
                place_words = place_name.split()
                # 두 단어 이상 (cand + 지점명) — 기존 dedup 체크 후 replace.
                #   예: verified='...현대시티아울렛 가산점 5층...' + place='현대시티아울렛 가산점 주차장'
                #     → replace 하면 '현대시티아울렛 가산점 주차장 가산점 5층' 중복 (2026-07-13 L-03190)
                # 2026-07-24 L-03379 fix: place_words 모든 단어가 verified 안에 이미 있으면 skip.
                #   기존 `already_has_branch` 는 `verified_words[i] == cand` 정확 매치라
                #   `미원스페셜티케미칼(주)` 처럼 접미 (주)/㈜ 붙은 형태 인식 못함.
                #   substring 검사로 확장 — verified 에 상호+지점 모두 있으면 replace 무의미.
                if len(place_words) >= 2 and all(w in verified_addr for w in place_words):
                    continue
                if len(place_words) >= 2:
                    place_second = place_words[1]
                    verified_words = verified_addr.split()
                    already_has_branch = False
                    for i, w in enumerate(verified_words):
                        if w == cand and i + 1 < len(verified_words):
                            if verified_words[i + 1] == place_second:
                                already_has_branch = True
                                break
                    if already_has_branch:
                        return verified_addr  # 이미 지점명 있음
                    return verified_addr.replace(cand, place_name, 1)
                # 단어 1개 & place_name != cand — 접두 지역명·오탈자 정정 케이스.
                #   예: cand='한강듀클래스' + place='김포한강듀클래스' → 접두 '김포' 부착 (2026-07-21 L-03316)
                # 2026-07-24 L-03361: verified 에 이미 접두어 포함된 상호가 있으면
                #   (대소문자 무시) replace 하지 말고 skip — 안 그러면 GS스틸타워 + gs스틸타워
                #   교체가 'GSgs스틸타워' 중복을 만듦.
                if place_name != cand:
                    if place_name.lower() in verified_addr.lower():
                        continue
                    return verified_addr.replace(cand, place_name, 1)
                continue  # 완전 동일 → 무의미
            # append 케이스: 원본에 법인 접두어 ((주)/㈜/주식회사) 있으면 유지
            # 2026-07-24 L-03372: 원본 '(주)아론' → POI place_name 은 '아론' 만 →
            #   append 시 접두어 소실. 원본 정보 보존을 위해 접두어 재부착.
            _m_prefix = re.search(
                r'((?:\(주\)|㈜|㈠|주식회사)\s*)' + re.escape(cand),
                original_text,
            )
            if _m_prefix:
                return f'{verified_addr} {_m_prefix.group(1)}{place_name}'.strip()
            return f'{verified_addr} {place_name}'.strip()
    return verified_addr


def _try_poi_fallback(text: str) -> Optional[str]:
    """카카오 verified 실패 케이스에서 POI(상호명) 검색으로 도로명 획득.

    조건 (2026-07-20 L-03292 최초 · 2026-07-21 L-03314 확장):
      - 시/도 없음 (매니저가 시/도 빼먹은 케이스만 타겟)
      - 상호 후보 하나가 POI place_name 정확 매치 (or 'cand ' 로 시작)
      - POI 결과 road_address_name 이 원본 힌트와 매치:
          · 구/동 있으면 → 지역 힌트 검증 (강한 필터)
          · 구/동 없으면 → 도로명 접두어 검증 (오탐 방지)
      - 구/동/도로명 다 없으면 skip (힌트 없이 검색 = 오탐 위험)

    Returns: normalize_display 결과 or None.
    """
    if not text:
        return None
    first_line = text.strip().split('\n', 1)[0].strip()
    # 시/도 있으면 skip (매니저 실수 케이스 대상 아님)
    if re.search(
        r'(?:서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주|고양|성남|수원|용인|안양|안산|광명|시흥|화성|평택|김포)'
        r'(?:특별시|광역시|특별자치시|특별자치도|도|시)?',
        first_line,
    ):
        return None
    candidates = _extract_shop_candidates(text)
    if not candidates:
        return None
    m_region = re.search(r'([가-힣]+동)|([가-힣]+구)', first_line)
    region = (m_region.group(1) or m_region.group(2)) if m_region else ''
    m_road = re.search(r'([가-힣]{2,}(?:대?로|길))\s*\d', first_line)
    road_prefix = m_road.group(1) if m_road else ''
    if not region and not road_prefix:
        return None  # 힌트 없으면 위험 → skip
    for cand in candidates[:5]:
        query = f'{cand} {region}'.strip() if region else f'{cand} {road_prefix}'.strip()
        results = _kakao_search_poi(query)
        if not results:
            continue
        for place_name, road_name in results:
            if not place_name or not road_name:
                continue
            # 정확 매칭: cand == place_name 또는 place_name 이 "cand " 로 시작
            if not (place_name == cand or place_name.startswith(cand + ' ')):
                continue
            # 지역 힌트 있으면 검증 (구/동 이 road_name 에 있어야 함)
            if region and region not in road_name.replace(' ', ''):
                continue
            # 구/동 없으면 도로명 접두어 검증 (POI road 에 원본 도로명 포함)
            if not region and road_prefix and road_prefix not in road_name:
                continue
            # 상호명(cand) 부착 — POI 매칭 성공 = 그 상호가 정답. 도로명 뒤에 붙임.
            return f'{normalize_display(road_name)} {cand}'
    return None


def resolve_address(
    text: str, regex_addr: Optional[str] = None, regex_level: str = ''
) -> Tuple[str, str]:
    """
    주소 확정 최종 함수. 호출 우선순위:

    1. 카카오 verified → ('도로명/지번 주소', 'verified')
    1b. POI fallback — 시/도 없이 구만 있는 케이스 → ('도로명 주소', 'verified')
    2. 정규식 결과 → (정규식 주소, 원래 level)
    3. 원문 첫 줄 (4~100자 + 한글 포함) → (첫 줄, 'raw')
    4. 다 실패 → ('', '')

    Returns:
        (주소, 신뢰도) 튜플. 신뢰도가 ''면 빈 결과.
    """
    # 1. 카카오 검증 시도
    verified = verify_address(text, regex_addr)
    if verified:
        addr, level = verified
        addr = _enrich_verified_address(addr, text, regex_addr)
        # 도로명주소에 법정동 중복 제거 (주소검색기 유입, 2026-07-27 L-03400)
        addr = _strip_redundant_legal_dong(addr)
        # tail 부착 후 후처리 (인접 유사 단어 dedup 등, 2026-07-22 ETC-b626fb)
        addr = _post_normalize_display(addr)
        return (addr, level)

    # 1b. POI fallback (2026-07-20 L-03292) — 시/도 빠진 케이스 상호 → 도로명
    poi_road = _try_poi_fallback(text)
    if poi_road:
        addr = _enrich_verified_address(poi_road, text, regex_addr)
        addr = _strip_redundant_legal_dong(addr)
        addr = _post_normalize_display(addr)
        return (addr, 'verified')

    # 2. 정규식 결과 (시도 prefix 정규화 적용 + 상호명 보강)
    if regex_addr:
        addr = normalize_display(regex_addr)
        addr = _enrich_verified_address(addr, text, regex_addr)
        return (addr, regex_level or 'regex')

    # 3. 원문 첫 줄 fallback — 엄격한 주소 패턴이 포함된 경우만
    # "세방정유라는 회사입니다" / "공장동 내에 ..." 같은 본문이 잘못 raw로 들어가는 것 방지
    if text:
        first_line = text.strip().split('\n', 1)[0].strip()
        first_line = re.sub(r'\s+', ' ', first_line)
        has_strict_address = re.search(
            r'(?:로|길)\s+\d|\d+(?:번지|호)'
            r'|(?:서울|부산|대구|인천|광주|대전|울산|세종'
            r'|경기|강원|충북|충남|전북|전남|경북|경남|제주)'
            r'\s+[가-힣]+(?:구|시|군|동)',
            first_line,
        )
        if (
            4 <= len(first_line) <= 100
            and re.search(r'[가-힣]', first_line)
            and has_strict_address
        ):
            return (first_line, 'raw')

    return ('', '')
