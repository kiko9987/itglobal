"""
카카오 로컬 API로 한국 주소 검증·정규화.

호출 순서 (verify_address):
1. 정규식 매칭 결과 (lead_helpers.extract_korean_address)
2. 원문 첫 줄 (메일 폼은 보통 첫 줄에 주소)
3. 본문에서 도로명+번지 패턴만 발췌
첫 매칭 성공한 결과를 사용. 모두 실패하면 None → 호출자가 fallback.

카카오 API 미설정 / 비활성 / 5초 타임아웃 → graceful 통과 (None 반환).
"""

import html
import json
import os
import re
import urllib.error
import urllib.parse
import urllib.request
from functools import lru_cache
from typing import List, Optional, Tuple

from dashboard.utils.logging_config import get_logger
from dashboard.services.lead_helpers import _normalize_road_spacing

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
    - 경기(수도권) + ○○시/군 → 도 prefix 제거 + "시/군" 제거 (수원·성남 등 익숙)
    - 비수도권 도(전북·전남·경북·경남·강원·충북·충남) + ○○시/군 → 도 유지 + "시/군"
      제거 ("전북 김제", 담당자 먼 지역 인지용, 2026-08-06 L-03646)

    >>> normalize_display('서울 강남구 테헤란로 152')
    '강남구 테헤란로 152'
    >>> normalize_display('인천 연수구 갯벌로 36')
    '인천 연수구 갯벌로 36'
    >>> normalize_display('광주 동구 충장로 1')
    '전남 광주 동구 충장로 1'
    >>> normalize_display('경기 광주시 경안로 100')
    '광주 경안로 100'
    >>> normalize_display('전북특별자치도 김제시 남북로 218')
    '전북 김제 남북로 218'
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
    elif ((first in ('광주', '광주광역시')
           # '전남광주통합특별시' 등 광주 포함 특별시 명칭도 광주광역시로 인식 → '전남 광주'
           #   (경기 광주시와 구분, L-03672). 광역시는 위 set, 특별시형만 추가 매치.
           or ('광주' in first and first.endswith('특별시')))
          and second.endswith('구')):
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
        elif second == '서귀포시':          # 제주도 서귀포시 → 서귀포 (docstring)
            result = ' '.join(['서귀포'] + rest)
        else:
            result = ' '.join(['제주'] + tokens[1:])
    elif first == '제주' and second == '제주시':
        result = ' '.join(['제주'] + rest)
    elif first == '제주' and second == '서귀포시':   # 제주 서귀포시 → 서귀포 (pre-existing 갭)
        result = ' '.join(['서귀포'] + rest)
    elif first in _PROV_SHORT:
        # 비수도권 도(전북·전남·경북·경남·강원·충북·충남)는 도 접두 유지 —
        #   담당자가 '김제'만 보면 먼 지역인지 감이 안 옴 → '전북 김제' (2026-08-06
        #   L-03646, 사용자 결정). 경기(수도권, 수원·성남 등 익숙)는 기존대로 제거.
        _keep_do = first != '경기'
        if second.endswith('시') or second.endswith('군'):
            _city = second[:-1]
            result = ' '.join(([first, _city] if _keep_do else [_city]) + rest)
        else:
            result = ' '.join(tokens if _keep_do else tokens[1:])
    elif (first.endswith('시') and len(first) >= 3
          and not first.endswith('광역시') and not first.endswith('특별시')
          # 단독 'XX시'(도 접두 없음) 축약 — 뒤가 구/읍/면뿐 아니라 도로명(로/길)·법정동
          #   (동/리/가)·번지(숫자)여도 시 제거 (2026-09-01 L-03867: 구 없는 시(동두천시)나
          #   구 생략 입력(수원시 매영로)이 '동두천시 아차노리로…'처럼 시가 남던 갭).
          and (second.endswith(('구', '읍', '면', '동', '리', '가', '로', '길'))
               or re.match(r'\d', second))):
        result = ' '.join([first[:-1]] + tokens[1:])
    else:
        result = addr

    return _post_normalize_display(result)


_SAME_SUFFIX_DEDUP = ('캠퍼스',)


# 기관·건물명 괄호 벗기기용 (L-03738) — 단일 토큰 + 강한 건물/기관 접미만.
_INST_PAREN_RE = re.compile(
    r'\s*\(([가-힣A-Za-z0-9]+(?:대학교|대학병원|대학|학교|병원|빌딩|타워|센터|캠퍼스|'
    r'문화관|회관|아트홀|미술관|박물관|도서관|체육관|경기장|아파트|오피스텔|프라자|'
    r'플라자|스퀘어|시티|타운|호텔|리조트|콘도|교회|성당|백화점|공장|연구원|연구소))\)'
)


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
    # (법정동 + 건물명) 무-콤마 괄호 → 건물명만 밖으로, 법정동 제거 (2026-08-18 L-03702).
    #   매니저가 '(갈현동 벧엘)' 처럼 '법정동 <건물>'을 괄호로 적은 케이스. 카카오/관행
    #   정식 괄호는 '(법정동, 건물)' **콤마형**이라 미매치(아래 콤마 정리로만 처리, 괄호
    #   유지). 번지형('(걸포동 172-1)')은 건물부가 숫자로 시작 → 미매치. 반드시 콤마
    #   제거(하단) 전에 실행해 콤마형을 무-콤마로 오인 unwrap 하지 않도록.
    #   법정동 형태: 'XX동'·'XX동N가'(성수동2가)·'XX로N가'(을지로3가)·'XX가/리' 지원
    #   (2026-08-18 L-03707). '로N가'는 반드시 숫자+가 있어야 매치 → 도로명(테헤란로) 제외.
    #   건물부는 1단계 중첩 괄호 허용 (2026-08-19 L-03716): '(정왕동 (주)삼인)' → '(주)삼인'
    #   (내부 '(주)'·'㈜'의 ')' 에서 안 끊기게). 숫자 시작(번지형)은 미매치(제거 규칙으로).
    #   2026-08-25 L-03771: **콤마형** '(고잔동, Win-Win프라자)' 도 unwrap (법정동 뒤
    #   구분자 `[,\s]+`) → 건물명만 남김. 법정동은 도로주소에 이미 있어 괄호로 남길 이유 없음.
    addr = re.sub(
        r'\s*\((?:[가-힣]+동(?:\d+가)?|[가-힣]+로\d+가|[가-힣]+(?:가|리))'
        r'[,\s]+((?!\d)[^()]*(?:\([^()]*\)[^()]*)*)\)',
        lambda m: ' ' + m.group(1).strip(), addr,
    )
    # 법정동 단독 괄호 제거 (2026-08-25 L-03779): 홈페이지 폼은 다음(카카오) 우편번호
    #   검색 위젯이 도로명주소를 '도로명 번지 (법정동)' 표준 표기로 반환 → 당근·전화(괄호
    #   없음)와 표기 불일치. 법정동은 도로명주소에 이미 포함돼 군더더기 → 제거해 통일.
    #   가드: **번지(숫자) 직후** 괄호만(위젯 signature — '16-8 (원창동)') + 건물 구역동
    #   (관리동·상가동·사무동 등)은 negative lookahead 로 제외(공장·복합건물 동 지정 보존).
    #   아파트 동(101동)·라틴(A동)·별관·(주) 는 애초에 [가-힣]+동 패턴 밖이라 미매치.
    addr = re.sub(
        r'(?<=\d)\s*\('
        r'(?:(?!관리|상가|사무|생활|기숙|별관|본관|신관|후생|복지|생산|공장|창고|주차)'
        r'[가-힣]+동(?:\d+가)?|[가-힣]+로\d+가|[가-힣]+(?:가|리))\)',
        '', addr)
    # 기관·건물명 괄호 벗기기 (2026-08-21 L-03738): 고객이 '(경희대학교)' 처럼 기관/건물
    #   명을 괄호로 감싼 것 → 괄호 제거. **단일 토큰**(공백·콤마 없음)이면서 강한 건물·기관
    #   접미로 끝날 때만 → 노트('(주차 …)')·(예정)·법정동·콤마형·(주) 는 미대상.
    addr = _INST_PAREN_RE.sub(r' \1', addr)
    # 건물명 뒤 중복어 '건물' 제거 (2026-08-25 L-03778): '프라임하우스 건물 1층' →
    #   '프라임하우스 1층'. 강한 건물 접미 뒤 단독 '건물'만(뒤 한글 lookahead 로
    #   '건물주'·'건물관리' 제외). 접미 없는 상호는 미대상(오제거 방지).
    addr = re.sub(
        r'((?:빌딩|하우스|타워|프라자|플라자|스퀘어|빌라|맨션|오피스텔|파크|타운|'
        r'시티|캐슬|팰리스))\s+건물(?![가-힣])',
        r'\1', addr)
    # 2026-07-24 L-03367: 매니저가 주소에 콤마 계속 넣는 습관 → `콤마+공백` 만 공백으로 치환.
    #   `3,4,7호` (호수 콤마+공백 없음) 는 유지, `한라시그마밸리2차, 1층` 만 정리.
    # 2026-07-27 L-03401: 숫자 사이 콤마 (호수 나열 `305, 306, 408호`) 는 보존.
    #   `31, 서울숲` (번지-건물 구분) 만 제거. 숫자,(공백)숫자 는 placeholder 로 보호 후 복원.
    addr = re.sub(r'(?<=\d),(?=\s*\d)', '\x00', addr)  # 숫자 사이 콤마 보호
    addr = re.sub(r',\s+', ' ', addr)                   # 나머지 콤마+공백 제거
    addr = addr.replace('\x00', ',')                    # 보호분 복원
    # 콤마+한글(공백 없음) → 공백 (2026-08-24 L-03752): '15층,뉴헤어의원'→'15층 뉴헤어의원'.
    #   기존 규칙은 '콤마+공백'만 처리해 공백 없는 콤마가 남았음. 숫자,숫자(호수 나열
    #   305,306호)는 위 \x00 보호로 복원돼 뒤가 숫자라 미매치(보존).
    addr = re.sub(r',(?=[가-힣])', ' ', addr)
    # 한글+숫자 층/호/관/동 붙임 분리 (2026-08-24 L-03754): '상가2238호'→'상가 2238호'.
    #   building_tail 은 base normalize(호/층 분리)를 안 거쳐 최종 단계에서 보강.
    #   'N동주민센터' 등 행정동 접미는 분리 금지(base 규칙과 동일). 이미 공백이면 무변.
    addr = re.sub(r'(?<=[가-힣])(\d+동)(?!주민|사무|행정|복지|자치)', r' \1', addr)
    addr = re.sub(r'(?<=[가-힣])(\d+(?:호|층|관|호실))', r' \1', addr)
    # ' 산 (숫자)' → ' 산(숫자)'
    addr = re.sub(r' 산 (\d)', r' 산\1', addr)
    # 구분 마침표 제거 (2026-08-29 L-03843): 고객이 번지와 상세 사이에 마침표를 구분자로
    #   넣은 것('42 . 지하1층'·'42.지하1층') → 공백. 끝 마침표(L-03769)는 아래 return 에서
    #   별도 처리. 소수점(1.5)은 뒤가 숫자라 미매치(가드).
    addr = re.sub(r'\s+[.．]\s+', ' ', addr)                    # ' . ' 구분자
    addr = re.sub(r'(?<=\d)[.．](?=\s*[가-힣])', ' ', addr)     # 번지.상세 (숫자 뒤 한글 앞)
    # 인접 유사 단어 dedup
    from difflib import SequenceMatcher
    tokens = addr.split()
    out = []
    for t in tokens:
        if out:
            prev = out[-1]
            # (0-a) 인접 동일 토큰 제거 (2026-08-25 L-03783): base 동-분리가 '자이101동'
            #   →'자이 101동' 으로 쪼개면 앞 단지명(자이)과 겹쳐 '자이 자이 101동'. 완전
            #   동일한 인접 토큰은 중복 → 뒤 것 skip. (짧은 단지명은 아래 접두 병합 3자
            #   가드 밖이라 이 규칙으로 커버.)
            if t == prev and re.search(r'[가-힣]', t):
                continue
            # (0-b) 단지명 반복 접두 병합 (2026-08-25 L-03783): 다음 위젯 building_name
            #   '종암2차 아이파크' unwrap 후, 고객이 상세에 '아이파크상가동'처럼 단지명을
            #   다시 붙여 '아이파크 아이파크상가동' 중복. cur 가 prev 로 시작하고 나머지가
            #   '…동' 건물 구역(상가동·관리동·N동)이면 **접두만** 제거 → '아이파크 상가동'
            #   (구역 정보 상가동 보존). prev 3자↑ 한글 & 나머지 한글+동 접미로 오제거 방지.
            if (len(prev) >= 3 and re.search(r'[가-힣]', prev)
                    and t != prev and t.startswith(prev)):
                _rem = t[len(prev):]
                if _rem.endswith('동') and re.search(r'[가-힣]', _rem):
                    t = _rem
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
    # 끝 마침표·공백 잔재 제거 (2026-08-25 L-03769): 고객이 '2ㅡ9.' 처럼 끝에 붙인 마침표
    #   가 '… 2-9 .' 로 남던 것 정리.
    return re.sub(r'[\s.．]+$', '', ' '.join(out))


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


# 카카오 building_name 끝의 순수 로마자 괄호 표기 제거 (L-03693): 카카오 등록 건물명이
#   '브릿지타워(BRIDGE TOWER)' 처럼 국문명+영문 로마자를 병기 → 방문 주소엔 노이즈.
#   끝 괄호 안이 **로마자/숫자/기호로만** 이뤄지고 로마자 1자 이상일 때만 제거
#   ('아이파크(IPARK)'→'아이파크'). '(주)'·'(101동)'·'(1234)' 등 국문·순수숫자는 보존.
_ROMAN_PAREN_TAIL_RE = re.compile(r"\s*\((?=[^)]*[A-Za-z])[A-Za-z0-9 .,&'\-]+\)\s*$")


def _strip_roman_paren(s: Optional[str]) -> Optional[str]:
    return _ROMAN_PAREN_TAIL_RE.sub('', s).strip() if s else s


@lru_cache(maxsize=512)
def _kakao_search_cached(query: str) -> Optional[dict]:
    url = KAKAO_ENDPOINT + '?' + urllib.parse.urlencode({'query': query})
    data = _kakao_get_json(url)  # _KakaoTransientError 는 lru_cache 미캐시
    if data is None:
        return None
    docs = data.get('documents', [])
    if not docs:
        return None
    doc = docs[0]
    # 카카오 API HTML escape 정규화 (building_name 등에 &amp; 잔존 방지, L-03583).
    for _sub in ('road_address', 'address'):
        _obj = doc.get(_sub)
        if isinstance(_obj, dict):
            for _k in ('building_name', 'address_name'):
                if isinstance(_obj.get(_k), str):
                    _obj[_k] = html.unescape(_obj[_k])
            # 건물명만 영문 로마자 괄호 제거 (주소 문자열엔 미적용)
            if isinstance(_obj.get('building_name'), str):
                _obj['building_name'] = _strip_roman_paren(_obj['building_name'])
    return doc


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
    # 카카오 API 는 place_name 등을 HTML escape 해 반환 ('케이&amp;케이…', L-03583).
    # 수신 즉시 unescape 해 다운스트림이 clean 텍스트만 보도록 (& → &amp; 잔존 방지).
    return tuple(
        (html.unescape(d.get('place_name', '') or ''),
         html.unescape(d.get('road_address_name', '') or ''))
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
        # title 에 <b> 태그 있음 (강조) → 제거 후 HTML 엔티티 unescape.
        # 순서 주의: 실제 <b> 태그 먼저 제거 → 그다음 &amp;/&lt; 등 unescape.
        # (unescape 먼저 하면 name 안 '&lt;' → '<' 이 태그 strip 에 오삭제될 수 있음)
        return tuple(
            (
                html.unescape(re.sub(r'<[^>]+>', '', d.get('title', '') or '')),
                html.unescape(d.get('roadAddress', '') or ''),
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


# 행안부 도로명주소 검색 API (2026-08-14 L-03671) — 권위 소스.
#   카카오 address.json·POI 가 0건인 실재 필지(백제고분로19길 13 등 미인덱싱 parcel)를
#   정부 공식 DB 로 verified 확인. roadAddr·jibunAddr·bdNm(건물명) 반환.
_JUSO_ENDPOINT = 'https://business.juso.go.kr/addrlink/addrLinkApi.do'


def _juso_key() -> str:
    return os.getenv('JUSO_CONFM_KEY', '').strip()


@lru_cache(maxsize=512)
def _juso_search_cached(query: str) -> tuple:
    """행안부 도로명주소 검색 — ((roadAddr, jibunAddr, bdNm), ...) 튜플.

    오류·빈결과·키 미설정은 () 반환 (resolve 파이프라인 절대 안 막음).
    roadAddr 는 '서울특별시 송파구 백제고분로19길 13 (잠실동)' 형태(법정동 괄호 포함).
    """
    key = _juso_key()
    if not key or not query.strip():
        return ()
    try:
        url = _JUSO_ENDPOINT + '?' + urllib.parse.urlencode({
            'confmKey': key, 'currentPage': 1, 'countPerPage': 10,
            'keyword': query.strip(), 'resultType': 'json',
        })
        with urllib.request.urlopen(url, timeout=5) as r:
            data = json.loads(r.read().decode('utf-8'))
        common = (data.get('results') or {}).get('common') or {}
        if str(common.get('errorCode', '')) not in ('0', ''):
            return ()
        juso = (data.get('results') or {}).get('juso') or []
        return tuple(
            (j.get('roadAddr', '') or '', j.get('jibunAddr', '') or '',
             j.get('bdNm', '') or '')
            for j in juso
        )
    except Exception as exc:
        logger.debug(f'[JUSO] {type(exc).__name__}: {query[:40]}')
        return ()


# 도로명+번지 패턴 (예: "갯벌로 36", "테헤란로 152", "강남대로 401-1", "꽃내음1길 19-22",
# "봉은사로 26길 12", "부천로431번길 16")
# - 한글 2글자 이상으로 도로명 시작 (1글자 "번길" 같은 오인 차단)
# - 도로명 중간/끝 숫자 허용 (꽃내음1길, 시흥대로14길)
# - 2단 도로명 옵션 (○○로 N길 12, 봉은사로 26길 12)
# - 도로명 + 숫자번길 옵션 (부천로431번길 16) — "번" 옵션
# 도로명+번지. `[가-힣]*` 을 \d* 뒤에 둬 '공단1대로'(한글+숫자+한글+로) 형태 지원
#   (2026-08-19 L-03716) — 기존엔 '공단1대로195번길'을 못 잡아 행안부 확인 불발.
_ROAD_PATTERN = re.compile(
    r'([가-힣]{2,}\d*[가-힣]*(?:로|길)(?:\s*\d+번?(?:로|길))?\s+\d+(?:-\d+)?)'
)


# 건물명/층/호 정보 추출용
# 의미 있는 tail 신호:
# - 숫자 + 동/층/호 (예: "1층", "102호", "302동")
# - prefix + 시설 키워드 (예: "DMC빌딩", "인하대학교", "광장힐스테이트")
_TAIL_SIGNAL = re.compile(
    r'(?:'
    r'\d+\s*(?:동|호|층|관|블록|블럭|단지)'
    # 숫자 없는 층 표기 (2026-08-14 L-03696): '지층 한국유통' 의 '지층'(지하층)이 숫자
    #   없어 tail 신호로 인정 못 받아 상호까지 통째 유실. 지하철/지하상가 등 복합어는
    #   뒤 한글 lookahead 로 제외 (독립 토큰만).
    r'|(?:지층|반지하|지하|옥탑|옥상)(?![가-힣])'
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

    # 법인 표기 '(주)/(유)/(사)/(재)' 는 flatten 하면 '주' 홀로 남아 어색
    #   ('에스티팜 시화공장(주)' → '…시화공장 주') → 원형 유지 (2026-08-06 L-03358).
    # '(예정)' 지정어도 괄호 유지 → 뒤 _mark_planned 가 'X (예정)' 로 정규화 (L-03680).
    if re.fullmatch(r'[주유사재]|예정', m.group(1).strip()):
        return tail

    before = tail[:m.start()].strip()
    after = tail[m.end():].strip()
    # 호수 나열 콤마 (숫자/호 뒤 + 숫자 앞: '301호,302호'·'305,306,408호') 는 보호 —
    #   split/공백화로 잃지 않게 placeholder 로 (2026-08-06 ETC-45c37b, 괄호 안에서도
    #   콤마 보존 = _post_normalize 의 '숫자 사이 콤마 보존'과 일관). '중계동, 건물'
    #   (동 뒤 콤마) 는 미보호 → 공백 정리.
    _inner = re.sub(r'(?<=[\d호]),(?=\s*\d)', '\x00', m.group(1))
    parts = [p.strip() for p in _inner.split(',')]
    if parts and re.match(
        r'^[가-힣]+(?:동|가|리)(?:\s+\d+(?:-\d+)?)?$', parts[0]
    ):
        parts = parts[1:]
    inner_flat = ' '.join(p for p in parts if p).strip()

    result = ' '.join(x for x in (before, inner_flat, after) if x)
    # before/after 의 호수 콤마도 보호 후 나머지 콤마(동,건물 등)만 공백으로 flatten
    result = re.sub(r'(?<=[\d호]),(?=\s*\d)', '\x00', result)
    result = re.sub(r'\s*,\s*', ' ', result)  # "B1, 위플레이스" → "B1 위플레이스"
    result = re.sub(r'\s+', ' ', result).strip().replace('\x00', ',')
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
            # 2026-08-06 L-03600: '예정지'(예정+지)는 계획 중 '장소' 표시라 절단 안 함
            #   (뒤 _mark_planned 가 '(예정)' 으로 변환). '예정'(지 없음)만 동사구
            #   ('설치 예정')로 보고 기존대로 절단.
            # '예정지'(장소) 및 '(예정)'(괄호 지정어)는 절단 안 함 — 계획 중 장소 표시.
            #   뒤 _mark_planned 가 'X (예정)' 로 정규화. '설치 예정'(동사구, 괄호X)만 절단.
            if sw == '예정' and 0 <= p and (
                    tail[p + 2:p + 3] == '지' or (p > 0 and tail[p - 1] == '(')):
                continue
            # 2026-08-06 ETC-f89e73: '회사' 가 법인 표기 '주식/유한/합자/합명회사'
            #   내부를 자르는 것 방지 — 매니저가 붙인 법인 표기 보존(사용자 결정).
            #   standalone '○○ 회사'(공백 앞)·다른 복합어는 기존대로 절단.
            if sw == '회사' and 0 <= p and tail[max(0, p - 2):p] in (
                    '주식', '유한', '합자', '합명'):
                continue
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
        # 평수(면적) 이후는 상세/문의라 절단 (2026-08-25 L-03774): '이마트…2층 25평
        #   인테리어공사시작단계' → '이마트…2층'. 공백+숫자+평(뒤 한글 아님) 지점에서 자름
        #   ('평택로'·'평화빌딩' 등 오절단 방지 위해 앞 공백+숫자 필수).
        _mp = re.search(r'\s\d+\s*평(?![가-힣])', tail)
        if _mp and _mp.start() < cut_pos:
            cut_pos = _mp.start()
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

    # 지역 접두(시/도 + 시/군/구) 추출 — 도로+번지가 여러 도시에 있을 때 카카오가
    #   엉뚱한 도시를 집는 것 방지 (2026-08-06 L-03659: '인천 부평구 경인로 789' 를
    #   지역 없는 '경인로 789' 로 조회 → 서울 영등포 경인로 789 오매칭). 지역+도로+번지
    #   후보를 지역 없는 것보다 먼저 시도.
    region_prefix = ''
    if first_line:
        m_reg = re.match(
            r'((?:서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|'
            r'전북|전남|경북|경남|제주)(?:특별시|광역시|특별자치시|특별자치도|도)?'
            r'(?:\s+[가-힣]+(?:시|군|구))*)',
            first_line,
        )
        if m_reg and m_reg.group(1).strip():
            region_prefix = m_reg.group(1).strip()

    # 도로명+번지 추출 — first_line 통째보다 우선 (2026-07-23 ETC-1ad649)
    # 카카오 verify 는 tail(상호명 등) 붙은 query 에서 번지를 잘못 근사하는 케이스 있음
    # 예: '장승배기로20길 46-3 송학대교회' → 46 리턴 (잘못됨) vs
    #     '장승배기로20길 46-3' → 46-3 리턴 (정확)
    # 순수 도로명+번지 후보를 우선 시도해 첫 성공이 정확하도록 함.
    if text:
        for m in _ROAD_PATTERN.finditer(text):
            road = m.group(1)
            # 지역+도로+번지 를 먼저 (도시 모호성 해소), 그다음 지역 없는 버전 (fallback)
            if region_prefix and region_prefix.replace(' ', '') not in road.replace(' ', ''):
                _push(f'{region_prefix} {road}')
            _push(road)

    if first_line:
        _push(first_line)

    if regex_addr:
        _push(regex_addr)

    if first_line and ',' in first_line:
        _push(first_line.split(',', 1)[0])

    return out


# 카카오 building_name 이 '건물'이 아니라 상가 입주 시설(교회·학원 등)인 접미 —
#   고객이 별도 상호를 준 경우 이 건물명 부착 skip (L-03627 열방교회). 종교·소규모
#   교육시설: 상가 건물의 대표 등록명으로 잡히나 실제론 한 입주 업체.
_KAKAO_BLDG_TENANT_RE = re.compile(r'(교회|성당|사찰|기도원|어린이집|유치원|학원|독서실|공부방)$')

# 고객이 아파트명을 축약·오기(‘독산동신도아파트’ ← 정식 ‘신도브래뉴아파트’)했을 때
#   공식명으로 승격 (L-03695). 원문 우선 원칙의 예외 — **아파트에 한해**, 이중 소스
#   (카카오 building_name == 행안부 bdNm)가 같은 아파트로 일치 + 고객명이 핵심 토큰
#   (‘신도’)을 공유할 때만. 다른 아파트를 지목했을 가능성(래미안≠신도브래뉴)은 토큰
#   불일치로 거절 → 원문 유지. 동/호 tail(‘101동 502호’)은 보존.
_APT_NAME_RE = re.compile(r'^([가-힣]+아파트)(\s.*)?$')
_LEGAL_DONG_PREFIX_RE = re.compile(r'^[가-힣]{2,3}동(?=[가-힣])')


def _maybe_upgrade_apartment_name(base: str, building_tail: str,
                                  kakao_bldg: str) -> Optional[str]:
    """고객 아파트 tail 을 이중소스 일치 공식명으로 승격. 조건 미충족 시 None(원문 유지)."""
    if not (building_tail and kakao_bldg and kakao_bldg.endswith('아파트')):
        return None
    m = _APT_NAME_RE.match(building_tail.strip())
    if not m:
        return None
    cust_name, rest = m.group(1), (m.group(2) or '')
    if cust_name == kakao_bldg:
        return None  # 이미 공식명
    # 행안부 bdNm 교차확인 (base = 정규화 도로명+번지)
    juso_bd = ''
    try:
        for _ra, _ji, _bd in (_juso_search_cached(base) or ()):
            if _bd:
                juso_bd = _bd.strip()
                break
    except Exception:
        return None
    if juso_bd != kakao_bldg:   # 이중 소스 불일치 → 승격 안 함
        return None
    # 핵심 토큰 공유: 고객명(법정동 접두 제거) 앞부분이 공식명에 포함돼야
    cust_core = _LEGAL_DONG_PREFIX_RE.sub('', cust_name)          # 독산동신도아파트→신도아파트
    cust_key = cust_core[:-3] if cust_core.endswith('아파트') else cust_core  # 신도
    if len(cust_key) >= 2 and cust_key in kakao_bldg:
        return f'{kakao_bldg}{rest}'
    return None


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

    def _compose(base: str, protect_name: str = '') -> str:
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
        # 2026-08-05 예외 (L-03553): 카카오 확정 건물명은 원문 존중 — 쪼갠 결과가
        #   카카오 verified(base) 또는 카카오 building_name(protect_name) 에 붙은 채로
        #   실재하면 분리 금지 ('타워팰리스' → '타워 팰리스' 오분리 방지). 매니저가
        #   붙여 쓴 '단지상가' 는 어느 쪽에도 없어 기존대로 분리됨(규칙 1 보존, 회귀 없음).
        #   ※ 타워팰리스는 tail_has_building 경로라 base 미포함 → building_name 으로 방어.
        #   ETC-3b713a 동 규칙과 동일 사상(카카오 정확값을 후처리가 망치는 것 방지).
        _protect_nospace = (
            (base or '').replace(' ', '') + '\x00' + (protect_name or '').replace(' ', '')
        )

        def _spacing_guard(m):
            tok = m.group(1)
            _fol = re.match(r'[가-힣]+', m.string[m.end():])
            _combined = tok + (_fol.group(0) if _fol else '')
            # 쪼개려는 단어(토큰+뒤 한글)가 카카오 base 또는 building_name 에 붙어
            # 실재 → 정식 건물명, 유지
            if _combined and _combined in _protect_nospace:
                return tok
            return tok + ' '

        result = re.sub(
            r'(단지|상가|아파트|빌딩|타워|오피스텔|맨션|빌라|하우스|클래스원)'
            r'(?=[가-힣])(?!대로|로|길|번길|동|층|호|관|번지)',
            _spacing_guard, result,
        )
        # 2. 한글 다음 숫자+동/호/층/관 — "○○상가101호" → "○○상가 101호"
        # 영문 제외 (2026-07-23 ETC-678632): "B1층" (지하 1층) 이 "B 1층" 으로 잘못 분리되는 케이스 방지.
        # 영문+숫자+층/호 는 대개 원본에 이미 공백 있음 ("SR타워 3층", "K타워 9층").
        # 동: 뒤에 시설명(주민센터·사무소·행정복지센터·자치센터)이 오면 행정동명 일부
        #   ('행당제1동주민센터')라 분리 금지 (2026-08-06 ETC-3b713a). 호/층/관은 기존대로.
        result = re.sub(r'(?<=[가-힣])(\d+동)(?!주민|사무|행정|복지|자치)', r' \1', result)
        result = re.sub(r'(?<=[가-힣])(\d+(?:호|층|관))', r' \1', result)
        # 2-b. 층/호/관 다음 한글 (부가 설명·시설 tail) — "7층복도" → "7층 복도"
        result = re.sub(r'(\d+동)(?!주민|사무|행정|복지|자치)([가-힣])', r'\1 \2', result)
        result = re.sub(r'(\d+(?:호|층|관|호실))([가-힣])', r'\1 \2', result)
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
        nonlocal building_tail  # 아파트명 승격 시 tail 교체 (L-03695)
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
            # 아파트명 축약·오기 승격 (L-03695): 고객 tail 이 아파트 근사치이고 카카오·
            #   행안부가 같은 아파트로 일치하면 공식명으로 tail 교체 → 아래 skip 로직이
            #   중복 없이 공식명 부착. 조건 미충족이면 building_tail 불변(원문 유지).
            if building_name.endswith('아파트') and building_tail:
                _apt_up = _maybe_upgrade_apartment_name(base, building_tail, building_name)
                if _apt_up:
                    building_tail = _apt_up
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
            # 카카오 building_name 이 입주 시설(교회·학원 등) 접미인데 고객이 이미 별도
            #   상호(tail)를 준 경우 = 상가 건물의 한 입주 업체를 건물명으로 잘못 받은 것.
            #   부착하면 '열방교회 도그버디 면목점'(교회+펫샵) 처럼 엉뚱 (2026-08-14 L-03627,
            #   사용자: 열방교회는 시장 상가 입주 교회). tail 없으면(그 시설 자체 방문) 유지.
            _bldg_is_tenant = bool(
                _KAKAO_BLDG_TENANT_RE.search(building_name)
                and (building_tail or '').strip()
            )
            if (
                building_name
                and not _bldg_is_tenant
                and building_name.lower() not in base.lower()
                # 원본 tail 에 이미 같은 건물명이 있으면 append 안 함 — 안 그러면 카카오
                #   건물명 + tail 의 건물명 이중 노출 ('서울랜드 후문 서울랜드 산타레스토랑',
                #   2026-08-04 L-03517). tail 은 _compose 가 항상 뒤에 붙이므로 중복됨.
                and building_name.replace(' ', '').lower()
                    not in (building_tail or '').replace(' ', '').lower()
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
            return (_compose(base, protect_name=building_name), 'verified')
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
    if not dong:
        return addr
    # 도로suffix 공백 보강(매산로3가→매산로 3가, 2026-08-04 L-03422)으로 법정동이 분리
    #   저장될 수 있어 원형·공백변형 둘 다 매칭. 도로명주소엔 법정동 불필요 → standalone 제거.
    _variants = [dong]
    _spaced = re.sub(r'([로길])(\d)', r'\1 \2', dong)
    if _spaced != dong:
        _variants.append(_spaced)
    _present = [d for d in _variants if d in addr]
    if not _present:
        return addr
    stripped = addr
    for d in _present:
        stripped = re.sub(
            rf'(?<![가-힣]){re.escape(d)}(?![가-힣])\s*', '', stripped,
        )
    stripped = re.sub(r'\s+', ' ', stripped).strip()
    return stripped or addr


def _strip_redundant_jibun(addr: str) -> str:
    """도로명주소에 딸려온 구 지번(번지)을 제거 — 카카오 jibun 본번-부번과 정확 일치할 때만.

    2026-08-06 L-03627: 고객이 도로명+지번을 함께 적음
    ('사가정로50길 51 열방교회 632-2 1층' — 632-2 는 면목동 632-2 지번).
    도로명+번지가 있으면 지번은 중복·불필요 → standalone 제거. _strip_redundant_legal_dong
    (법정동 제거)과 동일 사상, 지번 번지 버전.

    안전 가드:
    - 도로명+번지 있을 때만 (지번주소는 지번이 본질 → skip)
    - 카카오 jibun 본번(-부번) 과 정확 일치하는 토큰만
    - 도로명 번지와 우연히 같으면 보존(그게 번지)
    - 호/층/번지/동/관 접미 붙은 건 유닛번호라 보존 (632-2호 등)
    """
    if not addr:
        return addr
    m_road = re.search(r'[가-힣\d]+(?:로|길)\s*\d+(?:-\d+)?', addr)
    if not m_road:
        return addr
    doc = _kakao_search(m_road.group(0))  # lru_cache (legal_dong 과 공유)
    if not doc:
        return addr
    jibun = doc.get('address') or {}
    main_no = (jibun.get('main_address_no') or '').strip()
    sub_no = (jibun.get('sub_address_no') or '').strip()
    if not main_no:
        return addr
    jibun_num = f'{main_no}-{sub_no}' if sub_no and sub_no != '0' else main_no
    # 도로명 번지 추출 — 지번이 도로명 번지와 같으면 그건 번지라 보존
    road_num_m = re.search(r'(?:로|길)\s*(\d+(?:-\d+)?)', addr)
    if road_num_m and road_num_m.group(1) == jibun_num:
        return addr
    # standalone 지번 제거 (앞: 숫자/한글 아님, 뒤: 숫자/한글 or 호·층·번지·동·관 접미 아님)
    stripped = re.sub(
        rf'(?<![\d가-힣-]){re.escape(jibun_num)}(?![\d가-힣-]|\s*(?:호|층|호실|번지|동|관))\s*',
        '', addr,
    )
    stripped = re.sub(r'\s+', ' ', stripped).strip()
    return stripped or addr


# 강한 건물 접미(거의 항상 건물) — 파크/타운/시티/센터 등 모호어는 제외(입주사 오인)
_BLD_STRONG_SUFFIX_RE = re.compile(
    r'(타워|빌딩|오피스텔|프라자|플라자|스퀘어|캐슬|테크노|디팰리스|메가시티)'
)
# 입주사·부속시설 마커 — 있으면 건물명 아님(그 건물의 한 테넌트)
_POI_TENANT_MARK_RE = re.compile(
    r'(점$|지점|주차장|출입구|정문|후문|매장|약국|의원|병원|은행|마트|편의점|'
    r'충전소|ATM|커피|카페|식당|헬스|피트니스|\d호$|\d층$)'
)
# 번지 뒤 유닛 토큰(동/층/호/관/F) 판별 — 모두 유닛이면 '건물명 없는 민숭 주소'
_UNIT_TOKEN_RE = re.compile(
    r'^(?:[A-Za-z]?\d+(?:-\d+)?(?:동|층|호|호실|관)?|[A-Za-z]?\d+[Ff]|'
    r'지하\d*층?|[A-Z]동|B\d*)$'
)


def _enrich_building_by_road(verified_addr: str) -> str:
    """상호·카카오 building_name 둘 다 없는 '민숭한' 도로명주소에 건물명 부착 (L-03633).

    2026-08-06: 카카오 주소검색 API 가 building_name 을 빈값으로 주고 고객도 건물명
    없이 '번지 + 동/층/호'만 입력하면(다산지금로 202 B동 5F 0001호) 건물명이 통째
    누락. 큰 건물(다산 DIMC테라타워)인데도 주소API 엔 없음. 번지로 POI 조회해
    건물 접미 POI(입주사 마커 없는)가 **정확히 하나**면 부착.

    보수적 가드(입주사 오부착 방지):
    - 번지 뒤 토큰이 전부 유닛(동/층/호/F)일 때만 = 건물명이 진짜 없는 케이스
    - 강한 건물 접미(타워/빌딩/오피스텔/…)만, 입주사 마커(점/주차장/N호…) 제외
    - 같은 도로의 건물 후보가 정확히 1개일 때만 (여러 개면 모호 → skip)
    """
    if not verified_addr:
        return verified_addr
    m_road = re.search(
        r'[가-힣A-Za-z0-9]+(?:대?로|길)\s*\d+(?:-\d+)?', verified_addr
    )
    if not m_road:
        return verified_addr
    after_toks = verified_addr[m_road.end():].split()
    # 번지 뒤에 유닛이 아닌 토큰(=상호/건물명)이 이미 있으면 skip
    if after_toks and not all(_UNIT_TOKEN_RE.match(t) for t in after_toks):
        return verified_addr
    v_key = _road_key(verified_addr)
    try:
        results = _search_poi(m_road.group(0))
    except Exception:
        return verified_addr
    cands = []
    for pn, road in results:
        if not pn or _road_key(road) != v_key:
            continue
        if _BLD_STRONG_SUFFIX_RE.search(pn) and not _POI_TENANT_MARK_RE.search(pn):
            cands.append(pn)
    uniq = list(dict.fromkeys(cands))
    if len(uniq) != 1:
        return verified_addr  # 0개 or 모호(2+) → skip
    building = uniq[0]
    if building.replace(' ', '') in verified_addr.replace(' ', ''):
        return verified_addr
    # 번지 바로 뒤에 삽입 (동/층/호 앞)
    return (verified_addr[:m_road.end()] + f' {building}'
            + verified_addr[m_road.end():]).strip()


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
        # 2026-08-14 L-03650: '지하 1층'(공백)에서 '1층'만 캡처돼 '지하' 유실 → 지하 1층
        #   ≠ 1층(완전 다른 층, 오방문). 층 앞 '지하|지상' 접두 함께 캡처.
        # 층/호 여러 개(예 '지하2층. B207호') 모두 캡처 — 기존 search 는 첫 개만 잡아
        #   뒤 호수 유실 (2026-08-25 L-03772). finditer 로 순서대로 수집.
        _floor_re = re.compile(
            r'((?:(?:지하|지상)\s*)?[A-Za-z]?\d+(?:~\d+)?\s*(?:층|호|호실|관))(?![가-힣])')
        _fcands = [m.group(1).strip() for m in _floor_re.finditer(original_text)]
        _fcands = [c for c in _fcands if c and c not in verified_addr]
        if _fcands:
            cand = _fcands[0]
            # 원문이 '[층] [상호]' 순서(층이 상호 앞)면 그 순서 보존 (2026-08-01).
            #   '1층 피아노학원'(건물 1층에 입점한 한 층짜리 상호)을 '피아노학원 1층'
            #   (상호가 건물 통째·그 상호의 1층)으로 뒤집으면 의미가 완전히 달라짐.
            #   verified 끝 단어(상호, 숫자 아님)가 원문에서 층 바로 뒤에 오면 상호
            #   앞에 삽입해 어순 유지. 그 외(상호-층 순서·건물명 등)는 기존대로 맨 뒤.
            _vw = verified_addr.split()
            _last = _vw[-1] if _vw else ''
            if (_last and not re.search(r'\d', _last)
                    and re.search(rf'{re.escape(cand)}\s*{re.escape(_last)}',
                                  original_text)):
                verified_addr = ' '.join(_vw[:-1] + [cand, _last])
            else:
                verified_addr = f'{verified_addr} {cand}'
            # 나머지 층/호(B207호 등)는 순서대로 뒤에 부착
            for _c in _fcands[1:]:
                if _c not in verified_addr:
                    verified_addr = f'{verified_addr} {_c}'

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

    # 층↔건물명 어순 복원 (2026-08-03 L-03485): 카카오가 번지의 건물명을 미등록이면
    # verify 는 도로+번지만 주고, 층은 먼저(step 3) 붙고 건물명은 POI 로 맨 뒤에 붙어
    # '번지 2층 예전빌딩' 처럼 층-건물 역순이 됨. 원문이 '건물명 층'(건물-층) 순서였으면
    # 건물-층으로 복원. 원문이 층-건물('1층 피아노학원' 한 층 입점 상호)이면 유지
    # → 어순=의미 보존(9525637 사상을 POI 부착 경로까지 확장).
    _m_ord = re.search(
        r'(\d+(?:-\d+)?)\s+([A-Za-z]?\d+(?:~\d+)?(?:층|호|호실|관))\s+([가-힣][가-힣A-Za-z0-9]*)$',
        verified_addr,
    )
    if _m_ord:
        _fl, _bd = _m_ord.group(2), _m_ord.group(3)
        _bi, _fi = original_text.find(_bd), original_text.find(_fl)
        if 0 <= _bi < _fi:  # 원문에서 건물명이 층보다 앞 → 건물-층 순서 복원
            verified_addr = verified_addr[:_m_ord.start(2)] + f'{_bd} {_fl}'

    # 원문 끝 다단어 상호(상호+지점명) 보존 (2026-08-27 L-03811). verify 성공 경로는
    #   raw tail 을 버리고 POI 보강에만 의존하는데, step 1-b 상호 부착은 **한 단어**만
    #   잡아 '국면당 공세점'(상호+지점) 같은 다단어는 매치 실패 → POI 가 못 돌려주면
    #   유실(카카오 address.json 인덱싱 변동에 노출; verify=None 이면 _road_poi_fallback
    #   이 tail 보존). POI 이후, 원문 끝 '번지/호 + 다단어 상호'의 **상호(첫 단어)가 결과에
    #   전혀 없을 때만** 부착 → POI 가 이미 상호를 붙인 케이스(지점명 치환 등)는 중복 회피.
    _mshop2 = re.search(
        r'(?:[A-Za-z]?\d+(?:-\d+)?(?:호|층|번지|호실|관))\s+'
        r'([가-힣][가-힣A-Za-z0-9]{1,15}(?:\s+[가-힣][가-힣A-Za-z0-9]{1,15}){1,2})\s*$',
        original_text.split('\n')[-1].strip().rstrip('.'),
    )
    if _mshop2:
        _phrase = _mshop2.group(1).strip()
        _fw, _lw = _phrase.split()[0], _phrase.split()[-1]
        if (not re.search(r'(?:로|길|구|시|군|동|읍|면|층|호|번지)$', _lw)
                and _fw not in verified_addr
                and _phrase not in verified_addr):
            verified_addr = f'{verified_addr} {_phrase}'.strip()

    return verified_addr


# _STOP_WORDS 를 lead_helpers 재사용 (import 순환 방지 — 지연 import)
# 2026-07-24 L-03374 fix: '호|층' 제거 — '지호창호' 같이 상호가 '호' 로 끝나는 케이스가
#   후보에서 제외되던 버그. 아파트 부속 표기 (101호, 3층) 은 한글 시작 필터
#   ([가-힣][가-힣A-Za-z0-9]{1,14}) 로 이미 배제되므로 여기 유지 불필요.
_ADMIN_SUFFIX_RE = re.compile(r'(로|길|구|시|군|동|읍|면|리|번지|가|동로|번길)$')

# 시/도 이름(광역 행정구역) — 상호 후보에서 제외 (2026-09-04 L-03875). '서울'은 '시'로
#   안 끝나 _ADMIN_SUFFIX_RE 를 통과 → 순수 지번('서울 서초구 반포동 18-3')에서
#   유일한 상호 후보로 뽑혀 POI 검색 query('서울 서초구')가 'cand + 공백'으로 시작하는
#   엉뚱한 장소('서울 헌릉과 인릉')를 정확매치로 채택, 완전 다른 왕릉을 verified 로
#   반환하던 심각 오매칭. 시/도는 상호가 아니므로 후보에서 배제 (경기도·강원특별자치도
#   등 접미형 포함). fullmatch 라 '서울대입구'·'세종병원' 등 상호는 그대로 유지.
_SIDO_WORD_RE = re.compile(
    r'^(서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주)'
    r'(?:특별시|광역시|특별자치시|특별자치도|도)?$')


def _extract_region_hint(verified_addr: str) -> str:
    """verified 주소에서 지역 힌트 추출 (첫 시/군/구/광역시).

    2026-07-31 L-03254: 정규화 주소는 시 접미 없는 축약형(양주·시흥·화성)이 첫
    토큰이라 `시|군|구|도$` 매칭이 실패 → region 빈값 → POI 쿼리에 지역 없어
    전국 동명 상호가 나와 도로 불일치 → 지점명(송추점 등) 복원 실패. 구/동 우선,
    없으면(시+면/리 rural) 첫 토큰(시/도 축약형)으로 fallback.
    (road_key 일치가 append 를 gate 하므로 힌트 강화는 오탐 없이 매칭만 개선.)
    """
    words = verified_addr.split()
    for w in words:
        if re.search(r'(?:시|군|구|도)$', w):
            return w
    return words[0] if words else ''


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
        if _SIDO_WORD_RE.match(w):   # 시/도 이름은 상호 아님 (L-03875 헌인릉 오매칭 차단)
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
# 접미로만 판정하는 시설어 — substr 로 넣으면 오탐(글로비스⊃로비)이라 endswith 로.
#   '사무동로비'·'중앙로비' 등 로비/홀 부속공간 (2026-08-14 ETC-1765ea 사무동로비 오삽입).
_POI_FACILITY_SUFFIX = ('로비',)


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
        if any(wl.endswith(s) for s in _POI_FACILITY_SUFFIX):
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


# 서수 접두 '제'(제4공장·제2동) 정규화 — dedup 비교용 (2026-08-14 ETC-ad2710).
#   고객이 '4공장', 카카오 POI 가 '제4공장' 이면 같은 대상인데 dedup 이 다른 것으로 봐
#   '보우테이프 제4공장 4공장' 중복. 비교 시 '제'+숫자 접두를 제거해 동치 판정.
_JE_ORDINAL_RE = re.compile(r'제(?=\d)')


def _strip_je_ordinal(s: str) -> str:
    return _JE_ORDINAL_RE.sub('', s or '')


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
            # POI 정식명(지점명 포함)이 이미 verified 에 있으면 replace 시 지점명 이중
            #   부착 (2026-08-04 L-03451 '케이플라워마트 대화점 대화점', L-03486
            #   '힙춘향마라 계양점 계양점'). raw 끝에 이미 지점명이 있는데 _spaced 를
            #   _pname 으로 치환하면 뒤 지점명이 남아 중복 → 이미 있으면 skip (공백 무시).
            if _pn_ns in verified_addr.replace(' ', ''):
                continue
            return verified_addr.replace(_spaced, _pname, 1)

    # verified 에 이미 있는 후보는 우선순위 낮춤 (원문 신규 상호 먼저 시도)
    priority = (
        [c for c in candidates if c not in verified_addr]
        + [c for c in candidates if c in verified_addr]
    )
    for cand in priority[:5]:
        results = _search_poi(f'{cand} {region}'.strip())
        # cand 자체가 같은 도로의 정식 POI 로 실재하면 = 그 이름의 건물/장소가 존재.
        #   이때 'cand 로 끝나는' 접두 변형(PINS어반322, 슈퍼스타 어반322)은 같은
        #   건물의 다른 입점 업체 → endswith 로 치환/부착하면 건물명을 엉뚱한 테넌트로
        #   바꿈 (2026-08-04 L-03530 어반322 다세대 건물). 반대로 exact 가 없으면
        #   endswith 형(김포한강듀클래스)이 그 건물의 정식 전체명(접두 지역 포함)이라 허용.
        _has_exact = any(
            pn == cand and _road_key(rd) == v_key for pn, rd in results
        )
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
            #   2026-08-04 L-03530: (1) 접두가 '공백 분리'면 다른 업체('슈퍼스타 ')라 차단
            #     (붙은 접두 김포한강듀클래스만), (2) cand 가 exact POI 로 실재하면
            #     (어반322 건물) 접두형은 입점 업체이므로 endswith 자체를 차단.
            _pfx = (place_name[:-len(cand)]
                    if (len(cand) >= 4 and place_name.endswith(cand)) else None)
            # startswith('cand X'): cand 가 exact POI 로 실재(=건물/장소)하면 X 는 지점명이
            #   아니라 입점 업체(관악더행복마루 스크린파크골프장) → 지점(점/지점) 접미가
            #   아닌 한 append 금지 (2026-08-06 L-03541). exact 없으면 지점명 보강 유지
            #   (마성떡볶이 논현역점 — 단독 '마성떡볶이' POI 없어 _has_exact=False).
            _sw = place_name.startswith(cand + ' ')
            # 접두/접미 변형 채택 통합 규칙 (2026-08-06 L-03541 전수검사):
            #   • 매니저가 원문에 그 정식명을 썼으면(_pn_in_orig) 항상 유지
            #     (스포타임 엘타워·설원복지재단 안양의집 — 건물 exact 여도 원문 존중).
            #   • 아니면 cand 가 exact POI(건물)로 실재할 때만 차단 — 입점 업체를 지어
            #     붙이는 것(스크린파크골프장·롯데리아 서울랜드2호점·PINS어반322) 방지.
            #   • exact 없으면(마성떡볶이·한강듀클래스) 지점명·지역접두 보강 유지.
            #     endswith 는 공백 분리 접두(슈퍼스타 어반322=다른 업체) 제외(붙은 것만).
            _pn_in_orig = place_name.replace(' ', '') in original_text.replace(' ', '')
            _match = (
                place_name == cand
                or (_sw and (_pn_in_orig or not _has_exact))
                or (_pfx is not None and _pfx != ''
                    and (_pn_in_orig
                         or (not _has_exact and not _pfx.endswith(' '))))
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
                # 2026-08-14 ETC-ad2710: 서수 '제' 표기차 처리 — 고객 '4공장' 과 카카오
                #   공식명 '제4공장' 은 같은 공장. ①정확 일치면 유지 ②표기차(제 유무)만
                #   다르면 **공식명(제N공장)으로 승격**(중복 방지 + 만년로 제2공장 등과 표기
                #   통일) ③둘 다 없으면 지점명 부착.
                if len(place_words) >= 2 and all(w in verified_addr for w in place_words):
                    continue  # 상호+지점명 정확히 이미 있음
                if len(place_words) >= 2:
                    place_second = place_words[1]          # POI 공식 지점명 (제4공장)
                    _ps_je = _strip_je_ordinal(place_second)
                    verified_words = verified_addr.split()
                    already_has_branch = False
                    for i, w in enumerate(verified_words):
                        if w == cand and i + 1 < len(verified_words):
                            _vnext = verified_words[i + 1]
                            if _vnext == place_second:
                                already_has_branch = True
                                break
                            # 서수 '제' 표기차만 다르면 공식명으로 승격 (4공장→제4공장)
                            if _strip_je_ordinal(_vnext) == _ps_je:
                                return verified_addr.replace(_vnext, place_second, 1)
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
                    # 2026-08-06 ETC-f89e73: place_name 이 verified 에 '공백만 다른'
                    #   형태로 이미 있으면 재부착 시 중복 — verified '유니트 아이엔씨'
                    #   (매니저 띄어쓴 상호) + POI '유니트아이엔씨'(무공백 정식명) →
                    #   '아이엔씨'→'유니트아이엔씨' 치환이 '유니트 유니트아이엔씨' 중복.
                    #   무공백 비교로 이미 있으면 skip(같은 상호의 붙임/띄움 차이, 매니저
                    #   원형 유지). 한강듀클래스→김포한강듀클래스(접두 신규)는 무공백에도
                    #   없어 통과, JB↔제이비(독음)도 무공백 불일치라 아래 경계가드로 감.
                    if place_name.replace(' ', '').lower() \
                            in verified_addr.replace(' ', '').lower():
                        continue
                    # 2026-08-06 L-03597: cand 가 verified 에서 앞 토큰에 '붙어' 있으면
                    #   (예: 'JB미소빌딩' 안의 '미소빌딩') 이미 접두 수식된 건물명이라
                    #   POI 접두(제이비)를 재부착하면 'JB제이비미소빌딩' 중복 (JB↔제이비
                    #   = 영문/한글 독음 동일 건물). cand 가 단어 경계(문두·공백 뒤)로
                    #   '독립' 등장할 때만 접두 보강 — standalone '미소빌딩'→'제이비미소빌딩'
                    #   은 유지, glued 'JB미소빌딩' 은 원문 존중(skip). 한강듀클래스(공백
                    #   뒤 독립)→김포한강듀클래스 는 경계 매치라 정상 유지.
                    if not re.search(r'(?:^|\s)' + re.escape(cand), verified_addr):
                        continue
                    return verified_addr.replace(cand, place_name, 1)
                continue  # 완전 동일 → 무의미
            # POI 정식명이 공백 차이로 verified 에 이미 있으면 append 중복 방지
            #   (2026-08-06 ETC-3b713a: verified '행당제 1동 주민센터' + POI
            #   '행당제1동주민센터' → 중복. 무공백 비교로 skip).
            if place_name.replace(' ', '') in verified_addr.replace(' ', ''):
                continue
            # append 케이스: 원본에 법인 접두어 ((주)/㈜/주식회사) 있으면 유지
            # 2026-07-24 L-03372: 원본 '(주)아론' → POI place_name 은 '아론' 만 →
            #   append 시 접두어 소실. 원본 정보 보존을 위해 접두어 재부착.
            _m_prefix = re.search(
                r'((?:\(주\)|㈜|㈠|주식회사)\s*)' + re.escape(cand),
                original_text,
            )
            # 2026-08-06 ETC-de6041: 매니저가 상호 뒤 붙인 설명 접미(본사/본점/사옥/
            #   지사/본부)는 카카오 POI 에 없어 exact POI(place_name==cand) 부착 시
            #   유실됨 ('보우테이프 본사' → '보우테이프'). 공장/본사 구분 등 방문 맥락이라
            #   원문에 '상호+접미' 있으면 보존. 지점명(논현역점 등)은 카카오 POI 가
            #   이미 커버하므로 대상 외 (place_name==cand 인 exact 케이스로 한정 —
            #   place_name 에 이미 지점/공장명 있으면 접미 부착 안 함, 모순 방지).
            _sfx = ''
            if place_name == cand:
                _m_sfx = re.search(
                    re.escape(cand) + r'\s*(본사|본점|사옥|지사|본부)', original_text,
                )
                if _m_sfx:
                    _sfx = f' {_m_sfx.group(1)}'
            if _m_prefix:
                return f'{verified_addr} {_m_prefix.group(1)}{place_name}{_sfx}'.strip()
            return f'{verified_addr} {place_name}{_sfx}'.strip()
    return verified_addr


def _poi_fallback_by_gu(text: str, first_line: str) -> Optional[str]:
    """(B) 시/도는 있으나 카카오 주소 verify 가 실패한 케이스 POI 구제.

    2026-07-31 L-03473: 매니저가 지번(`인천 부평구 일신동 25`)으로 입력했으나 실제
    도로명은 `일신로 25`. 카카오 주소 API 는 이 지번을 0건 반환하지만, 건물명 POI
    (`송암노인요양원`)는 도로명(`인천 부평구 일신로 25`)을 돌려줌. 이때 지번의 '동'은
    도로명에 없으니(일신동≠일신로) **'구'가 지번↔도로명 공통 안정 단위** → 구로 검증.

    가드: 상호 후보가 POI place_name 정확 매치(or 'cand ' 로 시작) + POI 도로명에 구 포함.
    구가 없으면 힌트가 약해 오탐 위험 → skip.
    """
    # 입력에 도로명+번지가 이미 있으면 상호 POI 로 다른 지점 도로를 갈아치우지 않음
    #   (L-03679: '분당내곡로 131 … 포케올데이 판교점' → 상호검색이 서현역점 황새울로
    #   360번길 28 로 도로를 바꿔버림). startswith(cand+' ') 매칭이 느슨해 같은 브랜드
    #   다른 지점을 먼저 잡는 오매칭. 도로+번지가 있으면 _road_poi_fallback·Juso 가 그
    #   도로를 검증하도록 위임. (본 경로는 지번만 있고 도로명 없는 케이스용 — L-03473)
    if _ROAD_PATTERN.search(first_line):
        return None
    m_gu = re.search(r'([가-힣]{2,}구)', first_line)
    gu = m_gu.group(1) if m_gu else ''
    if not gu:
        return None
    candidates = _extract_shop_candidates(text)
    if not candidates:
        return None
    for cand in candidates[:5]:
        results = _kakao_search_poi(f'{cand} {gu}'.strip())
        if not results:
            continue
        for place_name, road_name in results:
            if not place_name or not road_name:
                continue
            if not (place_name == cand or place_name.startswith(cand + ' ')):
                continue
            if gu not in road_name.replace(' ', ''):
                continue
            return f'{normalize_display(road_name)} {cand}'
    return None


def _try_poi_fallback(text: str) -> Optional[str]:
    """카카오 verified 실패 케이스에서 POI(상호명) 검색으로 도로명 획득.

    조건 (2026-07-20 L-03292 최초 · 2026-07-21 L-03314 확장):
      - (A) 시/도 없음 (매니저가 시/도 빼먹은 케이스) — 아래 본 로직
      - (B) 시/도 있으나 지번↔도로명 불일치로 verify 실패 (2026-07-31 L-03473)
            → `_poi_fallback_by_gu` 로 위임 (구 단위 검증)
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
    # 입력에 도로명+번지가 이미 있으면 상호 POI 로 다른 지점 도로를 갈아치우지 않음
    #   (L-03644: '강남구 논현로159길 10 신사빌딩' → 서울 생략돼 A분기로 가서 다른
    #   신사빌딩 '언주로 817' 로 도로 변경). 도로+번지는 _road_poi_fallback·Juso 가
    #   검증하도록 위임. 본 함수는 도로 없이 상호만(또는 지번) 있는 케이스용.
    if _ROAD_PATTERN.search(first_line):
        return None
    # 시/도 있는데 여기 도달 = 카카오 주소 verify 실패 → (B) 구 기준 POI 구제로 위임.
    if re.search(
        r'(?:서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주|고양|성남|수원|용인|안양|안산|광명|시흥|화성|평택|김포)'
        r'(?:특별시|광역시|특별자치시|특별자치도|도|시)?',
        first_line,
    ):
        return _poi_fallback_by_gu(text, first_line)
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


# 방문 주소 필드에 한 줄로 붙여쓴 매니저 지시 노트 분리 (2026-08-04 L-03524).
#   기존 특이사항 이동(lead_sync)은 개행 분리만 처리 — 인라인('(현장은 3층) YG 소통
#   하세요')은 주소로 저장됨. 보수적 2-신호로 tail 지시문 + 노트성 괄호만 상담으로 이동.
_NOTE_VERB_END = re.compile(
    r'(하세요|해주세요|주세요|주십시오|하십시오|바랍니다|부탁드립니다|부탁드려요'
    r'|드립니다|드려요|요망|주시면|하시면|바람|주세용)(?:\s|$)'
)
_NOTE_TAIL_KW = ('소통', '연락', '요망', '참고', '주의')
# 시간 표현 (방문 스케줄 노트 감지용, 2026-08-06 ETC-4c47a2/1765ea)
_NOTE_TIME_RE = re.compile(r'오전|오후|새벽|정오|점심|저녁|밤|\d+\s*시(?![가-힣])')
_NOTE_PAREN_KW = _NOTE_TAIL_KW + (
    '현장', '확인', '문의', '주차', '위치',
    '정면', '오른쪽', '왼쪽', '좌측', '우측', '뒤편', '맞은편', '건너', '방향', '바라보',
    # 출입/도어 정보 (2026-08-14 L-03683): 괄호 안 '(비밀번호 7080)' 등은 주소가 아니라
    #   현장 출입 노트 → 상담 내용으로 분리. 괄호 내부에만 적용되고 주소성 괄호
    #   ('(101동)'·'(별관)')엔 안 나타나는 어휘라 오분리 위험 낮음.
    '비밀번호', '비번', '도어락', '현관', '출입', '호출',
)


def _note_has_signal(text: str) -> bool:
    """지시문 신호: 정중형 동사어미 or 독립 단어형 노트 키워드."""
    if _NOTE_VERB_END.search(text):
        return True
    return any(
        re.search(rf'(?:^|\s){re.escape(k)}(?:\s|$)', text) for k in _NOTE_TAIL_KW
    )


def _note_is_addr_token(t: str) -> bool:
    """주소 토큰이면 True (tail 노트 수집 정지 지점)."""
    if re.search(r'\d', t):                                    # 번지·층·호·동번호
        return True
    if re.search(r'(?:로|길)$', t):                            # 도로명
        return True
    if re.search(r'(?:빌딩|타워|아파트|상가|오피스텔|프라자|플라자|스퀘어|맨션|'
                 r'빌라|하우스|센터|파크|타운|시티|캐슬)$', t):  # 건물 접미
        return True
    return False


def split_address_notes(addr: str) -> Tuple[str, str]:
    """방문 주소에서 매니저 지시 노트(트레일링 지시문 + 노트성 괄호)를 분리.

    Returns (clean_addr, note_str). 보수적 — 노트 신호가 있을 때만 분리해 상호·주소
    오제거를 방지. 예: '한영빌딩 4층 (현장은 3층) YG 소통 하세요'
      → ('한영빌딩 4층', 'YG 소통 하세요 / 현장은 3층')
    """
    if not addr or not addr.strip():
        return addr, ''
    notes = []
    s = addr.strip()

    # 0) '/' 구분자 트레일링 노트 (2026-08-06 ETC-4c47a2·ETC-1765ea): 매니저가 주소란에
    #    '주소 / 오전 7시 현장설명 / 오후 1시 회수' 처럼 시간·지시를 '/' 로 붙임.
    #    기존엔 _extract_building_tail 이 '/' 뒤를 '잘라서' 변환 주소만 clean 하고
    #    노트는 카드 원본 아카이브에만 남아 상담/List/캔버스에서 유실됐음.
    #    첫 '/' 앞이 완결 주소(도로+번지 or 층/호/건물)이고 첫 노트 조각에 시간/지시
    #    신호가 있을 때만 뒤 전체를 노트로 이동 — '301호/302호' 같은 주소 연속은 미분리.
    if '/' in s:
        _head, _, _rest = s.partition('/')
        _head, _rest = _head.strip(), _rest.strip()
        _first_seg = _rest.split('/', 1)[0].strip()
        _head_is_addr = re.search(
            r'(?:로|길)\s*\d|\d+\s*(?:호|층|동|번지)|'
            r'(?:빌딩|타워|아파트|상가|오피스텔|프라자|플라자|스퀘어|맨션|빌라|'
            r'하우스|센터|시티|백화점|마트|병원|학교|공장)',
            _head,
        )
        _first_is_note = bool(_NOTE_TIME_RE.search(_first_seg)
                              or _note_has_signal(_first_seg))
        if _rest and _head_is_addr and _first_is_note:
            for _p in _rest.split('/'):
                _p = _p.strip()
                if _p:
                    notes.append(_p)
            s = _head

    # 1) 노트성 괄호 — 노트 신호(동사어미/키워드/방향어) 포함 시 이동, 주소성 괄호는 유지
    def _repl(m):
        inner = m.group(1).strip()
        if inner and (_note_has_signal(inner)
                      or any(k in inner for k in _NOTE_PAREN_KW)):
            notes.append(inner)
            return ' '
        return m.group(0)

    s = re.sub(r'\(([^)]*)\)', _repl, s)
    s = re.sub(r'\s+', ' ', s).strip()

    # 2) 트레일링 지시 클로즈 — 뒤에서부터 주소 토큰 전까지 모아 노트 신호 있으면 분리
    toks = s.split()
    cut = len(toks)
    for i in range(len(toks) - 1, -1, -1):
        if _note_is_addr_token(toks[i]):
            break
        cut = i
    tail = toks[cut:]
    if tail and _note_has_signal(' '.join(tail)):
        notes.insert(0, ' '.join(tail))
        s = ' '.join(toks[:cut]).strip()

    return s, ' / '.join(n for n in notes if n)


def _mark_planned(addr: str) -> str:
    """계획 중(미개업) 장소 표시 통일 — 'X예정지'/'X예정'(명사에 붙은) → 'X (예정)'.

    2026-08-06 L-03600 (사용자 요청): 매니저가 '중식당예정지'/'중식당예정'처럼
    아직 개업 전 장소를 상호 자리에 적으면 '중식당 (예정)' 으로 표기 통일.
      - '예정지'(장소 접미): 앞 명사에 붙든 띄든 → '(예정)'.
      - '예정'(지 없음): 앞 명사에 '붙은' 것만. '설치 예정'·'방문 예정' 같은
        동사구는 _TAIL_STOP_WORDS(구 '예정')가 이미 제거하므로 여기 안 남음
        (남더라도 공백 앞이라 미매치 → 영향 없음).
      - 재부착(_enrich m_shop 등)으로 생긴 'X X (예정)' 중복은 축약.

    ※ resolve_address 최종 단계에서만 호출 — 중간 함수(_flatten_paren_tail·_compose
      괄호 제거)가 '(예정)' 괄호를 벗기거나 재부착하는 것을 회피하기 위함.
    """
    if not addr or '예정' not in addr:
        return addr
    addr = re.sub(r'([가-힣]{2,})\s*예정지(?=\s|$)', r'\1 (예정)', addr)
    addr = re.sub(r'([가-힣]{2,})예정(?=\s|$)', r'\1 (예정)', addr)
    # 이미 괄호 친 '한의원(예정)'(공백 없이 붙은) → '한의원 (예정)' (L-03680). 고객·매니저가
    #   상호에 붙여 쓴 케이스. 앞이 한글일 때만 (숫자/공백은 정상 → 무변).
    addr = re.sub(r'([가-힣])\(예정\)', r'\1 (예정)', addr)
    # 파이프라인 재부착이 만든 'X X (예정)' 중복 축약 → 'X (예정)' (L-03600 실측)
    addr = re.sub(r'([가-힣]{2,})\s+\1\s+\(예정\)', r'\1 (예정)', addr)
    return re.sub(r'\s+', ' ', addr).strip()


def _jibun_road_fallback(text: str) -> Optional[str]:
    """순수 지번(도로명·건물 없이 '동 번지')을 카카오 keyword 로 도로명 구제 (L-03669).

    카카오 주소검색(address.json)은 순수 지번을 0건 반환하지만 keyword.json 은 그
    지번의 상호를 반환한다. POI 의 jibun 이 입력 지번과 '정확 일치'(같은 필지)하면
    그 POI 의 road_address 를 채택. 붙여쓴 지번('인계동1034-6번지2층')도 파싱.
    ※ 결과는 [추정] 유지용(비-verified level) — 근처 상호 기반 추론이라 매니저 확인 필요.
    """
    if not text:
        return None
    first = text.strip().split('\n', 1)[0]
    m = re.search(r'([가-힣]+동)\s*(\d+(?:-\d+)?)', first)
    if not m:
        return None
    dong, beonji = m.group(1), m.group(2)
    # 지번 뒤에 건물명/상호가 있으면(원일테크노2 등) 이 fallback 대상 아님 —
    #   순수 지번(+층/호)만. 건물 있으면 기존 regex/POI 경로가 처리(건물·호수 보존).
    #   L-03278 회귀 방지: '오정동 810-1 원일테크노2 4층 402호' 는 건물 있어 skip.
    _after = re.sub(r'번지', ' ', first[m.end():])
    _unit = re.compile(
        r'^(?:[A-Za-z]?\d+(?:-\d+)?(?:동|층|호|호실|관)?|[A-Za-z]?\d+[Ff]|'
        r'지하\d*층?|[A-Z]동)$'
    )
    if any(t and not _unit.match(t) for t in _after.split()):
        return None
    m_gu = re.search(r'([가-힣]{2,}구)', first)
    gu = m_gu.group(1) if m_gu else ''
    query = ' '.join(x for x in (gu, dong, beonji) if x)
    try:
        url = _KAKAO_POI_ENDPOINT + '?' + urllib.parse.urlencode(
            {'query': query, 'size': 5})
        data = _kakao_get_json(url)
    except _KakaoTransientError:
        return None
    if not data:
        return None
    # POI jibun 이 '…동 번지' 로 끝나야 = 같은 필지 (정확 일치 가드)
    _end = re.compile(re.escape(dong) + r'\s+' + re.escape(beonji) + r'$')
    for d in data.get('documents', []) or []:
        jibun = d.get('address_name') or ''
        road = d.get('road_address_name') or ''
        if road and _end.search(jibun):
            return normalize_display(road)
    return None


def _road_poi_fallback(text: str) -> Optional[str]:
    """도로명+번지가 카카오 주소검색(address.json) 0건이지만 POI 로 실재 확인되는
    케이스 구제 (2026-08-14 L-03667). 시골·고번지 도로명이 address.json 미인덱싱이면
    verify 실패 → raw/[확인필요]. POI(keyword)는 그 도로에 상호를 반환하고 그
    road_address 가 입력 도로명+번지와 '정확 일치'하면 도로명 주소를 채택.
    사용자 결정: verified (배지 제거) — 도로명+번지 자체가 실재 확인됨.
    """
    if not text:
        return None
    first = re.sub(r'\s+', ' ', text.strip().split('\n', 1)[0])
    m = _ROAD_PATTERN.search(first)
    if not m:
        return None
    road_beonji = m.group(1).strip()
    query = first[:m.end()].strip()        # 지역 포함 쿼리 (도시 모호 방지, L-03659)
    try:
        results = _kakao_search_poi_cached(query)
    except _KakaoTransientError:
        return None
    road_key = road_beonji.replace(' ', '')
    # 입력 지역 토큰(광적면 등) — POI road 에 있어야 (다른 도시 동명 도로 오탐 방지)
    region_toks = re.findall(r'[가-힣]{2,}(?:시|군|구|읍|면|동)', first)
    for pn, road in results:
        if not road:
            continue
        if road.replace(' ', '').endswith(road_key) and (
                not region_toks
                or any(rt in road.replace(' ', '') for rt in region_toks)):
            tail = first[m.end():].strip()
            base = normalize_display(road)
            return f'{base} {tail}'.strip() if tail else base
    return None


# 동+지번 코어. '…가' 법정동은 도로명형(종로1가·을지로3가·충정로3가류)이라 이름과
#   '가' 사이에 숫자가 낀다(충정로+3+가) → \d* 로 그 숫자를 흡수. 카카오 미인덱싱
#   필지를 juso 로 verify 하는 데 필요(L-03779 후속). 카카오 0건→여기 도달하므로
#   '충정로'(도로명) 오매칭 위험 없음(_ROAD_PATTERN 이 이미 None 인 케이스만 처리).
_JUSO_JIBUN_CORE_RE = re.compile(r'([가-힣]+\d*(?:동|리|가)\s*\d+(?:-\d+)?)')
_JUSO_REGION_RE = re.compile(
    r'((?:서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|'
    r'전북|전남|경북|경남|제주)(?:특별시|광역시|특별자치시|특별자치도|도)?'
    r'(?:\s+[가-힣]+(?:시|군|구))*)')


def _juso_fallback(text: str, regex_addr: Optional[str]) -> Optional[Tuple[str, str]]:
    """행안부 도로명주소 검색으로 코어 주소가 실재하는지 확인 (L-03671).

    카카오 address.json·POI 가 0건인데 실재하는 주소(백제고분로19길 13=잠실동 237-5,
    카카오 미인덱싱)를 정부 공식 DB 로 확인. 도로명+번지 또는 동+지번 '코어'를 추출해
    지역 포함 쿼리로 조회, 결과 roadAddr/jibunAddr 이 코어와 **경계 정확일치**(번지 뒤
    숫자·'-' 없음)할 때만 매치 → 퍼지·인접번지(876 vs 876-1) 오매칭 방지.

    Returns: (도로명 base, kind) — kind ∈ {'road','jibun'}. 매치 없으면 None.
      호출부는 road 케이스에 대해 **regex 문자열은 그대로 두고 level 만 verified 승격**
      (건물명·호수 유실 방지). base(정규화 도로명+bdNm)는 향후 jibun 치환용.
    """
    if not _juso_key():
        return None
    src = (regex_addr or '').strip() or (text or '').split('\n', 1)[0].strip()
    if not src:
        return None
    # 코어 주소 추출 (도로명+번지 우선, 없으면 동+지번)
    m_road = _ROAD_PATTERN.search(src) or _ROAD_PATTERN.search(text or '')
    if m_road:
        core, kind = m_road.group(1).strip(), 'road'
    else:
        m_jib = _JUSO_JIBUN_CORE_RE.search(src) or _JUSO_JIBUN_CORE_RE.search(text or '')
        if not m_jib:
            return None
        core, kind = m_jib.group(1).strip(), 'jibun'
    # 지역 접두 (도시 모호성·동명 도로 방지)
    region = ''
    m_reg = _JUSO_REGION_RE.match(src) or _JUSO_REGION_RE.match((text or '').strip())
    if m_reg:
        region = m_reg.group(1).strip()
    core_ns = core.replace(' ', '')
    # 지역 토큰 — 시/도 프리픽스가 없어도(‘송파구 …’) 코어 앞 시/군/구 를 잡음.
    #   쿼리 정밀도 + 다른 도시 동명 도로/지번 오매칭 가드 양쪽에 사용.
    _pre = src.split(core, 1)[0] if core in src else src
    _reg_toks = re.findall(r'[가-힣]{2,}(?:시|군|구)', region or _pre)
    if region and region.replace(' ', '') not in core_ns:
        query = f'{region} {core}'.strip()
    elif _reg_toks and _reg_toks[0] not in core_ns:
        query = f'{_reg_toks[0]} {core}'
    else:
        query = core
    results = _juso_search_cached(query)
    if not results and query != core:
        results = _juso_search_cached(core)
    if not results:
        return None
    _bound = re.compile(re.escape(core_ns) + r'(?![\d-])')
    # 도로/동 '이름' 토큰 (번지 제외) — 시작 경계 검증용. '판교로 393' → '판교로'.
    #   _bound(nospace)만으론 '판교로393' 이 '대왕판교로393' 안에 부분문자열로 매칭돼
    #   완전 다른 도로를 verified 로 오승격(L-03686). 이름이 결과의 **온전한 공백 토큰**
    #   으로 존재할 때만 통과 → '판교로' ⊄ '대왕판교로' 거절. '백제고분로19길'은 통과.
    _core_name = re.sub(r'\s*\d+(-\d+)?$', '', core).strip()
    # 경계+지역 통과 후보 수집 (같은 지번이 여러 도로에 걸친 경우 disambig 위해)
    _valid = []
    for road_addr, jibun_addr, bd in results:
        road_clean = re.sub(r'\s*\([^)]*\)\s*$', '', road_addr).strip()  # (법정동) 제거
        jibun_clean = re.sub(r'\s*\([^)]*\)\s*$', '', jibun_addr).strip()
        rk = road_clean.replace(' ', '')
        jk = jibun_clean.replace(' ', '')
        hit = ((kind == 'road' and _bound.search(rk))
               or (kind == 'jibun' and _bound.search(jk)))
        if not hit:
            continue
        # 시작 경계: 도로/동 이름이 부분문자열이 아니라 온전한 토큰이어야 함
        _clean_toks = (road_clean if kind == 'road' else jibun_clean).split()
        if _core_name and _core_name not in _clean_toks:
            continue
        # 지역 토큰 교차확인 (다른 도시 동명 도로/지번 방지)
        if _reg_toks and not any(rt in rk for rt in _reg_toks):
            continue
        _valid.append((road_clean, bd))
    if not _valid:
        return None
    # 같은 지번이 여러 도로에 걸침(L-03278 오정동 810-1 = 원일테크노Ⅱ/489번길 vs
    #   511번길) → 입력에 건물명이 있으면 그 bdNm 이 일치하는 결과 우선(Juso 반환 순서
    #   비의존). 로마숫자↔아라비아(원일테크노Ⅱ↔원일테크노2) 등가 판정.
    _src_key = _roman_to_arabic((text or '').replace(' ', '').lower())
    _chosen = None
    for road_clean, bd in _valid:
        _bk = _roman_to_arabic((bd or '').replace(' ', '').lower())
        if _bk and _bk in _src_key:
            _chosen = (road_clean, bd)
            break
    if _chosen is None:
        _chosen = _valid[0]
    road_clean, bd = _chosen
    base = normalize_display(road_clean)
    # Juso 건물명(bdNm) — 강한 건물 접미만 부착 (입주사/모호 접미 제외).
    #   이후 _enrich_verified_address 가 원문 층/호 tail 을 dedup 부착.
    if (bd and _BLD_STRONG_SUFFIX_RE.search(bd)
            and bd.replace(' ', '') not in base.replace(' ', '')):
        base = f'{base} {bd}'
    return (base, kind)


# 도로명 번호-길 공백 조인 (2026-08-14 L-03650): '언주로 107 길 27' → '언주로107길 27'.
#   고객이 '언주로107길'을 '언주로 107 길'로 띄어써 '언주로 107'(번지 107, 완전 다른
#   위치=현대2차아파트)로 오파싱되던 심각 버그. 길 뒤가 공백/끝일 때만 조인('길동'
#   같은 동명 오조인 방지). '번길'도 포함.
_ROAD_GIL_SPACE_RE = re.compile(r'([가-힣]+(?:대?로))\s+(\d+)\s+(번?길)(?=\s|$)')


def _join_road_gil(s: Optional[str]) -> Optional[str]:
    return _ROAD_GIL_SPACE_RE.sub(r'\1\2\3', s) if s else s


# 거래처 약칭 레지스트리 (2026-08-14 ETC-4feb23) — 자주 방문하는 주 거래처를 줄여
#   부르는 약칭('알만'=알만에이엠 고양점)의 **주소를 직접 등록**. 상호만 있고 위치
#   힌트(구/동/도로) 없는 케이스는 resolver 가 안전하게 POI 검색을 못 하므로(보수적),
#   등록된 거래처는 주소를 lookup 으로 확정. Redis hash `addr:partner_alias`
#   (field=약칭, value=정식 주소). 등록된 거래처로 한정 → 오탐 0. 60초 캐시·fails-open.
_PARTNER_ALIAS = {'data': {}, 'ts': 0.0}
_PARTNER_ALIAS_KEY = 'addr:partner_alias'


def _partner_alias_map() -> dict:
    import time as _t
    now = _t.time()
    if _PARTNER_ALIAS['data'] and now - _PARTNER_ALIAS['ts'] < 60:
        return _PARTNER_ALIAS['data']
    try:
        from dashboard.utils.redis_client import get_redis_client
        raw = get_redis_client().redis.hgetall(_PARTNER_ALIAS_KEY) or {}
        _PARTNER_ALIAS['data'] = {k: v for k, v in raw.items() if k and v}
        _PARTNER_ALIAS['ts'] = now
    except Exception:
        pass  # 장애 시 직전 캐시(없으면 빈 dict) 유지
    return _PARTNER_ALIAS['data']


def _partner_alias_lookup(text: Optional[str]) -> Optional[str]:
    """등록된 거래처 약칭이 입력에 있으면 등록 주소 반환 ('알만 고양점' → 등록주소).

    가장 긴 약칭 우선(‘알만 고양점’ > ‘알만’). 한글/영숫자 경계 매치(‘알만두’ 오탐
    방지). 입력의 층/호는 등록 주소 뒤에 부착. 매치 없으면 None.
    """
    if not text:
        return None
    m = _partner_alias_map()
    if not m:
        return None
    first = text.split('\n', 1)[0]
    for abbr in sorted(m, key=len, reverse=True):
        if not abbr:
            continue
        if re.search(r'(?<![가-힣A-Za-z0-9])' + re.escape(abbr)
                     + r'(?![가-힣A-Za-z0-9])', first):
            base = m[abbr].strip()
            _fl = re.search(
                r'((?:지하|지상)?\s*[A-Za-z]?\d+(?:~\d+)?\s*(?:층|호|호실|관|[Ff]))'
                r'(?![가-힣])', text)
            if _fl and _fl.group(1).replace(' ', '') not in base.replace(' ', ''):
                return f'{base} {_fl.group(1).strip()}'
            return base
    return None


# 시/군/구가 하위 행정구역·도로명에 붙은 글루(수원시매영로·수원시영통구매영로)에 공백 삽입
#   (2026-08-29 L-03839). 후보 생성 전용 — 실제 채택은 resolve_address 가 재귀 verify 로 게이팅.
#   [가-힣]{2,}? 비탐욕 → 연속 글루(성남시분당구정자로)의 각 경계를 개별 분리. lookahead
#   어간 2자+ 요구 → '청로'·'민로'(1자+로) 미매치라 시청로·시민로 오분리 1차 차단.
_SI_ROAD_GLUE_RE = re.compile(r'([가-힣]{2,}?(?:시|군|구))(?=[가-힣]{2,}(?:시|군|구|로|길))')


def _unglue_si_before_road(s: Optional[str]) -> Optional[str]:
    """시/군/구 뒤 도로명·하위 행정구역 글루에 공백 삽입한 후보 문자열 반환 (순수함수)."""
    return _SI_ROAD_GLUE_RE.sub(r'\1 ', s) if s else s


# 옛 '군' → '시' 승격 지명 후보 (포천군→포천시·양주군→양주시 등, 2026-09-01 L-03863).
#   행정 토큰 'XX군'(뒤 공백/끝)만 — 실제 채택은 resolve_address 가 재귀 verify 로 게이팅
#   (아직 군인 가평군·양평군은 verify 실패로 원본 유지).
_GUN_TOKEN_RE = re.compile(r'([가-힣]{2,})군(?=\s|$)')


def _promote_gun_to_si(s: Optional[str]) -> Optional[str]:
    """'XX군' 행정 토큰을 'XX시'로 치환한 후보 문자열 반환 (순수함수)."""
    return _GUN_TOKEN_RE.sub(r'\1시', s) if s else s


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
    # 0. 거래처 약칭 레지스트리 (ETC-4feb23) — 등록된 주 거래처 약칭('알만')이면 등록
    #   주소를 직접 확정(verified). 상호만 있고 위치힌트 없어 POI 로 못 푸는 케이스 구제.
    _pa = _partner_alias_lookup(text)
    if _pa:
        return (_mark_planned(_post_normalize_display(_pa)), 'verified')
    # 도로명+번지 붙여쓰기 정규화 (L-03675/L-03678) — kakao verify 도 정규화된 text 로
    #   조회하도록 resolve 진입 시 적용(extract 와 대칭). '상도로 13길4'→'상도로13길 4'.
    text = _normalize_road_spacing(text)
    regex_addr = _normalize_road_spacing(regex_addr)
    # 도로명 번호-길 공백 조인 (L-03650) — '언주로 107 길 27' → '언주로107길 27'
    text = _join_road_gil(text)
    regex_addr = _join_road_gil(regex_addr)

    # 시/군/구가 도로명·하위 행정구역에 붙은 글루 분리 (수원시매영로 → 수원시 매영로,
    #   2026-08-29 L-03839). 도로명에 청/흥 등이 섞이면 정적 분리는 오분리 위험(시청로·
    #   시흥대로) → 공백 삽입본을 재귀 resolve 해 **verified 될 때만 채택**(2차 안전망).
    #   lookahead 어간 2자+ 요구로 청로·민로류(시청로·시민로)는 1차 미매치. sub 이 멱등이라
    #   재귀 내부에선 _alt==text → 재발동 없음(무한재귀 방지).
    _alt_text = _unglue_si_before_road(text)
    if _alt_text != text:
        _alt_regex = _unglue_si_before_road(regex_addr) if regex_addr else regex_addr
        _alt_res = resolve_address(_alt_text, _alt_regex, regex_level)
        if _alt_res[1] == 'verified':
            return _alt_res

    # 옛 '군' → '시' 승격 지명 반영 (포천군→포천시, 2026-09-01 L-03863). 승격 지명은
    #   카카오/행안부가 현재명(시)만 인식 → 옛 '군' 지번은 변환 실패. 'XX군'→'XX시' 후보가
    #   verified 될 때만 채택(아직 군인 가평군 등은 verify 실패 → 원본 유지). sub 멱등→무한재귀 X.
    _gun_text = _promote_gun_to_si(text)
    if _gun_text != text:
        _gun_regex = _promote_gun_to_si(regex_addr) if regex_addr else regex_addr
        _gun_res = resolve_address(_gun_text, _gun_regex, regex_level)
        if _gun_res[1] == 'verified':
            return _gun_res

    # 1. 카카오 검증 시도
    verified = verify_address(text, regex_addr)
    if verified:
        addr, level = verified
        addr = _enrich_verified_address(addr, text, regex_addr)
        # 도로명주소에 법정동 중복 제거 (주소검색기 유입, 2026-07-27 L-03400)
        addr = _strip_redundant_legal_dong(addr)
        addr = _strip_redundant_jibun(addr)  # 도로명+지번 중복 시 구 지번 제거 (L-03627)
        addr = _enrich_building_by_road(addr)  # 민숭 주소에 건물명 부착 (L-03633)
        # tail 부착 후 후처리 (인접 유사 단어 dedup 등, 2026-07-22 ETC-b626fb)
        addr = _post_normalize_display(addr)
        addr = _mark_planned(addr)  # 'X예정지/X예정' → 'X (예정)' (L-03600)
        return (addr, level)

    # 1b-juso. 순수 지번 → 행안부 도로명 verified + 건물명 (2026-08-14 L-03673). 카카오
    #   keyword 구제(아래 1c, [추정])보다 **권위 우선** — `여의도동 15-24` → `은행로 3
    #   익스콘벤처타워`(verified). _juso_fallback 은 도로가 있으면 kind='road'(step 2에서
    #   처리)라, 여기엔 **도로 없는 순수 지번(+번지)만** 도달 → 경계 정확일치·지역 가드로
    #   지저분한 입력 오매칭(L-03280) 방어. 카카오 verified 와 동일 enrichment 체인.
    #   ★ 2026-09-04 L-03875: POI 퍼지 폴백(1c-poi)보다 **먼저** 실행. 순수 지번은 행안부
    #     (정부 공식 DB)가 권위 소스 — POI 가 같은 구의 엉뚱한 장소('반포동 18-3' → 내곡동
    #     헌인릉)를 정확매치로 verified 반환하던 오방문 버그를, juso 우선으로 원천 차단.
    _juso_j = _juso_fallback(text, regex_addr)
    if _juso_j and _juso_j[1] == 'jibun':
        addr = _enrich_verified_address(_juso_j[0], text, regex_addr)
        addr = _strip_redundant_legal_dong(addr)
        addr = _strip_redundant_jibun(addr)
        addr = _enrich_building_by_road(addr)
        addr = _post_normalize_display(addr)
        addr = _mark_planned(addr)
        return (addr, 'verified')

    # 1c-poi. POI fallback (2026-07-20 L-03292) — 시/도 빠진 케이스 상호 → 도로명.
    #   행안부(1b-juso) 미매치 시에만 도달(순수 지번은 위에서 이미 처리) → 상호 힌트가
    #   실재하는 케이스만 POI 로 구제.
    poi_road = _try_poi_fallback(text)
    if poi_road:
        addr = _enrich_verified_address(poi_road, text, regex_addr)
        addr = _strip_redundant_legal_dong(addr)
        addr = _post_normalize_display(addr)
        addr = _mark_planned(addr)
        return (addr, 'verified')

    # 1c. 순수 지번 → keyword 도로명 구제 (2026-08-13 L-03669). 행안부(1b-juso) 미매치
    #   시 fallback. 카카오 주소검색은 순수 지번 0건이지만 keyword 는 그 지번 상호를 줌
    #   → jibun 정확일치 도로명 채택. [추정] 유지(level='jibun_poi') — 근처 상호 기반.
    _jibun_road = _jibun_road_fallback(text)
    if _jibun_road:
        _floor = re.search(r'[A-Za-z]?\d+\s*(?:층|호|호실|관)', text or '')
        _addr = f'{_jibun_road} {_floor.group(0)}'.strip() if _floor else _jibun_road
        return (_mark_planned(_post_normalize_display(_addr)), 'jibun_poi')

    # 1d. 도로명+번지 POI 구제 (2026-08-14 L-03667). address.json 0건(시골·고번지
    #   미인덱싱)이지만 POI road_address 가 정확 일치하면 도로명 채택. 사용자 결정:
    #   verified(배지 제거) — 도로명+번지 자체가 실재 확인됨.
    _road = _road_poi_fallback(text)
    if _road:
        # 도로명만 확인된 민숭 주소에 POI 건물명 부착 (L-03650 다슬빌딩) — verified 경로와
        #   동일하게 강한 접미·입주사 마커 없는 건물 POI 가 유일할 때만.
        _road = _enrich_building_by_road(_road)
        return (_mark_planned(_post_normalize_display(_road)), 'verified')

    # 2. 정규식 결과 (시도 prefix 정규화 적용 + 상호명 보강)
    if regex_addr:
        addr = normalize_display(regex_addr)
        addr = _enrich_verified_address(addr, text, regex_addr)
        addr = _mark_planned(addr)
        _lv = regex_level or 'regex'
        # 행안부 도로명주소 검증 (2026-08-14 L-03671) — 카카오가 미인덱싱한 실재 도로+번지
        #   (백제고분로19길 13 등)를 정부 공식 DB 로 확인되면 verified 승격. **문자열은
        #   regex enriched 그대로 유지**(건물명·호수 유실 0) — level 만 상향해 [주소 확인
        #   필요]/[추정] 배지 제거. 도로 케이스만(jibun 은 상호 추론 여지 있어 보류).
        if _lv != 'verified':
            _juso_hit = _juso_fallback(text, regex_addr)
            if _juso_hit and _juso_hit[1] == 'road':
                _lv = 'verified'
        return (addr, _lv)

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
            # 2026-08-04 L-03398: raw fallback 도 시/도 정규화 적용 (verified·regex 경로만
            #   normalize_display 를 거쳐, verify 실패한 서울/광역시 주소는 '서울특별시'가
            #   그대로 남던 갭). 시/도 접두만 정리 — 건물·동·번지·호수는 보존.
            _raw = normalize_display(first_line)
            # 행안부 도로확인 (2026-08-18 L-03709): 시/도 없이 '김포…' 로 시작해
            #   extract=None → step2(regex) 를 못 타는 실재 도로+번지(카카오 미인덱싱
            #   신축 등)를 정부 공식 DB 로 verified 승격. 문자열은 그대로(정보 유실 0),
            #   level 만 상향해 오탐 [주소 확인 필요] 제거. 도로 케이스만(경계+토큰 정확
            #   일치라 없는 도로/퍼지는 미승격 — 판교로 393 등).
            _juso_hit3 = _juso_fallback(text, None)
            _lv3 = 'verified' if (_juso_hit3 and _juso_hit3[1] == 'road') else 'raw'
            return (_raw, _lv3)

    return ('', '')


# ─────────────────────────────────────────────────────────────
# 방문 모달 / PM 새 프로젝트 등록 공용 주소 정규화 (2026-09-03)
#   slack_bot._normalize_visit_address_if_verified 와 동일 규칙을 서비스 레이어로 승격.
#   PM 은 resolve_address_api → 이 함수 사용. (슬랙은 후속으로 이 함수를 쓰도록 수렴 예정)
# ─────────────────────────────────────────────────────────────

# 시/구 교차확인용 정규식 (오매칭 방어 — 여러 도시 공유 도로명에서 시/구가 바뀌면 경고)
_REGION_SIDO_RE = re.compile(
    r'\s*(서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|'
    r'전북|전남|경북|경남|제주)'
)
_REGION_GU_RE = re.compile(r'([가-힣]{2,}(?:구|군))(?:\s|$)')


def _addr_region_sig(addr: str) -> tuple:
    """주소에서 (시/도, 첫 구/군) 시그니처 추출 — 지역 교차확인용."""
    a = addr or ''
    m1 = _REGION_SIDO_RE.match(a)
    m2 = _REGION_GU_RE.search(a)
    return (m1.group(1) if m1 else '', m2.group(1) if m2 else '')


def _region_changed(raw: str, norm: str) -> bool:
    """입력(raw)과 정규화(norm)의 시/도·구가 '명시적으로' 다른지 (오매칭 감지, L-03659).
    한쪽이 시/구 정보 없으면(비교 불가) False(오탐 방지)."""
    rs, rg = _addr_region_sig(raw)
    ns, ng = _addr_region_sig(norm)
    if rg and ng and rg != ng:      # 구/군 양쪽 존재 + 다름 (부평구↔영등포구)
        return True
    if rs and ns and rs != ns:      # 시/도 양쪽 존재 + 다름 (인천↔서울)
        return True
    return False


def normalize_input_address(raw: str) -> dict:
    """방문 모달/PM 공용 주소 정규화 (slack _normalize_visit_address_if_verified 와 동일 파이프라인).

    extract_korean_address 로 주소 추출 → resolve_address(raw, 정규식주소, level) →
    레벨별 처리(verified 정정/동일 · jibun_poi [추정] · 미검증 raw 유지) + 시/구 교차확인.

    Returns dict:
      {
        'address': str,          # 저장할 주소 (verified/jibun_poi=정규화, 그 외=raw)
        'kind': str,             # 'normalized'|'same'|'estimated'|'failed'|'empty'
        'level': str,            # resolve_address 원본 level
        'region_changed': bool,  # 입력 대비 시/구 바뀜 (오매칭 의심 → ⚠️)
        'changed': bool,         # 원본과 값이 달라졌는지
      }
    """
    raw = (raw or '').strip()
    if not raw:
        return {'address': '', 'kind': 'empty', 'level': '', 'region_changed': False, 'changed': False}
    try:
        from dashboard.services.lead_helpers import extract_korean_address
        rx = extract_korean_address(raw)
        norm, lv = resolve_address(raw, rx[0] if rx else None, rx[1] if rx else '')
        if lv == 'verified' and norm:
            if norm != raw:
                return {'address': norm, 'kind': 'normalized', 'level': lv,
                        'region_changed': _region_changed(raw, norm), 'changed': True}
            return {'address': norm, 'kind': 'same', 'level': lv, 'region_changed': False, 'changed': False}
        if lv == 'jibun_poi' and norm and norm != raw:
            # 순수 지번 → keyword 도로명 구제 (근처 상호 기반, 매니저 재확인 유도 = [추정])
            return {'address': norm, 'kind': 'estimated', 'level': lv,
                    'region_changed': _region_changed(raw, norm), 'changed': True}
        # 미검증 — 도로명·번지가 카카오 미확인(오타 가능) → raw 유지
        return {'address': raw, 'kind': 'failed', 'level': lv or '', 'region_changed': False, 'changed': False}
    except Exception:
        return {'address': raw, 'kind': 'failed', 'level': '', 'region_changed': False, 'changed': False}
