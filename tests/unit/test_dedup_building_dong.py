# -*- coding: utf-8 -*-
"""법정동·건물명 중복 제거 회귀 테스트 (2026-08-04 L-03517·L-03422).

- L-03422: 카카오 법정동 '매산로3가'가 도로suffix 공백보강으로 '매산로 3가'로 분리
  저장돼 _strip_redundant_legal_dong 이 못 잡던 갭 → 공백변형도 매칭.
- L-03517: verify_address 가 카카오 건물명('서울랜드')을 붙이는데 원본 tail
  ('후문 서울랜드 산타레스토랑')에 이미 있어 '서울랜드 ... 서울랜드' 이중 노출 → tail 에
  이미 있으면 append skip.
_kakao_search 를 monkeypatch (네트워크 없이).
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services import address_resolver as ar


def test_strip_legal_dong_spaced(monkeypatch):
    """도로명과 겹치는 법정동(매산로3가)이 공백분리(매산로 3가)돼도 제거."""
    monkeypatch.setattr(ar, '_kakao_search',
                        lambda q: {'road_address': {'region_3depth_name': '매산로3가'}})
    assert ar._strip_redundant_legal_dong(
        '수원 팔달구 매산로 66 매산로 3가 4층') == '수원 팔달구 매산로 66 4층'
    assert ar._strip_redundant_legal_dong(
        '수원 팔달구 매산로 66 매산로3가 4층') == '수원 팔달구 매산로 66 4층'


def test_strip_legal_dong_preserves_building(monkeypatch):
    """건물명 안 법정동 부분매칭 보존 (회귀, 성수동2가 롯데캐슬)."""
    monkeypatch.setattr(ar, '_kakao_search',
                        lambda q: {'road_address': {'region_3depth_name': '성수동2가'}})
    assert ar._strip_redundant_legal_dong(
        '성동구 성수일로12길 52 성수동2가 롯데캐슬') == '성동구 성수일로12길 52 롯데캐슬'


def test_verify_no_duplicate_building(monkeypatch):
    """카카오 건물명이 원본 tail 에 이미 있으면 이중 부착 안 함 (서울랜드)."""
    monkeypatch.setattr(ar, '_kakao_key', lambda: 'test-key')  # 키 가드 통과
    monkeypatch.setattr(ar, '_kakao_search', lambda q: {
        'road_address': {
            'address_name': '경기 과천시 광명로 181',
            'building_name': '서울랜드',
            'region_3depth_name': '막계동',
        },
    })
    addr, lv = ar.verify_address('과천 광명로 181 후문 서울랜드 산타레스토랑')
    assert lv == 'verified'
    assert addr.count('서울랜드') == 1  # 이중 아님


def test_verify_no_duplicate_building_spaced(monkeypatch):
    """카카오 building_name 이 띄어쓴 형태(풍무 푸르지오)여도 무공백 비교로 tail 중복 감지 (L-03424).

    카카오 주소API building_name='풍무 푸르지오'(띄어쓴 비정식형) + 원본 tail
    '풍무푸르지오 아파트' → '풍무 푸르지오 풍무푸르지오 아파트' 이중. 무공백 비교로 skip.
    """
    monkeypatch.setattr(ar, '_kakao_key', lambda: 'test-key')
    monkeypatch.setattr(ar, '_kakao_search', lambda q: {
        'road_address': {
            'address_name': '경기 김포시 유현로 200',
            'building_name': '풍무 푸르지오',
            'region_3depth_name': '풍무동',
        },
    })
    addr, lv = ar.verify_address('김포 유현로 200 풍무푸르지오 아파트')
    assert lv == 'verified'
    assert '풍무 푸르지오 풍무푸르지오' not in addr
    assert addr.count('풍무푸르지오') == 1


def test_compose_no_split_admin_dong_facility(monkeypatch):
    """행정동 시설명(행당제1동주민센터)의 'N동'을 아파트 동으로 오인해 쪼개지 않음 (ETC-3b713a)."""
    monkeypatch.setattr(ar, '_kakao_key', lambda: 'k')
    monkeypatch.setattr(ar, '_kakao_search', lambda q: {
        'road_address': {
            'address_name': '서울 성동구 고산자로10길 18',
            'building_name': '행당제1동주민센터',
            'region_3depth_name': '행당동',
        },
    })
    addr, lv = ar.verify_address('성동구 고산자로10길 18 행당제1동주민센터 3층')
    assert lv == 'verified'
    assert '행당제1동주민센터' in addr
    assert '행당제 1동' not in addr  # 안 쪼개짐


def test_enrich_no_duplicate_spaced_building(monkeypatch):
    """POI 정식명이 공백 차이로 verified 에 이미 있으면 append 중복 방지 (ETC-3b713a B)."""
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('행당제1동주민센터', '서울 성동구 고산자로10길 18')])
    r = ar._enrich_with_poi('성동구 고산자로10길 18 행당제 1동 주민센터 3층',
                            '성동구 고산자로10길 18 행당제1동주민센터 3층')
    assert r.count('주민센터') == 1


def test_enrich_no_prefix_boost_when_glued(monkeypatch):
    """cand 가 verified 에서 앞 토큰에 붙어 있으면 접두 재부착 안 함 (L-03597).

    'JB미소빌딩'(=제이비미소빌딩, 영/한 독음 동일)에서 POI '제이비미소빌딩'이
    '미소빌딩' 접미로 매칭돼 '제이비' 재부착 → 'JB제이비미소빌딩' 중복되던 갭.
    """
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('제이비미소빌딩', '서울 강남구 논현로 841')])
    r = ar._enrich_with_poi('강남구 논현로 841 JB미소빌딩 2층 경희한의원',
                            '강남구 논현로 841 JB미소빌딩 2층 경희한의원')
    assert 'JB제이비미소빌딩' not in r
    assert 'JB미소빌딩' in r


def test_enrich_no_dup_on_spacing_variant(monkeypatch):
    """POI 무공백 정식명이 verified 에 공백형으로 이미 있으면 재부착 안 함 (ETC-f89e73).

    verified '유니트 아이엔씨'(매니저 띄어씀) + POI '유니트아이엔씨'(무공백) →
    '아이엔씨'→'유니트아이엔씨' 치환이 '유니트 유니트아이엔씨' 중복 만들던 갭.
    """
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('유니트아이엔씨', '서울 구로구 디지털로30길 28')])
    r = ar._enrich_with_poi('구로구 디지털로30길 28 마리오타워 1504호 유니트 아이엔씨',
                            '구로구 디지털로30길 28 마리오타워 1504호 유니트 아이엔씨 주식회사')
    assert '유니트 유니트아이엔씨' not in r
    assert r == '구로구 디지털로30길 28 마리오타워 1504호 유니트 아이엔씨'


def test_enrich_prefix_boost_when_standalone(monkeypatch):
    """standalone cand 는 접두 보강 유지 (회귀, 한강듀클래스→김포한강듀클래스 계열)."""
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('제이비미소빌딩', '서울 강남구 논현로 841')])
    r = ar._enrich_with_poi('강남구 논현로 841 미소빌딩 2층',
                            '강남구 논현로 841 미소빌딩 2층')
    assert '제이비미소빌딩' in r


def test_enrich_preserve_headoffice_suffix(monkeypatch):
    """exact POI 부착 시 매니저의 '본사' 접미 보존 (ETC-de6041).

    '보우테이프 본사' → POI exact '보우테이프' 부착하며 '본사' 유실되던 갭.
    """
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('보우테이프', '경기 화성시 만세구 향남읍 발안로 701-5')])
    r = ar._enrich_with_poi('화성 만세구 향남읍 발안로 701-5',
                            '화성 향남읍 발안로 701-5 보우테이프 본사')
    assert r.endswith('보우테이프 본사')


def test_enrich_suffix_only_on_exact(monkeypatch):
    """place_name != cand(지점/공장명 포함)이면 본사 접미 안 붙임 (모순 방지)."""
    # POI 가 '보우테이프 갈천공장'(지점형)만 반환 → cand '보우테이프' 와 다름
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('보우테이프 갈천공장', '경기 화성시 만세구 향남읍 발안로 701-5')])
    r = ar._enrich_with_poi('화성 만세구 향남읍 발안로 701-5',
                            '화성 향남읍 발안로 701-5 보우테이프 본사')
    assert '갈천공장 본사' not in r


def test_kakao_poi_html_unescape(monkeypatch):
    """카카오 POI place_name 의 HTML 엔티티(&amp;)를 수신 즉시 unescape (L-03583).

    카카오 keyword API 는 상호 '케이&케이베이스볼아카데미' 를 '케이&amp;케이…' 로
    escape 해 반환 → 주소에 &amp; 잔존 + line607 .upper() 가 amp→AMP 까지.
    """
    ar._kakao_search_poi_cached.cache_clear()
    monkeypatch.setattr(ar, '_kakao_get_json', lambda url: {
        'documents': [
            {'place_name': '케이&amp;케이베이스볼아카데미',
             'road_address_name': '인천 계양구 서운산업로 30'},
        ],
    })
    pois = ar._kakao_search_poi('케이케이베이스볼')
    assert pois[0][0] == '케이&케이베이스볼아카데미'
    assert '&amp;' not in pois[0][0]
    ar._kakao_search_poi_cached.cache_clear()


def test_kakao_search_html_unescape(monkeypatch):
    """카카오 주소검색 building_name 의 HTML 엔티티도 unescape."""
    ar._kakao_search_cached.cache_clear()
    monkeypatch.setattr(ar, '_kakao_get_json', lambda url: {
        'documents': [{
            'road_address': {
                'address_name': '인천 계양구 서운산업로 30',
                'building_name': 'B&amp;B타워',
            },
        }],
    })
    doc = ar._kakao_search('서운산업로 30')
    assert doc['road_address']['building_name'] == 'B&B타워'
    ar._kakao_search_cached.cache_clear()


def test_compose_no_split_kakao_building(monkeypatch):
    """카카오 확정 건물명(타워팰리스)을 스페이싱 규칙이 쪼개지 않음 (L-03553).

    카카오 building_name='타워팰리스'(정식 무공백)인데 규칙1이 '타워'+한글을
    분리해 '타워 팰리스'로 망치던 갭. base 에 붙어 실재하므로 유지.
    """
    monkeypatch.setattr(ar, '_kakao_key', lambda: 'k')
    monkeypatch.setattr(ar, '_kakao_search', lambda q: {
        'road_address': {
            'address_name': '서울 강남구 언주로30길 56',
            'building_name': '타워팰리스',
            'region_3depth_name': '도곡동',
        },
    })
    addr, lv = ar.verify_address('강남구 언주로30길 56 타워팰리스 제상가동 202호')
    assert lv == 'verified'
    assert '타워팰리스' in addr
    assert '타워 팰리스' not in addr  # 안 쪼개짐


def test_compose_still_splits_manager_typo(monkeypatch):
    """매니저가 붙여 쓴 '단지상가'(base 에 없음)는 기존대로 분리 (규칙1 보존)."""
    monkeypatch.setattr(ar, '_kakao_key', lambda: 'k')
    # 카카오는 도로+번지만 확정(건물명 없음), '단지상가'는 원본 tail
    monkeypatch.setattr(ar, '_kakao_search', lambda q: {
        'road_address': {
            'address_name': '서울 노원구 동일로 1234',
            'building_name': '',
            'region_3depth_name': '상계동',
        },
    })
    monkeypatch.setattr(ar, '_search_poi', lambda q: ())
    addr, lv = ar.verify_address('노원구 동일로 1234 상계주공아파트 단지상가')
    # base 에 '단지상가' 없음 → 여전히 분리
    assert '단지 상가' in addr


def test_strip_redundant_jibun(monkeypatch):
    """도로명+번지 있을 때 카카오 지번과 정확 일치하는 구 지번 제거 (L-03627)."""
    monkeypatch.setattr(ar, '_kakao_search', lambda q: {
        'road_address': {'address_name': '서울 중랑구 사가정로50길 51'},
        'address': {'main_address_no': '632', 'sub_address_no': '2'},
    })
    assert ar._strip_redundant_jibun('중랑구 사가정로50길 51 열방교회 632-2 1층') == \
        '중랑구 사가정로50길 51 열방교회 1층'


def test_strip_jibun_keeps_unit_number(monkeypatch):
    """호수(632-2호)는 유닛번호라 보존 — 지번과 숫자 같아도 접미로 구분."""
    monkeypatch.setattr(ar, '_kakao_search', lambda q: {
        'road_address': {'address_name': '서울 중랑구 사가정로50길 51'},
        'address': {'main_address_no': '632', 'sub_address_no': '2'},
    })
    r = ar._strip_redundant_jibun('중랑구 사가정로50길 51 열방빌딩 632-2호')
    assert '632-2호' in r


def test_strip_jibun_skips_when_equals_road_num(monkeypatch):
    """지번이 도로명 번지와 같으면(그게 번지) 보존."""
    monkeypatch.setattr(ar, '_kakao_search', lambda q: {
        'road_address': {'address_name': '서울 강남구 테헤란로 51'},
        'address': {'main_address_no': '51', 'sub_address_no': ''},
    })
    assert ar._strip_redundant_jibun('강남구 테헤란로 51 삼성빌딩') == \
        '강남구 테헤란로 51 삼성빌딩'


def test_enrich_building_by_road_bare(monkeypatch):
    """민숭 주소(번지+동/층/호)에 건물 접미 POI 1개면 부착 (L-03633)."""
    monkeypatch.setattr(ar, '_search_poi', lambda q: [
        ('현대테라타워DIMC', '경기 남양주시 다산지금로 202'),
        ('뽀로로테마파크', '경기 남양주시 다산지금로 202'),        # 접미 없음 → 무시
        ('GS25 다산테라타워점', '경기 남양주시 다산지금로 202'),   # 점 마커 → 무시
    ])
    r = ar._enrich_building_by_road('남양주 다산지금로 202 B동 5F 0001호')
    assert r == '남양주 다산지금로 202 현대테라타워DIMC B동 5F 0001호'


def test_enrich_building_skip_when_has_shop(monkeypatch):
    """번지 뒤 상호/건물 이미 있으면 skip (민숭 아님)."""
    monkeypatch.setattr(ar, '_search_poi', lambda q: [
        ('현대테라타워DIMC', '경기 남양주시 다산지금로 202')])
    v = '남양주 다산지금로 202 삼성빌딩 3층'
    assert ar._enrich_building_by_road(v) == v


def test_enrich_building_skip_when_ambiguous(monkeypatch):
    """건물 접미 POI 가 2개 이상이면 모호 → skip."""
    monkeypatch.setattr(ar, '_search_poi', lambda q: [
        ('A타워', '경기 남양주시 다산지금로 202'),
        ('B빌딩', '경기 남양주시 다산지금로 202')])
    v = '남양주 다산지금로 202 5층'
    assert ar._enrich_building_by_road(v) == v


def test_enrich_building_skip_tenant_only(monkeypatch):
    """건물 접미 POI 가 입주사 마커뿐이면 부착 안 함."""
    monkeypatch.setattr(ar, '_search_poi', lambda q: [
        ('스타벅스 판교타워점', '경기 성남시 분당구 판교역로 235'),
        ('올리브영 판교스퀘어점', '경기 성남시 분당구 판교역로 235')])
    v = '성남 분당구 판교역로 235 4층'
    assert ar._enrich_building_by_road(v) == v


def test_join_road_gil():
    """도로명 번호-길 공백 조인 — '언주로 107 길 27' → '언주로107길 27' (L-03650)."""
    J = ar._join_road_gil
    assert J('언주로 107 길 27 지하') == '언주로107길 27 지하'
    assert J('강남구 언주로 107 길 27') == '강남구 언주로107길 27'
    assert J('테헤란로 107 번길 16') == '테헤란로107번길 16'
    # 가드: '길동'(동명)·정상 번지·이미 붙은 것은 불변
    assert J('테헤란로 107 길동') == '테헤란로 107 길동'
    assert J('테헤란로 152 3층') == '테헤란로 152 3층'
    assert J('언주로107길 27') == '언주로107길 27'


def test_road_poi_fallback_exact(monkeypatch):
    """도로명+번지 address.json 0건이지만 POI road 정확일치 → 채택 (L-03667)."""
    monkeypatch.setattr(ar, '_kakao_search_poi_cached', lambda q: (
        ('한손커피', '경기 양주시 광적면 부흥로 876'),
        ('투타상사', '경기 양주시 광적면 부흥로 876'),
    ))
    r = ar._road_poi_fallback('양주 광적면 부흥로 876 1층')
    assert r == '양주 광적면 부흥로 876 1층'


def test_road_poi_fallback_region_guard(monkeypatch):
    """입력 지역 토큰(광적면)이 POI road 에 없으면 채택 안 함(다른 도시 동명 도로)."""
    monkeypatch.setattr(ar, '_kakao_search_poi_cached', lambda q: (
        ('딴가게', '서울 강남구 부흥로 876'),  # 광적면 아님
    ))
    assert ar._road_poi_fallback('양주 광적면 부흥로 876 1층') is None


def test_juso_fallback_road_exact(monkeypatch):
    """카카오 미인덱싱 도로+번지 → 행안부 경계 정확일치 → (base,'road') (L-03671)."""
    monkeypatch.setattr(ar, '_juso_key', lambda: 'K')
    monkeypatch.setattr(ar, '_juso_search_cached', lambda q: (
        ('서울특별시 송파구 백제고분로19길 13 (잠실동)',
         '서울특별시 송파구 잠실동 237-5', ''),
    ))
    assert ar._juso_fallback('송파구 백제고분로19길 13',
                             '송파구 백제고분로19길 13') == ('송파구 백제고분로19길 13', 'road')


def test_juso_fallback_beonji_boundary(monkeypatch):
    """인접 번지(876 vs 876-1) 경계 구분 — 876 입력에 876-1 채택 안 함."""
    monkeypatch.setattr(ar, '_juso_key', lambda: 'K')
    monkeypatch.setattr(ar, '_juso_search_cached', lambda q: (
        ('경기도 양주시 광적면 부흥로 876-1', '경기도 양주시 광적면 가납리 624-1', ''),
        ('경기도 양주시 광적면 부흥로 876', '경기도 양주시 광적면 가납리 627-2', ''),
    ))
    assert ar._juso_fallback('양주 광적면 부흥로 876',
                             '양주 광적면 부흥로 876') == ('양주 광적면 부흥로 876', 'road')


def test_juso_fallback_region_guard(monkeypatch):
    """지역 토큰(송파구) 불일치 결과는 채택 안 함(다른 도시 동명 도로)."""
    monkeypatch.setattr(ar, '_juso_key', lambda: 'K')
    monkeypatch.setattr(ar, '_juso_search_cached', lambda q: (
        ('부산광역시 사하구 백제고분로19길 13', '부산광역시 사하구 괴정동 1-1', ''),
    ))
    assert ar._juso_fallback('송파구 백제고분로19길 13',
                             '송파구 백제고분로19길 13') is None


def test_juso_fallback_no_key(monkeypatch):
    """키 없으면 네트워크 시도 없이 None."""
    monkeypatch.setattr(ar, '_juso_key', lambda: '')
    assert ar._juso_fallback('송파구 백제고분로19길 13',
                             '송파구 백제고분로19길 13') is None


def test_juso_fallback_jibun_bdnm(monkeypatch):
    """동+지번 입력 → jibun kind, 강한 접미 bdNm(익스콘벤처타워) base 부착."""
    monkeypatch.setattr(ar, '_juso_key', lambda: 'K')
    monkeypatch.setattr(ar, '_juso_search_cached', lambda q: (
        ('서울특별시 영등포구 은행로 3 (여의도동)',
         '서울특별시 영등포구 여의도동 15-24', '익스콘벤처타워'),
    ))
    base, kind = ar._juso_fallback('영등포구 여의도동 15-24', '영등포구 여의도동 15-24')
    assert kind == 'jibun'
    assert base == '영등포구 은행로 3 익스콘벤처타워'


def test_juso_fallback_jibun_multi_road_prefers_building(monkeypatch):
    """같은 지번이 여러 도로에 걸치면(오정동 810-1) 입력 건물명 매치 결과 우선 —
    Juso 반환 순서 비의존. 로마숫자↔아라비아 등가(원일테크노Ⅱ↔원일테크노2) (L-03278)."""
    monkeypatch.setattr(ar, '_juso_key', lambda: 'K')
    monkeypatch.setattr(ar, '_juso_search_cached', lambda q: (
        ('경기 부천시 오정구 신흥로511번길 13-39 (오정동)',
         '경기 부천시 오정구 오정동 810-1', ''),                      # 건물명 없는 게 첫 결과
        ('경기 부천시 오정구 신흥로489번길 56 (오정동)',
         '경기 부천시 오정구 오정동 810-1 원일테크노Ⅱ', '원일테크노Ⅱ'),
    ))
    base, kind = ar._juso_fallback('부천 오정구 오정동 810-1 원일테크노2 4층',
                                   '부천 오정구 오정동 810-1')
    assert kind == 'jibun'
    assert '신흥로489번길 56' in base   # 첫 결과(511번길) 아닌 건물명 매치 우선
    assert '원일테크노Ⅱ' in base


def test_resolve_jibun_juso_verified(monkeypatch):
    """순수 지번 → 행안부 도로명 verified + 건물명 (카카오 keyword [추정]보다 우선, L-03673)."""
    monkeypatch.setattr(ar, 'verify_address', lambda t, r=None: None)
    monkeypatch.setattr(ar, '_try_poi_fallback', lambda t: None)
    monkeypatch.setattr(ar, '_juso_key', lambda: 'K')
    monkeypatch.setattr(ar, '_juso_search_cached', lambda q: (
        ('서울특별시 영등포구 은행로 3 (여의도동)',
         '서울특별시 영등포구 여의도동 15-24 익스콘벤처타워', '익스콘벤처타워'),
    ))
    addr, lv = ar.resolve_address('영등포구 여의도동 15-24',
                                  '영등포구 여의도동 15-24', 'level7')
    assert lv == 'verified'   # jibun_poi([추정]) 아닌 verified
    assert addr == '영등포구 은행로 3 익스콘벤처타워'


def test_resolve_juso_upgrades_regex_to_verified(monkeypatch):
    """카카오·POI 전부 실패 + 행안부 도로 확인 → regex 문자열 유지 + verified 승격.

    핵심: 건물명·호수 유실 없이(문자열 불변) level 만 상향 → [주소 확인 필요] 배지 제거."""
    monkeypatch.setattr(ar, 'verify_address', lambda t, r=None: None)
    monkeypatch.setattr(ar, '_try_poi_fallback', lambda t: None)
    monkeypatch.setattr(ar, '_jibun_road_fallback', lambda t: None)
    monkeypatch.setattr(ar, '_road_poi_fallback', lambda t: None)
    monkeypatch.setattr(ar, '_juso_key', lambda: 'K')
    monkeypatch.setattr(ar, '_juso_search_cached', lambda q: (
        ('서울특별시 송파구 백제고분로19길 13 (잠실동)',
         '서울특별시 송파구 잠실동 237-5', ''),
    ))
    addr, lv = ar.resolve_address('송파구 백제고분로19길 13',
                                  '송파구 백제고분로19길 13', 'level3')
    assert lv == 'verified'
    assert addr == '송파구 백제고분로19길 13'


def test_resolve_juso_absent_keeps_regex_level(monkeypatch):
    """행안부도 못 찾으면 regex level 유지(가짜 도로 verified 위조 방지)."""
    monkeypatch.setattr(ar, 'verify_address', lambda t, r=None: None)
    monkeypatch.setattr(ar, '_try_poi_fallback', lambda t: None)
    monkeypatch.setattr(ar, '_jibun_road_fallback', lambda t: None)
    monkeypatch.setattr(ar, '_road_poi_fallback', lambda t: None)
    monkeypatch.setattr(ar, '_juso_key', lambda: 'K')
    monkeypatch.setattr(ar, '_juso_search_cached', lambda q: ())
    addr, lv = ar.resolve_address('없는길 99999', '없는길 99999', 'level3')
    assert lv == 'level3'


def test_road_poi_fallback_attaches_building(monkeypatch):
    """도로명+번지 POI 구제 결과에 건물 POI(다슬빌딩) 부착 (L-03650). address.json
    0건 도로명이 verified 채택돼도 강한 접미 건물 POI 가 유일하면 건물명 부착."""
    monkeypatch.setattr(ar, '_kakao_search_poi_cached', lambda q: (
        ('다슬빌딩', '서울 강남구 언주로107길 27'),
        ('썸에스테틱', '서울 강남구 언주로107길 27'),
    ))
    monkeypatch.setattr(ar, '_search_poi', lambda q: [
        ('다슬빌딩', '서울 강남구 언주로107길 27'),
        ('썸에스테틱', '서울 강남구 언주로107길 27'),
    ])
    road = ar._road_poi_fallback('강남구 언주로107길 27 지하')
    assert road == '강남구 언주로107길 27 지하'
    assert ar._enrich_building_by_road(road) == '강남구 언주로107길 27 다슬빌딩 지하'


def test_jibun_road_fallback_exact(monkeypatch):
    """순수 지번 → keyword POI 의 jibun 정확일치 도로명 채택 (L-03669)."""
    monkeypatch.setattr(ar, '_kakao_get_json', lambda url: {'documents': [
        {'place_name': '매일부동산', 'address_name': '경기 수원시 팔달구 인계동 1034-6',
         'road_address_name': '경기 수원시 팔달구 인계로108번길 27-23'},
    ]})
    assert ar._jibun_road_fallback('인계동1034-6번지2층') == '수원 팔달구 인계로108번길 27-23'


def test_jibun_road_fallback_skip_with_building(monkeypatch):
    """지번 뒤 건물명(원일테크노2) 있으면 skip — 건물·호수 보존 (L-03278 회귀방지)."""
    monkeypatch.setattr(ar, '_kakao_get_json', lambda url: {'documents': [
        {'address_name': '경기 부천시 오정구 오정동 810-1',
         'road_address_name': '경기 부천시 오정구 신흥로511번길 13-39'},
    ]})
    assert ar._jibun_road_fallback('경기도 부천시 오정동 810-1 원일테크노2 4층 402호') is None


def test_jibun_road_fallback_no_exact_match(monkeypatch):
    """POI jibun 이 입력 지번과 다르면(다른 필지) 채택 안 함."""
    monkeypatch.setattr(ar, '_kakao_get_json', lambda url: {'documents': [
        {'address_name': '경기 수원시 팔달구 인계동 999-9',  # 번지 다름
         'road_address_name': '경기 수원시 팔달구 다른로 1'},
    ]})
    assert ar._jibun_road_fallback('인계동 1034-6') is None


def test_flatten_paren_keeps_corporate_designator():
    """법인 표기 (주)/(유)/(사)/(재)는 flatten 안 함 — '주' 홀로 남기 방지 (L-03358)."""
    F = ar._flatten_paren_tail
    assert F('에스티팜 시화공장(주)') == '에스티팜 시화공장(주)'
    assert F('(주)가연결혼정보') == '(주)가연결혼정보'
    # 회귀: 층/지번 괄호 flatten 은 그대로
    assert F('타임빌딩(2층)') == '타임빌딩 2층'
    assert F('(가산동, 이앤씨드림타워7차)') == '이앤씨드림타워7차'


def test_flatten_paren_keeps_hosu_comma():
    """괄호 안 호수 나열 콤마 보존 (ETC-45c37b, '숫자 사이 콤마' 규칙과 일관)."""
    F = ar._flatten_paren_tail
    assert F('SG타워 3F (301호,302호)') == 'SG타워 3F 301호,302호'
    assert F('SG타워 (305,306,408호)') == 'SG타워 305,306,408호'
    # 콤마+공백형도 콤마 유지 (기존 '305, 306, 408호' 처리와 동일)
    assert F('타워(301호, 302호)') == '타워 301호, 302호'
    # 동,건물 콤마(호수 아님)는 기존대로 공백 정리
    assert F('(가산동, 이앤씨드림타워7차)') == '이앤씨드림타워7차'


def test_mark_planned_glued_예정지():
    """'X예정지'(붙은) → 'X (예정)' + 재부착 중복 축약 (L-03600)."""
    assert ar._mark_planned('크란츠테크노 지하 1층 중식당예정지') == '크란츠테크노 지하 1층 중식당 (예정)'
    # 파이프라인 재부착 'X X예정지' 중복도 'X (예정)' 로 축약
    assert ar._mark_planned('지하 1층 중식당 중식당예정지') == '지하 1층 중식당 (예정)'


def test_mark_planned_glued_예정():
    """'X예정'(지 없음, 붙은) → 'X (예정)'."""
    assert ar._mark_planned('지하 1층 카페예정') == '지하 1층 카페 (예정)'


def test_mark_planned_idempotent_and_verb():
    """이미 '(예정)'은 유지, '설치 예정'(동사구, 공백)은 미변환(스톱워드가 별도 제거)."""
    assert ar._mark_planned('지하 1층 중식당 (예정)') == '지하 1층 중식당 (예정)'
    # '설치 예정' 은 공백 앞 예정 → _mark_planned 미매치 (그대로; 실제론 스톱워드가 제거)
    assert ar._mark_planned('삼성빌딩 3층 설치 예정') == '삼성빌딩 3층 설치 예정'


def test_mark_planned_glued_paren():
    """이미 괄호 친 'X(예정)'(공백 없이 붙은) → 'X (예정)' (L-03680)."""
    assert ar._mark_planned('인하로 79 2층 한의원(예정)') == '인하로 79 2층 한의원 (예정)'
    # 앞이 공백이면 정상 → 무변 (멱등)
    assert ar._mark_planned('인하로 79 2층 한의원 (예정)') == '인하로 79 2층 한의원 (예정)'


def test_extract_building_tail_keeps_paren_planned():
    """verify 경로 tail 추출이 '(예정)' 괄호 지정어를 절단하지 않고 보존 (L-03680).
    '예정' stop-word 가 '(예정)'까지 잘라 _mark_planned 가 (예정)을 못 살리던 갭."""
    r = ar._extract_building_tail('인천 미추홀구 인하로 79 , 2층 한의원(예정)')
    assert '(예정)' in r   # '2층 한의원(예정)' — 예정 보존


def test_flatten_paren_keeps_planned():
    """_flatten_paren_tail 이 '(예정)' 괄호를 벗기지 않음 (뒤 _mark_planned 용)."""
    assert ar._flatten_paren_tail('2층 한의원(예정)') == '2층 한의원(예정)'
    # 법정동 괄호는 기존대로 flatten (회귀 방어)
    assert ar._flatten_paren_tail('건영아파트(중계동)') != '건영아파트(중계동)'


def test_extract_tail_keeps_예정지_not_truncated():
    """'예정지'(예정+지)는 스톱워드 절단 대상 아님 — verify tail 보존 (L-03600)."""
    r = ar._extract_building_tail('성남 둔촌대로 388 크란츠테크노 지하 1층 중식당예정지')
    assert '중식당예정지' in r  # 절단 전 (변환은 resolve 최종 _mark_planned)
    r2 = ar._extract_building_tail('성남 둔촌대로 388 크란츠테크노 지하 1층 중식당 예정지')
    assert '예정지' in r2  # 띄어쓴 예정지도 보존


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
