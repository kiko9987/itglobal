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


def test_enrich_prefix_boost_when_standalone(monkeypatch):
    """standalone cand 는 접두 보강 유지 (회귀, 한강듀클래스→김포한강듀클래스 계열)."""
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('제이비미소빌딩', '서울 강남구 논현로 841')])
    r = ar._enrich_with_poi('강남구 논현로 841 미소빌딩 2층',
                            '강남구 논현로 841 미소빌딩 2층')
    assert '제이비미소빌딩' in r


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


def test_flatten_paren_keeps_corporate_designator():
    """법인 표기 (주)/(유)/(사)/(재)는 flatten 안 함 — '주' 홀로 남기 방지 (L-03358)."""
    F = ar._flatten_paren_tail
    assert F('에스티팜 시화공장(주)') == '에스티팜 시화공장(주)'
    assert F('(주)가연결혼정보') == '(주)가연결혼정보'
    # 회귀: 층/지번 괄호 flatten 은 그대로
    assert F('타임빌딩(2층)') == '타임빌딩 2층'
    assert F('(가산동, 이앤씨드림타워7차)') == '이앤씨드림타워7차'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
