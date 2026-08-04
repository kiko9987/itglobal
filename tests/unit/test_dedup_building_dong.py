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
