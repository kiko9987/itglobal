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


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
