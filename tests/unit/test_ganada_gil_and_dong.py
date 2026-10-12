# -*- coding: utf-8 -*-
"""가나다길 도로 접미 · 아파트 동 번호 재부착 · 단지 내 입점업체 미부착 (L-04213).

'은평구 갈현로3나길 23 역촌센트레빌 102동 501호'가 ①입구 정규화에서 '갈현로 3 나길'로
쪼개져 엉뚱한 건물(은평자이더스타)로 오매칭, ②카카오 검증 주소에 동이 없어 '102동' 유실,
③'역촌센트레빌 공인중개사사무소' POI 의 '공인중개사사무소'가 지점명으로 붙던 갭.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar
from dashboard.services.lead_helpers import _normalize_road_spacing as N


@pytest.mark.parametrize('raw, expected', [
    ('은평구 갈현로3나길 23', '은평구 갈현로3나길 23'),     # 붙은 것 — 쪼개지 않음
    ('은평구 갈현로 3나길 23', '은평구 갈현로3나길 23'),    # 로 뒤 공백 — 붙임
    ('은평구 갈현로3나길23', '은평구 갈현로3나길 23'),      # 길+번지 붙여쓰기 — 분리
    ('성현로135번안길 80-7', '성현로135번안길 80-7'),      # 기존 번안길 회귀 없음
    ('가람로124봉명타워', '가람로 124 봉명타워'),           # 건물 붙여쓰기 분리 유지
])
def test_ganada_gil_spacing(raw, expected):
    assert N(raw) == expected


def test_join_road_gil_spaced_form():
    assert ar._join_road_gil('은평구 갈현로 3 나길 23') == '은평구 갈현로3나길 23'


def test_apartment_dong_reattached(monkeypatch):
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda addr, text: addr)
    raw = '은평구 갈현로3나길 23 역촌센트레빌 102동 501호'
    out = ar._enrich_verified_address('은평구 갈현로3나길 23 역촌센트레빌', raw, None)
    assert out == '은평구 갈현로3나길 23 역촌센트레빌 102동 501호'


def test_legal_dong_not_reattached(monkeypatch):
    # '성수2동'(법정동)의 '2동'은 동 번호로 오인하지 않음
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda addr, text: addr)
    out = ar._enrich_verified_address('성동구 아차산로 1', '성동구 성수2동 아차산로 1 3층', None)
    assert '2동' not in out.replace('성수2동', '') and out.endswith('3층')


def test_tenant_not_appended_when_complex_exists(monkeypatch):
    # 단지명+아파트 POI 가 같은 도로에 있으면 'cand X' 의 X(부동산)는 입점업체 → 미부착
    monkeypatch.setattr(ar, '_search_poi', lambda q: (
        ('역촌센트레빌아파트', '서울 은평구 갈현로3나길 23'),
        ('역촌센트레빌 공인중개사사무소', '서울 은평구 갈현로3나길 23'),
    ))
    va = '은평구 갈현로3나길 23 역촌센트레빌'
    out = ar._enrich_with_poi(va, va + ' 102동 501호')
    assert '공인중개사' not in out


def test_branch_still_appended_without_building(monkeypatch):
    # 건물 POI 가 없으면 지점명 보강은 기존대로(마성떡볶이 논현역점)
    monkeypatch.setattr(ar, '_search_poi', lambda q: (
        ('마성떡볶이 논현역점', '서울 강남구 학동로 1'),))
    va = '강남구 학동로 1 마성떡볶이'
    assert ar._enrich_with_poi(va, va).endswith('마성떡볶이 논현역점')


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
