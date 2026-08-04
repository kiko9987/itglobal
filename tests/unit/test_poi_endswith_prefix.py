# -*- coding: utf-8 -*-
"""POI endswith 접두 매치 가드 회귀 테스트 (2026-08-04 L-03530).

_enrich_with_poi 의 endswith 매치는 '접두 지역명 보강'(한강듀클래스→김포한강듀클래스)
용인데, 다세대 건물(어반322 지식산업센터)에선 같은 건물의 다른 입점 업체
(슈퍼스타/PINS/이마트24 어반322)까지 접두로 오인해 건물명을 엉뚱한 테넌트로 치환·부착.

두 가드로 방어 (_search_poi 를 monkeypatch — 네트워크 없이 매치 로직 검증):
  1. 접두가 '공백 분리'면 다른 업체('슈퍼스타 ')라 차단 (붙은 접두 김포한강듀클래스만).
  2. cand 가 같은 도로의 exact POI 로 실재하면(어반322 건물 존재) 접두형은 입점 업체
     이므로 endswith 자체를 차단. exact 가 없으면(한강듀클래스) 접두형이 정식 전체명이라 허용.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services import address_resolver as ar


def test_space_separated_tenant_not_appended(monkeypatch):
    """cand 뒤 '공백 접두' POI(다른 업체 '슈퍼스타 어반322')는 endswith 부착 안 함."""
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('슈퍼스타 어반322', '서울 영등포구 양평로25길 8')])
    r = ar._enrich_with_poi('영등포구 양평로25길 8 URBAN322',
                            '영등포구 양평로25길 8 어반322')
    assert '슈퍼스타' not in r


def test_exact_match_still_appends(monkeypatch):
    """정확 매치(place_name == cand)는 정상 부착 (회귀 방지)."""
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('어반322', '서울 영등포구 양평로25길 8')])
    r = ar._enrich_with_poi('영등포구 양평로25길 8 URBAN322',
                            '영등포구 양평로25길 8 어반322')
    assert '어반322' in r


def test_glued_tenant_blocked_when_exact_building_exists(monkeypatch):
    """L-03530 실측: 건물(exact 어반322)이 실재하면 붙은 접두 입점업체(PINS어반322) 차단.

    verified 에 이미 어반322 가 붙어있고(step 1-b m_shop 부착) _search_poi 가 exact
    어반322 + 입점업체 PINS어반322 를 함께 반환 → exact 가 있으므로 endswith 치환 금지.
    """
    monkeypatch.setattr(ar, '_search_poi', lambda q: [
        ('어반322', '서울 영등포구 양평로25길 8'),
        ('PINS어반322', '서울 영등포구 양평로25길 8'),
    ])
    r = ar._enrich_with_poi('영등포구 양평로25길 8 URBAN322 어반322',
                            '영등포구 양평로25길 8 어반322 어반322')
    assert 'PINS' not in r
    assert '어반322' in r


def test_glued_region_prefix_appends_when_no_exact(monkeypatch):
    """exact cand POI 없으면 접두 지역명 보강 유지 (한강듀클래스→김포한강듀클래스, L-03316)."""
    monkeypatch.setattr(ar, '_search_poi', lambda q: [
        ('김포한강듀클래스', '경기 김포시 김포한강로 100'),
    ])
    r = ar._enrich_with_poi('김포시 김포한강로 100 한강듀클래스',
                            '김포시 김포한강로 100 한강듀클래스')
    assert '김포한강듀클래스' in r


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
