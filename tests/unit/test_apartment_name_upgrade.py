# -*- coding: utf-8 -*-
"""고객 아파트명 축약·오기 → 공식명 승격 (L-03695).

'독산동신도아파트'(고객 축약) → '신도브래뉴아파트'(정식). 이중 소스(카카오 building_name
== 행안부 bdNm) 일치 + 핵심 토큰(신도) 공유 시에만. 다른 아파트 지목(래미안)·소스 불일치는
원문 유지. 네트워크 없이 _juso_search_cached 스텁.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.fixture
def juso_bd(monkeypatch):
    """_juso_search_cached 를 (roadAddr, jibunAddr, bdNm) 튜플 목록으로 대체."""
    def _install(bdnm):
        rows = ((f'서울특별시 금천구 시흥대로152길 25 (독산동, {bdnm})',
                 '서울특별시 금천구 독산동 291', bdnm),) if bdnm else ()
        monkeypatch.setattr(ar, '_juso_search_cached', lambda q: rows)
    return _install


BASE = '금천구 시흥대로152길 25'


def test_upgrade_when_dual_source_and_token_share(juso_bd):
    juso_bd('신도브래뉴아파트')
    # 고객 '독산동신도아파트', 카카오 building_name '신도브래뉴아파트'
    out = ar._maybe_upgrade_apartment_name(BASE, '독산동신도아파트', '신도브래뉴아파트')
    assert out == '신도브래뉴아파트'


def test_preserve_dong_ho_tail(juso_bd):
    juso_bd('신도브래뉴아파트')
    out = ar._maybe_upgrade_apartment_name(BASE, '독산동신도아파트 102동 501호', '신도브래뉴아파트')
    assert out == '신도브래뉴아파트 102동 501호'


def test_no_upgrade_when_juso_disagrees(juso_bd):
    juso_bd('행복아파트')  # 행안부는 다른 이름 → 이중소스 불일치
    assert ar._maybe_upgrade_apartment_name(BASE, '독산동신도아파트', '신도브래뉴아파트') is None


def test_no_upgrade_when_juso_empty(juso_bd):
    juso_bd(None)
    assert ar._maybe_upgrade_apartment_name(BASE, '독산동신도아파트', '신도브래뉴아파트') is None


def test_no_upgrade_when_token_mismatch(juso_bd):
    juso_bd('신도브래뉴아파트')
    # 고객이 '래미안'을 지목 — 다른 아파트일 수 있으므로 원문 유지
    assert ar._maybe_upgrade_apartment_name(BASE, '래미안아파트', '신도브래뉴아파트') is None


def test_no_upgrade_when_already_official(juso_bd):
    juso_bd('신도브래뉴아파트')
    assert ar._maybe_upgrade_apartment_name(BASE, '신도브래뉴아파트', '신도브래뉴아파트') is None


def test_no_upgrade_when_kakao_not_apartment(juso_bd):
    juso_bd('신도브래뉴빌딩')
    # 카카오 building_name 이 아파트가 아니면 미적용
    assert ar._maybe_upgrade_apartment_name(BASE, '독산동신도아파트', '신도브래뉴빌딩') is None


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
