# -*- coding: utf-8 -*-
"""층/호 여러 개 모두 부착 — '지하2층. B207호' 호수 유실 방지 (L-03772).

_enrich_verified_address 의 층/호 부착이 re.search(첫 개)라 '지하2층 B207호'에서 뒤 호수를
버렸음. finditer 로 순서대로 모두 부착. _search_poi 스텁으로 건물 보강은 배제.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.fixture(autouse=True)
def _no_poi(monkeypatch):
    monkeypatch.setattr(ar, '_search_poi', lambda q: [])


@pytest.mark.parametrize('verified, original, expected', [
    # 층 + 호 둘 다 부착
    ('성남 중원구 사기막골로62번길 37 스타타워',
     '성남시 중원구 상대원동 223-25번지 스타타워 지하2층. B207호',
     '성남 중원구 사기막골로62번길 37 스타타워 지하2층 B207호'),
    ('부천 오정구 신흥로489번길 56 원일테크노2',
     '부천시 오정동 810-1 원일테크노2 4층 402호',
     '부천 오정구 신흥로489번길 56 원일테크노2 4층 402호'),
    # 단일 층/호 — 회귀 없음
    ('성남 중원구 사기막골로62번길 37 스타타워',
     '스타타워 B207호', '성남 중원구 사기막골로62번길 37 스타타워 B207호'),
])
def test_multi_floor_ho_attached(verified, original, expected):
    assert ar._enrich_verified_address(verified, original, None) == expected


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
