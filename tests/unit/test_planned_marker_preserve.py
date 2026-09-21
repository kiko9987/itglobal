# -*- coding: utf-8 -*-
"""계획 중 장소 표시 '(예정)'·'예정지' 보존 — 전체 파이프라인 (L-04066).

extract_korean_address 가 주소 패턴 밖의 끝 괄호 노트 '(예정)'을 떨궈, 검증돼도 변환에서
(예정)이 사라지던 갭. extract 가 계획표시를 resolve 로 넘기고, resolve 의 _mark_planned 가
'X (예정)' 로 정규화. '설치 예정'(동사구)은 미대상. 네트워크는 monkeypatch 로 차단(hermetic:
kakao/juso 폴백 모두 None → 정규식 경로에서 _mark_planned 만 검증).
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar
from dashboard.services.lead_helpers import extract_korean_address as E


@pytest.fixture
def _no_net(monkeypatch):
    monkeypatch.setattr(ar, 'verify_address', lambda *a, **k: None)
    monkeypatch.setattr(ar, '_juso_fallback', lambda *a, **k: None)
    for fn in ['_try_poi_fallback', '_jibun_road_fallback', '_road_poi_fallback',
               '_dong_building_poi_fallback']:
        monkeypatch.setattr(ar, fn, lambda *a, **k: None)


def _pipeline(raw):
    rx = E(raw)
    return ar.resolve_address(raw, rx[0] if rx else None, rx[1] if rx else '')[0]


@pytest.mark.parametrize('raw, expected_has', [
    ('노원구 상계로1길 62-20 동한빌딩 3층 뷰티샵 (예정)', True),   # 괄호 노트
    ('강남구 테헤란로 152 3층 카페(예정)', True),                  # 붙임형 괄호
    ('강남구 테헤란로 152 3층 카페 예정지', True),                 # 예정지(명사 뒤)
    ('인천 서구 금정로 11 우미린클래스원 상가동 2층 건물 중 2층에 설치 예정입니다.', False),  # 동사구
    ('강남구 테헤란로 152 스타벅스', False),                       # 계획표시 없음
    ('강남구 테헤란로 152 3층 예정기획', False),                   # 상호 '예정기획' 오탐 방지
])
def test_planned_marker_end_to_end(raw, expected_has, _no_net):
    out = _pipeline(raw)
    assert ('(예정)' in out) is expected_has


def test_no_double_planned(_no_net):
    out = _pipeline('노원구 상계로1길 62-20 동한빌딩 3층 뷰티샵 (예정)')
    assert out.count('(예정)') == 1


def test_verified_path_keeps_planned(monkeypatch):
    # 검증(행안부 road)돼도 (예정) 유지 — regex_addr 에 (예정) 남으면 끝까지 보존.
    monkeypatch.setattr(ar, 'verify_address', lambda *a, **k: None)
    monkeypatch.setattr(ar, '_juso_fallback',
                        lambda *a, **k: ('노원구 상계로1길 62-20 동한빌딩', 'road'))
    raw = '노원구 상계로1길 62-20 동한빌딩 3층 뷰티샵 (예정)'
    rx = E(raw)
    addr, lv = ar.resolve_address(raw, rx[0], rx[1])
    assert lv == 'verified' and addr.endswith('(예정)')


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
