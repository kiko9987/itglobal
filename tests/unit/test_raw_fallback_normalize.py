# -*- coding: utf-8 -*-
"""raw fallback 경로 시/도 정규화 회귀 테스트 (2026-08-04 L-03398).

verified·regex 경로만 normalize_display 를 거쳐, 카카오 verify 실패로 raw 첫줄
fallback 된 서울/광역시 주소는 '서울특별시'·'경기도' 접두가 그대로 남던 갭.
raw 경로에도 normalize_display 적용(시/도만 정리, 건물·동·번지·호수 보존, level='raw' 유지).
verify_address·_try_poi_fallback 를 mock 해 raw 경로 강제.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services import address_resolver as ar


def _force_raw(monkeypatch):
    monkeypatch.setattr(ar, 'verify_address', lambda *a, **k: None)
    monkeypatch.setattr(ar, '_try_poi_fallback', lambda *a, **k: None)


def test_raw_fallback_strips_seoul(monkeypatch):
    """서울특별시 접두는 raw 경로에서도 제거, 나머지는 보존 + level='raw'."""
    _force_raw(monkeypatch)
    addr, lv = ar.resolve_address(
        '서울특별시 용산구 서빙고로 31 용산시티파크 2단지 202동 2805호', None, '')
    assert lv == 'raw'
    assert not addr.startswith('서울특별시')
    assert addr.startswith('용산구')
    # 건물·동·호수 보존
    assert '용산시티파크' in addr and '202동' in addr and '2805호' in addr


def test_raw_fallback_abbreviates_metro(monkeypatch):
    """광역시·시 접두 축약 (인천광역시→인천)."""
    _force_raw(monkeypatch)
    addr, lv = ar.resolve_address('인천광역시 부평구 어딘가로 5 3층', None, '')
    assert lv == 'raw'
    assert addr.startswith('인천 부평구')


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
