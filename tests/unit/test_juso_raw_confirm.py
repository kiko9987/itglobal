# -*- coding: utf-8 -*-
"""raw fallback 행안부 도로확인 — extract=None 실재도로 verified 승격 (L-03709).

'김포 양촌읍 삼도로 93'처럼 시/도(경기) 없이 시작하면 extract_korean_address=None →
step2(regex) 를 못 타 raw([주소 확인 필요])로 빠졌음. 카카오 미인덱싱(신축)이라도 행안부에
도로+번지가 실재하면 verified 승격. 없는/퍼지 도로는 미승격. 네트워크 스텁.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.fixture
def stub_apis(monkeypatch):
    """카카오=미인덱싱, 행안부=삼도로 93 실재 로 고정."""
    monkeypatch.setattr(ar, '_juso_key', lambda: 'STUB')
    monkeypatch.setattr(ar, '_kakao_search_cached', lambda q: None)
    monkeypatch.setattr(ar, '_kakao_search_poi_cached', lambda q: ())

    def _juso(q):
        if '삼도로' in q and '93' in q:
            return (('경기도 김포시 양촌읍 삼도로 93', '경기도 김포시 양촌읍 학운리 100', ''),)
        return ()
    monkeypatch.setattr(ar, '_juso_search_cached', _juso)


def _resolve(t):
    from dashboard.services.lead_helpers import extract_korean_address
    rx = extract_korean_address(t)
    return ar.resolve_address(t, rx[0] if rx else None, rx[1] if rx else '')


def test_extract_none_real_road_verified(stub_apis):
    addr, lv = _resolve('김포 양촌읍 삼도로 93')
    assert lv == 'verified'
    assert '삼도로 93' in addr


def test_nonexistent_road_stays_raw(stub_apis):
    # 행안부에 없는 도로+번지 → raw 유지([주소 확인 필요])
    addr, lv = _resolve('김포 양촌읍 없는길 12')
    assert lv != 'verified'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
