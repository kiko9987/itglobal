# -*- coding: utf-8 -*-
"""상호 뒤 전화번호가 붙어도 다단어 상호 보존 (L-04084).

_enrich_verified_address 의 다단어 상호 보존(_mshop2, L-03811)은 '$' 끝 앵커를 요구하는데,
고객이 상호 뒤에 전화를 붙이면('결헤어 미용실 02,308~0834') 앵커가 깨져 상호를 못 잡던 갭.
끝의 전화형 숫자 blob 만 제거 후 매치. _enrich_with_poi(네트워크)는 monkeypatch 로 차단.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.fixture(autouse=True)
def _no_poi(monkeypatch):
    # _enrich_with_poi 는 카카오 호출 → identity 로 차단(순수 regex 로직만 검증)
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda addr, text: addr)


def test_shop_kept_despite_trailing_phone():
    raw = '은평구 진관동 67,911동102호 결헤어 미용실 02,308~0834'
    out = ar._enrich_verified_address('은평구 진관3로 43-9', raw, '은평구 진관동 67')
    assert out == '은평구 진관3로 43-9 102호 결헤어 미용실'


def test_shop_kept_with_dashed_phone():
    raw = '수원 매영로 5 102호 국면당 공세점 010-1234-5678'
    out = ar._enrich_verified_address('수원 매영로 5', raw, '수원 매영로 5')
    assert out.endswith('국면당 공세점')


def test_no_phone_unchanged():
    # 전화 없이 상호가 끝 — 기존대로 보존
    raw = '수원 매영로 5 102호 국면당 공세점'
    out = ar._enrich_verified_address('수원 매영로 5', raw, '수원 매영로 5')
    assert out.endswith('국면당 공세점')


def test_trailing_beonji_not_stripped_as_phone():
    # 끝이 번지(43-9, 4자)면 전화로 오인 안 함 — 상호 없어 append 도 없음
    raw = '은평구 진관3로 43-9'
    out = ar._enrich_verified_address('은평구 진관3로 43-9', raw, '은평구 진관3로 43-9')
    assert out == '은평구 진관3로 43-9'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
