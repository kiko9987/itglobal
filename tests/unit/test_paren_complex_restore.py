# -*- coding: utf-8 -*-
"""원문 괄호 속 단지명 복원·번지 뒤 배치 (ETC-258cd3, L-03962 정책).

'권율로 1203번길 39-37(홍죽산업단지)'의 '(홍죽산업단지)'가 검증 주소에 없어 보강 단계에서
유실되던 갭 → 괄호 벗겨 도로+번지 바로 뒤에 삽입. 이미 있지만 상가·호 뒤 꼬리에 있으면 번지
뒤로 이동. 노트·법정동 괄호는 미대상. _enrich_with_poi(네트워크)는 monkeypatch.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.fixture(autouse=True)
def _no_poi(monkeypatch):
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda addr, text: addr)


def test_missing_complex_restored_after_beonji():
    raw = '양주시 백석읍 권율로 1203번길 39-37(홍죽산업단지)'
    out = ar._enrich_verified_address('양주 백석읍 권율로1203번길 39-37', raw, None)
    assert out == '양주 백석읍 권율로1203번길 39-37 홍죽산업단지'


def test_existing_complex_moved_before_shop():
    raw = '하남 미사강변서로 167-1 상가 1동 1층 103호 (미사강변도시 17단지)'
    va = '하남 미사강변서로 167-1 상가 1동 1층 103호 미사강변도시 17단지'
    out = ar._enrich_verified_address(va, raw, None)
    assert out.startswith('하남 미사강변서로 167-1 미사강변도시 17단지 상가')
    assert out.count('미사강변도시') == 1


@pytest.mark.parametrize('raw', [
    '고잔동 위례광장로 30 (비밀번호 7080)',   # 노트 괄호 — 단지 접미 아님
    '인천 서구 금정로 11 (불로동)',           # 법정동 괄호
])
def test_non_complex_paren_ignored(raw):
    va = '인천 서구 금정로 11' if '금정로' in raw else '위례광장로 30'
    out = ar._enrich_verified_address(va, raw, None)
    assert '비밀번호' not in out and '불로동' not in out


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
