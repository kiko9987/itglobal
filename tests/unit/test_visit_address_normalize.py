# -*- coding: utf-8 -*-
"""방문 모달 주소 정규화 헬퍼 회귀 테스트.

_normalize_visit_address_if_verified: 카카오 verified 면 정규화값, 아니면 raw 유지.
2026-08-01: 반환이 (주소, addr_note) 튜플로 변경 — 미verified(도로·번지 미확인)면
addr_note={'kind':'failed'} 로 방문 카드/답글에 '주소 확인 필요' 배지. 매니저가
도로명·번지를 잘못 받아적는(돌이킬 수 없는) 오배송 방어. resolve_address mock 로
네트워크 없이 분기만 검증.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.blueprints import slack_bot
from dashboard.services import address_resolver


def _patch_resolve(monkeypatch, ret):
    monkeypatch.setattr(address_resolver, 'resolve_address', lambda *a, **k: ret)


class TestVisitAddressNormalize:
    def test_verified_uses_normalized_no_badge(self, monkeypatch):
        _patch_resolve(monkeypatch, ('수원 권선구 덕영대로 1205 미래타운빌딩 501호', 'verified'))
        addr, note = slack_bot._normalize_visit_address_if_verified('수원시 권선구 덕영대로 1205 501호')
        assert addr == '수원 권선구 덕영대로 1205 미래타운빌딩 501호'
        assert note is None   # verified → 배지 없음

    @pytest.mark.parametrize('level', ['level3', 'level7', 'raw', 'failed', ''])
    def test_non_verified_keeps_raw_with_badge(self, monkeypatch, level):
        """미verified(도로·번지 미확인) → raw 유지 + failed addr_note (확인 필요 배지)."""
        _patch_resolve(monkeypatch, ('뭔가 다른 값', level))
        raw = '용인시동천동동천타워402호'
        addr, note = slack_bot._normalize_visit_address_if_verified(raw)
        assert addr == raw
        assert note == {'kind': 'failed', 'original': raw, 'normalized': ''}

    def test_empty_returns_empty_no_badge(self, monkeypatch):
        _patch_resolve(monkeypatch, ('', 'failed'))
        assert slack_bot._normalize_visit_address_if_verified('') == ('', None)
        assert slack_bot._normalize_visit_address_if_verified('   ') == ('', None)

    def test_exception_keeps_raw_no_badge(self, monkeypatch):
        """정규화 예외(카카오 장애) → raw 유지, 배지 없음(재상담 자체 안 막음)."""
        def _boom(*a, **k):
            raise RuntimeError('kakao down')
        monkeypatch.setattr(address_resolver, 'resolve_address', _boom)
        raw = '서울 강남구 역삼동 123-4'
        addr, note = slack_bot._normalize_visit_address_if_verified(raw)
        assert addr == raw
        assert note is None

    def test_verified_same_value_no_badge(self, monkeypatch):
        _patch_resolve(monkeypatch, ('서초구 잠원로8길 25', 'verified'))
        addr, note = slack_bot._normalize_visit_address_if_verified('서초구 잠원로8길 25')
        assert addr == '서초구 잠원로8길 25'
        assert note is None


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
