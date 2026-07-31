# -*- coding: utf-8 -*-
"""방문 모달 주소 정규화 헬퍼 회귀 테스트.

_normalize_visit_address_if_verified: 카카오 verified 일 때만 정규화값 사용,
아니면 raw 유지 (전화 sync 와 동일 사상, 2026-07-30). resolve_address 를 mock 해
네트워크 없이 분기 로직만 검증.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.blueprints import slack_bot
from dashboard.services import address_resolver


def _patch_resolve(monkeypatch, ret):
    monkeypatch.setattr(address_resolver, 'resolve_address', lambda *a, **k: ret)


class TestVisitAddressNormalize:
    def test_verified_uses_normalized(self, monkeypatch):
        _patch_resolve(monkeypatch, ('수원 권선구 덕영대로 1205 미래타운빌딩 501호', 'verified'))
        out = slack_bot._normalize_visit_address_if_verified('수원시 권선구 덕영대로 1205 501호')
        assert out == '수원 권선구 덕영대로 1205 미래타운빌딩 501호'

    @pytest.mark.parametrize('level', ['level3', 'level7', 'raw', 'failed', ''])
    def test_non_verified_keeps_raw(self, monkeypatch, level):
        _patch_resolve(monkeypatch, ('뭔가 다른 값', level))
        raw = '용인시동천동동천타워402호'
        assert slack_bot._normalize_visit_address_if_verified(raw) == raw

    def test_empty_returns_empty(self, monkeypatch):
        _patch_resolve(monkeypatch, ('', 'failed'))
        assert slack_bot._normalize_visit_address_if_verified('') == ''
        assert slack_bot._normalize_visit_address_if_verified('   ') == ''

    def test_exception_keeps_raw(self, monkeypatch):
        def _boom(*a, **k):
            raise RuntimeError('kakao down')
        monkeypatch.setattr(address_resolver, 'resolve_address', _boom)
        raw = '서울 강남구 역삼동 123-4'
        assert slack_bot._normalize_visit_address_if_verified(raw) == raw

    def test_verified_same_value_no_change(self, monkeypatch):
        _patch_resolve(monkeypatch, ('서초구 잠원로8길 25', 'verified'))
        assert slack_bot._normalize_visit_address_if_verified('서초구 잠원로8길 25') == '서초구 잠원로8길 25'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
