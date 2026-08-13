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
    def test_verified_different_returns_normalized_note(self, monkeypatch):
        """verified & 정정됨 → 정규화값 + normalized note (원본/변환 2줄 + ephemeral)."""
        norm = '수원 권선구 덕영대로 1205 미래타운빌딩 501호'
        raw = '수원시 권선구 덕영대로 1205 501호'
        _patch_resolve(monkeypatch, (norm, 'verified'))
        addr, note = slack_bot._normalize_visit_address_if_verified(raw)
        assert addr == norm
        assert note == {'kind': 'normalized', 'original': raw, 'normalized': norm}

    @pytest.mark.parametrize('level', ['level3', 'level7', 'raw', 'failed', ''])
    def test_non_verified_keeps_raw_with_badge(self, monkeypatch, level):
        """미verified(도로·번지 미확인) → raw 유지 + failed addr_note (확인 필요 배지)."""
        _patch_resolve(monkeypatch, ('뭔가 다른 값', level))
        raw = '용인시동천동동천타워402호'
        addr, note = slack_bot._normalize_visit_address_if_verified(raw)
        assert addr == raw
        assert note == {'kind': 'failed', 'original': raw, 'normalized': ''}

    def test_jibun_poi_returns_estimated(self, monkeypatch):
        """지번→keyword 도로명 구제(jibun_poi) → 도로명 채택 + estimated note (L-03673).

        온라인 카드처럼 방문/전화 카드도 [추정] 배지로 표시. verified 아니지만
        raw+확인필요로 접지 않고 도로명 변환값을 살림.
        """
        _patch_resolve(monkeypatch, ('영등포구 은행로 3', 'jibun_poi'))
        raw = '영등포구 여의도동 15-24'
        addr, note = slack_bot._normalize_visit_address_if_verified(raw)
        assert addr == '영등포구 은행로 3'
        assert note == {'kind': 'estimated', 'original': raw,
                        'normalized': '영등포구 은행로 3'}

    def test_jibun_poi_same_value_falls_through_failed(self, monkeypatch):
        """jibun_poi 인데 변환값이 원본과 같으면(이례) failed 로 폴백."""
        _patch_resolve(monkeypatch, ('영등포구 여의도동 15-24', 'jibun_poi'))
        raw = '영등포구 여의도동 15-24'
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


class TestRegionCrossCheck:
    """시/구 교차확인 — 입력과 다른 시/구로 변환되면 region_warn 플래그 (L-03659)."""

    @pytest.mark.parametrize('raw, norm, expected', [
        ('인천 부평구 경인로 789 서광빌딩', '영등포구 경인로 789 서광빌딩', True),
        ('인천 부평구 경인로 789', '인천 부평구 경인로 789 서광빌딩', False),
        ('강남구 테헤란로 152', '강남구 테헤란로 152 삼성빌딩', False),
        ('서울 강남구 테헤란로 152', '강남구 테헤란로 152', False),  # 서울 생략, 구 동일
        ('부평구 십정동 572-7', '인천 부평구 경인로 789 서광빌딩', False),  # 지번→도로, 구 동일
    ])
    def test_region_changed_detect(self, raw, norm, expected):
        assert slack_bot._region_changed(raw, norm) is expected

    def test_normalize_sets_region_warn(self, monkeypatch):
        """정규화로 시/구 바뀌면 addr_note 에 region_warn=True."""
        _patch_resolve(monkeypatch, ('영등포구 경인로 789 서광빌딩', 'verified'))
        addr, note = slack_bot._normalize_visit_address_if_verified(
            '인천 부평구 경인로 789 서광빌딩')
        assert note['kind'] == 'normalized'
        assert note.get('region_warn') is True

    def test_normalize_no_warn_same_region(self, monkeypatch):
        """같은 구로 정규화(건물명만 추가)면 region_warn 없음."""
        _patch_resolve(monkeypatch, ('인천 부평구 경인로 789 서광빌딩', 'verified'))
        addr, note = slack_bot._normalize_visit_address_if_verified(
            '인천 부평구 경인로 789')
        assert note['kind'] == 'normalized'
        assert not note.get('region_warn')


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
