# -*- coding: utf-8 -*-
"""층↔상호 어순 보존 회귀 테스트 (2026-08-01 ETC-db43ed).

어순이 의미를 결정한다:
  • '삼성전자 3층' = 건물 통째로 쓰는 상호의 특정 층 (상호-층)
  • '1층 피아노학원' = 건물 1층에 입점한 한 층짜리 상호 (층-상호)
_enrich_verified_address 의 층 부착이 층을 무조건 맨 뒤로 옮겨 '1층 피아노학원'을
'피아노학원 1층'으로 뒤집으면 두 의미가 뒤섞인다. 매니저 원문 어순을 보존한다.
_enrich_with_poi(네트워크)는 monkeypatch 로 무력화.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services import address_resolver as ar


def _enrich(monkeypatch, verified, original, regex=''):
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda addr, text: addr)
    return ar._enrich_verified_address(verified, original, regex)


class TestFloorShopOrder:
    def test_floor_before_shop_preserved(self, monkeypatch):
        """층-상호(한 층 입점) → 어순 유지, 뒤집지 않음."""
        r = _enrich(
            monkeypatch,
            '화성 동탄구 동탄산척로2나길 10-10 피아노학원',
            '화성 동탄구 동탄산척로2나길 10-10 , 1층 피아노학원',
        )
        assert r == '화성 동탄구 동탄산척로2나길 10-10 1층 피아노학원'

    def test_shop_before_floor_unchanged(self, monkeypatch):
        """상호-층(건물 통째 상호의 특정 층) → 그대로 유지."""
        r = _enrich(
            monkeypatch,
            '서울 강남구 테헤란로 1 삼성전자',
            '서울 강남구 테헤란로 1 삼성전자 3층',
        )
        assert r == '서울 강남구 테헤란로 1 삼성전자 3층'

    def test_floor_only_no_shop_appended_at_end(self, monkeypatch):
        """상호 없이 층만 → 맨 뒤 부착(기존 동작 유지)."""
        r = _enrich(
            monkeypatch,
            '서울 강남구 테헤란로 1',
            '서울 강남구 테헤란로 1 5층',
        )
        assert r == '서울 강남구 테헤란로 1 5층'

    def test_poi_building_after_floor_restored(self, monkeypatch):
        """카카오 건물명 미등록 → 층 먼저 + POI 건물명 뒤(번지 층 건물). 원문이
        건물-층이면 건물-층 복원 (L-03485)."""
        monkeypatch.setattr(ar, '_enrich_with_poi',
                            lambda a, t: '송파구 올림픽로35다길 32 2층 예전빌딩')
        r = ar._enrich_verified_address(
            '송파구 올림픽로35다길 32',
            '송파구 올림픽로35다길 32 예전빌딩 2층',
            '송파구 올림픽로35다길 32',
        )
        assert r == '송파구 올림픽로35다길 32 예전빌딩 2층'

    def test_poi_floor_first_tenant_preserved(self, monkeypatch):
        """원문이 층-건물(한 층 입점)이면 POI 경로여도 층-건물 보존(복원 안 함)."""
        monkeypatch.setattr(ar, '_enrich_with_poi',
                            lambda a, t: '화성 동탄구 동탄산척로2나길 10-10 1층 피아노학원')
        r = ar._enrich_verified_address(
            '화성 동탄구 동탄산척로2나길 10-10',
            '화성 동탄구 동탄산척로2나길 10-10 , 1층 피아노학원',
            '화성 동탄구 동탄산척로2나길 10-10',
        )
        assert r == '화성 동탄구 동탄산척로2나길 10-10 1층 피아노학원'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
