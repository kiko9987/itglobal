# -*- coding: utf-8 -*-
"""건물명 앞 (주) 제거 + 층↔건물 어순 복원 (L-03959).

'(주)뉴빛빌딩 2층' → '뉴빛빌딩 2층'(건물은 주식회사 아님, 건물접미로 끝날 때만 (주) 제거)
+ 층을 건물 뒤로. 접미 없는 상호 '(주)삼인'의 (주)는 보존하되 어순은 복원. _enrich_with_poi
는 monkeypatch 로 차단(hermetic).
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.fixture(autouse=True)
def _no_poi(monkeypatch):
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda v, o: v)


def test_jusa_building_strip_and_reorder():
    # 층-건물 역순 + (주) 건물접미 → 제거 + 건물-층 복원
    out = ar._enrich_verified_address(
        '의정부 산단로76번길 93 2층 (주)뉴빛빌딩',
        '경기 의정부시 용현동 524-4 (주)뉴빛빌딩 2층', None)
    assert out == '의정부 산단로76번길 93 뉴빛빌딩 2층'


def test_jusa_company_kept_but_reordered():
    # 접미 없는 상호 (주)삼인 → (주) 보존, 어순만 복원
    out = ar._enrich_verified_address(
        '시흥 공단1대로 38 2층 (주)삼인',
        '시흥 공단1대로 38 (주)삼인 2층', None)
    assert out == '시흥 공단1대로 38 (주)삼인 2층'


def test_building_no_jusa_reorder():
    out = ar._enrich_verified_address(
        '의정부 산단로76번길 93 2층 뉴빛빌딩',
        '의정부 산단로76번길 93 뉴빛빌딩 2층', None)
    assert out == '의정부 산단로76번길 93 뉴빛빌딩 2층'


def test_floor_shop_order_preserved():
    # 원문이 층-상호(1층 피아노학원)면 어순 유지 (역순 복원 안 함)
    out = ar._enrich_verified_address(
        '강남구 테헤란로 152 1층 피아노학원',
        '강남구 테헤란로 152 1층 피아노학원', None)
    assert out == '강남구 테헤란로 152 1층 피아노학원'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
