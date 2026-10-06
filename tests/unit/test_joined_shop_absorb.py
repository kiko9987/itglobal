# -*- coding: utf-8 -*-
"""결합 상호 POI 치환 시 뒤 토큰 흡수 + 일반 유형어만 남으면 미치환 (ETC-ed7d0d).

'힐스테이트 고덕 센트럴 3층' → POI '힐스테이트고덕센트럴아파트' 는 뒤 토큰 '센트럴'까지
포함해, 기존엔 '힐스테이트 고덕'만 치환해 '…아파트 센트럴 3층' 중복. 이제 포함된 뒤 토큰을
흡수하고, 남는 게 '아파트' 같은 일반 유형어뿐이면 원문 유지. 지점명이 남으면 기존대로 치환.
_search_poi 는 monkeypatch (hermetic).
"""
import sys
sys.path.insert(0, '.')

import dashboard.services.address_resolver as ar


def test_apartment_suffix_only_keeps_original(monkeypatch):
    monkeypatch.setattr(ar, '_search_poi', lambda q: (
        ('힐스테이트고덕센트럴아파트', '경기 평택시 고덕국제대로 77'),))
    va = '평택 고덕국제대로 77 힐스테이트 고덕 센트럴 3층'
    out = ar._enrich_with_poi(va, va)
    assert out == va
    assert out.count('센트럴') == 1


def test_branch_suffix_still_replaced(monkeypatch):
    # 지점명이 남는 원래 용도(남양가 양꼬치 → 남양가양꼬치 마곡점)는 유지
    monkeypatch.setattr(ar, '_search_poi', lambda q: (
        ('남양가양꼬치 마곡점', '서울 강서구 마곡중앙6로 11'),))
    va = '강서구 마곡중앙6로 11 남양가 양꼬치'
    out = ar._enrich_with_poi(va, va)
    assert out == '강서구 마곡중앙6로 11 남양가양꼬치 마곡점'


def test_absorbed_tokens_not_duplicated(monkeypatch):
    # POI 가 뒤 토큰 포함 + 지점명 → 흡수한 토큰은 중복 없이 치환
    monkeypatch.setattr(ar, '_search_poi', lambda q: (
        ('가나다라카페 역삼점', '서울 강남구 테헤란로 1'),))
    va = '강남구 테헤란로 1 가나 다라 카페 2층'
    out = ar._enrich_with_poi(va, va)
    assert out == '강남구 테헤란로 1 가나다라카페 역삼점 2층'


if __name__ == '__main__':
    import pytest
    sys.exit(pytest.main([__file__, '-v']))
