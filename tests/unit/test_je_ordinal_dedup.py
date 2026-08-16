# -*- coding: utf-8 -*-
"""서수 '제' 접두 dedup — POI '제4공장' + 고객 '4공장' 중복 방지 (ETC-ad2710).

고객이 '보우테이프 4공장', 카카오 POI 가 '보우테이프 제4공장' 이면 같은 공장인데
dedup 이 다른 것으로 봐 '보우테이프 제4공장 4공장' 중복되던 버그. '제'+숫자 접두를
정규화해 동치 판정. _search_poi 를 스텁으로 대체(네트워크 없음).
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


def test_strip_je_ordinal():
    assert ar._strip_je_ordinal('제4공장') == '4공장'
    assert ar._strip_je_ordinal('보우테이프 제4공장') == '보우테이프 4공장'
    assert ar._strip_je_ordinal('제2동 제3호') == '2동 3호'
    # '제'+비숫자(제일빌딩·제주)는 불변
    assert ar._strip_je_ordinal('제일빌딩') == '제일빌딩'
    assert ar._strip_je_ordinal('제주도청') == '제주도청'


@pytest.fixture
def poi_stub(monkeypatch):
    ROAD = '경기 화성시 만세구 양감면 정문송산로93번길 1'
    monkeypatch.setattr(ar, '_search_poi', lambda q: [('보우테이프 제4공장', ROAD)])
    monkeypatch.setattr(ar, '_extract_shop_candidates', lambda t: ['보우테이프'])
    return ROAD


def test_customer_number_no_dup(poi_stub):
    # 고객 '4공장' → POI '제4공장' 이 중복 부착되면 안 됨
    v = '화성 만세구 양감면 정문송산로93번길 1 보우테이프 4공장'
    assert ar._enrich_with_poi(v, '보우테이프 4공장') == v


def test_customer_je_number_no_dup(poi_stub):
    v = '화성 만세구 양감면 정문송산로93번길 1 보우테이프 제4공장'
    assert ar._enrich_with_poi(v, '보우테이프 제4공장') == v


def test_no_number_appends_disambiguation(poi_stub):
    # 고객이 번호 안 쓰면 POI 지점명(제4공장) 부착은 유지 (disambiguation)
    v = '화성 만세구 양감면 정문송산로93번길 1 보우테이프'
    assert ar._enrich_with_poi(v, '보우테이프') == v + ' 제4공장'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
