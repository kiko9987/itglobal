# -*- coding: utf-8 -*-
"""서수 '제' 표기차 처리 — 고객 '4공장' → 카카오 공식명 '제4공장' 승격 (ETC-ad2710).

고객이 '보우테이프 4공장', 카카오 POI 가 '보우테이프 제4공장' 이면 같은 공장. 예전엔
dedup 이 다른 것으로 봐 '보우테이프 제4공장 4공장' 중복. '제'+숫자 접두 정규화로 동치
판정 → 공식명(제N공장)으로 승격해 만년로 제2공장 등과 표기 통일. _search_poi 스텁.
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


ROAD = '경기 화성시 만세구 양감면 정문송산로93번길 1'
OFFICIAL = ROAD + ' 보우테이프 제4공장'


@pytest.fixture
def poi_stub(monkeypatch):
    monkeypatch.setattr(ar, '_search_poi', lambda q: [('보우테이프 제4공장', ROAD)])
    monkeypatch.setattr(ar, '_extract_shop_candidates', lambda t: ['보우테이프'])


def test_customer_number_upgraded_to_official(poi_stub):
    # 고객 '4공장' → 공식명 '제4공장' 으로 승격 (중복 없음)
    v = ROAD + ' 보우테이프 4공장'
    assert ar._enrich_with_poi(v, '보우테이프 4공장') == OFFICIAL


def test_customer_je_number_unchanged(poi_stub):
    v = ROAD + ' 보우테이프 제4공장'
    assert ar._enrich_with_poi(v, '보우테이프 제4공장') == OFFICIAL


def test_no_number_appends_official(poi_stub):
    # 번호 미기재 → 공식 지점명(제4공장) 부착 (disambiguation)
    v = ROAD + ' 보우테이프'
    assert ar._enrich_with_poi(v, '보우테이프') == OFFICIAL


def test_no_double_je_no_dup(poi_stub):
    # 어떤 경우에도 '제4공장 4공장' / '제4공장 제4공장' 중복 없음
    for cust in ['보우테이프 4공장', '보우테이프 제4공장', '보우테이프']:
        out = ar._enrich_with_poi(ROAD + ' ' + cust, cust)
        assert out.count('공장') == 1, out


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
