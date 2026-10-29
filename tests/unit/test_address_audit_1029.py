# -*- coding: utf-8 -*-
"""10-29 한 달치 주소 점검에서 나온 변환 결함.

- 지번+상호에서 부번 유실 (ETC-346b74)
- 숫자뿐인 카카오 건물명('101') 부착 (L-04130)
- 괄호 속 도로명+번지 반복 (L-04315)
- POI 이름의 법인 접두어 (L-04122)
- 폭 0 문자(\\u2060) 저장 (L-04167)
- 'B1층' 표기 통일 (L-04244)
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar
from dashboard.services.lead_helpers import _normalize_road_spacing as N


def _road_doc(addr, bldg='', main='', sub=''):
    return {'address_name': addr, 'address_type': 'ROAD_ADDR',
            'address': {'address_name': addr, 'main_address_no': main, 'sub_address_no': sub},
            'road_address': {'address_name': addr, 'building_name': bldg}}


def _jibun_doc(addr, main, sub=''):
    return {'address_name': addr, 'address_type': 'REGION_ADDR',
            'address': {'address_name': addr, 'main_address_no': main, 'sub_address_no': sub},
            'road_address': None}


@pytest.fixture
def kakao(monkeypatch):
    table = {}
    monkeypatch.setattr(ar, '_kakao_key', lambda: 'x')
    monkeypatch.setattr(ar, '_kakao_search', lambda q: table.get(q))
    monkeypatch.setattr(ar, '_search_poi', lambda q: ())
    return table


def test_jibun_sub_number_kept(kakao):
    kakao['관악구 신림동 463-88 서서울공조'] = _jibun_doc('서울 관악구 신림동 463', '463')
    kakao['관악구 신림동 463-88'] = _road_doc('서울 관악구 관천로 87', main='463', sub='88')
    r = ar.verify_address('관악구 신림동 463-88 서서울공조', '관악구 신림동 463-88 서서울공조')
    assert r[0].startswith('관악구 관천로 87')
    assert '463' not in r[0]


def test_jibun_sub_mismatch_rejected_when_no_tail(kakao):
    kakao['관악구 신림동 463-88'] = _jibun_doc('서울 관악구 신림동 463', '463')
    assert ar.verify_address('관악구 신림동 463-88', '관악구 신림동 463-88') is None


def test_jibun_matching_sub_untouched(kakao):
    kakao['관악구 신림동 463-88 서서울공조'] = _road_doc(
        '서울 관악구 관천로 87', main='463', sub='88')
    r = ar.verify_address('관악구 신림동 463-88 서서울공조', '관악구 신림동 463-88 서서울공조')
    assert r[0].startswith('관악구 관천로 87')


def test_numeric_building_name_not_appended(kakao):
    kakao['경기 고양시 일산동구 장항한강3로 80'] = _road_doc(
        '경기 고양시 일산동구 장항한강3로 80', bldg='101')
    r = ar.verify_address('경기 고양시 일산동구 장항한강3로 80 (장항동, 105) 13층',
                          '경기 고양시 일산동구 장항한강3로 80')
    assert ' 101 ' not in r[0] + ' '


@pytest.mark.parametrize('raw, expected', [
    ('남양주 천마산로 113 1층 천마산로 113', '남양주 천마산로 113 1층'),
    ('남양주 천마산로 113 1층 천마산로113', '남양주 천마산로 113 1층'),
    # 다른 번지·다른 도로는 그대로
    ('강남구 테헤란로 152 강남파이낸스센터 테헤란로 153', '강남구 테헤란로 152 강남파이낸스센터 테헤란로 153'),
    ('강서구 공항대로 248 대방빌딩 2층', '강서구 공항대로 248 대방빌딩 2층'),
])
def test_repeated_road_removed(raw, expected):
    assert ar._post_normalize_display(raw) == expected


def test_corp_prefix_stripped_from_poi(monkeypatch):
    monkeypatch.setattr(ar, '_kakao_search_poi', lambda q: ())
    monkeypatch.setattr(ar, '_naver_search_local', lambda q: (
        ('사회복지법인설원복지재단', '경기도 안양시 만안구 예술공원로117번길 32'),
        ('재단법인 아름다운재단', '서울 종로구 자하문로19길 6'),
        ('사단법인', ''),
    ))
    names = [pn for pn, _ in ar._search_poi('q')]
    assert names == ['설원복지재단', '아름다운재단', '사단법인']


@pytest.mark.parametrize('raw, expected', [
    ('인천 연수구 인천타워대로 257 C동 107호⁠', '인천 연수구 인천타워대로 257 C동 107호'),
    ('송도 ⁠아트포레⁠ 107호', '송도 아트포레 107호'),
    ('서초구 서초중앙로 142 삼하빌딩 B1층 101호', '서초구 서초중앙로 142 삼하빌딩 지하1층 101호'),
    ('강남구 테헤란로 152 SB1층', '강남구 테헤란로 152 SB1층'),   # 앞이 영문이면 미대상
])
def test_input_normalization(raw, expected):
    assert N(raw) == expected


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
