# -*- coding: utf-8 -*-
"""건물명 앞 단어가 지역·도로명과 겹쳐도 잘리지 않음 (L-04230 / L-03754).

①verify _compose 의 tail 중복 제거가 base 전체(지역·도로 포함)와 비교해, 건물명
'이천 센트레빌 레이크뷰'의 '이천'(지역 '이천'), '위례 한화 오벨리스크'의 '위례'(도로
'위례광장로')를 잘랐음. ②_enrich_with_poi 가 '오벨리스크'를 POI '위례오벨리스크'로 바꿔
이미 있는 '위례'와 중복.
"""
import sys
sys.path.insert(0, '.')

import dashboard.services.address_resolver as ar


def _kakao(addr, bldg):
    return lambda q: {'address_type': 'ROAD_ADDR', 'address_name': addr,
                      'road_address': {'address_name': addr, 'building_name': bldg}}


def test_tail_first_word_same_as_region_kept(monkeypatch):
    monkeypatch.setattr(ar, '_kakao_search', _kakao('경기 이천시 애련정로 61-10', '이천 센트레빌 레이크뷰'))
    monkeypatch.setattr(ar, '_kakao_key', lambda: 'x')
    monkeypatch.setattr(ar, '_search_poi', lambda q: ())
    t = '이천 애련정로 61-10 이천 센트레빌 레이크뷰 1층 112호'
    addr, _ = ar.verify_address(t, None)
    assert '이천 센트레빌 레이크뷰 1층 112호' in addr


def test_tail_first_word_in_road_name_kept(monkeypatch):
    monkeypatch.setattr(ar, '_kakao_search', _kakao('경기 성남시 수정구 위례광장로 104', ''))
    monkeypatch.setattr(ar, '_kakao_key', lambda: 'x')
    monkeypatch.setattr(ar, '_search_poi', lambda q: ())
    t = '성남 수정구 위례광장로 104 위례 한화 오벨리스크 2238호'
    addr, _ = ar.verify_address(t, None)
    assert '위례 한화 오벨리스크' in addr


def test_poi_prefix_not_added_when_already_present(monkeypatch):
    monkeypatch.setattr(ar, '_search_poi', lambda q: (
        ('위례오벨리스크', '경기 성남시 수정구 위례광장로 104'),))
    v = '성남 수정구 위례광장로 104 위례 한화 오벨리스크 2238호'
    assert ar._enrich_with_poi(v, v) == v


def test_poi_prefix_still_added_when_absent(monkeypatch):
    # 접두가 원문에 없으면 기존대로 정식명 보강 (L-03316 한강듀클래스 → 김포한강듀클래스)
    monkeypatch.setattr(ar, '_search_poi', lambda q: (
        ('김포한강듀클래스', '경기 김포시 김포한강1로 1'),))
    v = '김포 김포한강1로 1 한강듀클래스'
    assert ar._enrich_with_poi(v, v).endswith('김포한강듀클래스')


def test_existing_branch_not_duplicated(monkeypatch):
    # 원문 지점명('양재동점')이 있으면 POI 다른 표기('양재점') 미부착 (ETC-841a3c)
    monkeypatch.setattr(ar, '_search_poi', lambda q: (
        ('산들해 양재점', '서울특별시 서초구 강남대로 213 엘타워 지하1층'),))
    v = '서초구 강남대로 213 스포타임 엘타워 지하 1층 산들해 양재동점'
    assert ar._enrich_with_poi(v, v) == v


def test_branch_still_appended_when_absent(monkeypatch):
    monkeypatch.setattr(ar, '_search_poi', lambda q: (
        ('마성떡볶이 논현역점', '서울 강남구 학동로 1'),))
    v = '강남구 학동로 1 마성떡볶이'
    assert ar._enrich_with_poi(v, v).endswith('마성떡볶이 논현역점')
