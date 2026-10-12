# -*- coding: utf-8 -*-
"""'N가' 법정동(을지로3가·종로5가·성수동2가·명륜4가) 추출·표기 (2026-10-12).

'서울 중구 을지로3가 12'가 '을지로3'까지만 추출돼 '가 12'(지번)가 버려지고, 변환 단계에서
'을지로 3'(도로+번지처럼)으로 쪼개지던 갭. 실 리드 L-00301·L-02295·L-02975·L-03251 대상.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar
from dashboard.services.lead_helpers import extract_korean_address as X


@pytest.mark.parametrize('text, expected', [
    ('서울 중구 을지로3가 12', '서울 중구 을지로3가 12'),
    ('중구 을지로3가 296-27 2층', '중구 을지로3가 296-27 2층'),
    ('용산구 한강로2가 31', '용산구 한강로2가 31'),
    ('종로구 명륜4가 92 2층', '종로구 명륜4가 92 2층'),
    ('성동구 성수동2가 300-15', '성동구 성수동2가 300-15'),
    ('당산동1가 20', '당산동1가 20'),                   # 시·구 없는 지번(level7)
    ('중구 을지로3가', '중구 을지로3가'),               # 끝 '가'를 조사로 자르지 않음
])
def test_extract_keeps_ga_dong(text, expected):
    assert X(text)[0] == expected


@pytest.mark.parametrize('text, expected', [
    ('양천구 신정3동', '양천구 신정3동'),               # 기존 동 회귀 없음
    ('강남구 테헤란로 152 3층', '강남구 테헤란로 152 3층'),
    ('서울 중구 을지로 100', '서울 중구 을지로 100'),
])
def test_extract_no_regression(text, expected):
    assert X(text)[0] == expected


def test_enrich_does_not_split_ga_dong(monkeypatch):
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda addr, text: addr)
    t = '용산구 한강로2가 15-1 1층'
    assert ar._enrich_verified_address(t, t, t) == t


def test_enrich_still_spaces_road_number(monkeypatch):
    # 붙여쓴 도로+번지 공백 보강은 그대로 ('세월길2' → '세월길 2'), 가나다길도 유지
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda addr, text: addr)
    assert ar._enrich_verified_address('강서구 세월길2', '강서구 세월길2', None) == '강서구 세월길 2'
    t = '은평구 갈현로3가길 5'
    assert ar._enrich_verified_address(t, t, None) == t


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
