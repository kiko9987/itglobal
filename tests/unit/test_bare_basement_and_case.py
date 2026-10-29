# -*- coding: utf-8 -*-
"""단독 'B2' 지하층 인식 (ETC-b30611) + 결합형 POI 흡수 영문 대소문자 무시 (L-04404)."""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar
from dashboard.services.lead_helpers import _normalize_road_spacing as N


@pytest.mark.parametrize('raw, expected', [
    ('수원 영통구 도청로66번길 6 힐스테이트광교중앙역퍼스트 B2',
     '수원 영통구 도청로66번길 6 힐스테이트광교중앙역퍼스트 지하2층'),
    ('강남구 도곡로 126 오복빌딩 B1 위너골프연습장', '강남구 도곡로 126 오복빌딩 지하1층 위너골프연습장'),
    ('서초구 서초대로73길 9 타임빌딩 B1, 위플레이스', '서초구 서초대로73길 9 타임빌딩 지하1층, 위플레이스'),
    ('강남구 테헤란로 152 B2B 솔루션', '강남구 테헤란로 152 B2B 솔루션'),   # 영문 단어
    ('강남구 테헤란로 152 B-104호', '강남구 테헤란로 152 B-104호'),       # 호수 표기
    ('강남구 테헤란로 152 B201호', '강남구 테헤란로 152 B201호'),
])
def test_bare_basement(raw, expected):
    assert N(raw) == expected


def test_joined_absorb_case_insensitive(monkeypatch):
    monkeypatch.setattr(ar, '_search_poi', lambda q: (
        ('더샵송도트리플타워West오피스텔', '인천 연수구 송도과학로28번길 50'),))
    v = '인천 연수구 송도과학로28번길 50 더샵 송도트리플타워 WEST 142호'
    assert ar._enrich_with_poi(v, v) == v


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
