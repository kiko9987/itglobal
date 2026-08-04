# -*- coding: utf-8 -*-
"""POI endswith 매치 접두 가드 회귀 테스트 (2026-08-04 L-03530).

_enrich_with_poi 의 endswith 매치는 '접두 지역명 포함'(김포한강듀클래스) 케이스용인데,
공백으로 분리된 다른 업체 접두('슈퍼스타 어반322')까지 허용해 오부착했음. 붙은 접두만
허용(공백 분리 접두 차단). _search_poi 를 monkeypatch (네트워크 없이 매치 로직 검증).
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services import address_resolver as ar


def test_space_separated_tenant_not_appended(monkeypatch):
    """cand 뒤 '공백 접두' POI(다른 업체 '슈퍼스타 어반322')는 endswith 부착 안 함."""
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('슈퍼스타 어반322', '서울 영등포구 양평로25길 8')])
    r = ar._enrich_with_poi('영등포구 양평로25길 8 URBAN322',
                            '영등포구 양평로25길 8 어반322')
    assert '슈퍼스타' not in r


def test_exact_match_still_appends(monkeypatch):
    """정확 매치(place_name == cand)는 정상 부착 (회귀 방지)."""
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('어반322', '서울 영등포구 양평로25길 8')])
    r = ar._enrich_with_poi('영등포구 양평로25길 8 URBAN322',
                            '영등포구 양평로25길 8 어반322')
    assert '어반322' in r


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
