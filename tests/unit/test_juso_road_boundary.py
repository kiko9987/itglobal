# -*- coding: utf-8 -*-
"""_juso_fallback 도로명 시작 경계 — 부분문자열 오매칭 방지 (L-03686).

'판교로 393' 이 '대왕판교로 393'(완전 다른 도로) 안에 부분문자열로 매칭돼 없는 주소를
verified 로 오승격하던 버그. 도로/동 '이름'이 결과의 온전한 공백 토큰일 때만 통과하도록
강화. 네트워크 없이 _juso_search_cached 를 스텁으로 대체해 검증.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.fixture
def juso_stub(monkeypatch):
    """_juso_search_cached 를 고정 결과로 대체 + 키 존재 위장."""
    def _install(results):
        monkeypatch.setattr(ar, '_juso_key', lambda: 'STUB')
        monkeypatch.setattr(ar, '_juso_search_cached', lambda q: tuple(results))
    return _install


def test_road_substring_rejected(juso_stub):
    # 행안부가 '판교로 393' 을 '대왕판교로 393'(다른 도로)로 fuzzy 매칭 → 거절해야 함
    juso_stub([('경기도 성남시 분당구 대왕판교로 393 (백현동)',
                '경기도 성남시 분당구 백현동 407-3', '')])
    assert ar._juso_fallback('성남 분당구 판교로 393', '성남 분당구 판교로 393') is None


def test_exact_road_accepted(juso_stub):
    # 도로명이 온전한 토큰으로 일치하면 통과
    juso_stub([('경기도 성남시 분당구 대왕판교로 393 (백현동)',
                '경기도 성남시 분당구 백현동 407-3', '')])
    hit = ar._juso_fallback('성남 분당구 대왕판교로 393', '성남 분당구 대왕판교로 393')
    assert hit is not None and hit[1] == 'road'
    assert '대왕판교로 393' in hit[0]


def test_real_gil_road_accepted(juso_stub):
    # 카카오 미인덱싱 실재 도로(백제고분로19길)는 그대로 통과 (회귀 방지)
    juso_stub([('서울특별시 송파구 백제고분로19길 13 (잠실동)',
                '서울특별시 송파구 잠실동 237-5', '')])
    hit = ar._juso_fallback('송파구 백제고분로19길 13', '송파구 백제고분로19길 13')
    assert hit is not None and hit[1] == 'road'
    assert '백제고분로19길 13' in hit[0]


def test_shorter_road_not_matched_by_longer_input(juso_stub):
    # 역방향도 안전: '대왕판교로'는 결과 '판교로'에 매칭되면 안 됨
    juso_stub([('경기도 성남시 분당구 판교로 50 (삼평동)',
                '경기도 성남시 분당구 삼평동 680', '')])
    assert ar._juso_fallback('성남 분당구 대왕판교로 50', '성남 분당구 대왕판교로 50') is None


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
