# -*- coding: utf-8 -*-
"""카카오 조회 재시도 + 실패-비캐시 회귀 테스트 (2026-07-31 L-03476).

배경: 카카오 API 일시 실패(429 rate-limit/timeout/5xx)가 lru_cache 에 None 으로
캐시돼, 멀쩡한 주소가 재시작 전까지 계속 non-verified → '주소 확인 필요' 오배지.
_kakao_get_json 이 일시 실패 시 _KakaoTransientError raise → lru_cache 미캐시 →
다음 호출 재시도. 성공/유효-빈결과만 캐시.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services import address_resolver as ar


class TestKakaoRetryNoCacheFailure:
    def test_transient_not_cached_retries(self, monkeypatch):
        calls = []

        def fake(url):
            calls.append(1)
            if len(calls) == 1:
                raise ar._KakaoTransientError()
            return {'documents': [{'road_address': {'address_name': '인천 테스트로 1'},
                                   'address_name': '인천 테스트로 1'}]}
        monkeypatch.setattr(ar, '_kakao_get_json', fake)
        ar._kakao_search_cached.cache_clear()
        assert ar._kakao_search('쿼리A_uniq') is None       # 일시 실패 → None
        assert ar._kakao_search('쿼리A_uniq') is not None    # 재시도 → 성공
        assert len(calls) == 2                                # 실패가 캐시 안 됨

    def test_success_cached_once(self, monkeypatch):
        calls = []

        def fake(url):
            calls.append(1)
            return {'documents': [{'road_address': {'address_name': 'X'}, 'address_name': 'X'}]}
        monkeypatch.setattr(ar, '_kakao_get_json', fake)
        ar._kakao_search_cached.cache_clear()
        ar._kakao_search('쿼리B_uniq'); ar._kakao_search('쿼리B_uniq')
        assert len(calls) == 1                                # 성공은 캐시 → 1회

    def test_empty_result_cached(self, monkeypatch):
        calls = []

        def fake(url):
            calls.append(1)
            return {'documents': []}                           # 유효-빈결과
        monkeypatch.setattr(ar, '_kakao_get_json', fake)
        ar._kakao_search_cached.cache_clear()
        assert ar._kakao_search('쿼리C_uniq') is None
        assert ar._kakao_search('쿼리C_uniq') is None
        assert len(calls) == 1                                # 빈결과도 유효 → 캐시

    def test_poi_transient_not_cached(self, monkeypatch):
        calls = []

        def fake(url):
            calls.append(1)
            if len(calls) == 1:
                raise ar._KakaoTransientError()
            return {'documents': [{'place_name': 'P', 'road_address_name': 'R'}]}
        monkeypatch.setattr(ar, '_kakao_get_json', fake)
        ar._kakao_search_poi_cached.cache_clear()
        assert ar._kakao_search_poi('쿼리D_uniq') == ()
        assert ar._kakao_search_poi('쿼리D_uniq') == (('P', 'R'),)
        assert len(calls) == 2


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
