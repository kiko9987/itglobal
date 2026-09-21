# -*- coding: utf-8 -*-
"""행안부 juso 일시 실패는 lru_cache 에 캐시하지 않음 (L-04066, 카카오 L-03476 대칭).

juso 키가 트래픽 초과 시 errorCode E0001 '승인되지 않은 KEY'(승인된 키도 발생, 곧 회복)나
timeout/5xx 를 내는데, 기존엔 그 빈 결과 ()를 영구 캐시해 실재 주소가 재시작까지 미검증으로
굳음 → '주소 확인 필요' 오배지 위험. 이제 일시 실패는 _JusoTransientError 로 raise 되어
lru_cache 미오염, _juso_search 래퍼가 ()로 흡수. 네트워크는 monkeypatch 로 차단(hermetic).
"""
import sys
sys.path.insert(0, '.')

import json
import socket
import pytest
import dashboard.services.address_resolver as ar


def _resp(payload):
    class R:
        def __enter__(s): return s
        def __exit__(s, *a): return False
        def read(s): return json.dumps(payload).encode('utf-8')
    return R()


_OK = {'results': {'common': {'errorCode': '0', 'totalCount': '1'},
                   'juso': [{'roadAddr': '서울특별시 노원구 상계로1길 62-20 (상계동)',
                             'jibunAddr': '서울특별시 노원구 상계동 581-111',
                             'bdNm': '동한빌딩'}]}}
_E0001 = {'results': {'common': {'errorCode': 'E0001',
                                 'errorMessage': '승인되지 않은 KEY 입니다.'}, 'juso': None}}
_EMPTY = {'results': {'common': {'errorCode': '0', 'totalCount': '0'}, 'juso': None}}
_BADQ = {'results': {'common': {'errorCode': 'E0009'}, 'juso': None}}  # 검색어 오류 = 영구


@pytest.fixture(autouse=True)
def _key_and_clear(monkeypatch):
    monkeypatch.setattr(ar, '_juso_key', lambda: 'DUMMY')
    monkeypatch.setattr('time.sleep', lambda *a: None)
    ar._juso_search_cached.cache_clear()
    yield
    ar._juso_search_cached.cache_clear()


def test_e0001_transient_raises_and_not_cached(monkeypatch):
    monkeypatch.setattr(ar.urllib.request, 'urlopen', lambda *a, **k: _resp(_E0001))
    with pytest.raises(ar._JusoTransientError):
        ar._juso_get_json('q')
    assert ar._juso_search('q') == ()                       # 래퍼가 흡수
    assert ar._juso_search_cached.cache_info().currsize == 0  # 캐시 미오염


def test_timeout_transient_raises(monkeypatch):
    monkeypatch.setattr(ar.urllib.request, 'urlopen',
                        lambda *a, **k: (_ for _ in ()).throw(socket.timeout()))
    with pytest.raises(ar._JusoTransientError):
        ar._juso_get_json('q')
    assert ar._juso_search('q') == ()
    assert ar._juso_search_cached.cache_info().currsize == 0


def test_ok_result_returned_and_cached(monkeypatch):
    monkeypatch.setattr(ar.urllib.request, 'urlopen', lambda *a, **k: _resp(_OK))
    r = ar._juso_search('q')
    assert r and r[0][2] == '동한빌딩'
    assert ar._juso_search_cached.cache_info().currsize == 1  # 정상 결과는 캐시


def test_empty_is_permanent_cached(monkeypatch):
    # errorCode 0 + 0건 = 정상 무매치 → () 영구 캐시(재조회 무의미)
    monkeypatch.setattr(ar.urllib.request, 'urlopen', lambda *a, **k: _resp(_EMPTY))
    assert ar._juso_search('q') == ()
    assert ar._juso_search_cached.cache_info().currsize == 1


def test_search_term_error_permanent_cached(monkeypatch):
    # E0009(검색어 오류) 등 다른 errorCode = 영구 → () 캐시
    monkeypatch.setattr(ar.urllib.request, 'urlopen', lambda *a, **k: _resp(_BADQ))
    assert ar._juso_search('q') == ()
    assert ar._juso_search_cached.cache_info().currsize == 1


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
