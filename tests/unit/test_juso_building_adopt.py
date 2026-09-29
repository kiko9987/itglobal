# -*- coding: utf-8 -*-
"""지역 같을 때 juso 건물명(bdNm)이 고객 텍스트에 있으면 juso 기반 교정 (L-04119).

공백 건물명 '반도 유스퀘어'를 extract 가 '반도'로 쪼개 enrich 가 '반도 405호 반도유스퀘어'
처럼 중복·어순뒤바뀜 → juso bdNm 이 텍스트에 있을 때만 juso 기반으로 교정. bdNm 이 텍스트에
없으면(봉이랜드=juso 등록 다른 tenant) 미교정(오부착 방지). 네트워크 monkeypatch.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.fixture
def _stub(monkeypatch):
    monkeypatch.setattr(ar, 'verify_address', lambda *a, **k: None)
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda addr, text: addr)
    for fn in ['_try_poi_fallback', '_road_poi_fallback', '_dong_building_poi_fallback',
               '_jibun_road_fallback', '_partner_alias_lookup']:
        monkeypatch.setattr(ar, fn, lambda *a, **k: None)


def test_juso_building_in_text_adopted(_stub, monkeypatch):
    # bdNm '반도 유스퀘어'가 텍스트에 있음 → juso 기반 채택(어순·중복 교정)
    monkeypatch.setattr(ar, '_juso_fallback',
                        lambda *a, **k: ('고양 덕양구 삼송로 12 반도 유스퀘어', 'road'))
    raw = '고양시 덕양구 삼송로 12 반도 유스퀘어 405호'
    rx = '고양시 덕양구 삼송로 12 반도'   # extract 가 '반도'만 잡은 상태
    addr, lv = ar.resolve_address(raw, rx, 'level3')
    assert lv == 'verified'
    assert addr == '고양 덕양구 삼송로 12 반도 유스퀘어 405호'
    assert addr.count('반도') == 1   # 중복 없음


def test_juso_building_not_in_text_not_adopted(_stub, monkeypatch):
    # bdNm '봉이랜드'가 텍스트에 없음(다른 tenant) → 미채택, 고객 상호 유지
    monkeypatch.setattr(ar, '_juso_fallback',
                        lambda *a, **k: ('성북구 종암로 129 봉이랜드', 'road'))
    raw = '성북구 종암로 129 3층 호랑이신경외과'
    rx = '성북구 종암로 129 3층 호랑이신경외과'
    addr, lv = ar.resolve_address(raw, rx, 'level3')
    assert lv == 'verified'
    assert '봉이랜드' not in addr
    assert '호랑이신경외과' in addr


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
