# -*- coding: utf-8 -*-
"""verify 성공 경로에서 원문 끝 다단어 상호(상호+지점명) 보존 (L-03811).

카카오 verify 가 도로+번지만 주면(건물 미등록) 리졸버는 raw tail 을 버리고 POI 보강에
의존하는데, step 1-b 상호 부착은 한 단어만 잡아 '국면당 공세점' 같은 다단어 상호가
POI 미매치 시 유실됐다. POI 이후 상호(첫 단어)가 결과에 전혀 없을 때만 원문 상호구 부착.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.fixture
def no_poi(monkeypatch):
    # _enrich_with_poi 를 항등으로 — 다단어 보존 블록만 격리 검증 (네트워크 차단).
    monkeypatch.setattr(ar, '_enrich_with_poi', lambda v, o: v)


def test_multiword_shop_preserved(no_poi):
    out = ar._enrich_verified_address(
        '강남구 테헤란로 152', '강남구 테헤란로 152 106호 국면당 공세점', None)
    assert out == '강남구 테헤란로 152 106호 국면당 공세점'


def test_three_word_shop_preserved(no_poi):
    out = ar._enrich_verified_address(
        '용인 기흥구 탑실로 34', '용인 기흥구 탑실로 34 2층 국면당 공세점 별관', None)
    assert out == '용인 기흥구 탑실로 34 2층 국면당 공세점 별관'


def test_no_double_when_shop_already_present(no_poi):
    # verified 에 이미 상호(첫 단어)가 있으면 중복 부착 안 함
    out = ar._enrich_verified_address(
        '강남구 테헤란로 152 국면당 역삼점', '강남구 테헤란로 152 106호 국면당 공세점', None)
    assert '국면당 공세점' not in out
    assert out.count('국면당') == 1


def test_single_word_untouched_by_new_block(no_poi):
    # 단어 1개 상호는 기존 step 1-b 가 처리 (새 블록은 다단어만)
    out = ar._enrich_verified_address(
        '강남구 테헤란로 152', '강남구 테헤란로 152 106호 국면당', None)
    assert out == '강남구 테헤란로 152 106호 국면당'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
