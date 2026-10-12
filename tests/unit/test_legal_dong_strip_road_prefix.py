# -*- coding: utf-8 -*-
"""법정동 중복 제거가 같은 이름으로 시작하는 도로명을 자르지 않음 (ETC-ae8c97 / L-03981).

카카오 region_3depth '봉담읍 상리'가 도로명 '상리2길' 앞부분과 겹쳐
'화성 효행구 봉담읍 상리2길 55' → '화성 효행구 2길 55'로 잘리던 버그.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


def _fake(dong):
    return lambda q: {'road_address': {'region_3depth_name': dong}}


def test_road_named_after_legal_dong_kept(monkeypatch):
    monkeypatch.setattr(ar, '_kakao_search', _fake('봉담읍 상리'))
    a = '화성 효행구 봉담읍 상리2길 55 어반센터프라자 316호'
    assert ar._strip_redundant_legal_dong(a) == a


@pytest.mark.parametrize('dong, addr, expected', [
    ('성수동2가', '성동구 성수일로12길 52 성수동2가 3층', '성동구 성수일로12길 52 3층'),
    ('매산로3가', '수원 매산로 1 매산로 3가 2층', '수원 매산로 1 2층'),
    ('성수동2가', '성동구 성수동2가 성수일로12길 52', '성동구 성수일로12길 52'),
])
def test_standalone_legal_dong_still_stripped(monkeypatch, dong, addr, expected):
    monkeypatch.setattr(ar, '_kakao_search', _fake(dong))
    assert ar._strip_redundant_legal_dong(addr) == expected


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
