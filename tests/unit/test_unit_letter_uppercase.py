# -*- coding: utf-8 -*-
"""건물 구역/유닛 앞 라틴 소문자 대문자화 (L-03475 · L-04038).

'b-104호'(글자+대시+숫자+호)·'a101호'·'타워a동' 등 단위 접미 앞 라틴 1~2자는 대문자로.
건물명 중간 소문자(e편한세상·kt대리점)는 보존. 순수함수.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import _upper_unit_letter as U


@pytest.mark.parametrize('addr, expected', [
    ('강남구 삼성로72길 12 대치푸르지오써밋 상가 b-104호',
     '강남구 삼성로72길 12 대치푸르지오써밋 상가 B-104호'),   # L-04038 대시형
    ('상가 b-104호', '상가 B-104호'),
    ('타워a동', '타워A동'),          # L-03475 직결형
    ('건물 b호', '건물 B호'),
    ('a101호', 'A101호'),            # 숫자 직결형
    ('다산지금로 202 b동 5F 0001호', '다산지금로 202 B동 5F 0001호'),
])
def test_unit_letter_uppercased(addr, expected):
    assert U(addr) == expected


@pytest.mark.parametrize('addr', [
    'e편한세상 101동',              # e 는 건물명 일부(편 앞) — 보존
    '판교로 393 kt대리점 b동',      # kt 는 대리점 앞 — 보존 (b동만 대문자)
    '강남구 테헤란로 152 3층',      # 라틴 없음
])
def test_building_name_letters_preserved(addr):
    out = U(addr)
    # kt·e 는 소문자 유지
    assert 'kt' in out or 'KT' not in addr
    assert 'e편한' in out or 'e편한' not in addr


def test_kt_preserved_but_b_upper():
    assert U('판교로 393 kt대리점 b동') == '판교로 393 kt대리점 B동'


def test_e_prefix_preserved():
    assert U('e편한세상 101동') == 'e편한세상 101동'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
