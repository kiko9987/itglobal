# -*- coding: utf-8 -*-
"""주소 추출 tail stop-word 회귀 테스트.

배경: extract_korean_address 가 주소 끝 '전화연락 부탁드립니다' 같은 문의 문구를
흡수해 '구로구 고척동 전화' 처럼 '전화' 가 주소에 남던 leak (2026-07-30 G5,
당근 raw 감사에서 관측). _STOP_WORDS 에 '전화' 추가로 차단.
pure 함수만 — 카카오/네트워크 미접촉.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.lead_helpers import extract_korean_address


class TestPhoneTailStrip:
    @pytest.mark.parametrize('raw, expected', [
        ('구로구 고척동 전화연락 부탁드립니다 01067749555', '구로구 고척동'),
        ('서울 강남구 역삼동 123-4 전화주세요', '서울 강남구 역삼동 123-4'),
        ('인천 서구 원당동 전화 부탁', '인천 서구 원당동'),
        # 주의: '전화번호' 는 '호' 가 핵심 정규식 호수 유닛으로 매칭돼 별도 이슈(미해결).
        #   이번 fix 범위는 확장부 leak('전화연락/전화주세요/전화 부탁') 한정.
    ])
    def test_phone_tail_removed(self, raw, expected):
        r = extract_korean_address(raw)
        assert r is not None
        assert r[0] == expected

    @pytest.mark.parametrize('raw', [
        # 정상 주소 — 회귀 방어 (건물/상호 tail 유지, '전화' 미포함)
        '서울 양천구 신목로4길 9 1층',
        '분당구 대왕판교로660 유스페리스1',
        '경기도 양평군 양평읍 십리길 19 5층',
    ])
    def test_normal_address_not_broken(self, raw):
        r = extract_korean_address(raw)
        assert r is not None
        # '전화' 가 없던 정상 주소는 앞부분(시/구/동)이 그대로 보존
        assert r[0].split()[0] in raw
        assert '전화' not in r[0]


from dashboard.services.address_resolver import (
    _strip_personal_name, _extract_building_tail,
)


class TestPersonalNameStripShopGuard:
    """상호·업종 접미어는 사람 이름으로 오인해 자르면 안 됨 (L-03475 남양가 양꼬치)."""

    @pytest.mark.parametrize('tail', [
        '그랑트윈타워A동 남양가 양꼬치',   # 양꼬치(양=성씨) 유지
        '강서구 홍반점',                  # 홍반점(홍) 유지
        '마곡 김밥',                     # 김밥(김) 유지
        '역삼 오리족발',                  # 족발 유지
    ])
    def test_shop_suffix_kept(self, tail):
        assert _strip_personal_name(tail) == tail

    @pytest.mark.parametrize('tail, expected', [
        ('그로브리조트 정승종', '그로브리조트'),  # 실제 사람 이름은 여전히 제거
        ('ABC빌딩 김지수', 'ABC빌딩'),
    ])
    def test_real_name_stripped(self, tail, expected):
        assert _strip_personal_name(tail) == expected


class TestLatinUnitUppercase:
    """동·호 앞 라틴 소문자 대문자화 (L-03475 그랑트윈타워a동 → A동)."""

    def test_a_dong_uppercased_and_shop_kept(self):
        r = _extract_building_tail('서울시 강서구 마곡동799-7 그랑트윈타워a동 남양가 양꼬치')
        assert 'A동' in r
        assert 'a동' not in r
        assert '양꼬치' in r  # 상호 유지


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
