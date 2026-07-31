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


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
