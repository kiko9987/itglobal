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


from dashboard.services import address_resolver as _ar


class TestJoinedShopCandidates:
    """다단어 상호 결합 후보 (G1, L-03475)."""

    def test_joins_shop_words(self):
        r = _ar._joined_shop_candidates('서울시 강서구 마곡동799-7 그랑트윈타워a동 남양가 양꼬치')
        assert ('남양가양꼬치', '남양가 양꼬치') in r

    def test_addr_tokens_excluded(self):
        # 도로/번지/단위동 은 결합 대상 아님
        r = _ar._joined_shop_candidates('강서구 마곡중앙4로 10 그랑트윈타워a동 남양가 양꼬치')
        joined = [x[0] for x in r]
        assert '남양가양꼬치' in joined
        assert not any('동' in j or '로' in j for j in joined)


class TestPoiShopAdoption:
    """POI place_name 채택 (G1) — verified 실재 상호만, 정식명이 상호로 시작할 때만."""

    def _patch_poi(self, monkeypatch, results):
        monkeypatch.setattr(_ar, '_search_poi', lambda q: results)

    def test_adopts_official_place_name(self, monkeypatch):
        self._patch_poi(monkeypatch, [('남양가양꼬치 마곡점', '서울 강서구 마곡중앙4로 10')])
        out = _ar._enrich_with_poi(
            '강서구 마곡중앙4로 10 그랑트윈타워A동 남양가 양꼬치',
            '서울시 강서구 마곡동799-7 그랑트윈타워a동 남양가 양꼬치',
        )
        assert out == '강서구 마곡중앙4로 10 그랑트윈타워A동 남양가양꼬치 마곡점'

    def test_skip_when_place_name_not_startswith(self, monkeypatch):
        # 정식명이 우리 상호로 시작 안 하면 채택 X (다른 상호 오탐 방지)
        self._patch_poi(monkeypatch, [('엉뚱한식당 마곡점', '서울 강서구 마곡중앙4로 10')])
        raw = '강서구 마곡중앙4로 10 그랑트윈타워A동 남양가 양꼬치'
        assert _ar._enrich_with_poi(raw, '서울시 강서구 마곡동799-7 그랑트윈타워a동 남양가 양꼬치') == raw

    def test_skip_when_road_mismatch(self, monkeypatch):
        # 도로명 다르면 다른 위치 → 채택 X
        self._patch_poi(monkeypatch, [('남양가양꼬치 마곡점', '서울 강서구 다른대로 99')])
        raw = '강서구 마곡중앙4로 10 그랑트윈타워A동 남양가 양꼬치'
        assert _ar._enrich_with_poi(raw, '서울시 강서구 마곡동799-7 그랑트윈타워a동 남양가 양꼬치') == raw

    def test_skip_facility_second_word(self, monkeypatch):
        # place 두 번째 단어가 부속시설이면 채택 X
        self._patch_poi(monkeypatch, [('남양가양꼬치 주차장', '서울 강서구 마곡중앙4로 10')])
        raw = '강서구 마곡중앙4로 10 그랑트윈타워A동 남양가 양꼬치'
        assert _ar._enrich_with_poi(raw, '서울시 강서구 마곡동799-7 그랑트윈타워a동 남양가 양꼬치') == raw


class TestPoiFacilityFilter:
    """POI 상호 뒤 어느 단어든 부속시설이면 부착 skip (2026-07-30 근본 수정)."""

    @pytest.mark.parametrize('place_name, is_fac', [
        ('롯데마트 고양점 주차장', True),     # 주차장=3번째 (기존 버그: 통과)
        ('한국화훼농협 ATM 본점', True),      # ATM=2번째 (blacklist 신규)
        ('삼조빌딩 앞_옥외 102호', True),     # 옥외
        ('현대백화점 무역센터점 지하주차장', True),
        ('남양가양꼬치 마곡점', False),        # 정상 지점명
        ('마성떡볶이 논현역점', False),
        ('스타벅스 강남R점', False),
    ])
    def test_facility_detection(self, place_name, is_fac):
        assert _ar._poi_has_facility(place_name) is is_fac

    def test_enrich_skips_facility_at_pos3(self, monkeypatch):
        monkeypatch.setattr(_ar, '_search_poi',
                            lambda q: [('롯데마트 고양점 주차장', '고양 덕양구 충장로 150')])
        raw = '고양 덕양구 충장로 150 롯데마트 고양점 2층'
        assert _ar._enrich_with_poi(raw, raw) == raw  # 주차장 미부착

    def test_enrich_skips_atm(self, monkeypatch):
        monkeypatch.setattr(_ar, '_search_poi',
                            lambda q: [('한국화훼농협 ATM 본점', '고양 일산서구 대화로 362')])
        raw = '고양 일산서구 대화로 362 한국화훼농협'
        assert _ar._enrich_with_poi(raw, raw) == raw  # ATM 미부착


class TestMultiTokenDedup:
    """다토큰 near-dup 제거 (ETC-858578) — 정규화(공백·· 제거) 후 앞부분 substring."""

    def test_compound_building_dup_removed(self):
        # 카카오 '온수 어르신복지회관 ·보훈회관' 뒤 고객원문 '온수어르신복지회관' 중복 제거
        out = _ar._post_normalize_display(
            '구로구 부일로9길 111 온수 어르신복지회관 ·보훈회관 온수어르신복지회관')
        assert out == '구로구 부일로9길 111 온수 어르신복지회관 ·보훈회관'

    def test_exact_dup_removed(self):
        out = _ar._post_normalize_display('강남구 테헤란로 1 스타벅스강남점 스타벅스강남점')
        assert out == '강남구 테헤란로 1 스타벅스강남점'

    @pytest.mark.parametrize('addr', [
        '강서구 마곡중앙4로 10 그랑트윈타워A동 남양가양꼬치 마곡점',
        '수원 권선구 덕영대로 1205 미래타운빌딩 501호',
        '고양 일산서구 대화로 362 한국화훼농협 본점 케이플라워마트 대화점',
        '성남 분당구 대왕판교로 660 유스페이스1 지하 115호',
    ])
    def test_normal_unchanged(self, addr):
        assert _ar._post_normalize_display(addr) == addr

    def test_short_token_kept(self):
        # 4자 이하는 우연 substring 방지로 미검사 → '보훈회관'(4자) 유지
        out = _ar._post_normalize_display('구로구 부일로9길 111 온수 어르신복지회관 ·보훈회관')
        assert '보훈회관' in out


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
