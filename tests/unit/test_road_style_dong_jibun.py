# -*- coding: utf-8 -*-
"""도로명형 법정동('…로N가') 지번 주소의 juso verify 회귀 (L-03779 후속).

종로1~6가·을지로1~7가·충무로1~5가·충정로2·3가·신문로1·2가·남대문로1~5가 처럼 이름과
'가' 사이에 숫자가 낀 법정동('충정로'+'3'+'가')의 지번 주소를, _JUSO_JIBUN_CORE_RE 가
'<한글>동/리/가 <지번>'만 기대해 코어로 인식 못 해 juso 도달 실패 → raw 폴백하던 버그.

수정: 코어 정규식이 이름과 접미('동/리/가') 사이 숫자를 흡수(\\d*)하도록 확장.
네트워크 없이 _juso_search_cached·_juso_key 를 스텁으로 대체해 경계 함수만 검증.
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


class TestCoreRegex:
    def test_road_style_dong_jibun_core_extracted(self):
        """'충정로3가 183-1' 이 지번 코어로 잡혀야 함 (버그의 핵심 지점)."""
        m = ar._JUSO_JIBUN_CORE_RE.search('서울 서대문구 충정로3가 183-1 1층')
        assert m is not None and m.group(1) == '충정로3가 183-1'

    @pytest.mark.parametrize('text,expected', [
        ('서울 종로구 종로1가 24', '종로1가 24'),
        ('서울 중구 을지로3가 275', '을지로3가 275'),
        ('서울 중구 남대문로5가 120-3', '남대문로5가 120-3'),
    ])
    def test_various_road_style_dong(self, text, expected):
        m = ar._JUSO_JIBUN_CORE_RE.search(text)
        assert m is not None and m.group(1) == expected

    def test_plain_dong_core_still_extracted(self):
        """숫자 없는 일반 '…동' 지번은 그대로 잡혀야 함 (회귀 방지)."""
        m = ar._JUSO_JIBUN_CORE_RE.search('인천 계양구 효성동 66-16')
        assert m is not None and m.group(1) == '효성동 66-16'

    def test_road_style_dong_not_matched_by_road_pattern(self):
        """'충정로3가'(법정동)은 도로명 패턴엔 안 잡혀야 함 (오매칭 유지 방지)."""
        assert ar._ROAD_PATTERN.search('서울 서대문구 충정로3가 183-1 1층') is None


class TestJusoVerify:
    def test_충정로3가_jibun_verified(self, juso_stub):
        """카카오 0건이어도 juso 로 도로명 확인 → jibun 매치."""
        juso_stub([('서울특별시 서대문구 경기대로 38 (충정로3가)',
                    '서울특별시 서대문구 충정로3가 183-1', '')])
        hit = ar._juso_fallback('서울 서대문구 충정로3가 183-1 1층',
                                '서대문구 충정로3가 183-1 1층')
        assert hit is not None and hit[1] == 'jibun'
        assert hit[0] == '서대문구 경기대로 38'

    def test_plain_dong_jibun_still_verified(self, juso_stub):
        """일반 '…동' 지번의 juso 확인이 여전히 동작 (회귀 방지)."""
        juso_stub([('인천광역시 계양구 효성로 123 (효성동)',
                    '인천광역시 계양구 효성동 66-16', '')])
        hit = ar._juso_fallback('인천 계양구 효성동 66-16', '계양구 효성동 66-16')
        assert hit is not None and hit[1] == 'jibun'
        assert '효성로 123' in hit[0]

    def test_road_name_input_uses_road_kind(self, juso_stub):
        """도로명 입력('경기대로 38')은 road 경로로 verified (jibun 오인 안 함)."""
        juso_stub([('서울특별시 서대문구 경기대로 38 (충정로3가)',
                    '서울특별시 서대문구 충정로3가 183-1', '')])
        hit = ar._juso_fallback('서울 서대문구 경기대로 38', '서대문구 경기대로 38')
        assert hit is not None and hit[1] == 'road'
        assert '경기대로 38' in hit[0]


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
