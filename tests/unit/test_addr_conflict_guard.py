# -*- coding: utf-8 -*-
"""전화 등록 같은 번호·다른 주소 dedup 가드 회귀 테스트 (2026-08-07 L-03367).

배경: JK 가 성남 등자로 리드(010-8247-4206)와 같은 번호로 센부동산(다른 현장)을
전화 등록 → 연락처 dedup 이 기존 성남 리드를 센부동산 값으로 덮어써 성남 소실.

가드: 같은 번호라도 방문 주소의 도로명+번지 core 가 둘 다 있고 서로 다르면
병합하지 않고 신규 발번(+⚠️ 경고). core 추출 실패(빈값)면 보수적으로 병합 허용
(수동+워크플로 이중입력·당근 enrichment 등 기존 dedup 케이스 보존).

_addr_core (순수 함수) + conflict 판정 규칙만 검증 — 네트워크·시트 없음.
"""
import sys
sys.path.insert(0, '.')

from dashboard.services.lead_sync import _addr_core


def _conflict(old_addr, new_addr):
    """가드가 병합을 막는 조건 = 양쪽 core 존재 + 불일치."""
    old_core, new_core = _addr_core(old_addr), _addr_core(new_addr)
    return bool(old_core and new_core and old_core != new_core)


# ── _addr_core: 행정구역 prefix·건물·상세 제거하고 도로명/동+번지 core 만 ──
def test_addr_core_drops_prefix_and_detail():
    assert _addr_core('성남 수정구 등자로 56') == '등자로56'
    assert _addr_core('수정구 등자로 56 3층') == '등자로56'
    assert _addr_core('등자로 56') == '등자로56'


def test_addr_core_road_and_building():
    assert _addr_core('서울 강남구 테헤란로 152 강남파이낸스센터') == '테헤란로152'
    assert _addr_core('남양주 진건읍 진건오남로 77 심미에셈빌 1층') == '진건오남로77'


def test_addr_core_jibun_dong():
    assert _addr_core('경기 부천시 오정동 810-1 원일테크노2 4층') == '오정동810-1'


def test_addr_core_empty_and_dash():
    assert _addr_core('') == ''
    assert _addr_core('-') == ''
    assert _addr_core('상호만 있고 주소 불명') == ''  # 도로명/동+번지 없음


# ── conflict 판정 ──
def test_conflict_different_site_blocks_merge():
    """L-03367 재현: 등자로 vs 판교로 = 다른 현장 → 병합 차단."""
    assert _conflict('성남 수정구 등자로 56', '성남 판교로 20') is True


def test_no_conflict_same_place_different_format():
    """같은 곳 다른 서식(prefix·층) → 병합 허용."""
    assert _conflict('성남 수정구 등자로 56', '등자로 56 3층') is False
    assert _conflict('등자로 56', '등자로 56') is False


def test_no_conflict_when_one_addr_missing():
    """한쪽 주소 없음(=enrichment 케이스) → 병합 허용."""
    assert _conflict('성남 수정구 등자로 56', '') is False
    assert _conflict('', '성남 판교로 20') is False


def test_no_conflict_when_core_unparseable():
    """core 추출 실패 → 비교 불가 → 보수적으로 병합 허용."""
    assert _conflict('상호만 있고 주소 불명', '판교로 20') is False
    assert _conflict('판교로 20', '동 이름만') is False
