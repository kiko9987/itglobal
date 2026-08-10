# -*- coding: utf-8 -*-
"""방문 주소 인라인 노트 분리 회귀 테스트 (2026-08-04 L-03524).

매니저가 방문 주소 필드에 한 줄로 붙여쓴 지시 노트('(현장은 3층) YG 소통 하세요')를
상담 내용으로 이동. 보수적 2-신호(정중형 동사어미 + 노트성 괄호)라 상호·관리사무소·
정상 주소는 건드리지 않음. split_address_notes 는 순수 함수 (네트워크 없음).
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.address_resolver import split_address_notes as S


def test_trailing_instruction_and_note_paren():
    """지시문(하세요) + 노트성 괄호(현장) 동시 분리 (L-03524)."""
    clean, note = S('서초구 강남대로 97길 32 한영빌딩 4층 (현장은 3층) YG 소통 하세요')
    assert clean == '서초구 강남대로 97길 32 한영빌딩 4층'
    assert '현장은 3층' in note and 'YG 소통 하세요' in note


def test_trailing_instruction_yomang():
    """'~요망' 지시문 분리 (L-03361형)."""
    assert S('서초구 학동로 102 방문 전 연락 요망') == ('서초구 학동로 102', '방문 전 연락 요망')


def test_note_paren_direction():
    """방향어 괄호 분리, '지하'(층 표기)는 주소로 유지 (L-03442형)."""
    clean, note = S('성북구 성북로 91 지하(정면 바라보고 오른쪽 문)')
    assert clean == '성북구 성북로 91 지하'
    assert note == '정면 바라보고 오른쪽 문'


@pytest.mark.parametrize('addr', [
    '김포 유현로 200 풍무푸르지오아파트 관리사무소',   # 관리사무소=방문대상, 유지
    '과천 광명로 181 후문 서울랜드 산타레스토랑',       # 상호, 유지
    '강남구 학동로 102 마성떡볶이',                      # 상호, 유지
    '인천 계양구 봉오대로 712 1층 힙춘향마라 계양점',    # 지점명, 유지
    '서울 강서구 마곡중앙로 161-8 그랑트윈타워 A동 남양가양꼬치 마곡점',
    '영등포구 양평로25길 8 URBAN322',
])
def test_no_false_positive(addr):
    """상호·관리사무소·정상 주소는 노트 분리 안 함 (오제거 방지)."""
    clean, note = S(addr)
    assert clean == addr and note == ''


def test_slash_time_note_moved():
    """'/' 뒤 시간·지시 노트 상담 이관 (ETC-4c47a2)."""
    clean, note = S('영등포구 경인로 775 에이스하이테크시티 2동 / 오전 7시 현장설명 / 오후 1시 철거제품 회수 및 결산')
    assert clean == '영등포구 경인로 775 에이스하이테크시티 2동'
    assert note == '오전 7시 현장설명 / 오후 1시 철거제품 회수 및 결산'


def test_slash_time_note_single():
    """'/' 뒤 단일 시간 노트 (ETC-1765ea)."""
    clean, note = S('서울 양천구 목동동로 309 행복한 백화점 5F / 오전 7시')
    assert clean == '서울 양천구 목동동로 309 행복한 백화점 5F'
    assert note == '오전 7시'


def test_slash_address_continuation_not_split():
    """'/' 로 이은 주소 연속(301호/302호)은 분리 안 함 — 시간·지시 신호 없음."""
    assert S('강남구 테헤란로 152 삼성빌딩 301호/302호') == (
        '강남구 테헤란로 152 삼성빌딩 301호/302호', '')


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
