# -*- coding: utf-8 -*-
"""캔버스2 중복 배정 감지 회귀 테스트 (2026-08-04).

같은 건이 여러 담당/섹션에 걸리면 둘 다 DM 나가는 것 방지 — 5시 체크·/일정확인에
경고로 노출(감지·플래그만, 배정/DM 은 안 건드림). 매니저가 캔버스 정리하도록 유도.
  - phone 있는 배정: 같은 전화가 다른 담당에 2건+ → 중복
  - phone 없는 태스크(등기발송): 같은 주소가 개인섹션 배정 + 온라인 태스크 양쪽 → 중복
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.visit_assignment_sync import _detect_assignment_duplicates


def test_phone_duplicate_across_assignees():
    """같은 전화가 서로 다른 담당에 2건 → 중복."""
    asg = [
        {'phone_digits': '01011112222', 'assign': ['MW'], 'address': '서초구 A'},
        {'phone_digits': '01011112222', 'assign': ['SJ'], 'address': '서초구 A'},
    ]
    dups = _detect_assignment_duplicates(asg, [])
    assert any(d['kind'] == 'phone' and d['key'] == '01011112222'
               and d['assignees'] == ['MW', 'SJ'] for d in dups)


def test_address_duplicate_assignment_and_online_task():
    """등기발송: 같은 주소가 개인섹션 배정(TH+SD) + 온라인 태스크 → 중복 (L-03485형 X, 등기)."""
    asg = [{'phone_digits': '', 'assign': ['TH', 'SD'], 'address': '근처 가까운 우체국'}]
    onl = [{'addr': '근처 가까운 우체국', 'raw': 'JK 점심때 (SB) ...'}]
    dups = _detect_assignment_duplicates(asg, onl)
    assert any(d['kind'] == 'address' and '우체국' in d['key']
               and set(d['assignees']) == {'TH+SD', '온라인'} for d in dups)


def test_no_dup_covisit_single_line():
    """공동 방문(TH+SD 한 줄)은 중복 아님."""
    asg = [{'phone_digits': '01033334444', 'assign': ['TH', 'SD'], 'address': 'X'}]
    assert _detect_assignment_duplicates(asg, []) == []


def test_no_dup_different_phones_same_building():
    """같은 건물 다른 호수(전화 다름)는 중복 아님 (별개 방문)."""
    asg = [
        {'phone_digits': '01011110000', 'assign': ['MW'], 'address': '강남빌딩 2층'},
        {'phone_digits': '01022220000', 'assign': ['SJ'], 'address': '강남빌딩 3층'},
    ]
    assert _detect_assignment_duplicates(asg, []) == []


def test_no_dup_same_assignee_same_addr():
    """같은 주소 같은 담당(phone없음)은 중복 아님(uniq 1)."""
    asg = [{'phone_digits': '', 'assign': ['SD'], 'address': '우체국'}]
    onl = [{'addr': '우체국', 'raw': 'SD ...'}]
    # 온라인 태스크는 '온라인' 라벨이라 SD 와 다름 → 이 케이스는 중복으로 잡힘.
    # 순수 동일담당 확인: 배정 2건 모두 SD
    asg2 = [{'phone_digits': '', 'assign': ['SD'], 'address': '우체국'},
            {'phone_digits': '', 'assign': ['SD'], 'address': '우체국'}]
    assert _detect_assignment_duplicates(asg2, []) == []


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
