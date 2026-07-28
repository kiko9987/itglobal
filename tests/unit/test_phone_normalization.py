#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""전화번호 앞자리 0 탈락 복원·정합성 판정 회귀 테스트.

배경: 구글시트 셀이 숫자 서식이면 '01091501411' 입력이 1091501411(앞 0 탈락)로
저장됨. 과거 ETC-690f80 이 이 형태('1091501411')로 저장돼 캔버스2 매칭이 조용히
실패 → 담당자 방문 배정 누락 사고. 그 방지책의 회귀 테스트.

pure 모듈(lead_helpers)만 import — Redis/Flask 미접촉.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.lead_helpers import (
    normalize_phone,
    restore_leading_zero,
    is_valid_phone,
    is_valid_phone_digits,
)


class TestRestoreLeadingZero:
    """앞자리 0 탈락 복원 — 유효 번호가 되는 경우만 복원."""

    @pytest.mark.parametrize('stripped, restored', [
        ('1091501411', '01091501411'),   # 실제 사고 케이스 (휴대폰)
        ('1044445555', '01044445555'),   # 휴대폰 010
        ('1198765432', '01198765432'),   # 011 (11자리)
        ('212345678', '0212345678'),     # 서울 02 (10자리)
        ('25581105', '025581105'),       # 서울 02 (9자리)
        ('317771234', '0317771234'),     # 경기 031
        ('7012345678', '07012345678'),   # 070
    ])
    def test_repairs_stripped(self, stripped, restored):
        assert restore_leading_zero(stripped) == restored

    @pytest.mark.parametrize('already_ok', [
        '01091501411', '0212345678', '025581105', '07012345678',
    ])
    def test_leaves_valid_untouched(self, already_ok):
        assert restore_leading_zero(already_ok) == already_ok

    @pytest.mark.parametrize('garbage', [
        '12345',          # 너무 짧음
        '9876543',        # 복원해도 유효 아님
        '1234567890',     # 012...는 유효 프리픽스 아님
        '8210991501411',  # 국제표기(82) — 복원 대상 아님
        '2026',           # 연도
        '',               # 빈값
    ])
    def test_does_not_falsely_repair(self, garbage):
        """복원해서 유효 번호가 안 되면 원본 그대로 (잘못된 조용한 정정 방지)."""
        assert restore_leading_zero(garbage) == garbage


class TestNormalizePhoneWithRepair:
    """normalize_phone 이 복원 후 하이픈 포맷까지."""

    @pytest.mark.parametrize('inp, out', [
        ('1091501411', '010-9150-1411'),   # 사고 케이스 end-to-end
        ('212345678', '02-1234-5678'),
        ('317771234', '031-777-1234'),
        ('01091501411', '010-9150-1411'),  # 정상
        ('010-9150-1411', '010-9150-1411'),
        ('025581105', '02-558-1105'),
    ])
    def test_normalize(self, inp, out):
        assert normalize_phone(inp) == out

    @pytest.mark.parametrize('nonphone', ['만료됨', '-', 'nan', ''])
    def test_nonphone_returns_empty(self, nonphone):
        assert normalize_phone(nonphone) == ''


class TestIsValidPhone:
    """정합성 체크용 — '값은 있는데 형태 이상' 감지. 복원 가능하면 유효로 본다."""

    @pytest.mark.parametrize('valid', [
        '010-9150-1411', '01091501411', '1091501411',  # 복원 가능 → 유효
        '02-558-1105', '212345678', '031-777-1234',
    ])
    def test_valid(self, valid):
        assert is_valid_phone(valid) is True

    @pytest.mark.parametrize('invalid', [
        '', '-', 'abc', '12345', '1234567890', '전화없음',
    ])
    def test_invalid(self, invalid):
        assert is_valid_phone(invalid) is False

    def test_digits_predicate(self):
        assert is_valid_phone_digits('01091501411') is True
        assert is_valid_phone_digits('1091501411') is False  # 0 없는 raw digits


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
