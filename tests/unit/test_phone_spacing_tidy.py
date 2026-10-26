# -*- coding: utf-8 -*-
"""공백 섞인 전화번호 회귀 테스트 (2026-10-26 ETC-587764 청담 유니크클리닉).

사고: 워크플로 입력 '010- 4327-1111'(하이픈 뒤 공백)이 시트·캔버스1에 그대로 찍히고 JW가
캔버스2로 옮겨 적음 → 캔버스2 전화 인식 실패 → 매칭 실패 → 담당(MW) 일정확정 DM 누락.
고침: ①시트에 쓸 때 번호만인 유효 번호는 표준 표기로 정리 ②캔버스2 인식은 공백 1칸 허용.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.lead_helpers import tidy_phone_for_sheet
import dashboard.services.visit_assignment_sync as vas


@pytest.mark.parametrize('raw,expected', [
    ('010- 4327-1111', '010-4327-1111'),     # 사고 케이스
    ('010 4327 1111', '010-4327-1111'),
    ('010 -4327- 1111', '010-4327-1111'),
    ('01043271111', '010-4327-1111'),
    ('010-4327-1111', '010-4327-1111'),      # 이미 정상
    ('02 333 8150', '02-333-8150'),
    ('-', '-'),                              # 번호 아님 → 그대로
    ('/', '/'),
    ('', ''),
    ('공항로424번길 72 112호', '공항로424번길 72 112호'),   # 주소 오기재(L-04220) → 건드리지 않음
    ('010-1234-5678 / 02-333-8150', '010-1234-5678 / 02-333-8150'),  # 번호 2개
    ('010-123-45', '010-123-45'),            # 형태 이상 → 그대로(감지 대상)
    (None, None),
])
def test_tidy_phone_for_sheet(raw, expected):
    assert tidy_phone_for_sheet(raw) == expected


@pytest.mark.parametrize('line,digits', [
    ('(MW) 10월 27일 / 010- 4327-1111 / 강남구 도산대로67길 7 파크빌딩 4층 / 실외기', '01043271111'),
    ('(MW) 10월 27일 / 010 4327 1111 / 강남구 도산대로67길 7', '01043271111'),
    ('(MW) 10월 27일 / 010-4327-1111 / 강남구 도산대로67길 7', '01043271111'),
    ('(당) 10월 13~14일 / 010-9226-8475 / 인천 미추홀구 독정이로 82', '01092268475'),
])
def test_canvas2_phone_regex_tolerates_spaces(line, digits):
    m = vas._PHONE_RE.search(line)
    assert m is not None
    assert vas._normalize_phone(m.group(0)) == digits


@pytest.mark.parametrize('line', [
    '(SH) 8월 14일 / - / 인천 미추홀구 인하로 79 2층 한의원(예정) / 현재공실',
    '(YM) 10월 26~27일 / - / 용인 수지구 현암로125번길 7 / 20마력 실외기 5층 발코니',
])
def test_canvas2_phone_regex_no_false_match(line):
    """전화 없는 줄에서 날짜·번지·층수를 전화로 오인하지 않음."""
    assert vas._PHONE_RE.search(line) is None
