# -*- coding: utf-8 -*-
"""대체수금(상계) 블록 단위 파싱 (2026-10-15 R4101-JSH).

잔금 칸에 은행 입금(10,400,000원)이 있는데 현장 장비 매입분 400,000원을 상계로 덧붙이면,
예전엔 '대체 수금'이 메모 어디에든 있으면 메모 통째로 대체수금 한 건이 되어 은행 입금이
사라졌고, '상계'로 쓰면 아예 못 읽었다. 경영지원이 어떤 표현을 쓸지 몰라 여러 형태를 받는다.
배포 전 전 프로젝트(3,728) 메모를 구/신 파서로 비교 — 결과 차이 0건.
"""
import sys

sys.path.insert(0, '.')

import pytest

from dashboard.services.payment_sync import _parse_notes

DEP = '2026/10/12\n농협, 10/12 10:45\n352-****-1682-33\n입금 10,400,000원\n김영은'


def _brief(ps):
    return [(p['stage'], p['date_md'], p['amount'], p['partner'], bool(p.get('is_substitute')))
            for p in ps]


@pytest.mark.parametrize('extra, date', [
    ('\n\n2026/10/15\n현장 장비 매입 상계 400,000원', '10/15'),
    ('\n\n2026/10/15\n장비 매입금 400,000원 대체 수금', '10/15'),
    ('\n\n10월 15일 장비대 차감 400,000원', '10/15'),
    ('\n\n10/15 장비대 40만 원 상계', '10/15'),
    ('\n\n장비 매입 상계 40만원', '-'),
    # 빈 줄 없이 입금 문자 바로 아래 붙여 쓴 경우
    ('\n2026/10/15\n장비 매입 상계 400,000원', '10/15'),
    ('\n400,000원 대체 수금(장비 매입)', '-'),
])
def test_deposit_plus_offset_same_cell_keeps_both(extra, date):
    ps = _parse_notes(['', '', DEP + extra], stage_vals={'잔금': 10_800_000})
    assert _brief(ps) == [('잔금', '10/12', 10_400_000, '김영은', False),
                          ('잔금', date, 400_000, '대체수금', True)]   # 상계가 단계 마지막(카드 '새 입금')


def test_offset_alone_in_empty_stage():
    ps = _parse_notes(['', '2026/10/15\n현장 장비 매입 상계 400,000원', ''], stage_vals={'중도금': 400_000})
    assert _brief(ps) == [('중도금', '10/15', 400_000, '대체수금', True)]


@pytest.mark.parametrize('memo', [
    '2024-06-07\n입금 165,000원\n최지훈(상계아산내과',      # 상호 속 '상계'
    '입금일: 2024-03-29\n입금자: 백승필T 도급비에서 차감',   # 라벨 양식(입금 단어) → 기존대로
])
def test_offset_word_in_bank_or_label_block_not_substitute(memo):
    ps = _parse_notes(['', '', memo], stage_vals={'잔금': 165_000})
    assert not any(p.get('is_substitute') for p in ps)


def test_existing_single_block_substitute_unchanged():
    # 기존 규칙(메모 전체 기준) 그대로 — G3776-SH·R2244-TH 형
    ps = _parse_notes(['', '', 'G3794-SH 리베이트 250,000원 대체수금'], stage_vals={'잔금': 250_000})
    assert _brief(ps) == [('잔금', '-', 250_000, '대체수금', True)]


def test_offset_without_amount_not_substitute():
    assert _parse_notes(['', '장비 상계처리', ''], stage_vals={'중도금': 400_000}) == []
