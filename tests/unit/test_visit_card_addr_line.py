# -*- coding: utf-8 -*-
"""방문 카드 원본/변환 주소 라인 — 정규화 후 동일하면 단일 라인 (L-03696).

방문 예약 전환 시 이미 정규화된 주소가 재정규화돼도 실제 변경이 없는데, 원본↔변환을
raw 문자열로 비교해 공백 등 미세차만으로 '원본/변환' 두 줄이 떠 '자동 변환' 오해를
유발하던 이슈. 공백 정규화 후 동일하면 '방문 주소' 한 줄만.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.blueprints.slack_bot as sb


A = '안산 단원구 선부광장남로 109 주공아파트 상가 지층 한국유통 선부점'


def _addr_lines(addr_note, visit_address):
    body, _ = sb._build_visit_notice_blocks(
        lead_no='L-TEST', category_display='온라인(당근)', initial='SD',
        visit_date='2026-08-18', name='변세일', contact='010-0000-0000',
        visit_address=visit_address, consultation='x', addr_note=addr_note,
    )
    return [l for l in body.split('\n') if '주소' in l]


def test_identical_orig_conv_single_line():
    lines = _addr_lines({'kind': 'normalized', 'original': A}, A)
    assert len(lines) == 1
    assert lines[0].startswith('>방문 주소 :')
    assert '변환 주소' not in lines[0]


def test_whitespace_only_diff_collapses_to_single_line():
    # raw 는 이중공백으로 다르지만 정규화 후 동일 → 한 줄
    orig = '안산  단원구 선부광장남로 109 주공아파트 상가 지층 한국유통 선부점'
    lines = _addr_lines({'kind': 'normalized', 'original': orig}, A)
    assert len(lines) == 1 and lines[0].startswith('>방문 주소 :')


def test_real_conversion_shows_two_lines():
    lines = _addr_lines({'kind': 'normalized', 'original': '선부광장남로109 지층 한국유통'}, A)
    assert len(lines) == 2
    assert lines[0].startswith('>*원본 주소*') and lines[1].startswith('>*변환 주소*')


def test_failed_shows_confirm_badge():
    lines = _addr_lines({'kind': 'failed', 'original': ''}, A)
    assert len(lines) == 1 and '주소 확인 필요' in lines[0]


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
