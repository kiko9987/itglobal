# -*- coding: utf-8 -*-
"""공사 정보 수정 diff 거짓 변경 방지 회귀 테스트 (2026-08-07 R3906-TH).

배경: 시트에 '-' 로 저장된 도급 구분·시공자가, 모달 빈칸('') 제출과 '' != '-'
거짓 diff 나서 → 안 건드린 필드가 '변경'으로 잡히고 → perform_edit가 '-'를 ''로
덮어쓰며 '공사 내용 수정 알림' 카드가 뜸. (변경 내역 "- → -" 가 그 흔적.)

_norm_edit_val 로 '-'·''·공백을 같은 빈값 취급 → 거짓 diff 제거. 실제 값
변경/삭제는 그대로 감지.
"""
import sys
sys.path.insert(0, '.')

from dashboard.blueprints.slack_bot import _norm_edit_val


def _changed(old, new):
    """수정 제출 diff 판정 = 정규화 후 다름."""
    return _norm_edit_val(new) != _norm_edit_val(old)


# ── 거짓 diff 방지 (R3906-TH 재현) ──
def test_dash_vs_empty_not_a_change():
    assert _changed('-', '') is False   # 시트 '-' vs 모달 빈칸
    assert _changed('', '-') is False
    assert _changed('-', '   ') is False
    assert _changed(None, '') is False


def test_norm_variants():
    assert _norm_edit_val('-') == ''
    assert _norm_edit_val('  -  ') == ''
    assert _norm_edit_val('') == ''
    assert _norm_edit_val(None) == ''
    assert _norm_edit_val(' 직도급 ') == '직도급'


# ── 실제 변경은 여전히 감지 ──
def test_real_change_detected():
    assert _changed('-', '직도급') is True        # 빈값 → 실제 값
    assert _changed('직도급', '') is True          # 실제 값 → 삭제(빈값)
    assert _changed('직도급', '-') is True         # 실제 값 → '-'(=삭제)
    assert _changed('18평 2대', '25평 1대') is True  # 값 교체
    assert _changed('갑', '을') is True
