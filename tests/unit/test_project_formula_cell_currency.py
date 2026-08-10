# -*- coding: utf-8 -*-
"""편집 저장 응답에서 수식 금액 셀 오파싱 방지 회귀 테스트 (2026-08-10 R3826-MJ).

배경: 제품대/자재비/기타비는 VLOOKUP 수식 셀. PUT 편집 저장 시 current_values 를
FORMULA render 로 읽어(날짜/수식 보존용) 이 셀들이 수식 문자열로 들어오는데,
_build_updated_project_from_values 의 safe_parse_currency 가 수식 안 행 참조
($A3827)를 금액(3827)으로 오추출 → 저장 직후 카드에 행번호가 금액처럼 뜨고
순익도 왜곡(총액 - (3827*3 + 도급비)). 새로고침하면 시트 실제값(0)으로 복귀했으나
저장 직후 표시가 잘못됨.

가드: 통화 필드 값이 수식('='로 시작)이면 빈값 처리 → 카드 정상('-') + 순익 정확.
사용자 입력 금액(도급비=969000)은 그대로 보존.
"""
import sys
sys.path.insert(0, '.')

from dashboard.blueprints.projects import _build_updated_project_from_values

# VLOOKUP 수식 (행 3827 참조) — 실제 시트 제품대/자재비/기타비 셀 형태
_FORMULA = ('=IF($C3827<>"", IFERROR(VLOOKUP($A3827, '
            'INDIRECT("\'"&$C3827&"\'!A:AB"), 28, FALSE), ""), "")')

_FIELD_INDEX = {
    '총액 1': 0, '제품대': 1, '도급비': 2, '자재비': 3, '기타비': 4,
    '순익': 5, '마진율': 6, '부가세': 7,
}


def _build(total, product, contract, material, other, vat='FALSE'):
    vals = [total, product, contract, material, other,
            '=R3827-(AB3827+AC3827+AD3827+AE3827)', '=x', vat]
    return _build_updated_project_from_values(vals, _FIELD_INDEX)


def test_formula_cost_cells_become_empty_not_row_number():
    r = _build('4900000', _FORMULA, '969000', _FORMULA, _FORMULA)
    assert r['제품대'] == ''      # 3827 아님
    assert r['자재비'] == ''
    assert r['기타비'] == ''


def test_manual_amount_preserved():
    r = _build('4900000', _FORMULA, '969000', _FORMULA, _FORMULA)
    assert r['도급비'] == '969000'


def test_net_profit_correct_with_formula_cost_cells():
    """순익 = 총액 - (제품대+도급비+자재비+기타비). 수식 셀=0 취급 → 3,931,000."""
    r = _build('4900000', _FORMULA, '969000', _FORMULA, _FORMULA)
    assert r['순익'] == '3931000'   # 3,919,519 (오파싱) 아님


def test_real_amounts_still_parsed():
    """수식 아닌 실제 금액은 정상 파싱 (₩·콤마 제거)."""
    r = _build('4900000', '3827', '969000', '0', '₩1,000')
    assert r['제품대'] == '3827'    # 실제 입력값이면 그대로
    assert r['기타비'] == '1000'
    # 순익 = 4,900,000 - (3827+969000+0+1000) = 3,926,173
    assert r['순익'] == '3926173'
