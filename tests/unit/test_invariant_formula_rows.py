# -*- coding: utf-8 -*-
"""정합성 점검 — 수식 칸이 다른 행을 계산 (2026-10-28 G4153 4154행 수식이 4153행을 가리킨 사고)."""
from dashboard.services import invariant_checks as ic

F2L = {'총액 2': 'U', '미수금': 'Y', '순익': 'AJ', '마진율': 'AK'}
GOOD = {
    'U': '=IF(T4155=TRUE, S4155+FLOOR(S4155*0.1,1), S4155)',
    'Y': '=IF(ABS($U4155-$V4155-$W4155-$X4155)<2, 0, $U4155-$V4155-$W4155-$X4155)',
    'AJ': '=S4155-(AF4155+AG4155+AH4155+AI4155)',
    'AK': '=IF(OR(S4155=0, AJ4155=0), 0, AJ4155/S4155)',
}


def test_detects_formulas_pointing_at_row_above():
    bad = {k: v.replace('4155', '4153') for k, v in GOOD.items()}      # 4154행인데 4153행을 가리킴
    rows = [(4155, 'G4154-JW', GOOD), (4154, 'G4153-JW', bad)]
    found = ic.formula_row_mismatches(rows, f2l=F2L)
    assert [c for _, c, _ in found] == ['G4153-JW'] * 4
    assert '미수금(Y4154) 수식이 4153행을 계산' in found[1][2]
    assert found[0][0] == 'formula_row:G4153-JW:총액 2:4154'


def test_ignores_function_names_strings_other_tabs_and_whole_columns():
    cells = {
        'A': '=LOG10(S10)+ATAN2(S10, 1)',             # 함수 이름 속 숫자
        'B': '=IF(S10>0, "A1 메모 B2", "")',          # 따옴표 안 글자
        'C': "='거래처'!A5",                           # 다른 탭
        'D': '=SUM(A:A)-S10',                         # 전체 열
        'E': '=$S$10*1.1',                            # 절대 참조
        'F': 1234,                                    # 수식 아님
    }
    assert ic.formula_row_mismatches([(10, 'G1', cells)], f2l=F2L) == []


def test_check_wraps_violations():
    ctx = {'formula_mismatches': [('formula_row:G1:미수금:5', 'G1', '미수금(Y5) 수식이 4행을 계산')]}
    v = ic.check_formula_rows(ctx)
    assert len(v) == 1 and v[0]['check'] == 'formula_row' and v[0]['title'] == 'G1'
    assert ic.check_formula_rows({}) == []
    assert ic.check_formula_rows in ic.INVARIANTS and 'formula_row' in ic._CHECK_LABEL
