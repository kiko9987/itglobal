# -*- coding: utf-8 -*-
"""엑셀·CSV 로 내보낼 때 직원·고객이 쓴 글이 수식으로 실행되지 않게 (2026-11-13 보안 점검).

- xlsx(openpyxl): '=' 로 시작하는 글을 openpyxl 이 수식 칸으로 저장한다 → 행을 넣은 뒤 글 칸으로 되돌린다.
  (값은 그대로 보이고, 실행만 안 됨)
- csv: 엑셀이 = + - @ 탭·줄바꿈으로 시작하는 칸을 수식으로 읽는다 → 앞에 ' 를 붙인다.
  '-' 한 글자(빈 값 표시)는 그대로.
"""

_CSV_FORMULA_LEADS = ('=', '+', '-', '@', '\t', '\r', '\n')


def csv_safe(value):
    if isinstance(value, str) and value and value[0] in _CSV_FORMULA_LEADS and value != '-':
        return "'" + value
    return value


def csv_safe_row(values):
    return [csv_safe(v) for v in values]


def xlsx_append_text(ws, values):
    """ws.append 와 같지만, 넣은 행에서 수식으로 잡힌 글 칸을 글 칸으로 되돌린다."""
    ws.append(values)
    for cell in ws[ws.max_row]:
        if cell.data_type == 'f':
            cell.data_type = 's'


# ── 구글 시트 USER_ENTERED 쓰기 공통 방지 (2026-11-13 보안 점검 2차) ──────────────────
#
# 공사현황·A/S 시트는 USER_ENTERED 로 쓴다. 리드에서 막아 둔 글('=…')을 다시 읽으면 ' 가 빠진 채 오고,
# 그 값을 공사 등록·거래처명 채우기 등에서 그대로 쓰면 수식이 되살아났다(2차 점검 #12·보류 A).
# 칸마다 고치지 않고 GoogleSheetsManager 의 쓰기 함수에서 한 번에 막는다.
#   - 코드가 일부러 넣는 계산식(총액 2·미수금·순익·마진율·부담자 총액 2)은 계산 함수·칸 주소·숫자만 쓰므로
#     그대로 둔다(_is_plain_calc_formula). 글자열("…")·다른 탭(!)·HYPERLINK/IMAGE/IMPORT 같은 함수는 막는다.
#   - 숫자 칸의 음수('-5,000')와 빈 값 표시('-')는 그대로.
import logging
import re

_logger = logging.getLogger(__name__)

_CALC_FUNCS = frozenset({
    'IF', 'IFERROR', 'ABS', 'OR', 'AND', 'NOT', 'MIN', 'MAX', 'FLOOR', 'CEILING', 'MOD', 'SUM', 'N',
    'ROUND', 'ROUNDDOWN', 'ROUNDUP', 'INT',
})
_CELL_REF = re.compile(r'\$?[A-Za-z]{1,3}\$?\d+')
_FUNC_CALL = re.compile(r'([A-Za-z_][A-Za-z0-9_.]*)\s*\(')
_NUMERIC_NEG = re.compile(r'^-\s*[\d,]*\.?\d+\s*$')
_SHEET_LEADS = ('=', '+', '@', '\t', '\r', '\n')


def _is_plain_calc_formula(text: str) -> bool:
    """계산 함수·칸 주소·숫자·연산자만으로 된 수식인지 (코드가 만드는 계산식 모양)."""
    body = text[1:]
    if '"' in body.replace('""', ''):        # 글자열 인자(URL 등) — 빈 글자 "" 만 허용
        return False
    if any(ch in body for ch in "!'{}[];\\@#`~^|\n\r\t"):  # 다른 탭 참조·배열·로캘 구분자 등
        return False
    for name in _FUNC_CALL.findall(body):
        if name.upper() not in _CALC_FUNCS:
            return False
    rest = _FUNC_CALL.sub('(', body)
    rest = _CELL_REF.sub('', rest)
    rest = re.sub(r'\b(?:TRUE|FALSE)\b', '', rest, flags=re.I)
    return not re.search(r'[A-Za-z가-힣]', rest)


def sheet_cell_safe(value):
    """USER_ENTERED 로 쓸 한 칸. 수식으로 실행될 글이면 앞에 ' 를 붙여 글자로 저장(시트에는 안 보임)."""
    if not isinstance(value, str) or not value:
        return value
    lead = value[0]
    if lead == '=':
        if _is_plain_calc_formula(value):
            return value
        _logger.warning(f"[SHEET_SAFE] 수식 모양 글을 글자로 저장: {value[:40]!r}")
        return "'" + value
    if lead in _SHEET_LEADS:
        return "'" + value
    if lead == '-' and value != '-' and not _NUMERIC_NEG.match(value):
        return "'" + value
    return value


def sheet_row_safe(values):
    return [sheet_cell_safe(v) for v in values] if isinstance(values, (list, tuple)) else values


def sheet_grid_safe(grid):
    """[[…], […]] 2차원 값."""
    return [sheet_row_safe(r) for r in grid] if isinstance(grid, (list, tuple)) else grid


def sheet_data_safe(data):
    """batchUpdate 의 data 목록 [{'range':…, 'values': [[…]]}]."""
    out = []
    for d in data or []:
        if isinstance(d, dict) and 'values' in d:
            d = dict(d, values=sheet_grid_safe(d['values']))
        out.append(d)
    return out
