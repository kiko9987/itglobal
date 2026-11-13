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
