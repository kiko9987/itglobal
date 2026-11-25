# -*- coding: utf-8 -*-
"""자금일보 자동 초안(11-26) — 상세 내역 줄 만들기(유입 구분·매출수금 완료/중·전자결재 짝짓기·지난 기록 제안)와 시트 추가."""
import sys
from collections import Counter
from datetime import date

sys.path.insert(0, '.')

import openpyxl
import pytest
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

from dashboard.services import fund_daily as fd

HINTS = {'in': {fd.norm('이자'): Counter({('이자수익', ''): 5})},
         'out': {fd.norm('유현세무법인'): Counter({('지급수수료', '기장료'): 9}),
                 fd.norm('최만호'): Counter({('지급수수료', 'R4294-JW 앵글'): 1, ('지급수수료', 'R4100-JW 철거'): 1})}}


def test_deposit_rows_from_project():
    txs = [{'code': '9012', 'direction': 'in', 'amount': 3_000_000, 'counterparty': '(주)영진바이오크',
            'project_code': 'R4259-MJ', 'at': '2026-11-25 10:00'},
           {'code': '9012', 'direction': 'in', 'amount': 220_000, 'counterparty': '주식회사조이토이',
            'project_code': 'R4328-TH', 'at': '2026-11-25 11:00'},
           {'code': '1011', 'direction': 'in', 'amount': 1_234, 'counterparty': '이자', 'project_code': '', 'at': '2026-11-25 12:00'},
           {'code': '1011', 'direction': 'in', 'amount': 50_000, 'counterparty': '모르는사람', 'project_code': '',
            'at': '2026-11-25 13:00'}]
    projects = {'R4259-MJ': {'유입 구분': '홈페이지', '미수금': 0.0}, 'R4328-TH': {'유입 구분': '홈페이지', '미수금': 2_200_000}}
    ins, outs = fd.build_rows(txs, HINTS, projects)
    assert [(r['C'], r['D'], r['G'], r['H']) for r in ins[:2]] == [
        ('홈페이지', 'R4259-MJ 매출수금완료', '영진바이오크', 3_000_000), ('홈페이지', 'R4328-TH 매출수금 중', '주식회사조이토이', 220_000)]
    assert (ins[2]['C'], ins[2]['check']) == ('이자수익', False)
    assert (ins[3]['C'], ins[3]['check']) == ('', True)                       # 못 찾으면 비우고 확인 필요
    assert outs == []


def test_outflows_match_approval_then_hints():
    txs = [{'code': '1011', 'direction': 'out', 'amount': 2_492_665, 'counterparty': '오세영', 'project_code': '', 'at': '2026-11-25 15:00'},
           {'code': '9012', 'direction': 'out', 'amount': 165_000, 'counterparty': '유현세무법인', 'project_code': '', 'at': '2026-11-25 16:00'},
           {'code': '농협', 'direction': 'out', 'amount': 300_000, 'counterparty': '최만호', 'project_code': '', 'at': '2026-11-25 17:00'},
           {'code': '농협', 'direction': 'out', 'amount': 1_000_000, 'counterparty': '농협 출금', 'project_code': '', 'at': '2026-11-25 18:00'}]
    approvals = [{'id': 82, 'amount': 2_492_665, 'project_code': '프로젝트0', 'cost_category': '자재비', 'content': '전기 자재구매',
                  'partner': '세광전기', 'payee_name': '오세영', 'requester_name': '조성헌', 'done_at': '2026-11-25 14:59:49'},
                 {'id': 99, 'amount': 2_492_665, 'project_code': 'R1-X', 'cost_category': '자재비', 'content': '사흘 뒤',
                  'partner': '딴곳', 'requester_name': '', 'done_at': '2026-11-28 10:00:00'}]
    _, outs = fd.build_rows(txs, HINTS, {}, approvals, lambda n: 'JSH' if n == '조성헌' else '')
    assert {k: outs[0][k] for k in ('I', 'J', 'K', 'N')} == {'I': '1011', 'J': '외상매입', 'K': '0 JSH 전기 자재구매', 'N': '세광전기'}
    assert (outs[1]['J'], outs[1]['K']) == ('지급수수료', '기장료')
    assert (outs[2]['J'], outs[2]['K']) == ('지급수수료', '')                 # 공사마다 다른 적요는 계정만 제안
    assert len(outs) == 3                       # 농협에서 받는 사람 없이 나간 돈은 자금일보에서 뺌(11-26 부장님)


def _template(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '24'
    for r in fd.ACCT_ROWS:
        ws[f'G{r}'] = f"=SUM('23'!L{r}:M{r})"
        ws[f'L{r}'] = f'=G{r}+H{r}-J{r}'
    for r in range(fd.FIRST_ROW, fd.LAST_ROW + 1):
        ws.merge_cells(f'D{r}:F{r}')
        ws.merge_cells(f'K{r}:M{r}')
        ws.row_dimensions[r].hidden = r > 22
    ws['B19'], ws['H19'] = 1011, 5
    ws['H128'] = '=SUM($H$19:$H$127)'
    ws.conditional_formatting.add('B19:H127', CellIsRule(operator='equal', formula=['1011'], fill=PatternFill('solid', fgColor='DDFFDD')))
    wb.save(path)


def test_add_day_sheet(tmp_path):
    p = str(tmp_path / '26.11 자금일보.xlsx')
    _template(p)
    ins = [{'B': '9012', 'C': '홈페이지', 'D': 'R4259-MJ 매출수금완료', 'G': '영진', 'H': 3_000_000, 'check': False}]
    outs = [{'I': '1011', 'J': '', 'K': '', 'N': '모름', 'O': 1000, 'check': True}] * 3
    assert fd.add_day_sheet(p, p, date(2026, 11, 25), ins, outs) == 'made'
    wb = openpyxl.load_workbook(p)
    ws = wb['25']
    assert wb.sheetnames == ['24', '25']
    assert ws['G6'].value == "=SUM('24'!L6:M6)" and ws['H128'].value == '=SUM($H$19:$H$127)'
    assert ws['B19'].value == 9012 and ws['D19'].value == 'R4259-MJ 매출수금완료' and ws['H19'].value == 3_000_000
    assert ws['O21'].value == 1000 and ws['J21'].comment is not None and ws['B22'].value is None
    assert not ws.row_dimensions[19].hidden and not ws.row_dimensions[23].hidden and ws.row_dimensions[24].hidden
    assert len(ws.conditional_formatting) == 1 and 'D19:F19' in {str(r) for r in ws.merged_cells.ranges}
    assert fd.add_day_sheet(p, p, date(2026, 11, 25), ins, outs) == 'exists'        # SB 가 만든 시트는 안 건드림
    assert fd.add_day_sheet(p, p, date(2026, 11, 25), ins, outs, suffix=' 자동') == 'made'
    assert '25 자동' in openpyxl.load_workbook(p).sheetnames


def test_prev_business_day(monkeypatch):
    from dashboard.services import company_holidays as ch
    monkeypatch.setattr(ch, 'is_holiday', lambda d: d == date(2026, 12, 25))
    assert fd.prev_business_day(date(2026, 11, 26)) == date(2026, 11, 25)
    assert fd.prev_business_day(date(2026, 11, 30)) == date(2026, 11, 27)       # 월요일 → 금요일
    assert fd.prev_business_day(date(2026, 12, 28)) == date(2026, 12, 24)       # 크리스마스 건너뜀


def test_off_by_default(monkeypatch):
    monkeypatch.delenv('FUND_DAILY_MODE', raising=False)
    assert fd.run() == {'status': 'off'}
