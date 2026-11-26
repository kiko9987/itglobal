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
    nh = [{'code': '농협', 'direction': 'in', 'amount': 1, 'counterparty': '누구', 'project_code': '', 'at': '2026-11-25 09:00'},
          {'code': '농협', 'direction': 'in', 'amount': 2, 'counterparty': '고객', 'project_code': 'R4259-MJ', 'at': '2026-11-25 09:01'}]
    ins, _ = fd.build_rows(nh, HINTS, projects)
    assert [r['H'] for r in ins] == [2]                  # 농협 입금은 공사 입금으로 지정된 것만


def test_outflows_match_approval_then_hints():
    txs = [{'code': '1011', 'direction': 'out', 'amount': 2_492_665, 'counterparty': '오세영', 'project_code': '', 'at': '2026-11-25 15:00'},
           {'code': '9012', 'direction': 'out', 'amount': 165_000, 'counterparty': '유현세무법인', 'project_code': '', 'at': '2026-11-25 16:00'},
           {'code': '농협', 'direction': 'out', 'amount': 300_000, 'counterparty': '최만호', 'project_code': '', 'at': '2026-11-25 17:00'},
           {'code': '농협', 'direction': 'out', 'amount': 1_000_000, 'counterparty': '농협 출금', 'project_code': '', 'at': '2026-11-25 18:00'}]
    approvals = [{'id': 82, 'amount': 2_492_665, 'project_code': '프로젝트0', 'cost_category': '자재비', 'content': '전기 자재구매',
                  'partner': '세광전기', 'payee_name': '오세영', 'requester_name': '조성헌', 'grn': 'G', 'done_at': '2026-11-25 14:59:49'},
                 {'id': 84, 'amount': 300_000, 'project_code': 'R4294-JW', 'cost_category': '도급비', 'content': '앵글',
                  'partner': '최만호', 'requester_name': '조성헌', 'grn': 'N', 'done_at': '2026-11-25 15:16:00'},
                 {'id': 99, 'amount': 2_492_665, 'project_code': 'R1-X', 'cost_category': '자재비', 'content': '사흘 뒤',
                  'partner': '딴곳', 'requester_name': '', 'done_at': '2026-11-28 10:00:00'}]
    _, outs = fd.build_rows(txs, HINTS, {}, approvals, lambda n: 'JSH' if n == '조성헌' else '')
    assert {k: outs[0][k] for k in ('I', 'J', 'K', 'N')} == {'I': '1011', 'J': '외상매입', 'K': '0 JSH 전기 자재구매', 'N': '세광전기'}
    assert (outs[1]['J'], outs[1]['K']) == ('지급수수료', '기장료')
    assert (outs[2]['N'], outs[2]['K']) == ('최만호', 'R4294-JW 앵글')       # 농협은 전자결재 N 출금과 짝일 때만
    assert len(outs) == 3                       # 짝 없는 농협 출금은 자금일보에서 뺌(11-26 부장님)


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


def test_report_balances_reads_latest_sheet_before_today(tmp_path):
    p = str(tmp_path / '26.11 자금일보.xlsx')
    wb = openpyxl.Workbook()
    wb.active.title = '24'
    for name, base in (('24', 100), ('25', 200)):
        ws = wb[name] if name in wb.sheetnames else wb.create_sheet(name)
        for i, r in enumerate(fd.ACCT_ROWS):
            ws[f'D{r}'], ws[f'L{r}'] = [1011, 1012, 1013, 1014, 9011, 9012, 4011, 5011][i], base + i
    wb.create_sheet('26')                                                    # 오늘 시트(아직 값 없음)
    wb.save(p)
    d, vals = fd.report_balances(p, date(2026, 11, 26))
    assert d == date(2026, 11, 25) and vals['1012'] == 201 and vals['5011'] == 207
    d, vals = fd.report_balances(p, date(2026, 11, 25))
    assert d == date(2026, 11, 24) and vals['1011'] == 100
    assert fd.report_balances(p, date(2026, 11, 24)) == (None, {})


# ── 새 달 파일(11-26 SB '엑셀 원본이 만들어져야' → 서버가 첫 영업일에 만들어 둠) ──
CODES = [1011, 1012, 1013, 1014, 9011, 9012, 4011, 5011]


def _prev_month(path, last_missing=False):
    wb = openpyxl.Workbook()
    wb.active.title = '27'
    for name, base in (('27', 100), ('30', 1000)):
        ws = wb[name] if name in wb.sheetnames else wb.create_sheet(name)
        for i, r in enumerate(fd.ACCT_ROWS):
            ws[f'D{r}'] = CODES[i]
            ws[f'G{r}'] = "=SUM('[1]27'!L6:M6)" if name == '27' else f"=SUM('27'!L{r}:M{r})"
            ws[f'L{r}'] = None if (last_missing and name == '30' and i == 7) else base + i   # 엑셀 저장값 대신 숫자
        for r in range(fd.FIRST_ROW, fd.LAST_ROW + 1):
            ws.merge_cells(f'D{r}:F{r}')
        ws['B19'], ws['H19'] = 1011, 5
        ws['H128'] = '=SUM($H$19:$H$127)'
    wb.save(path)


def test_make_month_file_first_sheet_from_last_balances(tmp_path):
    prev, out = str(tmp_path / '26.11 자금일보.xlsx'), str(tmp_path / '26.12 자금일보.xlsx')
    _prev_month(prev)
    res = fd.make_month_file(prev, out, date(2026, 12, 1))
    assert res == {'status': 'made', 'missing': [], 'from_sheet': '30'}
    wb = openpyxl.load_workbook(out)
    ws = wb['1']
    assert wb.sheetnames == ['1']                                            # 지난달 시트는 하나도 안 남김
    assert [ws[f'G{r}'].value for r in fd.ACCT_ROWS] == [1000 + i for i in range(8)]   # 11/30 잔액을 숫자로
    assert ws['G6'].comment is not None and ws['B19'].value is None and ws['H128'].value == '=SUM($H$19:$H$127)'
    # 다음 날 07:00 — 상세가 빈 첫 시트는 그 자리에 채움(전일잔액 그대로), 채운 뒤엔 다시 안 건드림
    ins = [{'B': '9012', 'C': '홈페이지', 'D': 'R1-A 매출수금완료', 'G': '가', 'H': 7, 'check': False}]
    assert fd.add_day_sheet(out, out, date(2026, 12, 1), ins, []) == 'made'
    ws = openpyxl.load_workbook(out)['1']
    assert ws['B19'].value == 9012 and ws['G6'].value == 1000
    assert fd.add_day_sheet(out, out, date(2026, 12, 1), ins, []) == 'exists'


def test_make_month_file_marks_missing_balance(tmp_path):
    prev, out = str(tmp_path / 'p.xlsx'), str(tmp_path / 'o.xlsx')
    _prev_month(prev, last_missing=True)
    res = fd.make_month_file(prev, out, date(2026, 12, 1))
    assert res['missing'] == ['5011']
    ws = openpyxl.load_workbook(out)['1']
    assert ws['G13'].value is None and ws['G13'].comment is not None


def test_copy_mode_first_day_uses_same_day_sheet(tmp_path):
    prev, out = str(tmp_path / 'p.xlsx'), str(tmp_path / 'o.xlsx')
    _prev_month(prev)
    fd.make_month_file(prev, out, date(2026, 12, 1))
    assert fd.add_day_sheet(out, out, date(2026, 12, 1), [], [], suffix=' 자동') == 'made'
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ['1', '1 자동'] and wb['1 자동']['G7'].value == 1001


def test_ensure_month_file(tmp_path, monkeypatch):
    from dashboard.services import company_holidays as ch
    monkeypatch.setattr(ch, 'is_holiday', lambda d: d == date(2027, 1, 1))
    monkeypatch.setenv('FUND_DAILY_COPY_DIR', str(tmp_path / 'copy'))
    hint = tmp_path / 'hints'
    hint.mkdir()
    _prev_month(str(hint / '26.11 자금일보.xlsx'))
    files = {'26.11 자금일보.xlsx': {'id': 'x', 'modifiedTime': 't'}}
    assert fd.first_business_day(2027, 1) == date(2027, 1, 4)
    assert fd.ensure_month_file(date(2026, 11, 30), 'copy', None, 'f', files, str(hint)) == 'exists'     # 이달 파일 있음
    assert fd.ensure_month_file(date(2026, 12, 1), 'copy', None, 'f', files, str(hint)) == 'made'
    assert (tmp_path / 'copy' / '26.12 자금일보 (새 달 자동).xlsx').exists()
    assert fd.ensure_month_file(date(2026, 12, 2), 'copy', None, 'f', files, str(hint)) == 'exists'
    assert fd.ensure_month_file(date(2027, 1, 1), 'copy', None, 'f', files, str(hint)) == 'not_yet'   # 1/1 휴일 → 첫 영업일 1/4
    assert fd.ensure_month_file(date(2027, 1, 4), 'copy', None, 'f', files, str(hint)) == 'no_prev'    # 12월 파일이 드라이브에 없음


def test_match_combined_approvals_same_payee():
    """11-26 이종원 개인경비 #87 24,200 + #88 3,850 → 출금 28,050 한 번."""
    outs = [{'code': '9012', 'amount': 28_050, 'at': '2026-11-26 15:00', 'counterparty': '이종원'},
            {'code': '9012', 'amount': 17_050, 'at': '2026-11-26 15:00', 'counterparty': '오세영'}]
    ap = [{'id': 87, 'amount': 24_200, 'grn': 'R', 'payee_name': '이종원', 'project_code': 'R4294-JW', 'content': '22M 스틸파이프 구매',
           'done_at': '2026-11-26 14:58:31'},
          {'id': 88, 'amount': 3_850, 'grn': 'R', 'payee_name': '이종원', 'project_code': 'R4294-JW', 'content': '전기 부속 구매',
           'done_at': '2026-11-26 14:58:33'},
          {'id': 85, 'amount': 17_050, 'grn': 'G', 'payee_name': '오세영', 'content': '전기 자재', 'done_at': '2026-11-26 14:58:15'}]
    m = fd.match_approvals(outs, ap)
    assert m[0]['ids'] == [87, 88] and m[0]['amount'] == 28_050 and m[0]['project_code'] == 'R4294-JW'
    assert m[0]['content'] == '22M 스틸파이프 구매 외 1건'
    assert 1 not in m                                   # 17,050 은 G 통장 결재라 R 통장 출금과 짝 아님
    # 받는 분이 다르면 더하지 않음
    ap2 = [dict(ap[0]), dict(ap[1], payee_name='다른사람')]
    assert fd.match_approvals(outs[:1], ap2) == {}
