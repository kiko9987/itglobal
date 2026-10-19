# -*- coding: utf-8 -*-
"""[부담자 설정] 입력 → 칸 배정·형식 기록·메모 윗줄('◆') + 입금 파서가 '◆' 줄을 무시하는지 (2026-10-19)."""
import sys

sys.path.insert(0, '.')

import pytest

from dashboard.services import payer_split as ps
from dashboard.services import payment_sync as pay

DEP = '2026/10/10 15:52\n입금 5,509,680원\n이희승\n452***38801011\n기업'
REC = {'프로젝트 코드': 'G4146-JW', '총액 2': '14258200', '계약금': '5509680', '중도금': '', '잔금': '',
       '계약금 부담자': '-', '중도금 부담자': '-', '잔금 부담자': '-'}
LL = {'role': '임대인', 'name': '이희승', 'biz_no': '1221541208', 'amount': '6,887,100', 'method': '계산서'}
TT = {'role': '임차인', 'name': '주식회사 제이케이디학원', 'amount': 7371100, 'method': '계산서',
      'email': 'jagangn@naver.com'}


# ── 입금 파서: '◆' 줄 무시 ──
def test_parser_ignores_payer_line_only_note():
    assert pay._parse_notes(['', '', '◆ 임차인 주식회사 제이케이디학원']) == []


def test_parser_same_payments_with_payer_line():
    plain = pay._parse_notes([DEP, '', ''])
    marked = pay._parse_notes([f'◆ 임대인 이희승\n\n{DEP}', '', ''])
    glued = pay._parse_notes([f'◆ 임대인 이희승\n{DEP}', '', ''])   # 빈 줄 없이 붙여도
    assert plain and marked == plain and glued == plain
    assert pay._hash_payments(marked) == pay._hash_payments(plain)


def test_baseline_unchanged_by_payer_line():
    vals = [''] * 30
    vals[20] = 5509680
    a = pay.row_baseline_mapping(vals, [DEP, '', ''])
    b = pay.row_baseline_mapping(vals, [f'◆ 임대인 이희승\n\n{DEP}', '', '◆ 임차인 제이케이디'])
    assert a == b


def test_special_notes_skip_payer_line():
    text = '헤더\n' + pay._SEP + '\n미수금 : 0원'
    # '상계'가 들어간 상호라도 '◆' 줄은 특이사항이 아님
    assert pay._append_special_notes(text, ['◆ 임대인 상계물산', '', '']) == text


# ── 메모 윗줄 ──
def test_apply_note_line_add_replace_remove():
    a = ps.apply_note_line(DEP, '◆ 임대인 이희승')
    assert a == f'◆ 임대인 이희승\n\n{DEP}'
    assert ps.apply_note_line(a, '◆ 임대인 이희승') == a                     # 다시 해도 같음
    assert ps.apply_note_line(a, '◆ 임차인 다른사람').startswith('◆ 임차인 다른사람\n\n2026/10/10')
    assert ps.apply_note_line(a, '') == DEP                                    # 해제
    assert ps.apply_note_line('', '◆ 임차인 X') == '◆ 임차인 X'
    assert ps.apply_note_line('◆ 임차인 X', '') == ''


def test_format_roundtrip():
    s = ps.format_payer_cell({**LL, 'amount': 6887100})
    assert s == '임대인 · 이희승 · 122-15-41208 · 6,887,100 · 계산서'
    p = ps.parse_payer_cell(s)
    assert (p['name'], p['biz_no'], p['amount']) == ('이희승', '122-15-41208', 6887100)
    s2 = ps.format_payer_cell({**TT})
    assert ps.parse_payer_cell(s2)['email'] == 'jagangn@naver.com'


# ── 칸 배정 ──
def test_plan_assigns_first_contract_second_balance():
    plan, err, warn = ps.plan_payers([LL, TT], REC)
    assert not err and not warn
    assert plan['계약금']['name'] == '이희승' and plan['잔금']['name'] == '주식회사 제이케이디학원'


def test_plan_keeps_existing_person_column_even_if_order_changes():
    rec = dict(REC, **{'계약금 부담자': ps.format_payer_cell({**LL, 'amount': 6887100}),
                       '잔금 부담자': ps.format_payer_cell(TT)})
    plan, err, _ = ps.plan_payers([TT, LL], rec)        # 순서를 바꿔 보내도
    assert not err and plan['계약금']['name'] == '이희승' and plan['잔금']['name'] == '주식회사 제이케이디학원'


def test_plan_errors():
    _, err, _ = ps.plan_payers([{**LL, 'name': ''}], REC)
    assert any('사업자명' in e for e in err)
    _, err, _ = ps.plan_payers([{**LL, 'amount': ''}, TT], REC)
    assert any('부담액' in e for e in err)
    _, err, _ = ps.plan_payers([LL, TT, {**TT, 'name': 'A'}, {**TT, 'name': 'B'}], REC)
    assert any('3명' in e for e in err)
    _, err, _ = ps.plan_payers([{**LL, 'stage': '잔금'}, {**TT, 'stage': '잔금'}], REC)
    assert any('두 사람' in e for e in err)
    # 계약금 칸에 이미 입금이 있는데 그 칸을 비우면 오류(돈 주인 미상)
    _, err, _ = ps.plan_payers([{**LL, 'stage': '중도금'}, {**TT, 'stage': '잔금'}], REC)
    assert any('계약금 칸에 이미 입금' in e for e in err)


def test_plan_total_warning_and_clear():
    _, err, warn = ps.plan_payers([LL, {**TT, 'amount': 7000000}], REC)
    assert not err and warn and '다릅니다' in warn[0]
    plan, err, warn = ps.plan_payers([], REC)           # 해제 — 입금 있어도 허용
    assert plan == {} and not err and not warn
    plan, err, _ = ps.plan_payers([{'role': '임대인'}], REC)   # 이름·금액 없는 빈 줄은 무시
    assert plan == {} and not err


# ── 시트 기록 ──
class _Mgr:
    def __init__(self):
        self.cells = {'AL10': '-', 'AM10': '-', 'AN10': '-', 'U10': 5509680}
        self.notes = {'U10': DEP}
        self.writes = 0

    def find_row_by_project_code(self, *a, **k):
        return 10

    def get_field_to_letter(self):
        return {'계약금': 'U', '중도금': 'V', '잔금': 'W',
                '계약금 부담자': 'AL', '중도금 부담자': 'AM', '잔금 부담자': 'AN'}

    def get_cell_value(self, sid, sn, cell):
        return self.cells.get(cell, '')

    def update_cell_value(self, sid, sn, cell, v):
        self.cells[cell] = v; self.writes += 1
        return True

    def get_cell_note(self, sid, sn, cell):
        return self.notes.get(cell, '')

    def update_cell_note(self, sid, sn, cell, note):
        self.notes[cell] = note; self.writes += 1
        return True


@pytest.fixture
def mgr(monkeypatch):
    m = _Mgr()
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: m)
    monkeypatch.setenv('GOOGLE_SHEET_ID', 'sid')
    monkeypatch.setenv('GOOGLE_SHEET_NAME', '공사 현황')
    return m


def test_save_writes_cells_and_note_lines_idempotent(mgr):
    plan, err, _ = ps.plan_payers([LL, TT], REC)
    changed = ps.save_payers('G4146-JW', plan)
    assert mgr.cells['AL10'] == '임대인 · 이희승 · 122-15-41208 · 6,887,100 · 계산서'
    assert mgr.cells['AN10'].startswith('임차인 · 주식회사 제이케이디학원')
    assert mgr.cells['AM10'] == '-'
    assert mgr.notes['U10'] == f'◆ 임대인 이희승\n\n{DEP}'
    assert mgr.notes['W10'] == '◆ 임차인 주식회사 제이케이디학원'
    assert 'V10' not in mgr.notes
    assert set(changed) == {'계약금 부담자', '잔금 부담자', '계약금_메모', '잔금_메모'}
    n = mgr.writes
    assert ps.save_payers('G4146-JW', plan) == {} and mgr.writes == n     # 다시 저장해도 쓰기 없음
    # 입금 파서 결과는 그대로(카드 영향 없음)
    assert pay._parse_notes([mgr.notes['U10'], '', mgr.notes['W10']]) == pay._parse_notes([DEP, '', ''])


def test_save_empty_keeps_old_depositor_values(mgr):
    """옛 공사(옛 '입금자명' 값)에서 빈 채로 저장해도 옛 값은 지우지 않는다."""
    mgr.cells['AL10'] = '25/02/11 고희원 R'
    changed = ps.save_payers('R2087-SH', {})
    assert mgr.cells['AL10'] == '25/02/11 고희원 R' and '계약금 부담자' not in changed


def test_save_clear_removes_lines(mgr):
    plan, _, _ = ps.plan_payers([LL, TT], REC)
    ps.save_payers('G4146-JW', plan)
    changed = ps.save_payers('G4146-JW', {})
    assert mgr.cells['AL10'] == '-' and mgr.cells['AN10'] == '-'
    assert mgr.notes['U10'] == DEP and mgr.notes['W10'] == ''
    assert '계약금 부담자' in changed and '잔금_메모' in changed
