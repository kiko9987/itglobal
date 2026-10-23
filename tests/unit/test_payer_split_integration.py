# -*- coding: utf-8 -*-
"""나눠 내는 공사(부담자 칸) — 입금 지정·계산서 칸 자동채움·경고·추정·등록증·정합성 연동 (2026-10-19)."""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb

HANA = '[Web발신]\n하나,09/30, 17:02\n255******31304\n입금 3,014,000원\n대연이엔지주식회'


class _FakeMgr:
    """시트 매니저 대역 (test_nh_account_invoice_fill 과 같은 모양)."""
    F2L = {'계약금': 'U', '중도금': 'V', '잔금': 'W', '계산서': 'Y', '총액 1': 'R', '총액 2': 'T',
           '수금 확인': 'AD', '계약금 계산서': 'Z', '중도금 계산서': 'AA', '잔금 계산서': 'AB'}

    def __init__(self, **cells):
        self.cells = {'R10': 12000000, 'T10': 13200000, 'U10': 1200000, 'V10': 0, 'W10': 0,
                      'Y10': '계약금 - 미발행', 'Z10': '미발행', 'AA10': '-', 'AB10': '',
                      'AD10': 'FALSE'}
        self.cells.update(cells)
        self.notes = {}

    def find_row_by_project_code(self, *a, **k):
        return 10

    def get_cell_value(self, sid, sn, cell):
        return self.cells.get(cell, '')

    def update_cell_value(self, sid, sn, cell, v):
        self.cells[cell] = v
        return True

    def get_cell_note(self, sid, sn, cell):
        return self.notes.get(cell, '')

    def update_cell_note(self, sid, sn, cell, note):
        self.notes[cell] = note
        return True

    def get_field_to_letter(self):
        return dict(self.F2L)

REC_303 = {
    '프로젝트 코드': 'G4146-JW', '총액 1': '12962000', '총액 2': '14258200', '미수금': '8748520',
    '계약금': '5509680', '중도금': '0', '잔금': '0', '사업자명': '이희승',
    '계약금 부담자': '임대인 · 이희승 · 122-15-41208 · 6,887,100 · 계산서',
    '중도금 부담자': '-',
    '잔금 부담자': '임차인 · 주식회사 제이케이디학원 · 405-81-13558 · 7,371,100 · 계산서',
}
PLAIN = {'프로젝트 코드': 'R4101-JSH', '총액 1': '11000000', '총액 2': '12000000', '미수금': '3000000',
         '계약금': '9000000', '중도금': '0', '잔금': '0', '계약금 부담자': '25/02/11 고희원 R'}


class _PayerMgr(_FakeMgr):
    F2L = dict(_FakeMgr.F2L, **{'계약금 부담자': 'AL', '중도금 부담자': 'AM', '잔금 부담자': 'AN'})


@pytest.fixture
def pmgr(monkeypatch):
    m = _PayerMgr()
    monkeypatch.setenv('GOOGLE_SHEET_ID', 'sid')
    monkeypatch.setenv('GOOGLE_SHEET_NAME', '공사 현황')
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: m)
    import dashboard.utils.user_database as udb     # ⚠️ 운영 users.db 감사 로그 차단
    monkeypatch.setattr(udb, 'get_audit_repository',
                        lambda: type('A', (), {'log_action': lambda *a, **k: None})())
    return m


def test_cash_payer_column_fills_n_ipgeum_even_from_hana(pmgr):
    """계산서 없이 현금으로 내는 부담자가 하나 통장으로 보내도 '미발행' 대신 'N입금'."""
    pmgr.cells['AN10'] = '임차인 · 이동엽 · 3,252,000 · 현금'
    ok, *_ = sb._commit_intake_to_sheet('G4144-JW', '잔금', 3014000, HANA, 'U_SB')
    assert ok and pmgr.cells['AB10'] == 'N입금'


def test_card_payer_and_invoice_payer(pmgr):
    pmgr.cells['AN10'] = '임차인 · 누구 · 3,252,000 · 카드'
    ok, *_ = sb._commit_intake_to_sheet('G4144-JW', '잔금', 3014000, HANA, 'U_SB')
    assert ok and pmgr.cells['AB10'] == '카드'


def test_invoice_payer_keeps_mibalhaeng(pmgr):
    pmgr.cells['AN10'] = '임차인 · 누구 · 3,252,000 · 계산서'
    ok, *_ = sb._commit_intake_to_sheet('G4144-JW', '잔금', 3014000, HANA, 'U_SB')
    assert ok and pmgr.cells['AB10'] == '미발행'


def test_short_final_warning_skipped_for_split_only():
    w_split = sb._intake_warnings('G4146-JW', 5_896_880, {}, '', records=[REC_303], stage='잔금')
    assert not [w for w in w_split if w['kind'] == 'short_final']
    w_plain = sb._intake_warnings('R4101-JSH', 1_000_000, {}, '', records=[PLAIN], stage='잔금')
    assert [w for w in w_plain if w['kind'] == 'short_final']


def test_intake_view_shows_payers_and_preselects(monkeypatch):
    import dashboard.services.project_service as psv
    monkeypatch.setattr(psv, 'get_project_records', lambda: [REC_303])
    sms = '[Web발신]\n하나,10/20, 10:00\n255******31304\n입금 1,377,420원\n이희승'
    view = sb._build_payment_intake_view(
        'iid', 'C1', '1.0', sms,
        selected_project={'text': {'type': 'plain_text', 'text': 'G4146-JW'}, 'value': 'G4146-JW'},
        project_details={'biz': '이희승'})
    stage = next(b for b in view['blocks'] if b.get('block_id') == 'stage')
    # 2026-10-20: 사람 칸만(부담자 없는 중도금 칸은 못 고름) + 남은 돈, 라벨 '누가 낸 돈인가요', 차수 칸 필수
    assert [o['value'] for o in stage['element']['options']] == ['계약금', '잔금']
    texts = [o['text']['text'] for o in stage['element']['options']]
    assert texts[0].startswith('임대인 이희승 · 남은 돈') and texts[1].startswith('임차인 주식회사 제이케이디학원')
    assert stage['label']['text'] == '누가 낸 돈인가요'
    assert stage['element']['initial_option']['value'] == '계약금'
    inst = next(b for b in view['blocks'] if b.get('block_id') == 'installment')
    assert [o['value'] for o in inst['element']['options']] == ['계약금', '중도금', '잔금', '일시불']
    assert not inst.get('optional')
    assert any('나눠 내는 공사' in str(b) for b in view['blocks'])


def test_intake_view_plain_project_unchanged(monkeypatch):
    import dashboard.services.project_service as psv
    monkeypatch.setattr(psv, 'get_project_records', lambda: [PLAIN])
    view = sb._build_payment_intake_view(
        'iid', 'C1', '1.0', HANA,
        selected_project={'text': {'type': 'plain_text', 'text': 'R4101-JSH'}, 'value': 'R4101-JSH'},
        project_details={'biz': 'x'})
    stage = next(b for b in view['blocks'] if b.get('block_id') == 'stage')
    assert [o['text']['text'] for o in stage['element']['options']] == ['계약금', '중도금', '잔금']
    assert 'initial_option' not in stage['element']


def test_search_finds_project_by_payer_name(monkeypatch):
    import dashboard.services.project_service as psv
    monkeypatch.setattr(psv, 'get_project_records', lambda: [REC_303, PLAIN])
    opts = sb._build_payment_intake_options('제이케이디')
    assert [o['value'] for o in opts] == ['G4146-JW']


def test_pin_guess_returns_code_and_role():
    from dashboard.services.pin_remind import _guess_project
    assert _guess_project('이희승', 1_377_420, [REC_303]) == 'G4146-JW 임대인'


def test_invariant_split_check():
    from dashboard.services.invariant_checks import check_split_payers
    assert check_split_payers({'records': [REC_303, PLAIN]}) == []
    bad = dict(REC_303, **{'총액 2': '14858200'})
    v = check_split_payers({'records': [bad]})
    assert len(v) == 1 and v[0]['check'] == 'split_payer' and '부담액 합계' in v[0]['detail']


def test_split_payer_license_choice(monkeypatch):
    import dashboard.services.project_service as psv
    import dashboard.services.business_license_handler as blh
    monkeypatch.setattr(psv, 'get_project_records', lambda: [REC_303, PLAIN])
    monkeypatch.setattr(blh, '_own_canonical', lambda code, drive=None: ('F_OWN', 'jpg'))
    monkeypatch.setattr(blh, '_file_biz', lambda fid: blh._norm_biz('이희승') if fid == 'F_OWN' else '')
    asked = []

    def _find(code, payer):
        asked.append(payer['name'])
        if payer['name'] == '주식회사 제이케이디학원':
            return {'content': b'x', 'file_name': 'G3999-XX 사업자등록증.pdf', 'mimetype': 'application/pdf'}
        return None
    monkeypatch.setattr(blh, 'find_payer_license', _find)
    # 프로젝트 등록증 = 임대인 것 → 임대인 요청은 기존 동작
    assert sb._split_payer_license('G4146-JW', '이희승') is None
    # 임차인 요청 → 임차인 등록증(다른 파일) — 이 공사·부담자 이름으로 붙임
    kind, lic = sb._split_payer_license('G4146-JW', '주식회사 제이케이디 학원')
    assert kind == 'use' and lic['file_name'] == 'G4146-JW 사업자등록증_주식회사 제이케이디학원.pdf'
    assert lic['content'] == b'x' and asked == ['주식회사 제이케이디학원']
    # 임차인 등록증을 못 찾으면 붙이지 않음(엉뚱한 첨부 방지) — 부담자를 돌려줘 거래처 탭을 번호로도 찾게
    monkeypatch.setattr(blh, 'find_payer_license', lambda code, payer: None)
    kind, payer = sb._split_payer_license('G4146-JW', '주식회사 제이케이디 학원')
    assert kind == 'skip' and payer['name'] == '주식회사 제이케이디학원'
    # OCR 이 '명:이현주'처럼 읽힌 프로젝트 등록증도 그 부담자 것으로 인정
    monkeypatch.setattr(blh, '_file_biz', lambda fid: '명:이희승' if fid == 'F_OWN' else '')
    assert sb._split_payer_license('G4146-JW', '이희승') is None
    # 평범한 공사·부담자와 무관한 사업자명은 기존 동작
    assert sb._split_payer_license('R4101-JSH', '아무개') is None
    assert sb._split_payer_license('G4146-JW', '전혀다른회사') is None
