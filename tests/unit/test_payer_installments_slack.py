# -*- coding: utf-8 -*-
"""나눠 내는 공사 — 슬랙 입금 지정·확인 기록·계산서 연동 (2026-10-20 차수·사람별 상태)."""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb
from tests.unit.test_payer_split_integration import _PayerMgr, HANA, REC_303

LESSOR = '임대인 · 갑상사 · 1,200,000 · 계산서'
TENANT = '임차인 · 을학원 · 3,014,000 · 계산서'


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


# ── 확인 기록(_commit_intake_to_sheet) ──
def test_commit_writes_installment_line_inside_block(pmgr):
    pmgr.cells.update({'AL10': LESSOR, 'AN10': TENANT})
    ok, *_ = sb._commit_intake_to_sheet('G9999-JW', '잔금', 3014000, HANA, 'U_SB',
                                        note_extra=['⚠️ 테스트 경고 (지정 JW · 확인 SB)'], installment='중도금')
    assert ok
    note = pmgr.notes['W10']
    lines = note.split('\n')
    assert '◇ 차수 중도금' in lines
    assert lines.index('◇ 차수 중도금') < lines.index('⚠️ 테스트 경고 (지정 JW · 확인 SB)')
    # 입금 파서는 차수 줄을 입금자로 읽지 않는다
    from dashboard.services.payment_sync import _parse_notes
    pays = _parse_notes(['', '', note], {'잔금': 3014000})
    assert pays and pays[0]['partner'] == '대연이엔지주식회' and pays[0]['amount'] == 3014000


def test_commit_split_fills_dash_column_even_if_other_column_issued(pmgr):
    # G4146 사고 모양: 다른 칸에 '발행'이 있어도 이 사람 칸 '-'는 계산서 대상(미발행)으로 채움
    pmgr.cells.update({'AL10': LESSOR, 'AN10': TENANT, 'Z10': '발행', 'AB10': '-'})
    ok, *_ = sb._commit_intake_to_sheet('G9999-JW', '잔금', 3014000, HANA, 'U_SB', installment='일시불')
    assert ok and pmgr.cells['AB10'] == '미발행'


def test_commit_plain_project_keeps_dash_covered(pmgr):
    # 일반 공사는 예전 그대로 — 다른 칸 '발행'이면 '-'(통합 발행 포함) 유지
    pmgr.cells.update({'Z10': '발행', 'AB10': '-'})
    ok, *_ = sb._commit_intake_to_sheet('R4101-JSH', '잔금', 3014000, HANA, 'U_SB')
    assert ok and pmgr.cells['AB10'] == '-'


def test_commit_split_y_summary_shows_uninvoiced_other_payer(pmgr):
    # 임대인 완납·계산서 없음('-'), 임차인 계산서 발행 → 예전 규칙이면 '잔금 - 발행완료'로 임대인 미발행이 숨음
    pmgr.cells.update({'AL10': LESSOR, 'AN10': TENANT, 'U10': 1200000, 'Z10': '-', 'AB10': '발행'})
    pmgr.notes['Y10'] = '2026-10-15 잔금 발행 3,014,000원 (2,740,000원 VAT 별도) · 을학원'
    ok, *_ = sb._commit_intake_to_sheet('G9999-JW', '잔금', 3014000, HANA, 'U_SB', installment='일시불')
    assert ok and pmgr.cells['Y10'] == '계약금 - 미발행'


def test_card_settlement_uses_payer_remaining():
    class M:
        def get_cell_value(self, *a):
            return ''
    # 임차인 부담 1,000,000 중 0 받음, 카드 순입금 970,000(3% 수수료) → 그 사람 남은 돈 기준이면 보정
    tgt = sb._card_settlement_target(M(), 's', 'n', 10, 'W', 0, 970000, '입금 970,000원\nBC카드 12345678',
                                     snap={'T': 5000000, 'U': 0, 'V': 0, 'W': 0}, payer_amount=1000000)
    assert tgt == 1000000
    # 공사 전체 기준이면(다른 사람 몫 남음) 보정 안 됨
    assert sb._card_settlement_target(M(), 's', 'n', 10, 'W', 0, 970000, '입금 970,000원\nBC카드 12345678',
                                      snap={'T': 5000000, 'U': 0, 'V': 0, 'W': 0}) is None


# ── 확인 때 사람 칸 결정 ──
def test_resolve_intake_payer_legacy_and_plain(monkeypatch):
    from dashboard.services.payer_split import stage_payers
    payers = stage_payers({'계약금 부담자': '임대인 · 이현주 · 8,031,100 · 계산서',
                           '잔금 부담자': '임차인 · 민선홍(서울아산가정의학) · 7,701,100 · 계산서'})
    monkeypatch.setattr(sb, '_intake_split_payers', lambda code: payers if code == 'G4158-JW' else {})
    r = sb._resolve_intake_payer('G4158-JW', '중도금', 6424880, '이현주가정의에어')
    assert r['stage'] == '계약금' and r['moved'] and r['installment'] == '중도금' and r['who'] == '임대인 이현주'
    r = sb._resolve_intake_payer('G4158-JW', '잔금', 1540220, '민선홍', installment='잔금')
    assert r['stage'] == '잔금' and not r['moved'] and r['installment'] == '잔금'
    r = sb._resolve_intake_payer('G4158-JW', '중도금', 100000, '모르는사람')
    assert r['error']
    assert sb._resolve_intake_payer('R4101-JSH', '중도금', 1, 'x') == {
        'stage': '중도금', 'installment': '', 'moved': False, 'error': '', 'who': ''}


# ── 지정 경고 ──
def test_intake_warnings_payer_over_and_mismatch():
    # 임대인 이희승 남은 돈 1,377,420인데 2,000,000을 임대인으로 지정 → 사람 과입금
    w = sb._intake_warnings('G4146-JW', 2000000, {'partner': '이희승'}, '', records=[REC_303], stage='계약금')
    assert any(x['kind'] == 'payer_over' for x in w)
    # 입금자는 임대인 이희승인데 임차인 칸으로 지정 → 확인 경고
    w = sb._intake_warnings('G4146-JW', 1000000, {'partner': '이희승'}, '', records=[REC_303], stage='잔금')
    assert any(x['kind'] == 'payer_mismatch' for x in w)
    w = sb._intake_warnings('G4146-JW', 1000000, {'partner': '이희승'}, '', records=[REC_303], stage='계약금')
    assert not any(x['kind'] in ('payer_over', 'payer_mismatch') for x in w)


# ── 카드 표기 ──
def test_stage_display_and_pending_card():
    assert sb._stage_display('잔금') == '잔금'
    assert sb._stage_display('잔금', '임차인 김지민', '중도금') == '임차인 김지민(잔금 칸) · 차수 중도금'
    blocks = sb._build_intake_pending_blocks('iid', 'G4159-JW', '잔금', 1177880, '입금 1,177,880원', '',
                                             who='임차인 김지민', installment='중도금')
    assert '임차인 김지민(잔금 칸) · 차수 중도금' in blocks[0]['text']['text']


# ── 통합 분할 창 ──
def test_split_modal_row_switches_to_payer_and_installment(monkeypatch):
    from dashboard.services.payer_split import stage_payers
    payers = stage_payers(REC_303)
    monkeypatch.setattr(sb, '_intake_split_payers', lambda code: payers if code == 'G4146-JW' else {})
    pre = [{'project_code': 'G4146-JW', 'project_option': {'text': {'type': 'plain_text', 'text': 'G4146-JW'},
                                                           'value': 'G4146-JW'}, 'stage': '', 'amount': 1377420},
           {'project_code': 'R4101-JSH', 'project_option': None, 'stage': '잔금', 'amount': 1826440}]
    view = sb._build_split_modal_view('iid', 'C', '1', 3203860, '입금 3,203,860원\n이희승\n기업', 2, pre)
    ids = [b.get('block_id') for b in view['blocks']]
    assert 'split_row_0_inst' in ids and 'split_row_1_inst' not in ids
    st0 = next(b for b in view['blocks'] if b.get('block_id') == 'split_row_0_stage')
    assert [o['value'] for o in st0['element']['options']] == ['계약금', '잔금']
    assert st0['element']['initial_option']['value'] == '계약금'          # 입금자 이희승 = 임대인
    p0 = next(b for b in view['blocks'] if b.get('block_id') == 'split_row_0_project')
    assert p0.get('dispatch_action') is True


def test_split_state_to_rows_reads_installment():
    state = {'split_row_0_project': {'split_project_0': {'selected_option': {'value': 'G4146-JW'}}},
             'split_row_0_stage': {'split_stage_0': {'selected_option': {'value': '계약금'}}},
             'split_row_0_amount': {'split_amount_0': {'value': '1,377,420'}},
             'split_row_0_inst': {'split_inst_0': {'selected_option': {'value': '잔금'}}}}
    r = sb._split_state_to_rows(state, 1)[0]
    assert r['installment'] == '잔금' and r['has_inst'] and r['amount'] == 1377420


# ── 계산서 ──
def test_invoice_view_values_and_block_id_with_suffix():
    view = {'state': {'values': {'amt__f': {'value': {'value': '6,701,000'}},
                                 'stages': {'invoice_payer_pick': {'selected_option': {'value': '잔금'}}}}},
            'blocks': [{'block_id': 'stages'}, {'block_id': 'amt__f'}]}
    vals = sb._invoice_view_values(view)
    assert vals['amt']['value'] == '6,701,000'
    assert vals['stages']['selected_option']['value'] == '잔금'
    assert sb._invoice_block_id(view, 'amt') == 'amt__f'


def test_invoice_dedup_key_per_payer(monkeypatch):
    monkeypatch.setattr(sb, '_intake_split_payers', lambda code: {'잔금': {}} if code == 'G4146-JW' else {})
    assert sb._invoice_dedup_key('G4146-JW', '잔금') == 'invoice_request_dedup:G4146-JW:잔금'
    assert sb._invoice_dedup_key('R4101-JSH', '잔금') == 'invoice_request_dedup:R4101-JSH'


def test_split_invoice_modal_view_payer_first():
    opts = [{'text': {'type': 'plain_text', 'text': '임대인 이희승'}, 'value': '계약금'}]
    v = sb._build_invoice_modal_view('G4146-JW', '이희승', '주소', '', '', '{}', payer_opts=opts)
    assert 'submit' not in v and v['blocks'][1]['block_id'] == 'stages'
    v = sb._build_invoice_modal_view('G4146-JW', '이희승', '주소', '1,252,200', 'a@b.c', '{}',
                                     payer_opts=opts, payer_stage='계약금')
    ids = [b.get('block_id') for b in v['blocks']]
    assert 'amt__c' in ids and 'email__c' in ids and ids.count('stages') == 1 and 'submit' in v


def test_project_issued_invoice_per_payer(monkeypatch):
    import dashboard.services.project_service as psv
    rec = dict(REC_303, **{'계약금 계산서': '발행', '잔금 계산서': '발행', '계산서_메모': (
        '2026-10-10 계약금 발행 6,887,100원 (6,261,000원 VAT 별도) · 이희승\n'
        '2026-10-15 잔금 발행 7,371,100원 (6,701,000원 VAT 별도) · 주식회사 제이케이디 학원'),
        '부가세': 'TRUE'})
    monkeypatch.setattr(psv, 'get_project_records', lambda: [rec])
    allp = sb._project_issued_invoice('G4146-JW')
    assert allp['issued_gross'] == 14258200 and allp['payers']['계약금']['issued'] == 6887100
    one = sb._project_issued_invoice('G4146-JW', stage='잔금')
    assert one['issued_gross'] == 7371100 and one['total2'] == 7371100 and one['remaining_supply'] == 0
    # 임차인 부담액 7,371,100인데 또 발행하면 사람 기준 초과 경고
    w = sb._invoice_amount_warning('G4146-JW', '1000000', 'sep', stage='잔금')
    assert '부담액' in w


def test_mark_invoice_split_no_full_dash_and_split_y(pmgr, monkeypatch):
    # 총액1과 같은 금액으로 요청해도(나눠 내는 공사) 다른 사람 칸을 '-'로 덮지 않음
    pmgr.cells.update({'AL10': LESSOR, 'AN10': '임차인 · 을학원 · 12,000,000 · 계산서',
                       'U10': 1200000, 'Z10': '미발행', 'AB10': ''})
    sb._mark_invoice_issued_in_sheet('G9999-JW', '잔금', invoice_amt='12000000', vat_val='sep', biz='을학원')
    assert pmgr.cells['AB10'] == '발행' and pmgr.cells['Z10'] == '미발행'
    assert '잔금 발행' in pmgr.notes['Y10']
    assert pmgr.cells['Y10'] == '계약금 - 미발행'      # 임대인 완납·미발행이 요약에 드러남
