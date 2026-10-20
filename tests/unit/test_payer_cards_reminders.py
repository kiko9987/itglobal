# -*- coding: utf-8 -*-
"""나눠 내는 공사 — 수금 카드 문구·수금완료 칸·미발행 리마인드·반환 짝짓기 (2026-10-20)."""
import time

from dashboard.services import payment_sync as ps
from dashboard.services import invoice_collection_remind as icr
from dashboard.services import refund_match as rm
from dashboard.services.payer_split import stage_payers

LESSOR = '임대인 · 이현주 · 1,802,350 · 계산서'
TENANT = '임차인 · 김지민(다정약국) · 1,472,350 · 계산서'
PAYERS = stage_payers({'계약금 부담자': LESSOR, '잔금 부담자': TENANT})


def _pays():
    return [
        {'stage': '계약금', 'date_md': '10/17', 'date_year': '2026', 'amount': 1441880, 'partner': '이현주약국에어컨',
         'bank': '기업', 'installment': '중도금'},
        {'stage': '잔금', 'date_md': '10/17', 'date_year': '2026', 'amount': 1177880, 'partner': '다정약국.김지민',
         'bank': '기업', 'installment': '중도금'},
    ]


def test_memo_block_parses_installment_and_keeps_partner():
    p = ps._parse_memo_block('2026/10/17 19:55\n입금 1,177,880원\n다정약국.김지민\n452***38801011\n기업\n◇ 차수 중도금')
    assert p['installment'] == '중도금' and p['partner'] == '다정약국.김지민' and p['amount'] == 1177880


def test_stage_card_split_title_history_and_payer_line(monkeypatch):
    monkeypatch.setattr(ps, '_split_payers_for', lambda project: PAYERS if project == 'G4159-JW' else {})
    pays = _pays()
    txt = ps._build_stage_with_history_message(
        stage='잔금', project='G4159-JW', address='주소', last_payment=pays[1], all_payments=pays,
        invoice_value='', total_r=0, total_t=3274700, unpaid=654940,
        stage_sheet_vals={'계약금': 1441880, '중도금': 0, '잔금': 1177880})
    assert '임차인 김지민(다정약국) 입금* · 차수 중도금 (잔금 칸)' in txt
    assert '이현주약국에어컨' not in txt                 # 다른 사람(임대인) 입금은 섞지 않음
    assert '중도금  26/10/17' in txt                     # 이력 줄 머리 = 차수
    assert '임차인 김지민(다정약국) · 부담 1,472,350원 · 받음 1,177,880원 · 남음 294,470원' in txt


def test_stage_card_plain_unchanged(monkeypatch):
    monkeypatch.setattr(ps, '_split_payers_for', lambda project: {})
    pays = [dict(p, installment='') for p in _pays()]
    txt = ps._build_stage_with_history_message(
        stage='잔금', project='R1-XX', address='주소', last_payment=pays[1], all_payments=pays,
        invoice_value='', total_r=0, total_t=3274700, unpaid=0,
        stage_sheet_vals={'계약금': 1441880, '중도금': 0, '잔금': 1177880})
    assert '*잔금 입금*' in txt and '이현주약국에어컨' in txt and '부담' not in txt


def test_complete_message_labels_and_payer_lines(monkeypatch):
    monkeypatch.setattr(ps, '_split_payers_for', lambda project: PAYERS)
    txt = ps._build_complete_message('G4159-JW', '주소', _pays(), '', 3274700,
                                     stage_sheet_vals={'계약금': 1802350, '중도금': 0, '잔금': 1472350})
    assert '임대인 중도금  26/10/17' in txt and '임차인 중도금  26/10/17' in txt
    assert '임대인 이현주 · 부담 1,802,350원 · 받음 1,802,350원 · 완납' in txt


def test_completes_project_split_any_column(monkeypatch):
    monkeypatch.setattr(ps, '_split_payers_for', lambda project: PAYERS if project == 'S' else {})
    assert ps._completes_project('계약금', 'S', 0) is True        # 임대인이 마지막으로 완납
    assert ps._completes_project('계약금', 'P', 0) is False       # 일반 공사는 잔금만
    assert ps._completes_project('잔금', 'P', 0) is True
    assert ps._completes_project('잔금', 'S', 100) is False


def test_rerender_split_complete_only_on_remembered_column(monkeypatch):
    monkeypatch.setattr(ps, '_split_payers_for', lambda project: PAYERS)
    monkeypatch.setattr(ps, '_split_complete_stage', lambda project: '계약금')
    c = {'unpaid': 0, 'address': '주소', 'invoice': '', 'total_t': 3274700, 'total_r': 0}
    vals = {'계약금': 1802350, '중도금': 0, '잔금': 1472350}
    assert '*수금완료*' in ps._build_stage_card_text('계약금', 'G4159-JW', _pays(), c, vals)
    assert '*수금완료*' not in ps._build_stage_card_text('잔금', 'G4159-JW', _pays(), c, vals)


# ── 미발행 리마인드 ──
def _row(**kw):
    r = {'프로젝트 코드': 'G4159-JW', '총액 2': 3274700, '미수금': 0, '사업자명': '-',
         '계약금': 1802350, '중도금': 0, '잔금': 1472350, '계약금 부담자': LESSOR, '중도금 부담자': '-',
         '잔금 부담자': TENANT, '계약금 계산서': '-', '중도금 계산서': '-', '잔금 계산서': '발행',
         '계산서_메모': '2026-10-18 잔금 발행 1,472,350원 (1,338,500원 VAT 별도) · 다정약국'}
    r.update(kw)
    return r


def test_remind_split_payer_uninvoiced_not_hidden_by_other_issued():
    bs = icr._bill_stages(_row())
    assert bs['계약금'] == '미발행' and bs['잔금'] == '발행'      # 예전 규칙이면 '-' covered 로 숨음
    assert not icr._has_issued(_row())
    txt = icr._unbilled(_row())
    assert '임대인 이현주 1,802,350원' in txt and '임차인 발행 완료' in txt


def test_remind_split_pending_while_unpaid():
    bs = icr._bill_stages(_row(**{'계약금': 1441880, '미수금': 360470}))
    assert bs['계약금'] == '발행예정'


# ── 반환 짝짓기 ──
def test_refund_match_split_column_overpay():
    vals = ['G4159-JW'] + [''] * 18 + [3274700, 1441880, 0, 2472350] + [''] * 14 + [LESSOR, '-', TENANT]
    assert vals[19] == 3274700 and vals[22] == 2472350 and vals[39] == TENANT
    note_w = '2026/10/15 10:00\n입금 1,000,000원\n김지민\n기업'
    rows = [{'code': 'G4159-JW', 'row': 4160, 'vals': vals, 'notes': ['', '', note_w]}]
    out = rm.find_candidates({'amount': 1000000, 'partner': '김지민', 'bank': '기업', 'ts': time.time()}, rows)
    assert len(out) == 1 and out[0]['stage'] == '잔금' and out[0]['kind'] == 'overpay'
    assert '부담액 1,472,350원' in out[0]['reason']
