# -*- coding: utf-8 -*-
"""2026-11-25 G4160-JW — 카드 중간 결제는 수금 지정 때 실결제(수수료 빼기 전)를 받는다."""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb

CARD_MEMO = '2026/11/24 05:53\n입금 2,018,800원\nKB97390776\n452***38801011\n기업'
BANK_MEMO = '2026/11/24 19:22\n입금 800,000원\n오남(필라테스)\n452***38801011\n기업'


@pytest.fixture
def records(monkeypatch):
    recs = []
    import dashboard.services.project_service as ps
    monkeypatch.setattr(ps, 'get_project_records', lambda *a, **k: recs)
    return recs


def _plain(code, total2, paid=0):
    return {'프로젝트 코드': code, '총액 2': total2, '계약금': 0, '중도금': 0, '잔금': paid}


def test_middle_card_payment_needs_gross(records):
    # 남은 돈 3,000,000 중 2,018,800 카드 순입금 — 자동 보정 기준이 없어 실결제를 꼭 받음
    records.append(_plain('G9001-JW', 3000000))
    gross, err = sb._card_gross_check('G9001-JW', '잔금', 2018800, '', CARD_MEMO)
    assert gross is None and '카드로 결제한 금액' in err


def test_final_card_payment_auto(records):
    # 남은 돈 2,060,000 을 카드로 다 냄(순입금 2% 적음) — 기록 때 자동으로 맞추니 비워도 됨
    records.append(_plain('G9002-JW', 2060000))
    assert sb._card_gross_check('G9002-JW', '잔금', 2018800, '', CARD_MEMO) == (None, None)


def test_bank_deposit_not_asked(records):
    records.append(_plain('G9003-JW', 3000000))
    assert sb._card_gross_check('G9003-JW', '잔금', 800000, '', BANK_MEMO) == (None, None)


def test_gross_value_checked():
    assert sb._card_gross_check('X', '잔금', 2018800, '2,060,000', CARD_MEMO) == (2060000, None)
    g, err = sb._card_gross_check('X', '잔금', 2018800, '2000000', CARD_MEMO)
    assert g is None and '작을 수 없습니다' in err
    g, err = sb._card_gross_check('X', '잔금', 2018800, '20600000', CARD_MEMO)   # 0 하나 더
    assert g is None and '다시 확인' in err


def test_split_payer_uses_person_remaining(records, monkeypatch):
    # 나눠 내는 공사 — 공사 전체가 아니라 그 사람이 남은 돈 기준
    records.append(_plain('G4160-JW', 11963700))
    import dashboard.services.payer_split as psp
    monkeypatch.setattr(psp, 'stage_payers', lambda rec: {'잔금': {'remaining': 5711849}})
    g, err = sb._card_gross_check('G4160-JW', '잔금', 2018800, '', CARD_MEMO)
    assert g is None and err
    monkeypatch.setattr(psp, 'stage_payers', lambda rec: {'잔금': {'remaining': 2060000}})
    assert sb._card_gross_check('G4160-JW', '잔금', 2018800, '', CARD_MEMO) == (None, None)


def test_modal_shows_gross_field_only_for_card():
    def blocks(text, **kw):
        return sb._build_payment_intake_view('i', 'c', 't', text, **kw)['blocks']
    card = [b for b in blocks(CARD_MEMO, card_gross_value='2,060,000') if b.get('block_id') == 'card_gross']
    assert card and card[0]['optional'] and card[0]['element']['initial_value'] == '2,060,000'
    assert not [b for b in blocks(BANK_MEMO) if b.get('block_id') == 'card_gross']


def test_split_modal_gross_per_row(monkeypatch):
    monkeypatch.setattr(sb, '_intake_split_payers', lambda code: {})
    v = sb._build_split_modal_view('i', 'c', 't', 4037600, CARD_MEMO, 2,
                                   [{'card_gross': 2060000}, {}])
    gross = [b for b in v['blocks'] if str(b.get('block_id', '')).endswith('_gross')]
    assert [b['block_id'] for b in gross] == ['split_row_0_gross', 'split_row_1_gross']
    assert gross[0]['element']['initial_value'] == '2,060,000'
    v2 = sb._build_split_modal_view('i', 'c', 't', 800000, BANK_MEMO, 2, [])
    assert not [b for b in v2['blocks'] if str(b.get('block_id', '')).endswith('_gross')]


def test_split_state_reads_gross():
    state = {'split_row_0_gross': {'split_gross_0': {'value': ' 2,060,000 '}}}
    rows = sb._split_state_to_rows(state, 2)
    assert rows[0]['card_gross'] == '2,060,000' and rows[0]['has_gross']
    assert rows[1]['card_gross'] == '' and not rows[1]['has_gross']


def test_split_cards_show_gross(monkeypatch):
    monkeypatch.setattr(sb, '_resolve_manager_initial', lambda u: 'JW')
    splits = [{'project_code': 'G4160-JW', 'stage': '잔금', 'amount': 2018800, 'card_gross': 2060000},
              {'project_code': 'G9001-JW', 'stage': '잔금', 'amount': 1000000, 'card_gross': 0}]
    for blocks in (sb._build_intake_split_pending_blocks('i', splits, 3018800, 'u', CARD_MEMO),
                   sb._build_intake_split_done_blocks(splits, 3018800, 'u', 'u', CARD_MEMO)):
        text = blocks[0]['text']['text']
        assert text.count('💳 카드 실결제') == 1 and '41,200원' in text


@pytest.fixture
def cardpay(monkeypatch):
    reqs = []
    import dashboard.services.e_approval as ea
    monkeypatch.setattr(ea, 'find_by_project', lambda form, code: [r for r in reqs if r['code'] == code])
    return reqs


def test_cardpay_prefill_single_amount(cardpay):
    # G4160: 카드결제 요청 #27·#55 둘 다 2,060,000 → 그 금액으로 미리 채움
    cardpay += [{'id': 55, 'amount': 2060000, 'code': 'G4160-JW'}, {'id': 27, 'amount': 2060000, 'code': 'G4160-JW'},
                {'id': 9, 'amount': 5000000, 'code': 'G4160-JW'}]   # 다른 결제(범위 밖)는 무시
    amt, note = sb._cardpay_gross_hint('G4160-JW', 2018800)
    assert amt == 2060000 and '#55' in note and '#27' in note and '#9' not in note


def test_cardpay_ambiguous_or_none(cardpay):
    cardpay += [{'id': 1, 'amount': 2060000, 'code': 'X'}, {'id': 2, 'amount': 2080000, 'code': 'X'}]
    amt, note = sb._cardpay_gross_hint('X', 2018800)
    assert amt is None and '맞는 금액' in note
    assert sb._cardpay_gross_hint('Y', 2018800) == (None, '')


def test_modal_prefills_from_cardpay(cardpay):
    cardpay.append({'id': 55, 'amount': 2060000, 'code': 'G4160-JW'})
    v = sb._build_payment_intake_view('i', 'c', 't', CARD_MEMO,
                                      selected_project={'value': 'G4160-JW', 'text': {'type': 'plain_text', 'text': 'G4160-JW'}})
    g = [b for b in v['blocks'] if b.get('block_id') == 'card_gross'][0]
    assert g['element']['initial_value'] == '2,060,000'
    assert any('#55' in str(b) for b in v['blocks'] if b.get('type') == 'context')


def test_split_row_prefills_from_cardpay(cardpay, monkeypatch):
    monkeypatch.setattr(sb, '_intake_split_payers', lambda code: {})
    cardpay.append({'id': 55, 'amount': 2060000, 'code': 'G4160-JW'})
    v = sb._build_split_modal_view('i', 'c', 't', 4037600, CARD_MEMO, 2,
                                   [{'project_code': 'G4160-JW', 'amount': 2018800}, {}])
    g = [b for b in v['blocks'] if b.get('block_id') == 'split_row_0_gross'][0]
    assert g['element']['initial_value'] == '2,060,000'


def test_cards_show_gross_and_fee():
    line = sb._card_gross_line(2060000, 2018800)
    assert '2,060,000원' in line and '41,200원' in line
    assert sb._card_gross_line(0, 2018800) == ''
