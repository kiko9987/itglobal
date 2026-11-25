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


def test_cards_show_gross_and_fee():
    line = sb._card_gross_line(2060000, 2018800)
    assert '2,060,000원' in line and '41,200원' in line
    assert sb._card_gross_line(0, 2018800) == ''
