# -*- coding: utf-8 -*-
"""나눠 내는 공사 — 부담액 부가세 별도 표기·'현금 · 부가세 없음' 표시·합계 점검 (2026-10-21).

칸 금액은 그대로 '받을 돈'(부가세 포함). PM 창이 부가세 별도로 보여 주고 입력받아 환산해 보낸다.
현금은 공사마다 다름(G4148 임대인 600만 = 부가세 없음, G4151·G4152 임차인 현금 = 부가세 포함) → 사람별 표시.
"""
from unittest.mock import patch

import pandas as pd
from flask import Flask

from dashboard.services import payer_split as ps
from dashboard.blueprints import invoice_mgmt as im

G4148 = {'프로젝트 코드': 'G4148-JW', '총액 1': 16042000, '총액 2': 17646200, '부가세': True,
         '계약금': 0, '중도금': 0, '잔금': 0,
         '계약금 부담자': '임대인 · 유미숙 · 6,000,000 · 현금',
         '중도금 부담자': '-', '잔금 부담자': '임차인 · 주식회사 제이케이디학원 · 11,046,200 · 계산서'}
G4148_NOVAT = dict(G4148, **{'계약금 부담자': '임대인 · 유미숙 · 6,000,000 · 현금 · 부가세 없음'})
G4151 = {'총액 1': 11842000, '총액 2': 13026200,
         '계약금 부담자': '임대인 · 남상철 · 6,050,000 · 계산서',
         '잔금 부담자': '임차인 · 권영진(지담학원) · 6,976,200 · 현금'}
G4144 = {'총액 1': 13252000, '총액 2': 13252000,
         '계약금 부담자': '임대인 · 임기환 · 10,000,000 · 현금',
         '잔금 부담자': '임차인 · 이동엽(클라이맥스국어학원) · 3,252,000 · 현금'}


def test_parse_vat_free_token_only_with_cash():
    p = ps.parse_payer_cell('임대인 · 유미숙 · 6,000,000 · 현금 · 부가세 없음')
    assert p['vat_free'] is True and p['name'] == '유미숙' and p['method'] == '현금' and p['note'] == ''
    assert ps.parse_payer_cell('임대인 · 유미숙 · 6,000,000 · 현금')['vat_free'] is False
    # 계산서·카드는 부가세가 붙으므로 표시가 있어도 무시
    assert ps.parse_payer_cell('임대인 · 유미숙 · 6,000,000 · 계산서 · 부가세없음')['vat_free'] is False


def test_format_round_trip_keeps_token():
    p = ps.parse_payer_cell('임대인 · 유미숙 · 6,000,000 · 현금 · 부가세 없음')
    cell = ps.format_payer_cell(p)
    assert cell == '임대인 · 유미숙 · 6,000,000 · 현금 · 부가세 없음'
    assert ps.parse_payer_cell(cell)['vat_free'] is True
    assert ps.format_payer_cell(dict(p, method='계산서')).endswith('· 계산서')   # 계산서면 표시 안 붙음


def test_supply_conversion():
    r = ps.vat_ratio(G4148)
    assert abs(r - 1.1) < 1e-9 and ps.vat_ratio(G4144) == 1.0 and ps.vat_ratio({}) == 1.0
    payers = ps.stage_payers(G4148_NOVAT)
    assert ps.payer_supply(payers['계약금'], r) == 6000000          # 부가세 없음 → 그대로
    assert ps.payer_supply(payers['잔금'], r) == 10042000           # 11,046,200 / 1.1


def test_share_check_vat_basis():
    ok = list(ps.stage_payers(G4151).values())          # 현금도 부가세 포함 — 별도 합 = 총액 1
    assert ps.share_check(ok, G4151) == ''
    assert ps.share_check(list(ps.stage_payers(G4144).values()), G4144) == ''   # 부가세 없는 공사
    # G4148 표시 없음 → 임대인 600만을 부가세 포함으로 보고 별도 합이 총액 1 과 안 맞음
    msg = ps.share_check(list(ps.stage_payers(G4148).values()), G4148)
    assert msg.startswith('부담액 합계(VAT 별도) 15,496,545원') and '총액 1' in msg
    # 표시 있음 → 별도 합이 총액 1 과 맞음. 총액 2 는 시스템이 그 몫을 빼고 계산하므로(10-21 test_payer_total2)
    #   '받을 돈 ≠ 총액 2' 경고는 더 이상 없음
    assert ps.share_check(list(ps.stage_payers(G4148_NOVAT).values()), G4148_NOVAT) == ''


def test_split_problems_uses_same_check():
    assert not any('부담액 합계' in p for p in ps.split_problems(G4148_NOVAT))
    assert any('부담액 합계' in p for p in ps.split_problems(G4148))
    assert not any('부담액 합계' in p for p in ps.split_problems(dict(G4151, **{'계약금': 0, '잔금': 0})))


def test_plan_payers_keeps_vat_free_for_cash_only():
    rows = [{'role': '임대인', 'name': '유미숙', 'amount': '6000000', 'method': '현금', 'vat_free': True},
            {'role': '임차인', 'name': '제이케이디학원', 'amount': '11046200', 'method': '계산서', 'vat_free': True}]
    plan, errs, warns = ps.plan_payers(rows, dict(G4148, **{'계약금 부담자': '-', '잔금 부담자': '-'}))
    assert not errs
    assert plan['계약금']['vat_free'] is True and plan['잔금']['vat_free'] is False
    assert ps.format_payer_cell(plan['계약금']).endswith('현금 · 부가세 없음')
    assert warns == []      # 별도 합 = 총액 1, 총액 2 는 저장 때 시스템이 맞춤


def test_payers_api_gives_total1_ratio_and_flag():
    app = Flask(__name__)
    with patch('dashboard.services.project_service.load_data', return_value=pd.DataFrame([G4148_NOVAT])), \
         patch.object(im, '_payer_rec_from_sheet', side_effect=AssertionError('시트 읽으면 안 됨')):
        with app.test_request_context('/api/invoice/payers/G4148-JW'):
            resp = im.api_payers_get.__wrapped__('G4148-JW')
            resp = resp[0] if isinstance(resp, tuple) else resp
            data = resp.get_json()['data']
    assert data['total1'] == 16042000 and abs(data['vat_ratio'] - 1.1) < 1e-9
    assert [(p['stage'], p['vat_free']) for p in data['payers']] == [('계약금', True), ('잔금', False)]
