# -*- coding: utf-8 -*-
"""[부담자 설정] 창 열기(GET /api/invoice/payers/<code>) — 캐시 한 행 우선, 캐시에 없을 때만 시트 (2026-10-20)."""
from unittest.mock import patch

import pandas as pd
import pytest
from flask import Flask

from dashboard.blueprints import invoice_mgmt as im

REC = {'프로젝트 코드': 'G4148-JW', '총액 2': 17646200.0, '계약금': 5000000.0, '중도금': float('nan'), '잔금': 0.0,
       '계약금 부담자': '임대인 · 홍길동 · 123-45-67890 · 8,000,000 · 계산서',
       '중도금 부담자': '-', '잔금 부담자': '임차인 · 주식회사 테스트 · 9,646,200 · 계산서'}
OTHER = dict(REC, **{'프로젝트 코드': 'G4147-JW', '계약금 부담자': '-', '잔금 부담자': '-'})
LOAD = 'dashboard.services.project_service.load_data'


@pytest.fixture
def app():
    return Flask(__name__)


def _get(app, code):
    with app.test_request_context(f'/api/invoice/payers/{code}'):
        resp = im.api_payers_get.__wrapped__(code)
        resp = resp[0] if isinstance(resp, tuple) else resp
        return resp.get_json(), resp.status_code


def test_open_uses_cached_row_without_sheet_read(app):
    df = pd.DataFrame([OTHER, REC])
    with patch(LOAD, return_value=df), \
         patch.object(im, '_payer_rec_from_sheet', side_effect=AssertionError('시트 읽으면 안 됨')):
        body, status = _get(app, 'G4148-JW')
    assert status == 200
    data = body['data']
    assert data['total2'] == 17646200
    assert data['stage_values'] == {'계약금': 5000000, '중도금': 0, '잔금': 0}   # NaN → 0
    assert [(p['stage'], p['role'], p['amount'], p['received']) for p in data['payers']] == [
        ('계약금', '임대인', 8000000, 5000000), ('잔금', '임차인', 9646200, 0)]


def test_cached_row_code_with_spaces(app):
    df = pd.DataFrame([dict(REC, **{'프로젝트 코드': ' G4148-JW '})])
    with patch(LOAD, return_value=df), \
         patch.object(im, '_payer_rec_from_sheet', side_effect=AssertionError('시트 읽으면 안 됨')):
        body, status = _get(app, 'G4148-JW')
    assert status == 200 and len(body['data']['payers']) == 2


def test_open_falls_back_to_sheet_when_not_cached(app):
    # 막 등록한 공사 — 캐시엔 없고 시트엔 있음
    with patch(LOAD, return_value=pd.DataFrame([OTHER])), \
         patch.object(im, '_payer_rec_from_sheet', return_value=dict(REC, **{'프로젝트 코드': 'G9999-JW'})) as sheet:
        body, status = _get(app, 'G9999-JW')
    assert status == 200 and sheet.call_count == 1
    assert len(body['data']['payers']) == 2


@pytest.mark.parametrize('loaded', [None, pd.DataFrame()])
def test_open_empty_cache_falls_back_to_sheet(app, loaded):
    with patch(LOAD, return_value=loaded), \
         patch.object(im, '_payer_rec_from_sheet', return_value=dict(REC)) as sheet:
        body, status = _get(app, 'G4148-JW')
    assert status == 200 and sheet.call_count == 1


def test_open_cache_error_falls_back_to_sheet(app):
    with patch(LOAD, side_effect=RuntimeError('redis down')), \
         patch.object(im, '_payer_rec_from_sheet', return_value=dict(REC)) as sheet:
        body, status = _get(app, 'G4148-JW')
    assert status == 200 and sheet.call_count == 1


def test_open_not_found_anywhere_is_404(app):
    with patch(LOAD, return_value=pd.DataFrame([OTHER])), \
         patch.object(im, '_payer_rec_from_sheet', return_value=None):
        body, status = _get(app, 'X0000-NO')
    assert status == 404


# ── PM 계산서 요청 창의 이미 발행 요약(GET /api/invoice/issued/<code>) — 나눠 내는 공사는 사람별 (2026-10-20) ──
RECS = 'dashboard.services.project_service.get_project_records'
ISSUED_REC = dict(REC, **{'총액 1': 16042000, '부가세': True, '계약금 계산서': '발행', '중도금': 0,
                          '계산서_메모': '2026-10-18 계약금 발행 8,000,000원 · 홍길동'})


def _issued(app, code):
    with app.test_request_context(f'/api/invoice/issued/{code}'):
        resp = im.api_invoice_issued.__wrapped__(code)
        resp = resp[0] if isinstance(resp, tuple) else resp
        return resp.get_json()['data']


def test_issued_split_returns_per_payer(app):
    with patch(LOAD, return_value=pd.DataFrame([ISSUED_REC])), patch(RECS, return_value=[ISSUED_REC]):
        data = _issued(app, 'G4148-JW')
    assert data['split'] is True and data['warn'] == '' and data['remaining_supply'] is None
    lessor, tenant = data['payers']['계약금'], data['payers']['잔금']
    assert lessor['issued'] == 8000000 and lessor['remaining_supply'] == 0 and '전액 발행됨' in lessor['warn']
    assert tenant['issued'] == 0 and tenant['warn'] == ''
    assert tenant['remaining_supply'] == round(9646200 * 16042000 / 17646200)   # 부담액의 공급가 전부
    assert [it['stage'] for it in data['active']] == ['계약금']


def test_issued_plain_project_has_no_split_keys(app):
    plain = dict(ISSUED_REC, **{'계약금 부담자': '-', '잔금 부담자': '-'})
    with patch(LOAD, return_value=pd.DataFrame([plain])), patch(RECS, return_value=[plain]):
        data = _issued(app, 'G4148-JW')
    assert 'split' not in data and 'payers' not in data
    assert data['has_issued'] is True and data['remaining_supply'] is not None
