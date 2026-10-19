# -*- coding: utf-8 -*-
"""[부담자 설정] 창 열기(GET /api/invoice/payers/<code>) — 캐시 우선, 캐시에 없을 때만 시트 (2026-10-20)."""
from unittest.mock import patch

import pytest
from flask import Flask

from dashboard.blueprints import invoice_mgmt as im

REC = {'프로젝트 코드': 'G4148-JW', '총액 2': 17646200, '계약금': 5000000, '중도금': 0, '잔금': 0,
       '계약금 부담자': '임대인 · 홍길동 · 123-45-67890 · 8,000,000 · 계산서',
       '중도금 부담자': '-', '잔금 부담자': '임차인 · 주식회사 테스트 · 9,646,200 · 계산서'}


@pytest.fixture
def app():
    return Flask(__name__)


def _get(app, code):
    with app.test_request_context(f'/api/invoice/payers/{code}'):
        resp = im.api_payers_get.__wrapped__(code)
        resp = resp[0] if isinstance(resp, tuple) else resp
        return resp.get_json(), resp.status_code


def test_open_uses_cache_without_sheet_read(app):
    with patch('dashboard.services.project_service.get_project_records', return_value=[dict(REC)]), \
         patch.object(im, '_payer_rec_from_sheet', side_effect=AssertionError('시트 읽으면 안 됨')):
        body, status = _get(app, 'G4148-JW')
    assert status == 200
    data = body['data']
    assert data['total2'] == 17646200
    assert data['stage_values'] == {'계약금': 5000000, '중도금': 0, '잔금': 0}
    assert [(p['stage'], p['role'], p['amount'], p['received']) for p in data['payers']] == [
        ('계약금', '임대인', 8000000, 5000000), ('잔금', '임차인', 9646200, 0)]


def test_open_falls_back_to_sheet_when_not_cached(app):
    # 막 등록한 공사 — 캐시엔 없고 시트엔 있음
    with patch('dashboard.services.project_service.get_project_records', return_value=[]), \
         patch.object(im, '_payer_rec_from_sheet', return_value=dict(REC, **{'프로젝트 코드': 'G9999-JW'})) as sheet:
        body, status = _get(app, 'G9999-JW')
    assert status == 200 and sheet.call_count == 1
    assert len(body['data']['payers']) == 2


def test_open_cache_error_falls_back_to_sheet(app):
    with patch('dashboard.services.project_service.get_project_records', side_effect=RuntimeError('redis down')), \
         patch.object(im, '_payer_rec_from_sheet', return_value=dict(REC)) as sheet:
        body, status = _get(app, 'G4148-JW')
    assert status == 200 and sheet.call_count == 1


def test_open_not_found_anywhere_is_404(app):
    with patch('dashboard.services.project_service.get_project_records', return_value=[]), \
         patch.object(im, '_payer_rec_from_sheet', return_value=None):
        body, status = _get(app, 'X0000-NO')
    assert status == 404
