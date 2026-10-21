# -*- coding: utf-8 -*-
"""슬랙 입금 지정 모달·계산서 창에 '현금 · 부가세 없음' 표시 (2026-10-21, 305호 G4148)."""
import json
from unittest.mock import patch, MagicMock

from dashboard.services import payer_split as ps

G4148 = {'프로젝트 코드': 'G4148-JW', '총액 1': 16042000, '총액 2': 17046200, '부가세': True,
         '계약금': 0, '중도금': 0, '잔금': 0, '현장 주소': '인천 부평구 부영로189번길 51 오남프라자 305호',
         '계약금 부담자': '임대인 · 유미숙 · 6,000,000 · 현금 · 부가세 없음',
         '중도금 부담자': '-', '잔금 부담자': '임차인 · 주식회사 제이케이디학원 · 11,046,200 · 계산서'}
RECS = 'dashboard.services.project_service.get_project_records'


def test_payer_line_marks_vat_free():
    payers = ps.stage_payers(G4148)
    assert ps.payer_line(payers['계약금']) == '임대인 유미숙 · 부담 6,000,000 · 현금·부가세 없음 · 남음 6,000,000'
    assert ps.payer_line(payers['잔금']) == '임차인 주식회사 제이케이디학원 · 부담 11,046,200 · 남음 11,046,200'


def test_project_details_amount_shows_vat_free_part():
    from dashboard.services.as_service import get_project_details
    with patch(RECS, return_value=[G4148]):
        assert get_project_details('G4148-JW')['amount'] == '16,042,000원 (VAT 별도 · 그중 6,000,000원은 부가세 없음)'
    plain = dict(G4148, **{'계약금 부담자': '-', '잔금 부담자': '-'})
    with patch(RECS, return_value=[plain]):
        assert get_project_details('G4148-JW')['amount'] == '16,042,000원 (VAT 별도)'


def test_intake_modal_context_shows_vat_free():
    from dashboard.blueprints import slack_bot as sb
    with patch(RECS, return_value=[G4148]):
        view = sb._build_payment_intake_view(
            'i1', 'C1', '1.0', '입금 6,000,000원\n유미숙',
            selected_project={'text': {'type': 'plain_text', 'text': 'G4148-JW'}, 'value': 'G4148-JW'},
            project_details={'amount': '16,042,000원 (VAT 별도 · 그중 6,000,000원은 부가세 없음)'})
    txt = json.dumps(view, ensure_ascii=False)
    assert '그중 6,000,000원은 부가세 없음' in txt
    assert '계약금 칸 = 임대인 유미숙 · 부담 6,000,000 · 현금·부가세 없음' in txt


def test_slack_invoice_modal_cash_payer_note_and_label():
    from dashboard.blueprints import slack_bot as sb
    client = MagicMock()
    payers = ps.stage_payers(G4148)
    with patch(RECS, return_value=[dict(G4148, 계산서_메모='')]), \
         patch('dashboard.blueprints.invoice_mgmt._payer_invoice_prefill', return_value={}), \
         patch.object(sb, '_partner_status_warn', return_value=''):
        sb._build_split_invoice_modal(client, 'V1', 'G4148-JW', payers, '-', '주소', '16,042,000', [],
                                      payer_stage='계약금')
    view = json.dumps(client.views_update.call_args.kwargs['view'], ensure_ascii=False)
    assert '현금으로 내기로 돼 있습니다' in view and '부담액 (부가세 없음)' in view
    assert '임대인 유미숙 · 부담 6,000,000 · 현금·부가세 없음' in view          # 고르는 목록
    with patch(RECS, return_value=[dict(G4148, 계산서_메모='')]), \
         patch('dashboard.blueprints.invoice_mgmt._payer_invoice_prefill', return_value={}), \
         patch.object(sb, '_partner_status_warn', return_value=''):
        sb._build_split_invoice_modal(client, 'V1', 'G4148-JW', payers, '-', '주소', '16,042,000', [],
                                      payer_stage='잔금')
    view2 = json.dumps(client.views_update.call_args.kwargs['view'], ensure_ascii=False)
    assert '현금으로 내기로' not in view2 and '부담액 (부가세 포함)' in view2
