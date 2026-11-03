# -*- coding: utf-8 -*-
"""경영지원 건의 (2026-11-03) — 수금 카드 유입 구분 줄, 공사 재개 카드 금액 줄·매출 0원 요청 철회."""
import json

from dashboard.services import payment_sync as ps


def _row(inflow, partner=''):
    return {'프로젝트 코드': 'G1-XX', '유입 구분': inflow, '거래처명': partner}


def test_inflow_line(monkeypatch):
    cases = [
        (_row('당근', '온라인'), ['유입 구분 : 온라인 (당근)']),
        (_row('거래처', '메이드바이'), ['유입 구분 : 거래처 (메이드바이)']),
        (_row('거래처', '온라인'), ['유입 구분 : 거래처']),           # 이름 모르면 '거래처'만
        (_row('소개'), ['유입 구분 : 거래처 (소개)']),
        (_row('-'), []),
        ({}, []),
    ]
    for row, want in cases:
        monkeypatch.setattr(ps, '_project_cache_row', lambda project, _r=row: _r)
        assert ps._inflow_line('G1-XX') == want


def test_stage_card_has_inflow_above_address(monkeypatch):
    monkeypatch.setattr(ps, '_project_cache_row', lambda project: _row('홈페이지'))
    pays = [{'stage': '잔금', 'date_md': '11/03', 'date_year': '2026', 'amount': 1000000,
             'partner': '홍길동', 'bank': '기업'}]
    txt = ps._build_stage_with_history_message(
        stage='잔금', project='G1-XX', address='강남구 학동로 1', last_payment=pays[0], all_payments=pays,
        invoice_value='', total_r=0, total_t=1000000, unpaid=0,
        stage_sheet_vals={'계약금': 0, '중도금': 0, '잔금': 1000000})
    # 구분선 바로 아래 첫 줄 (공사 확정 카드와 같은 순서, 11-03 SB 테스트 후 결정)
    assert f'{ps._SEP}\n유입 구분 : 온라인 (홈페이지)\n주소 : 강남구 학동로 1' in txt


def test_resume_amount_line():
    from dashboard.blueprints.slack_bot import _resume_amount_line
    p = {'총액 1': 3400000, '부가세': True}
    assert _resume_amount_line(p, 0) == '💵 금액 : 0원  →  3,400,000원 (VAT 별도) · 취소 전 금액으로 복원'
    assert _resume_amount_line(p, 3400000) == '💵 금액 : 3,400,000원 (VAT 별도) · 매출 0원 반영 전이라 그대로'
    assert '그대로' in _resume_amount_line(p, None)


def test_withdraw_cancel_money_requests(monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    store = {
        'project_cancel_money_req:C1:1.0': json.dumps({'code': 'G1-XX', 'card_text':
            '⠀\n💰 *[공사 취소 — 매출 0원 반영 요청]*  `G1-XX`'}),
        'project_cancel_money_req:C1:2.0': json.dumps({'code': 'G2-YY', 'card_text': 'x'}),
        'project_cancel_money_req:C1:1.0:proc': '1',
    }

    class _R:
        def scan_iter(self, match, count=0):
            return list(store)

        def get(self, k):
            return store.get(k)

        def delete(self, k):
            store.pop(k, None)

    updated = []

    class _P:
        def chat_update(self, channel, ts, text, blocks):
            updated.append((channel, ts, text))

    monkeypatch.setattr('dashboard.utils.redis_client.get_redis_client', lambda: type('C', (), {'redis': _R()})())
    monkeypatch.setattr(sb, '_project_client', lambda: _P())
    sb._withdraw_cancel_money_requests('G1-XX', 'TH')
    assert 'project_cancel_money_req:C1:1.0' not in store           # pending 삭제 → 핀 리마인드에서도 빠짐
    assert 'project_cancel_money_req:C1:2.0' in store               # 다른 공사는 그대로
    assert len(updated) == 1 and updated[0][:2] == ('C1', '1.0')
    assert '재개되어 요청 철회' in updated[0][2] and '재개 : TH' in updated[0][2]
