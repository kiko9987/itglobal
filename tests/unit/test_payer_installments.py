# -*- coding: utf-8 -*-
"""나눠 내는 공사 — 차수·사람별 계산서 상태·입금 칸 결정 (2026-10-20 오남 입금 칸 어긋남 계기)."""
from dashboard.services import payer_split as ps

LESSOR = '임대인 · 이현주 · 8,031,100 · 계산서'
TENANT = '임차인 · 민선홍(서울아산가정의학) · 7,701,100 · 계산서'


def _rec(**kw):
    r = {'계약금 부담자': LESSOR, '중도금 부담자': '-', '잔금 부담자': TENANT,
         '계약금': 0, '중도금': 0, '잔금': 0, '총액 2': 15732200,
         '계약금 계산서': '-', '중도금 계산서': '-', '잔금 계산서': '-'}
    r.update(kw)
    return r


# ── 차수 줄 ──
def test_installment_line_roundtrip():
    assert ps.installment_line('중도금') == '◇ 차수 중도금'
    assert ps.installment_line('일시불') == '◇ 차수 일시불'
    assert ps.installment_line('기타') == ''
    block = '2026/10/17 19:55\n입금 1,177,880원\n다정약국.김지민\n◇ 차수 중도금\n기업'
    assert ps.parse_installment(block) == '중도금'
    assert ps.parse_installment('입금 1원') == ''


# ── 입금 칸 결정 ──
def test_resolve_new_designation_payer_column():
    payers = ps.stage_payers(_rec())
    r = ps.resolve_deposit_column(payers, '잔금', '다정약국.김지민', 1177880, installment='중도금')
    assert r == {'stage': '잔금', 'installment': '중도금', 'moved': False, 'error': ''}


def test_resolve_new_designation_rejects_non_payer_column():
    payers = ps.stage_payers(_rec())
    r = ps.resolve_deposit_column(payers, '중도금', '아무개', 1000000, installment='중도금')
    assert r['stage'] == '' and '부담자가 없습니다' in r['error']


def test_resolve_legacy_designation_moves_to_named_payer():
    # G4158: 임차인 민선홍 돈을 들어온 순서대로 '계약금'으로 지정(10/17) → 잔금 칸, 지정 단계는 차수로
    payers = ps.stage_payers(_rec())
    r = ps.resolve_deposit_column(payers, '계약금', '민선홍(서울아산가', 6160880)
    assert r == {'stage': '잔금', 'installment': '계약금', 'moved': True, 'error': ''}
    # G4158: 임대인 이현주 돈을 '중도금'(부담자 없는 칸)으로 지정 → 계약금 칸
    r = ps.resolve_deposit_column(payers, '중도금', '이현주가정의에어', 6424880)
    assert r['stage'] == '계약금' and r['moved'] and r['installment'] == '중도금'


def test_resolve_legacy_proxy_payer_keeps_designated_payer_column():
    # 대리 입금(입금자명이 부담자와 다름) — 지정 칸이 사람 칸이면 그대로
    payers = ps.stage_payers(_rec())
    r = ps.resolve_deposit_column(payers, '잔금', '남상원', 3000000)
    assert r == {'stage': '잔금', 'installment': '', 'moved': False, 'error': ''}


def test_resolve_legacy_unknown_payer_non_payer_column_is_error():
    payers = ps.stage_payers(_rec())
    r = ps.resolve_deposit_column(payers, '중도금', '남상원', 3000000)
    assert r['stage'] == '' and '[재지정]' in r['error']


# ── 사람별 계산서 상태 ──
def _p(amount, received, method='계산서'):
    return {'amount': amount, 'received': received, 'method': method}


def test_bill_state_basic():
    assert ps.payer_bill_state(_p(100, 0), '-')['state'] == 'none'
    assert ps.payer_bill_state(_p(100, 80), '미발행')['state'] == 'pending'
    assert ps.payer_bill_state(_p(100, 100), '미발행')['state'] == 'needed'
    assert ps.payer_bill_state(_p(100, 100), '-')['state'] == 'needed'          # '-' 라도 covered 아님
    assert ps.payer_bill_state(_p(100, 100), '카드')['state'] == 'card'
    assert ps.payer_bill_state(_p(100, 100), 'N입금')['state'] == 'cash'
    assert ps.payer_bill_state(_p(100, 100, '현금'), '미발행')['state'] == 'cash'


def test_bill_state_issued_full_and_preissued():
    st = ps.payer_bill_state(_p(1000, 800), '발행', 1000, True)
    assert st['state'] == 'issued' and st['preissued'] and st['pending'] == 200
    st = ps.payer_bill_state(_p(1000, 1000), '발행', 1000, True)
    assert st['state'] == 'issued' and not st['preissued']
    # 줄 없는 옛·수기 '발행' = 전액 발행으로
    assert ps.payer_bill_state(_p(1000, 1000), '발행')['state'] == 'issued'


def test_bill_state_partial_installment_invoice():
    st = ps.payer_bill_state(_p(1000, 800), '발행', 800, True)       # 1차분만 발행, 아직 덜 받음
    assert st['state'] == 'partial' and st['to_issue'] == 200 and not st['need']
    st = ps.payer_bill_state(_p(1000, 1000), '발행', 800, True)      # 다 받았는데 200 미발행
    assert st['state'] == 'partial' and st['need']


def test_bill_states_from_invoice_lines():
    rec = _rec(**{'계약금': 8031100, '잔금': 7701100, '계약금 계산서': '-', '잔금 계산서': '발행'})
    lines = [{'stage': '잔금', 'amt': 7701100}]
    st = ps.payer_bill_states(rec, lines)
    assert st['잔금']['state'] == 'issued'
    assert st['계약금']['state'] == 'needed'      # 예전 규칙이면 '-' covered → 숨던 미발행


# ── 요약(Y) ──
def test_split_y_summary_shows_missing_payer_invoice():
    rec = _rec(**{'계약금': 8031100, '잔금': 7701100, '잔금 계산서': '발행'})
    st = ps.payer_bill_states(rec, [{'stage': '잔금', 'amt': 7701100}])
    assert ps.split_y_summary(st) == '계약금 - 미발행'      # 예전: '잔금 - 발행완료'


def test_split_y_summary_progress_and_done():
    rec = _rec(**{'계약금': 6424880, '잔금': 6160880, '계약금 계산서': '미발행', '잔금 계산서': '미발행'})
    assert ps.split_y_summary(ps.payer_bill_states(rec, [])) == '계약금 - 미발행'
    rec = _rec(**{'계약금': 8031100, '잔금': 7701100, '계약금 계산서': '발행', '잔금 계산서': '발행'})
    lines = [{'stage': '계약금', 'amt': 8031100}, {'stage': '잔금', 'amt': 7701100}]
    assert ps.split_y_summary(ps.payer_bill_states(rec, lines)) == '잔금 - 발행완료'
    assert ps.split_y_summary(ps.payer_bill_states(_rec(), [])) == '미발행'
    # 일부 발행이지만 받은 돈은 다 발행됨 → 미발행 아님
    rec = _rec(**{'계약금': 6424880, '계약금 계산서': '발행'})
    st = ps.payer_bill_states(rec, [{'stage': '계약금', 'amt': 6424880}])
    assert st['계약금']['state'] == 'partial' and ps.split_y_summary(st) == '계약금 - 발행완료'


def test_split_y_summary_cash_payer():
    rec = _rec(**{'계약금 부담자': '임대인 · 임기환 · 10,000,000 · 현금', '계약금': 8000000,
                  '계약금 계산서': 'N입금', '총액 2': 17701100})
    assert ps.split_y_summary(ps.payer_bill_states(rec, [])) == '계약금 - N입금'


# ── 정합성 ──
def test_split_problems_overpay_and_missing_payer_line():
    rec = _rec(**{'계약금': 9000000, '계약금_메모': '2026/10/17\n입금 9,000,000원', '잔금_메모': None})
    probs = ps.split_problems(rec)
    assert any('더 받음' in p for p in probs)
    assert any('계약금 칸 메모 맨 윗줄' in p for p in probs)
    assert any('잔금 칸 메모 맨 윗줄' in p for p in probs)
    ok = _rec(**{'계약금_메모': '◆ 임대인 이현주\n\n2026/10/17', '잔금_메모': '◆ 임차인 민선홍(서울아산가정의학)'})
    assert not any('메모 맨 윗줄' in p for p in ps.split_problems(ok))


def test_plan_payers_role_error_is_generic():
    _, errs, _ = ps.plan_payers([{'name': '○○의원', 'amount': '1000000'}], {})
    assert errs and '예: 임대인·병원·본점' in errs[0]
