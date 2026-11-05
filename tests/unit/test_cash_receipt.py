"""세금계산서 요청 창의 현금영수증 (2026-11-05) — 창·발급 번호·시트 메모 줄."""
from dashboard.blueprints import slack_bot as sb


def test_cash_receipt_kind():
    assert sb._cash_receipt_kind('010-1234-5678') == '소득공제용'
    assert sb._cash_receipt_kind('01012345678') == '소득공제용'
    assert sb._cash_receipt_kind('124-81-00998') == '지출증빙용'      # 검증 숫자가 맞는 사업자번호
    assert sb._cash_receipt_kind('123-45-67890') == ''                # 검증 숫자 틀림(오타) — 11-05 점검
    assert sb._cash_receipt_kind('02-1234-5678') == ''                # 서울 일반전화 10자리
    assert sb._cash_receipt_kind('') == '' and sb._cash_receipt_kind('1234') == ''


def _ids(view):
    return [b.get('block_id') for b in view['blocks'] if b.get('block_id')]


def test_modal_doc_switch_blocks():
    tax = sb._build_invoice_modal_view('G4049-JW', '강원포장', '광명', '1,000,000', 'a@b.c', '{}')
    assert tax['title']['text'] == '세금계산서 발행 요청'
    assert _ids(tax)[0] == 'doc' and 'email' in _ids(tax) and 'vat' in _ids(tax)
    cash = sb._build_invoice_modal_view('G4049-JW', '홍길동', '광명', '1,100,000', '', '{}', doc='cash')
    assert cash['title']['text'] == '현금영수증 발행 요청'
    ids = _ids(cash)
    # 이메일·VAT·수정발행 칸 없음, 금액·발급 번호는 새 칸(접미사)이라 종류를 바꾸면 값이 새로 채워진다
    assert 'amt__x' in ids and 'crno__x' in ids and 'stages' in ids
    assert not any(i.split('__')[0] in ('email', 'vat', 'replace') for i in ids)
    vals = sb._invoice_view_values({'state': {'values': {'crno__x': {'value': {'value': '010'}}}}})
    assert vals['crno'] == {'value': '010'}
    # 나눠 내는 공사 창엔 종류 고르기가 없다(세금계산서만)
    split = sb._build_invoice_modal_view('G1', 'a', 'b', '1', '', '{}', payer_opts=[], doc='cash')
    assert 'doc' not in _ids(split)


def test_cash_mode_stage_label_and_issued_warn():
    """현금영수증을 고르면 발행 단계 이름·설명과 '이미 발행' 안내도 현금영수증·합계 기준으로(11-05)."""
    cash = sb._build_invoice_modal_view('G1', '홍길동', '광명', '1,100,000', '', '{}', doc='cash')
    stage = next(b for b in cash['blocks'] if b.get('block_id') == 'stages')
    assert stage['label']['text'] == '현금영수증 발행 단계' and '이 현금영수증이' in stage['hint']['text']
    tax = sb._build_invoice_modal_view('G1', 'a', 'b', '1', '', '{}')
    assert next(b for b in tax['blocks'] if b.get('block_id') == 'stages')['label']['text'] == '계산서 발행 단계'
    s = {'total2': 3300000, 'issued_gross': 1100000, 'issued_supply': 1000000, 'remaining_supply': 2000000,
         'uncertain': False, 'active': [{}]}
    assert '남은 공급가액 2,000,000원' in sb._fmt_issued_warn(s)
    c = sb._fmt_issued_warn(s, doc='cash')
    assert '남은 금액 2,200,000원(VAT 포함)이 현금영수증 발행 금액에 입력' in c and '공급가' not in c
    done = dict(s, issued_gross=3300000, issued_supply=3000000, remaining_supply=0)
    assert '수정발행' in sb._fmt_issued_warn(done) and '수정발행' not in sb._fmt_issued_warn(done, doc='cash')


def test_cash_memo_line_parses_like_invoice():
    """시트 메모 줄은 계산서와 같은 '발행' 줄 — 괄호 안 표시만 달라 발행액 합산·사업자 이름이 그대로 읽힌다."""
    memo = ('2026-11-05 잔금 발행 1,100,000원 (VAT 포함 · 현금영수증) · 홍길동\n'
            '2026-11-05 계약금 발행 2,200,000원 (2,000,000원 VAT 별도 · 현금영수증) · 강원포장')
    items = sb._active_invoice_lines(memo)
    assert [(i['stage'], i['amt'], i['biz']) for i in items] == [
        ('잔금', 1100000, '홍길동'), ('계약금', 2200000, '강원포장')]


def test_dedup_key_separates_cash(monkeypatch):
    monkeypatch.setattr(sb, '_intake_split_payers', lambda code: None)
    assert sb._invoice_dedup_key('G1', '잔금') == 'invoice_request_dedup:G1'
    assert sb._invoice_dedup_key('G1', '잔금', 'cash') == 'invoice_request_dedup:G1:cash'
