"""나눠 내는 공사 — '○○ 부담자' 칸 해석·입금 칸 추정·정합성 (2026-10-19 오남프라자)."""
from dashboard.services import payer_split as ps


def _rec(**kw):
    base = {'프로젝트 코드': 'G4146-JW', '총액 2': '14258200',
            '계약금': '5509680', '중도금': '0', '잔금': '0',
            '계약금 부담자': '임대인 · 이희승 · 122-15-41208 · 6,887,100 · 계산서',
            '중도금 부담자': '-',
            '잔금 부담자': '임차인 · 주식회사 제이케이디학원 · 405-81-13558 · 7,371,100 · 계산서'}
    base.update(kw)
    return base


def test_parse_full_cell():
    p = ps.parse_payer_cell('임대인 · 이희승 · 122-15-41208 · 6,887,100 · 계산서')
    assert p['role'] == '임대인' and p['name'] == '이희승'
    assert p['biz_no'] == '122-15-41208' and p['amount'] == 6_887_100 and p['method'] == '계산서'


def test_parse_any_order_email_and_cash():
    p = ps.parse_payer_cell('임차인 · 3,252,000원 · 이동엽 · 현금')
    assert (p['name'], p['amount'], p['method']) == ('이동엽', 3_252_000, '현금')
    p2 = ps.parse_payer_cell('임차인 | 기종원(위너스골프) | bell8517@nate.com | 18471200')
    assert p2['email'] == 'bell8517@nate.com' and p2['method'] == '계산서' and p2['amount'] == 18_471_200
    assert ps.parse_payer_cell('임차인 · 권영진 · 6,976,200 · 계산서 x')['method'] == '현금'


def test_old_depositor_values_are_not_payers():
    for old in ('25/02/11 고희원 R', '(주)넥스로 입금및 발행', '1/4 190만원별도 세금계산서 발행',
                '-', '', None, '24/06/11 홍재영 G', '임대인 · 이희승', '6,887,100 · 이희승 · 계산서'):
        assert ps.parse_payer_cell(old) is None, old


def test_stage_payers_received_and_remaining():
    pays = ps.stage_payers(_rec())
    assert set(pays) == {'계약금', '잔금'}
    assert pays['계약금']['received'] == 5_509_680 and pays['계약금']['remaining'] == 1_377_420
    assert pays['잔금']['remaining'] == 7_371_100
    assert ps.stage_label('잔금', pays['잔금']) == '잔금(임차인)'
    assert ps.stage_label('중도금', None) == '중도금'
    assert ps.payer_line(pays['계약금']) == '임대인 이희승 · 부담 6,887,100 · 남음 1,377,420'
    assert not ps.is_split({'계약금 부담자': '25/02/11 고희원 R'})


def test_match_deposit_by_name_and_amount():
    pays = ps.stage_payers(_rec())
    assert ps.match_deposit_stage(pays, '이희승', 1_377_420) == '계약금'
    assert ps.match_deposit_stage(pays, '(주)제이케이디학원', 5_896_880) == '잔금'
    # 이름 안 맞아도 남은 돈과 정확히 같으면
    assert ps.match_deposit_stage(pays, '김현준', 7_371_100) == '잔금'
    # 이름 안 맞고 금액도 애매하면 추정 안 함
    assert ps.match_deposit_stage(pays, '홍길동', 1_000_000) == ''


def test_guess_split_payer_same_landlord_many_units():
    """같은 임대인(이현주)이 여러 호실 — 금액이 남은 돈 안에 드는 곳이 하나면 그 호실."""
    def unit(code, total, landlord_amt, got=0):
        return {'프로젝트 코드': code, '총액 2': str(total), '계약금': str(got),
                '계약금 부담자': f'임대인 · 이현주 · {landlord_amt:,} · 계산서',
                '잔금 부담자': f'임차인 · 다른사람 · {total - landlord_amt:,} · 계산서'}
    recs = [unit('G4156-JW', 3_879_700, 2_104_850), unit('G4157-JW', 20_363_200, 10_346_600),
            unit('G4158-JW', 15_732_200, 8_031_100), unit('G4160-JW', 11_843_700, 6_251_850)]
    assert ps.guess_split_payer('이현주이비인임대', 8_145_280, recs) == 'G4157-JW 임대인'
    # 1차 뒤 2차: 남은 돈과 같은 곳으로 좁힘
    recs[1]['계약금'] = '8145280'
    assert ps.guess_split_payer('이현주이비인에어', 2_201_320, recs) == 'G4157-JW 임대인'
    assert ps.guess_split_payer('이현주', 1_000_000, recs) == ''


def test_payer_for_biz():
    pays = ps.stage_payers(_rec())
    assert ps.payer_for_biz(pays, '주식회사 제이케이디 학원')['stage'] == '잔금'
    assert ps.payer_for_biz(pays, '', '1221541208')['stage'] == '계약금'
    assert ps.payer_for_biz(pays, '다른회사') is None


def test_split_problems():
    assert ps.split_problems(_rec()) == []
    bad = ps.split_problems(_rec(**{'총액 2': '14858200', '중도금': '100000'}))
    assert any('부담액 합계' in b for b in bad) and any('중도금 칸에 입금' in b for b in bad)
    assert any('형식 확인' in b for b in ps.split_problems(_rec(**{'중도금 부담자': '임대인 · 이름만'})))
    # 평범한 공사(부담자 없음)는 문제 없음 — 옛 입금자명 값도 무시
    assert ps.split_problems({'계약금 부담자': '25/02/11 고희원 R', '계약금': '1000000'}) == []
