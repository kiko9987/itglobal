"""나눠 내는 공사 — '○○ 부담자' 칸 해석·입금 칸 추정·정합성 (2026-10-19 오남프라자)."""
from dashboard.services import payer_split as ps


def _rec(**kw):
    base = {'프로젝트 코드': 'G4146-JW', '총액 2': '14258200', '사업자명': '이희승 / 주식회사 제이케이디학원',
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


# ── 화면용 사업자명 (2026-10-28 G4155 — 카드에 임차인 상호 하나만 보였음) ──

def test_biz_display_shows_both_payers_in_stage_order():
    assert ps.biz_display(_rec(사업자명='이희승')) == '임대인 이희승 / 임차인 주식회사 제이케이디학원'
    assert ps.biz_display(_rec(사업자명='-')) == '임대인 이희승 / 임차인 주식회사 제이케이디학원'


def test_biz_display_plain_project_keeps_sheet_name():
    rec = _rec(사업자명='주식회사 칸도', **{'계약금 부담자': '', '잔금 부담자': '김연종'})   # 옛 입금자명 값은 부담자 아님
    assert ps.biz_display(rec) == '주식회사 칸도'
    assert ps.biz_display({'사업자명': None}) == ''


def test_project_card_shows_both_payers():
    from dashboard.services.project_slack_notifier import _build_message
    text = _build_message(_rec(사업자명='위너스 골프존', **{'유입 구분': '거래처', '거래처명': '오남프라자'}), 'G4146-JW')
    assert '🏢 사업자명 : 임대인 이희승 / 임차인 주식회사 제이케이디학원' in text


def test_as_details_biz_display(monkeypatch):
    import dashboard.services.project_service as psvc
    from dashboard.services.as_service import get_project_details
    monkeypatch.setattr(psvc, 'get_project_records', lambda force_refresh=False: [_rec(사업자명='이희승')])
    d = get_project_details('G4146-JW')
    assert d['biz'] == '이희승'                                            # 원래 값은 그대로(다른 곳 대조용)
    assert d['biz_display'] == '임대인 이희승 / 임차인 주식회사 제이케이디학원'


def test_deposit_from_own_payer_is_not_misplacement(monkeypatch):
    """임대인 이현주가 여러 호실에 걸친 경우 — 다른 호실 사업자명이 '이현주'여도 이 공사 부담자면 경고 안 함 (2026-10-28)."""
    import dashboard.services.project_service as psvc
    from dashboard.services.payment_sync import _partner_misplacement_warning
    lessor = '임대인 · 이현주 · 414-02-62330 · 1,802,350 · 계산서'
    recs = [
        {'프로젝트 코드': 'G4156-JW', '사업자명': '-', '거래처명': '오남프라자',
         '계약금 부담자': lessor, '잔금 부담자': '임차인 · 홍은희(루카수학) · 1,472,350 · 계산서'},
        {'프로젝트 코드': 'G4158-JW', '사업자명': '이현주', '거래처명': '오남프라자',
         '계약금 부담자': lessor, '잔금 부담자': '임차인 · 민선홍(서울아산가정의학) · 1,472,350 · 계산서'},
        {'프로젝트 코드': 'G9000-AA', '사업자명': '홍길동상사', '거래처명': '홍길동상사'},
    ]
    monkeypatch.setattr(psvc, 'get_project_records', lambda force_refresh=False: recs)
    assert _partner_misplacement_warning('G4156-JW', '이현주') == ''
    assert _partner_misplacement_warning('G4156-JW', '홍은희') == ''
    assert '다른 프로젝트' in _partner_misplacement_warning('G4156-JW', '홍길동상사')   # 진짜 오배치는 그대로



# ── 시트 사업자명 = 부담자 이름 모음 (2026-10-28) ──────────────────

class _BizMgr:
    def __init__(self, biz, ver='3'):
        self.cells = {'F10': biz, 'AT10': ver}
    def get_cell_value(self, sid, sn, a): return self.cells.get(a, '')
    def update_cell_value(self, sid, sn, a, v): self.cells[a] = v; return True


def test_split_biz_value_and_marker():
    payers = ps.stage_payers(_rec())
    assert ps.split_biz_value(payers) == '이희승 / 주식회사 제이케이디학원'
    assert ps.is_split_biz_value('이희승 / 주식회사 제이케이디학원')
    assert not ps.is_split_biz_value('주식회사 칸도') and not ps.is_split_biz_value('A/B상사')


def test_sync_biz_cell_writes_names_and_bumps_version():
    f2l = {'사업자명': 'F', '_version': 'AT'}
    m = _BizMgr('위너스 골프존')
    out = ps.sync_biz_cell(m, 's', 'n', 10, f2l, ps.stage_payers(_rec()))
    assert out == {'사업자명': '이희승 / 주식회사 제이케이디학원', '_version': '4'}
    assert ps.sync_biz_cell(m, 's', 'n', 10, f2l, ps.stage_payers(_rec())) == {}      # 같으면 안 씀
    assert ps.sync_biz_cell(m, 's', 'n', 10, f2l, {}) == {'사업자명': '-', '_version': '5'}   # 해제 → '-'
    m2 = _BizMgr('주식회사 칸도')
    assert ps.sync_biz_cell(m2, 's', 'n', 10, f2l, {}) == {}                          # 원래 상호는 그대로


def test_split_problems_flags_hand_edited_biz():
    assert ps.split_problems(_rec()) == []
    assert any('사업자명 칸' in p for p in ps.split_problems(_rec(사업자명='위너스 골프존')))


def test_ocr_does_not_touch_split_biz(monkeypatch):
    import dashboard.services.project_service as psvc
    from dashboard.services.business_license_handler import _maybe_update_business_name
    monkeypatch.setattr(psvc, 'get_project_records', lambda force_refresh=False: [_rec()])
    assert _maybe_update_business_name('G4146-JW', '주식회사 제이케이디학원') == ('split', '이희승 / 주식회사 제이케이디학원')


def test_partner_rule_ignores_split_biz(monkeypatch):
    from dashboard.services import partner_names as pn
    called = []
    monkeypatch.setattr(pn, 'promote_previous', lambda *a, **k: called.append(a) or 1)
    assert pn.apply_place_on_biz_change('G4146-JW', '이희승 / 주식회사 제이케이디학원') == 0
    assert called == []


def test_company_name_search_excludes_split_value(monkeypatch):
    import pandas as pd
    import dashboard.services.project_service as psvc
    from dashboard.blueprints.slack_bot import _search_company_names
    monkeypatch.setattr(psvc, 'load_data', lambda force_refresh=False: pd.DataFrame(
        {'사업자명': ['이희승 / 주식회사 제이케이디학원', '주식회사 제이케이디학원']}))
    assert [o['value'] for o in _search_company_names('제이케이디')] == ['주식회사 제이케이디학원']


def test_misplacement_uses_payer_names_of_other_split_projects(monkeypatch):
    import dashboard.services.project_service as psvc
    from dashboard.services.payment_sync import _partner_misplacement_warning
    recs = [_rec(), {'프로젝트 코드': 'G9000-AA', '사업자명': '홍길동상사', '거래처명': '홍길동상사'}]
    monkeypatch.setattr(psvc, 'get_project_records', lambda force_refresh=False: recs)
    w = _partner_misplacement_warning('G9000-AA', '이희승')
    assert 'G4146-JW(이희승)' in w                                   # 모음 글자 말고 그 사람 이름으로 표시
