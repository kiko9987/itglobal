# -*- coding: utf-8 -*-
"""수금 카드·리마인드 4종 — 시트·문자·고객 입력 값이 슬랙 mrkdwn 호출/위장 링크가 되지 않게 (2026-11-14 보안 점검).

'<!channel>' 이 그대로 들어가면 채널 전체 호출, '<https://악성|견적서 확인>' 은 위장 링크가 된다.
값만 & < > 를 바꾸고, 코드가 일부러 만든 링크(<permalink|바로가기>)는 그대로 둔다.
"""
from datetime import date

import pytest

from dashboard.services import absent_remind as ar
from dashboard.services import invoice_collection_remind as icr
from dashboard.services import payment_sync as ps
from dashboard.services import pin_remind as pr

EVIL = '<!channel> <https://x|y> &'
EVIL_ESC = '&lt;!channel&gt; &lt;https://x|y&gt; &amp;'


def _assert_safe(txt: str):
    assert '&lt;!channel&gt;' in txt
    assert '<!channel>' not in txt
    assert '<https://x|y>' not in txt


# ── payment_sync (수금 카드) ──────────────────────────────────

@pytest.fixture
def _ps_plain(monkeypatch):
    monkeypatch.setattr(ps, '_split_payers_for', lambda project: {})
    monkeypatch.setattr(ps, '_inflow_line', lambda project: [])
    monkeypatch.setattr('dashboard.services.project_service.get_project_records', lambda *a, **k: [])


def _pay(**kw):
    p = {'stage': '잔금', 'date_md': '11/14', 'date_year': '2026', 'amount': 1000000,
         'partner': EVIL, 'bank': '기업', 'installment': ''}
    p.update(kw)
    return p


def test_stage_message_escapes_address_construction_partner(_ps_plain):
    txt = ps._build_stage_message(
        stage='계약금', project='G1234-AB', address=EVIL, payment=_pay(stage='계약금'),
        invoice_value='', total_r=0, total_t=2000000, unpaid=1000000, construction=EVIL)
    _assert_safe(txt)
    assert f'주소 : {EVIL_ESC}' in txt
    assert f'공사내용 : {EVIL_ESC}' in txt
    assert f'입금자 : {EVIL_ESC}' in txt


def test_stage_history_message_escapes_history_line(_ps_plain):
    pays = [_pay(stage='계약금'), _pay()]
    txt = ps._build_stage_with_history_message(
        stage='잔금', project='G1234-AB', address=EVIL, last_payment=pays[1], all_payments=pays,
        invoice_value='', total_r=0, total_t=2000000, unpaid=0,
        stage_sheet_vals={'계약금': 1000000, '중도금': 0, '잔금': 1000000}, construction=EVIL)
    _assert_safe(txt)
    assert f'1,000,000원  {EVIL_ESC}' in txt          # [누적 이력] 줄의 입금자


def test_complete_and_unified_messages_escape(_ps_plain):
    pays = [_pay()]
    txt = ps._build_complete_message('G1234-AB', EVIL, pays, '', 1000000,
                                     stage_sheet_vals={'계약금': 0, '중도금': 0, '잔금': 1000000},
                                     construction=EVIL)
    _assert_safe(txt)
    uni = ps._build_unified_stage_message(
        stage='잔금', projects=[{'code': 'G1234-AB', 'address': EVIL, 'sheet_val': 500000},
                               {'code': 'G1235-AB', 'address': '정상 주소', 'sheet_val': 500000}],
        payments=pays, invoice_value='', is_complete=False)
    _assert_safe(uni)
    assert f'입금자 : {EVIL_ESC}' in uni


def test_split_payer_title_and_status_escape(monkeypatch):
    payers = {'잔금': {'role': '임차인', 'name': EVIL, 'amount': 1000000, 'received': 0},
              '계약금': {'role': '<!here>', 'name': '이현주', 'amount': 1000000, 'received': 0}}
    monkeypatch.setattr(ps, '_split_payers_for', lambda project: payers)
    monkeypatch.setattr(ps, '_inflow_line', lambda project: [])
    monkeypatch.setattr('dashboard.services.project_service.get_project_records', lambda *a, **k: [])
    pays = [_pay(partner='정상입금자', installment='<!channel>')]
    txt = ps._build_stage_with_history_message(
        stage='잔금', project='G1234-AB', address='주소', last_payment=pays[0], all_payments=pays,
        invoice_value='', total_r=0, total_t=2000000, unpaid=1500000,
        stage_sheet_vals={'계약금': 0, '중도금': 0, '잔금': 500000})
    _assert_safe(txt)
    assert '<!here>' not in txt and '&lt;!here&gt; 1,000,000원' in txt   # 미수금 줄 '누구 몫'


def test_inflow_line_escapes_partner_name(monkeypatch):
    monkeypatch.setattr(ps, '_project_cache_row', lambda project: {'유입 구분': '거래처', '거래처명': EVIL})
    monkeypatch.setattr('dashboard.services.partner_names.is_partner_name', lambda s: True)
    out = ps._inflow_line('G1234-AB')
    assert out and '<!channel>' not in out[0] and '&lt;!channel&gt;' in out[0]


def test_misplacement_warning_and_special_notes_escape(monkeypatch):
    recs = [{'프로젝트 코드': 'G1234-AB', '사업자명': '내사업자'},
            {'프로젝트 코드': 'G9999-ZZ', '사업자명': '악성상사 <!channel>'}]
    monkeypatch.setattr('dashboard.services.project_service.get_project_records', lambda *a, **k: recs)
    w = ps._partner_misplacement_warning('G1234-AB', '악성상사 <!channel>')
    assert w and '<!channel>' not in w and '&lt;!channel&gt;' in w
    txt = ps._append_special_notes('카드\n' + ps._SEP + '\n끝', ['', '', '중고매입 상계 <!channel> <https://x|y>'])
    _assert_safe(txt)


def test_overdue_and_search_escape(monkeypatch):
    monkeypatch.setattr(ps, 'find_overdue_unpaid', lambda days=30: [
        {'project': 'G1234-AB', 'address': EVIL, 'unpaid': 1000, 'last_date': '2026-01-01'}])
    _assert_safe(ps.build_overdue_message())


def test_plain_text_unchanged(_ps_plain):
    """보통 글(한글·괄호·점)은 그대로 — 이스케이프는 & < > 만."""
    txt = ps._build_stage_message(
        stage='계약금', project='G1234-AB', address='경기 화성시 동탄대로 1 (101호)',
        payment=_pay(stage='계약금', partner='(주)아이티.김철수'),
        invoice_value='', total_r=0, total_t=2000000, unpaid=1000000, construction='시스템에어컨 2대')
    assert '주소 : 경기 화성시 동탄대로 1 (101호)' in txt
    assert '입금자 : (주)아이티.김철수' in txt and '공사내용 : 시스템에어컨 2대' in txt


# ── invoice_collection_remind (계산서/수금 리마인드) ─────────────

def _icr_row(**kw):
    r = {'프로젝트 코드': 'G1234-AB', '사업자명': EVIL, '현장 주소': EVIL, '총액 2': 1000000,
         '계약금': 0, '중도금': 0, '잔금': 1000000, '미수금': 0, '수금 확인': 'TRUE',
         '잔금 계산서': '미발행', '공사 종료': '2026-01-01'}
    r.update(kw)
    return r


def test_invoice_monthly_and_collection_text_escape(monkeypatch):
    monkeypatch.setattr(icr, '_split_states', lambda row: {})
    monkeypatch.setattr(icr, '_card_links', lambda codes: {'G1234-AB': 'https://slack.example/p1'})
    bk = icr.classify([_icr_row()])
    assert bk['issue_collected']
    txt = icr.build_monthly_text(bk)
    _assert_safe(txt)
    assert '<https://slack.example/p1|바로가기>' in txt        # 코드가 만든 링크는 그대로
    assert '사업자 정보 없음' in icr._who(_icr_row(사업자명='-'))
    col = icr.build_collection_text({'ar_issued': [_icr_row(미수금=500000)], 'ar_uninvoiced': []})
    _assert_safe(col)


def test_invoice_unbilled_split_payer_names_escape(monkeypatch):
    ss = {'잔금': {'state': 'needed', 'issued': 0,
                  'payer': {'role': '임차인', 'name': EVIL, 'received': 1000000}}}
    monkeypatch.setattr(icr, '_split_states', lambda row: ss)
    txt = icr._unbilled(_icr_row())
    _assert_safe(txt)


# ── pin_remind (미처리 정산 핀 리마인드) ─────────────────────────

def test_pin_remind_text_escapes_summaries_and_moves(monkeypatch):
    import dashboard.services.fund_transfer as ft
    monkeypatch.setattr(ft, 'ENTITY', {'R': '글로벌그룹', 'G': '글로벌'}, raising=False)
    data = {
        # 슬랙에서 읽은 핀 원문 — 이미 &lt; 로 온 것과 날 것('<!channel>') 둘 다
        'invoices': [{'summary': '11/14 G 1,000원 ' + EVIL, 'permalink': 'https://slack.example/p2'}],
        'others': [{'summary': 'A &amp; B &lt;b&gt;', 'permalink': ''}],
        'intakes': [{'summary': '11/14 G 1,000원 ' + EVIL, 'permalink': '', 'state': '미지정',
                     'guess': 'G4157-JW <!channel>'}],
        'moves': [{'code': 'G1234-AB', 'stage': '잔금', 'amount': 1000, 'partner': EVIL, 'bank': '하나',
                   'from': 'R', 'to': 'G', 'date': '2026/11/14'}],
    }
    txt = pr.build_pin_remind_text(data)
    _assert_safe(txt)
    assert '<https://slack.example/p2|바로가기>' in txt          # 코드가 만든 링크는 그대로
    assert '• A &amp; B &lt;b&gt;' in txt                        # 슬랙 원문 엔티티는 두 번 바뀌지 않음
    assert '&amp;amp;' not in txt and '&amp;lt;' not in txt


# ── absent_remind (09시 부재중·미완료·견적요청) ─────────────────

def test_absent_remind_escapes_name_platform_phone(monkeypatch):
    monkeypatch.setattr(ar, '_reminder_ordinal', lambda d, t: 1)
    lead = {'리드 No': 'L-1', '상담 시간': '2026.11.13. 09:18', '상태': '부재중',
            '고객명': '<!channel>', '플랫폼': '<https://x|y>', '고객 연락처': '010 &',
            '상담 내용': ''}
    txt, _ = ar.build_remind_text([lead], {'x': [lead]}, quote_pending={'y': [lead]},
                                  callback={'A': [dict(lead, **{'상담 내용': '내일 다시 연락 <!channel>'})]})
    _assert_safe(txt)
    assert '[&lt;https://x|y&gt;] &lt;!channel&gt;' in txt
    assert '010 &amp;' in txt
