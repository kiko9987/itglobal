# -*- coding: utf-8 -*-
"""ITG 사업자 통장 레지스트리 회귀 테스트 (2026-09-30 하도급지킴이 계좌 추가).

핵심 보장: 사업자는 은행명이 아니라 계좌로 판정 — 기업은행이라도 하도급지킴이 계좌
(452-047865-04-039)는 글로벌그룹 R. 기존 계좌 판정(G/R/N)은 그대로.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.itg_accounts import match_account

# 가정 양식(인터넷뱅킹 개통 전, 사용자 제공) — 기존 기업은행 양식과 동일
SUBCONTRACT_SMS = (
    "[Web발신]\n2026/09/29 16:18\n입금 6,503,000원\n"
    "예원초하도급지킴이\n452***86504039\n기업"
)
GLOBAL_IBK_SMS = (
    "[Web발신]\n2026/09/29 13:46\n입금 7,480,000원\n"
    "(주)아이티플레이\n452***38801011\n기업"
)


class TestMatchAccount:
    @pytest.mark.parametrize('text,bank,code', [
        ('452***38801011', '기업', 'G'),
        ('452-039388-01-011', '기업', 'G'),
        ('452***86504039', '기업', 'R'),          # 하도급지킴이 — 기업인데 R
        ('452-047865-04-039', '기업', 'R'),
        ('452****6504039', '기업', 'R'),          # 마스킹 폭이 달라도 끝자리로 판정
        ('255******31304', '하나', 'R'),
        ('255-910014-31304', '하나', 'R'),
        ('352-****-1682-33', '농협', 'N'),
    ])
    def test_registered(self, text, bank, code):
        a = match_account(text)
        assert a is not None and a.bank == bank and a.code == code

    @pytest.mark.parametrize('text', [
        '452***99999999',          # 같은 은행 다른(개인) 계좌
        '110-1234-5678',
        '입금 452,000원',           # 금액의 452 오탐 금지
        '452-***-****-039',        # 보이는 끝자리 3자리 — 모호, 판정 안 함
        '',
        None,
    ])
    def test_not_registered(self, text):
        assert match_account(text) is None


class TestIntakeFilterAndHeader:
    def test_subcontract_passes_business_filter(self):
        from dashboard.services.sms_intake import has_business_account
        assert has_business_account(SUBCONTRACT_SMS)
        assert has_business_account(GLOBAL_IBK_SMS)
        assert not has_business_account('입금 100,000원\n홍길동\n452***99999999\n기업')

    def test_intake_header_label(self):
        from dashboard.blueprints.sms_inbound import _build_intake_blocks
        from dashboard.services.sms_intake import parse_preview
        sub = _build_intake_blocks('x', SUBCONTRACT_SMS, parse_preview(SUBCONTRACT_SMS))
        assert '기업은행 (글로벌그룹)' in sub[0]['text']['text']
        glo = _build_intake_blocks('x', GLOBAL_IBK_SMS, parse_preview(GLOBAL_IBK_SMS))
        assert '기업은행 (글로벌)' in glo[0]['text']['text']
        assert '(글로벌그룹)' not in glo[0]['text']['text']


class TestPaymentCode:
    def test_memo_parse_subcontract(self):
        from dashboard.services.payment_sync import _parse_memo_block
        blk = _parse_memo_block(SUBCONTRACT_SMS)
        assert blk['bank'] == '기업'
        assert blk['acct_code'] == 'R'
        assert blk['amount'] == 6_503_000
        assert blk['partner'] == '예원초하도급지킴이'

    def test_card_code_R_for_subcontract_G_for_global(self):
        from dashboard.services.payment_sync import _parse_memo_block, _resolve_payment_code
        sub = _parse_memo_block(SUBCONTRACT_SMS)
        glo = _parse_memo_block(GLOBAL_IBK_SMS)
        assert _resolve_payment_code('', sub['bank'], sub['partner'], sub['acct_code']) == 'R'
        assert _resolve_payment_code('', glo['bank'], glo['partner'], glo['acct_code']) == 'G'

    def test_legacy_callers_without_acct_code_unchanged(self):
        from dashboard.services.payment_sync import _resolve_payment_code
        assert _resolve_payment_code('', '기업') == 'G'
        assert _resolve_payment_code('', '하나') == 'R'
        assert _resolve_payment_code('', '농협') == 'N'
        assert _resolve_payment_code('', '기업', '현금 수령 (YG)', 'R') == 'N'   # 현금 우선

    def test_preview_carries_acct_code(self):
        from dashboard.services.sms_intake import parse_preview
        assert parse_preview(SUBCONTRACT_SMS)['acct_code'] == 'R'


class TestPinRemind:
    def test_deposit_grn(self):
        from dashboard.services.pin_remind import _deposit_grn, _fmt_deposit_line
        assert _deposit_grn(SUBCONTRACT_SMS) == 'R'
        assert _deposit_grn(GLOBAL_IBK_SMS) == 'G'
        line = _fmt_deposit_line({'date_md': '09/29', 'bank': '기업', 'acct_code': 'R',
                                  'amount': 6_503_000, 'partner': '예원초하도급지킴이'})
        assert line == '09/29 R 6,503,000원 예원초하도급지킴이'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
