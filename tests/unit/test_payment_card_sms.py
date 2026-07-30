# -*- coding: utf-8 -*-
"""카드결제 SMS 포맷 파싱 회귀 테스트 (2026-07-29).

배경: 카드 파싱(_is_card_payment/_CARD_PARTNER_RE)은 승인번호가 **자기 줄 단독**일 때만
동작했음. 은행 SMS 가 '금액원 승인번호'(같은 줄)로 오면 접두 없는 금액이 안 떨어져
입금자 추출 실패 → _is_card_payment False → 미완성 → 수금완료 카드 미발송
(R3795 '현703011838', R3829 '삼성 204108778', Y='카드결제').

수정: ①접두 없는 'X원' 제거 ②'잔액' 줄 skip ③_CARD_BRAND_RE 브랜드+승인번호 인식
④계산식 제거 후 빈 괄호 정리. 실데이터 2000-3920행 회귀 = 의도 3건 + 무해 3건.
pure 함수만.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.payment_sync import (
    _parse_notes, _is_card_payment, _resolve_payment_code)


def _w(memo, val=1000000):
    res = _parse_notes(['', '', memo], stage_vals={'잔금': val})
    return next((p for p in (res or []) if p.get('stage') == '잔금'), {})


class TestCardSms:
    def test_contiguous_approval_token(self):
        """'현703011838' (카드사+승인번호 연속) → 입금자 추출 + 카드 인식."""
        p = _w('2026-07-20\n2,889,150원 현703011838\n하나', 2889150)
        assert p['partner'] == '현703011838'
        assert _is_card_payment('카드결제', p['partner']) is True

    def test_brand_space_approval(self):
        """'삼성 204108778' (브랜드+공백+승인번호) → 카드 인식."""
        p = _w('2026-07-20\n3,059,100원 삼성 204108778\n하나', 3059100)
        assert p['partner'] == '삼성 204108778'
        assert _is_card_payment('카드결제', p['partner']) is True

    def test_card_gate_requires_card_y(self):
        """Y가 카드결제/혼합 아니면 브랜드여도 카드 아님 (게이트 유지)."""
        assert _is_card_payment('잔금', '삼성 204108778') is False


class TestBankTransferRegression:
    def test_balance_line_skipped(self):
        """은행 SMS '잔액 X원' 줄은 입금자로 오추출 안 됨 → 진짜 거래처 추출."""
        memo = ('2026/07/03 13:10\n입금 330,000원\n잔액 192,703,519원\n'
                '(주)와이디와이\n452***38801011\n기업')
        p = _w(memo, 330000)
        assert p['partner'] == '(주)와이디와이'
        assert p['bank'] == '기업'
        assert _is_card_payment('잔금', p['partner']) is False  # 이체지 카드 아님

    def test_cash_calc_no_empty_parens(self):
        """'현금 X원 (계산식)' → 빈 괄호 없이 '현금', N 코드."""
        p = _w('2026-01-06\n현금 6,800,000원 (136*50000)', 6800000)
        assert p['partner'] in ('현금', '현금 수령')  # '현금 ()' 아님
        assert _resolve_payment_code('', p.get('bank', ''), p['partner']) == 'N'

    def test_normal_jeokyo_unchanged(self):
        """일반 '적요 거래처' 메모는 영향 없음 (회귀 방지)."""
        p = _w('일시 07/10, 14:00\n입금 1,000,000원\n계좌번호 255***31304\n적요 김철수', 1000000)
        assert p['partner'] == '김철수'


class TestPayerLabelJunk:
    """입금자 라벨에 날짜 junk('일시 02/20, 17:23')가 들어가면 거부 → 적요/이름 fallback (R3239)."""

    def test_junk_label_falls_through_to_jeokyo(self):
        """junk 라벨 무시하고 적요 사용 — 계약금·잔금 슬롯 일관."""
        memo = ('입금일: \n입금자: 일시 02/20, 17:23\n입금 300,000원\n'
                '계좌번호 255******31304\n적요 주식회사제우스')
        for stage, notes in (('계약금', [memo, '', '']), ('잔금', ['', '', memo])):
            res = _parse_notes(notes, stage_vals={'계약금': 300000, '중도금': 300000, '잔금': 300000})
            p = next((x for x in res if x.get('stage') == stage), {})
            assert p.get('partner') == '주식회사제우스', f'{stage}: {p.get("partner")!r}'

    def test_junk_label_no_jeokyo_stuck(self):
        """junk 라벨 + 적요 없음 → partner 빈값 (stuck → STUCK 체크로 잡힘, 조용히 틀리지 않음)."""
        memo = '입금일: \n입금자: 일시 02/20, 17:23\n입금 300,000원\n계좌번호 255******31304'
        res = _parse_notes([memo, '', ''], stage_vals={'계약금': 300000})
        assert (res[0].get('partner') or '') in ('', '-')

    def test_valid_label_kept(self):
        """정상 입금자 라벨(현금·현금 수령)은 그대로 유지 (회귀 방지)."""
        for lp in ('현금', '현금 수령'):
            res = _parse_notes([f'입금자: {lp}\n입금 500,000원', '', ''],
                               stage_vals={'계약금': 500000})
            assert res[0].get('partner') == lp


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
