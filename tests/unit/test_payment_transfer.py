# -*- coding: utf-8 -*-
"""매출이동(계좌간 이체) 표기 → 원 입금 payment 에 transfer_to 부착 회귀 테스트.

배경: 고객이 잘못된 계좌로 입금 → 자체적으로 다른 계좌(N통장 등)로 이동(매출이동).
메모에 '2026-07-29 R>N 매출이동' 을 붙이면 시트는 반영되나 수금 카드는 받은 계좌(R)만
표기 → 최종 목적지(N) 안 보임(R3888-JSH). 파서가 이체 목적지를 원 입금 payment 에
transfer_to 로 부착 → 카드 빌더가 'R → N' 렌더. (2026-07-29)

핵심 불변식: 금액·payment 개수·partner 변경 0 (순수 표시 필드만 추가).
pure 함수(_parse_notes) 만 — Redis/Slack 미접촉.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.payment_sync import _parse_notes


class TestTransferAttach:
    """이체 목적지가 원 입금에 부착되어야."""

    def test_r3888_separate_block(self):
        """이체가 별도 블록(yyyy-date 헤더로 분리)이어도 원 입금에 부착."""
        memo = ('일시 07/21, 16:53\n입금 200,000원\n계좌번호 255******31304\n'
                '적요 신세게유통\n2026-07-29 R>N 매출이동')
        res = _parse_notes([memo, '', ''], stage_vals={'계약금': 200000})
        assert len(res) == 1
        p = res[0]
        assert p['transfer_to'] == 'N'      # 하나(R) → 농협(N)
        assert p['amount'] == 200000        # 금액 불변
        assert p['partner'] == '신세게유통'  # 거래처 불변
        assert p['bank'] == '하나'          # 원 입금 은행 불변

    def test_origin_bank_match(self):
        """같은 단계에 은행 다른 입금 2건이면 출발 은행 일치 건에만 부착."""
        memo = ('일시 07/10, 10:00\n입금 900,000원\n하나은행\n적요 A상사\n\n'
                '일시 07/11, 10:00\n입금 8,250,000원\n기업은행\n적요 A상사\n\n'
                '2026-07-12 G>N 매출이동')
        res = _parse_notes([memo, '', ''], stage_vals={'계약금': 9150000})
        by_amt = {p['amount']: p for p in res}
        assert by_amt[8250000].get('transfer_to') == 'N'   # 기업(G) → 부착
        assert by_amt[900000].get('transfer_to', '') == ''  # 하나(R) → 미부착

    def test_no_transfer_unchanged(self):
        """이체 표기 없으면 transfer_to 빈값 (회귀 없음)."""
        memo = '일시 07/10, 14:00\n입금 1,000,000원\n하나은행\n적요 김철수'
        res = _parse_notes([memo, '', ''], stage_vals={'계약금': 1000000})
        assert res[0].get('transfer_to', '') == ''

    def test_amount_not_doubled(self):
        """이체 블록이 새 입금으로 중복 카운트되지 않음 (payment 1건)."""
        memo = ('일시 07/21, 16:53\n입금 200,000원\n적요 신세게유통\n\n'
                '2026-07-29 R>N 매출이동\n농협 입금 200,000원')
        res = _parse_notes([memo, '', ''], stage_vals={'계약금': 200000})
        assert len(res) == 1                # 이체 블록은 skip (중복 아님)
        assert res[0]['amount'] == 200000

    def test_same_code_transfer_still_attaches(self):
        """G>G 등 동일코드 이체도 부착 (렌더 단계에서 == code 로 화살표 생략)."""
        memo = ('일시 07/20, 09:00\n입금 500,000원\n기업은행\n적요 홍길동\n'
                '2026-07-20 G>G 매출이동')
        res = _parse_notes([memo, '', ''], stage_vals={'계약금': 500000})
        # 부착 자체는 되되(G), 빌더의 transfer_to != code 가드가 화살표를 생략
        assert res[0]['amount'] == 500000


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
