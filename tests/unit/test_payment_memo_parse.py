# -*- coding: utf-8 -*-
"""입금 메모 파싱 — 거래처(partner) 추출 회귀 테스트.

배경: 매니저 수기 메모 '농협 입금 800,000원 \\n 2026-07-29 한미침례교회' 에서
거래처가 날짜와 같은 줄에 있어 skip 되어 partner='' → 미완성 메모 가드가 매 폴링
skip → 수금완료 미발송(G3814-MS). 파서가 'yyyy-mm-dd 거래처' 한 줄 형식도 잡도록
수정한 것의 회귀 보호. (2026-07-29)

pure 함수(_parse_notes) 만 — Redis/Slack 미접촉.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.payment_sync import _parse_notes


def _partner(memo: str, val: int = 800000) -> str:
    res = _parse_notes(['', '', memo], stage_vals={'잔금': val})
    return (res[0].get('partner', '') if res else '') or ''


class TestPartnerExtraction:
    """거래처가 잡혀야 하는 형식 (발송 가능)."""

    @pytest.mark.parametrize('memo, expected', [
        # 이번 수정 대상 — 수기 'yyyy-mm-dd 거래처' 한 줄
        ('농협 입금 800,000원 \n2026-07-29 한미침례교회\n', '한미침례교회'),
        # 은행알림 MM/DD + 계좌 + 거래처 (G3808-YG)
        ('농협 입금4,600,000원\n07/04 11:43 352-****-1682-33 푸드스케치', '푸드스케치'),
        # 표준 은행알림 5줄 (yyyy/mm/dd HH:MM / 입금 / 거래처 / 계좌 / 은행)
        ('2026/07/29 10:47\n입금 800,000원\n한미침례교회\n452***38801011\n기업', '한미침례교회'),
        # 거래처 별도 줄
        ('농협 입금 800,000원\n한미침례교회\n07/29', '한미침례교회'),
        # KATOK 표준 (MM/DD N 금액원 거래처)
        ('07/29 N 800,000원 한미침례교회', '한미침례교회'),
    ])
    def test_partner_found(self, memo, expected):
        assert _partner(memo) == expected

    def test_manager_summary_format(self):
        # 매니저 요약형: 'yyyy-mm-dd 금액원 거래처 은행'
        res = _parse_notes(['', '', '2025-12-05 300,000원 김연종 농협'],
                           stage_vals={'잔금': 300000})
        assert res and res[0].get('partner') == '김연종'


class TestNoFalsePartner:
    """거래처가 없어야 하는 형식 (partner 빈값 — 미완성 가드가 skip)."""

    @pytest.mark.parametrize('memo', [
        '2026/07/29 10:47',                       # 순수 날짜+시간
        '2026-07-29',                             # 순수 날짜
        '2026/07/29 10:47 452-****-38801011',     # 날짜+시간+계좌만
        '입금 800,000원\n2026/07/29 입금 800,000원',  # 날짜+입금+금액 (거래처 아님)
    ])
    def test_no_partner(self, memo):
        p = _partner(memo)
        assert not p or p in ('', '-'), f'예상치 못한 partner: {p!r}'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
