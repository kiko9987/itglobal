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


class TestSubstitutePayment:
    """대체 수금(리베이트 상계 등 비현금 대체입금) 인식 — 2026-09-29 G3984-YG 제보.

    표준 입금 양식이 아니라 블록 파서로는 payment 0개였다. 좁은 트리거('대체 수금')로만
    한 건의 payment(is_substitute) 로 인식한다. '상계/리베이트' 단독 단어는 오탐·flood
    위험으로 트리거에서 제외.
    """

    def test_g3984_substitute_parsed(self):
        # G3984-YG 실제 중도금 메모 (리베이트 상계 → 대체 수금)
        memo = (
            '프로젝트 / 리베이트 3% / 타공\n'
            'G1798-YG/915,000/390,000\n'
            'G2014-YG/585,000/300,000\n'
            '총 5,280,000원 대체 수금'
        )
        # baseline 경로(stage_vals 없이) 에서도 금액 추출돼야 폴러가 감지 가능
        res = _parse_notes(['', memo, ''])
        mid = [p for p in res if p.get('stage') == '중도금']
        assert len(mid) == 1, f'중도금 대체수금 1건 기대, 실제 {res}'
        assert mid[0]['is_substitute'] is True
        assert mid[0]['amount'] == 5280000
        assert mid[0]['partner'] == '대체수금'

    def test_amount_before_keyword(self):
        # '총 X원' 없이 'X원 대체 수금' 만 있어도 추출
        res = _parse_notes(['', '3,300,000원 대체 수금', ''])
        mid = [p for p in res if p.get('stage') == '중도금']
        assert mid and mid[0]['is_substitute'] and mid[0]['amount'] == 3300000

    def test_narrow_trigger_no_false_positive(self):
        # '상계'/'리베이트' 단독은 대체수금으로 오인하면 안 됨 (상계아산내과 등 상호 오탐)
        for memo in ('입금일: 2024-06-07\n입금자: 최지훈(상계아산내과)',
                     '2025년 12월 삼한 매입금에서 상계처리',
                     '리베이트 정산 예정'):
            res = _parse_notes(['', memo, ''], stage_vals={'중도금': 165000})
            assert not any(p.get('is_substitute') for p in res), f'오탐: {memo!r} → {res}'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
