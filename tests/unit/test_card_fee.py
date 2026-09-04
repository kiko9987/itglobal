# -*- coding: utf-8 -*-
"""카드 결제 수수료 자동 계산 회귀 테스트 (2026-09-01 G3954-TH 계기).

두 축:
  1. _card_fee_line — 표시. 3% 얹어 청구한 설치 건은 '실결제/3%/수수료',
     ITG가 수수료 흡수한 기타 건은 '실결제/수수료'(3% 생략).
  2. _card_settlement_target — 기록. 카드 순입금이 계약잔액보다 ≤4% 작으면
     실결제(계약잔액)로 기록해 미수금 0. 4% 초과·비카드·초과입금은 건드리지 않음.

pure/mocked — 시트 IO 는 FakeMgr 로 대체.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.payment_sync import (
    _card_fee_line, _is_itg_card_deposit, _is_card_payment, _parse_notes,
    _fmt_payment_date, _parse_memo_block, _collection_gross_fee, _collection_fee_line,
)
from dashboard.blueprints.slack_bot import _card_settlement_target, _collection_settlement_target

# G3954-TH 실제 카드 정산 메모 (통장 순입금)
CARD_MEMO = '2026/09/01 07:21\n입금 443,003원\nSHC0117935\n452***38801011\n기업'
BANK_MEMO = '2026/09/01 07:21\n입금 443,003원\n프레임플러스\n452***38801011\n기업'


class TestCardFeeLine:
    def test_absorbed_fee_no_3pct(self):
        # 기타 건: 실결제=계약잔액(453,200), 순입금 443,003 → 수수료만, 3% 없음
        out = _card_fee_line('잔금', 453200, 443003, 453200,
                             {'계약금': 0, '중도금': 0, '잔금': 453200})
        assert out == '(실결제 453,200원 / 카드 수수료 10,197원)'
        assert '3%' not in out

    def test_surcharge_shows_3pct(self):
        # 설치 건: 실결제(1,030,000)가 계약잔액(1,000,000)보다 큼 → 3% 표기 유지
        out = _card_fee_line('잔금', 1030000, 1009000, 1000000,
                             {'계약금': 0, '중도금': 0, '잔금': 1030000})
        assert out == '(실결제 1,030,000원 / 3% 30,000원 / 카드 수수료 21,000원)'

    def test_no_fee_when_equal(self):
        assert _card_fee_line('잔금', 443003, 443003, 453200, {'잔금': 443003}) == ''

    def test_no_fee_when_deposit_higher(self):
        assert _card_fee_line('잔금', 400000, 410000, 453200, {'잔금': 400000}) == ''

    def test_other_stages_reduce_outstanding(self):
        # 계약금 200,000 이미 납 → 잔금이 커버할 계약잔액 = 453,200-200,000=253,200.
        # 실결제 253,200 = 계약잔액 → 흡수(3% 없음)
        out = _card_fee_line('잔금', 253200, 247000, 453200,
                             {'계약금': 200000, '중도금': 0, '잔금': 253200})
        assert out == '(실결제 253,200원 / 카드 수수료 6,200원)'


class FakeMgr:
    """get_cell_value 만 흉내내는 최소 시트 매니저."""
    def __init__(self, cells):
        self.cells = cells

    def get_cell_value(self, sid, sn, cell):
        return self.cells.get(cell)


SID, SN, ROW = 'sheet', '공사현황', 3955


def _target(cells, col, old_num, deposit, memo):
    return _card_settlement_target(FakeMgr(cells), SID, SN, ROW, col, old_num, deposit, memo)


class TestCardSettlementTarget:
    BASE = {'T3955': 453200, 'U3955': 0, 'V3955': 0, 'W3955': 0}

    def test_card_within_4pct_returns_outstanding(self):
        # G3954: 순입금 443,003, 계약잔액 453,200, 차이 10,197 (2.25%) ≤ 4% → 실결제 453,200
        assert _target(self.BASE, 'W', 0, 443003, CARD_MEMO) == 453200

    def test_non_card_partner_ignored(self):
        # 일반 입금자(프레임플러스) → 카드 아님 → None (일반 합산)
        assert _target(self.BASE, 'W', 0, 443003, BANK_MEMO) is None

    def test_gap_over_4pct_ignored(self):
        # 순입금이 계약잔액보다 10% 작음 = 부분결제로 보고 보정 안 함
        assert _target(self.BASE, 'W', 0, 407880, CARD_MEMO) is None  # gap 45,320 > 18,128

    def test_exact_payment_not_touched(self):
        # 순입금 == 계약잔액 (gap 0) → 보정 안 함 (수수료 없는 카드 or 정확 입금)
        assert _target(self.BASE, 'W', 0, 453200, CARD_MEMO) is None

    def test_overpaid_not_touched(self):
        # 순입금 > 계약잔액 → gap 음수 → None
        assert _target(self.BASE, 'W', 0, 460000, CARD_MEMO) is None

    def test_no_contract_total_ignored(self):
        cells = {'T3955': 0, 'U3955': 0, 'V3955': 0, 'W3955': 0}
        assert _target(cells, 'W', 0, 443003, CARD_MEMO) is None

    def test_other_stage_reduces_outstanding(self):
        # 계약금 200,000 납 → 잔금 계약잔액 253,200. 순입금 247,000 (차이 6,200 ≤ 4%) → 253,200
        cells = {'T3955': 453200, 'U3955': 200000, 'V3955': 0, 'W3955': 0}
        assert _target(cells, 'W', 0, 247000, CARD_MEMO) == 253200


class TestItgCardMerchant:
    """통장 카드매출 정산 적요면 카드로 확정 (Y열 불문). 실제 적요 16종(2026-09-01 제공)."""

    @pytest.mark.parametrize('partner', [
        # 글로벌 (G)
        '720364972BC', 'BC-720364972', 'KB97390776', 'NH15415440', 'SHC0117935',
        '롯데59366153', '삼성카드962', '하나90242344', '하나90718642', '현108017094',
        # 그룹 (R)
        '현703011838', '삼성204108778', 'KB12194399', '745389850B',
        '롯데26648732', '신한14322878',
    ])
    def test_real_settlement_descriptors_detected(self, partner):
        assert _is_itg_card_deposit(partner)

    @pytest.mark.parametrize('partner', [
        '프레임플러스', 'SK텔레콤', '㈜시프트업', '', '기업', '홍길동',
        '452388010',   # 사업자 계좌 앞자리 — 가맹점 아님
        '주식회사제우스',
    ])
    def test_non_merchants_ignored(self, partner):
        assert not _is_itg_card_deposit(partner)

    def test_is_card_payment_without_y(self):
        # Y열이 비어도(미수정) 카드매출 적요면 카드로 인식
        assert _is_card_payment('', 'SHC0117935')
        assert _is_card_payment('미발행', '하나90242344')
        assert _is_card_payment('', '롯데59366153')   # 앞자리 절삭형
        assert _is_card_payment('', '삼성카드962')      # 브랜드+카드
        # 일반 입금자는 Y 없으면 카드 아님
        assert not _is_card_payment('', '프레임플러스')


class TestNoteAnnotationParsesSafely:
    """자동보정 노트 주석('카드 실결제/수수료' 2줄)이 금액 파서를 흔들면 안 됨 —
    파서는 순입금(입금 X원)만 읽어야 수수료(실결제W − 순입금) 계산이 맞다."""

    NOTE = ('2026/09/01 07:21\n입금 443,003원\nSHC0117935\n452***38801011\n기업\n'
            '카드 실결제 453,200원\n수수료 10,197원')

    def test_amount_stays_deposit(self):
        p = _parse_notes([self.NOTE], stage_vals={'잔금': 453200})
        assert len(p) == 1                     # 유령 항목 없음
        assert p[0]['amount'] == 443003        # 순입금(실결제·수수료 줄 무시)
        assert p[0]['partner'] == 'SHC0117935'
        assert not p[0].get('is_refund')


class TestPaymentDateYear:
    """누적 이력·헤드라인 날짜에 연도 표기 (분납 다년 구분, G1897-MW 계기)."""

    def test_fmt_with_year(self):
        assert _fmt_payment_date({'date_md': '09/03', 'date_year': '2025'}) == '25/09/03'
        assert _fmt_payment_date({'date_md': '01/10', 'date_year': '2026'}) == '26/01/10'

    def test_fmt_without_year(self):
        assert _fmt_payment_date({'date_md': '09/03'}) == '09/03'
        assert _fmt_payment_date({'date_md': '09/03', 'date_year': ''}) == '09/03'
        assert _fmt_payment_date({'date_md': '-'}) == '-'

    def test_parse_captures_year_slash(self):
        p = _parse_memo_block('2025/02/18 20:48\n입금 4,000,000원\n디자인TOV\n452***38801011\n기업')
        assert p['date_md'] == '02/18' and p['date_year'] == '2025'

    def test_parse_captures_year_label(self):
        p = _parse_memo_block('입금일: 2025-01-10\n입금 1,330,000원\n디자인TOV\n기업')
        assert p['date_md'] == '01/10' and p['date_year'] == '2025'

    def test_korean_date_no_year(self):
        p = _parse_memo_block('6월15일\n입금 100,000원\n홍길동\n기업')
        assert p['date_md'] == '06/15' and p['date_year'] == ''

    def test_mmdd_only_no_year(self):
        p = _parse_memo_block('09/03\n입금 76,900원\n고려/디자인TOV\n기업')
        assert p['date_md'] == '09/03' and p['date_year'] == ''


class TestCollectionFee:
    """추심(고려신용정보) 수수료 — 순입금 → 실추심(gross)·수수료 (21%+VAT=23.1%)."""

    def test_gross_fee(self):
        assert _collection_gross_fee('고려/디자인TOV', 153800) == (200000, 46200)
        assert _collection_gross_fee('고려/디자인TOV', 76900) == (100000, 23100)
        assert _collection_gross_fee('고려신용정보', 76900) == (100000, 23100)

    def test_not_collection(self):
        assert _collection_gross_fee('디자인TOV', 153800) is None
        assert _collection_gross_fee('프레임플러스', 100000) is None
        assert _collection_gross_fee('고려/디자인TOV', 0) is None

    def test_fee_line(self):
        assert _collection_fee_line('고려/디자인TOV', 153800) == '(실추심 200,000원 / 수수료 46,200원)'
        assert _collection_fee_line('디자인TOV', 153800) == ''

    def test_settlement_target_records_gross(self):
        # 추심 순입금 → 단계값 = old + 실추심(gross)
        memo = '2026/09/03 12:47\n입금 76,900원\n고려/디자인TOV\n452***38801011\n기업'
        assert _collection_settlement_target(memo, 76900, 0) == 100000
        assert _collection_settlement_target(memo, 76900, 400000) == 500000
        # 비추심 → None
        bank = '2026/09/03 12:47\n입금 76,900원\n디자인TOV\n452***38801011\n기업'
        assert _collection_settlement_target(bank, 76900, 0) is None

    def test_note_annotation_parses_to_net(self):
        # '실추심/수수료' 주석이 붙어도 파서는 순입금만 읽어야 (누적 합·수수료 계산 정확)
        note = ('2026/09/03 12:47\n입금 76,900원\n고려/디자인TOV\n452***38801011\n기업\n'
                '실추심 100,000원\n수수료 23,100원')
        p = _parse_notes([note], stage_vals={'잔금': 100000})
        assert len(p) == 1 and p[0]['amount'] == 76900


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
