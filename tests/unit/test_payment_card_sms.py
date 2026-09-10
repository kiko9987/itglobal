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
    _parse_notes, _is_card_payment, _resolve_payment_code, _ensure_note_year)


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
        """Y가 카드결제/혼합 아니고 ITG 승인번호도 아니면 카드 아님 (게이트 유지)."""
        assert _is_card_payment('잔금', '홍길동') is False
        # ITG 가맹점 승인번호(적요)면 Y 불문 카드 확정 (2026-09 _is_itg_card_deposit 도입)
        assert _is_card_payment('잔금', '삼성 204108778') is True


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


class TestYStatusNewFormat:
    """Y열 '{단계} - {상태}' 신형식 상태 추출 회귀 (2026-09-09).

    배경: 계산서 다단계 도입으로 Y('계산서')가 'N입금'/'카드결제'(구) → '잔금 - N입금'/
    '잔금 - 카드결제'(신)로 바뀌었는데, _resolve_payment_code 는 'N입금' 정확일치,
    _is_card_payment 는 ('카드결제','혼합') 정확일치로 판정 → 신형식 매칭 실패.
    결과: 은행 없는 N입금(현금)이 기본값 G 로 오표기(G4026-MS·G4059-MS 'N 분할지정→G' 제보),
    카드결제도 미인식. _y_status 로 상태 토큰 추출해 해소.
    """

    def test_n_deposit_new_format_resolves_N(self):
        # 은행 표식 없으면 Y 상태('N입금')로 N. (신형식 '잔금 - N입금')
        assert _resolve_payment_code('잔금 - N입금', '', '라은정(색담)') == 'N'

    def test_bank_takes_precedence_over_y(self):
        # 은행이 있으면 각 입금 은행 코드 우선 (하나→R, 기업→G) — Y fallback 아님.
        assert _resolve_payment_code('잔금 - N입금', '하나', '김철수') == 'R'
        assert _resolve_payment_code('잔금 - N입금', '기업', '김철수') == 'G'

    def test_card_new_format_detected(self):
        # 신형식 '잔금 - 카드결제' + 카드사명 → 카드 인식.
        assert _is_card_payment('잔금 - 카드결제', '비씨카드') is True

    def test_legacy_bare_format_unchanged(self):
        # 구형식(바로 상태)도 그대로 동작 (하위호환).
        assert _resolve_payment_code('N입금', '', '홍길동') == 'N'
        assert _is_card_payment('카드결제', '삼성 204108778') is True
        # 비-N입금 신형식은 은행 없으면 기본 G 유지.
        assert _resolve_payment_code('잔금 - 미발행', '', '홍길동') == 'G'


class TestEnsureNoteYear:
    """저장 시점에 노트 날짜 연도 삽입 (근본책, 2026-09-10 사용자 제안).

    입금 문자→시트 기록 시 연도를 박아두면 표시 때 추측이 불필요 — 다년 분납도 각 건이
    기록 당시 연도로 확정. 기록 시점 연도(=지금)는 확실하므로 오표기 없음.
    """

    def test_prepends_current_year_when_missing(self):
        from datetime import datetime
        out = _ensure_note_year('입금 800,000원\n09/08 하나\n라은정')
        assert out.split('\n')[0] == f'{datetime.now().year}/09/08'

    def test_keeps_existing_year(self):
        # 은행 SMS 풀날짜(연도 포함)면 그대로 (중복 삽입 안 함).
        memo = '2026-09-08\n입금 800,000원\n라은정'
        assert _ensure_note_year(memo) == memo

    def test_no_date_unchanged(self):
        # 날짜 토큰 자체가 없으면 손대지 않음.
        memo = '입금 500,000원'
        assert _ensure_note_year(memo) == memo


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
