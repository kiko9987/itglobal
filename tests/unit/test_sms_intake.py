# -*- coding: utf-8 -*-
"""은행 입금 SMS 인입 순수 로직 회귀 테스트.

핵심 보장:
  1. 잔액 라인 제거 — 통장 잔고가 슬랙·시트로 절대 넘어가지 않음 (프라이버시)
     · 기업(공백형)  '잔액 144,153,377원'  — 입금액 바로 아래
     · 하나(무공백형) '잔액239,612,486원'   — 맨 끝줄
  2. '[Web발신]' 머리말 제거하되 거래처·금액·은행은 보존
  3. 입금 문자 판별 (광고·인증 문자 배제)
  4. 중복 판별 해시 안정성

pure 함수만 — Flask/Slack/Redis 미접촉.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.sms_intake import (
    strip_balance, looks_like_payment, dedup_hash, has_business_account,
    parse_cash_amount, looks_like_cash, normalize_cash_layout,
)

# 실제 원본 구조 (금액은 샘플). 사용자 제공 원문 기준.
IBK_SMS = (  # 기업은행 — 잔액 공백형, 입금액 바로 아래
    "[Web발신]\n"
    "2026/08/10 14:49\n"
    "입금 407,000원\n"
    "잔액 144,153,377원\n"
    "㈜시프트업\n"
    "452***38801011\n"
    "기업"
)
HANA_SMS = (  # 하나은행 — 잔액 무공백형, 맨 끝줄
    "[Web발신]\n"
    "하나,08/07, 15:33\n"
    "255******31304\n"
    "입금5,115,000원\n"
    "(주)클리어윈코\n"
    "잔액239,612,486원"
)


class TestStripBalance:
    def test_ibk_balance_removed(self):
        out = strip_balance(IBK_SMS)
        assert '잔액' not in out
        assert '144,153,377' not in out          # 잔고 노출 0
        assert '입금 407,000원' in out            # 입금액 보존
        assert '㈜시프트업' in out                # 거래처 보존
        assert '기업' in out                      # 은행 보존
        assert '[Web발신]' in out                 # 머리말 보존 (2026-08-13 요청)

    def test_hana_balance_removed(self):
        out = strip_balance(HANA_SMS)
        assert '잔액' not in out
        assert '239,612,486' not in out
        assert '입금5,115,000원' in out
        assert '(주)클리어윈코' in out
        assert '하나' in out

    def test_no_trailing_blank_lines(self):
        # 하나: 잔액이 마지막 줄 → 제거 후 끝에 빈 줄 남지 않아야
        out = strip_balance(HANA_SMS)
        assert out == out.strip()
        assert not out.endswith('\n')

    @pytest.mark.parametrize('bal', [
        '잔액 144,153,377원',
        '잔액239,612,486원',
        '잔고 1,000,000원',
        '현재잔액 500,000원',
        '출금가능액 12,345원',
        '이체후잔액 99,999원',
    ])
    def test_balance_variants(self, bal):
        out = strip_balance(f"입금 100,000원\n{bal}\n홍길동")
        assert bal.split()[0][:2] not in out or '잔' not in out  # 잔액/잔고 라인 사라짐
        assert '홍길동' in out
        assert '입금 100,000원' in out

    def test_partner_not_falsely_stripped(self):
        # 거래처명이 '잔' 으로 시작해도 금액이 안 붙으면 보존
        out = strip_balance("입금 100,000원\n잔다르크상사\n기업")
        assert '잔다르크상사' in out

    def test_empty(self):
        assert strip_balance('') == ''
        assert strip_balance(None) == ''


class TestLooksLikePayment:
    @pytest.mark.parametrize('text', [IBK_SMS, HANA_SMS, '입금 100,000원 홍길동'])
    def test_payment(self, text):
        assert looks_like_payment(text)

    @pytest.mark.parametrize('text', [
        '[Web발신] 인증번호 123456 입력하세요',
        '광고) 오늘의 특가 세일! 클릭',
        '',
        '택배가 도착했습니다',
    ])
    def test_not_payment(self, text):
        assert not looks_like_payment(text)


class TestBusinessAccount:
    """ITG 사업자 통장(452/255/352)만 통과 — 개인 입금 배제."""

    NH_SMS = '농협 입금3,825,000원\n08/11 14:48 352-****-1682-33 장동선'

    @pytest.mark.parametrize('text', [IBK_SMS, HANA_SMS, NH_SMS])
    def test_business_passes(self, text):
        assert has_business_account(text)

    @pytest.mark.parametrize('text', [
        '입금 100,000원\n홍길동\n452***99999999\n기업',   # 같은 은행 다른(개인) 계좌
        '입금 50,000원\n개인이체\n110-1234-5678',           # 타행 개인
        '입금 30,000원 광고성',
        '',
    ])
    def test_personal_or_unknown_blocked(self, text):
        assert not has_business_account(text)


class TestDedupHash:
    def test_stable(self):
        assert dedup_hash('1544', IBK_SMS) == dedup_hash('1544', IBK_SMS)

    def test_differs_by_text(self):
        # 서로 다른 입금 = 다른 키
        assert dedup_hash('1544', IBK_SMS) != dedup_hash('1544', HANA_SMS)

    def test_same_across_phones_ignores_sender(self):
        # 같은 문자를 3명이 받음 → 발신번호 표기 달라도 같은 키 (3중 수신 dedup 핵심)
        assert dedup_hash('1661', IBK_SMS) == dedup_hash('', IBK_SMS)
        assert dedup_hash('15881111', IBK_SMS) == dedup_hash('1661', IBK_SMS)

    def test_robust_to_whitespace_and_header(self):
        # 안드/아이폰 앱이 공백·'[Web발신]' 머리말을 다르게 보내도 같은 키
        a = "[Web발신]\n입금 100,000원\n홍길동"
        b = "입금 100,000원   홍길동"           # 머리말 제거 + 공백 변형
        assert dedup_hash('x', a) == dedup_hash('y', b)

    def test_len_16(self):
        assert len(dedup_hash('x', 'y')) == 16


class TestCashAmount:
    """현금 금액 파싱 — 콤마·만·억·천만·백만·복합(2천5백만) 다양한 표기."""

    @pytest.mark.parametrize('text,expected', [
        ('8월25일 권태훈 매니저 현금 240만원 YG 수령 완료', 2_400_000),
        ('현금 2,400,000원 안기성 YG수령', 2_400_000),
        ('08/25 현금 2,400,000원 안기성', 2_400_000),
        ('현금 240만원', 2_400_000),
        ('현금 2400만원 수령', 24_000_000),
        ('현금 1,500,000원', 1_500_000),
        ('현금 1억2천만원', 120_000_000),
        ('현금 1억', 100_000_000),
        ('현금 1억5000만원', 150_000_000),
        ('현금 500만원 YG', 5_000_000),
        ('현금 5백만원', 5_000_000),
        ('현금 2천만원', 20_000_000),
        ('현금 1000만원 수령완료', 10_000_000),
        ('현금 2천5백만원', 25_000_000),
        ('현금 3천5백만원 YG수령', 35_000_000),
        ('현금 240 만원', 2_400_000),
        ('현금 2,400,000 YG', 2_400_000),
        ('현금 3억', 300_000_000),
        ('현금 1천만원', 10_000_000),
    ])
    def test_amounts(self, text, expected):
        assert parse_cash_amount(text) == expected

    def test_no_amount(self):
        assert parse_cash_amount('현금 받았어요') == 0
        assert parse_cash_amount('') == 0


class TestCashDetectAndNormalize:
    def test_looks_like_cash(self):
        assert looks_like_cash('8월25일 현금 240만원 YG 수령')
        assert looks_like_cash('현금 2,400,000원')
        assert not looks_like_cash('현금 받았어요')       # 금액 없음
        assert not looks_like_cash('입금 500,000원 홍길동')  # '현금' 없음(은행)

    def test_normalize_layout(self):
        # 자유문장 → '입금 X원 / 현금 수령' 표준 메모 (파서 호환)
        out = normalize_cash_layout('8월25일 권태훈 매니저 현금 240만원 YG 수령 완료')
        assert out == '08/25 입금 2,400,000원\n현금 수령'
        # 날짜 없으면 날짜 생략
        out2 = normalize_cash_layout('현금 2,400,000원 안기성')
        assert out2 == '입금 2,400,000원\n현금 수령'
        # 인식 실패 시 원문 유지
        assert normalize_cash_layout('현금 받았어요') == '현금 받았어요'

    def test_bank_not_treated_as_cash(self):
        # 사업자계좌 있는 은행 문자는 현금으로 오인 안 함 (라우팅은 계좌 유무로 분리)
        bank = '[Web발신]\n기업 입금 500,000원\n홍길동\n452***38801011'
        assert has_business_account(bank)
        # '현금' 단어가 없으면 looks_like_cash 자체가 False
        assert not looks_like_cash(bank)


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
